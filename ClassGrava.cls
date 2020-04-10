VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function CTBConfig_Grava_Trans(ByVal objCTBConfig As ClassCTBConfig, alComando() As Long) As Long
'grava registro em CTBConfig. Apenas altera o conteúdo, se não existir dá erro.
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, tCTBConfig As typeCTBConfig

On Error GoTo Erro_CTBConfig_Grava_Trans

    'Inicializar buffer
    tCTBConfig.sConteudo = String(STRING_CTBCONFIG_CONTEUDO, 0)

    'Lê o registo
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Conteudo FROM CTBConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, tCTBConfig.sConteudo, objCTBConfig.sCodigo, objCTBConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 44663

    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44664

    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then

        lErro = Comando_Executar(alComando(2), "INSERT INTO CTBConfig (Codigo, FilialEmpresa, Descricao, Tipo, Conteudo) VALUES (?,?,?,?,?)", objCTBConfig.sCodigo, objCTBConfig.iFilialEmpresa, objCTBConfig.sDescricao, objCTBConfig.iTipo, objCTBConfig.sConteudo)
        If lErro <> AD_SQL_SUCESSO Then Error 44665

    Else

        'Se o conteúdo foi modificado
        If objCTBConfig.sConteudo <> tCTBConfig.sConteudo Then
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE CTBConfig SET Conteudo = ?", alComando(1), objCTBConfig.sConteudo)
            If lErro <> AD_SQL_SUCESSO Then Error 44666
        End If

    End If
    
    CTBConfig_Grava_Trans = SUCESSO

    Exit Function

Erro_CTBConfig_Grava_Trans:

    CTBConfig_Grava_Trans = Err

    Select Case Err

        Case 44663, 44664
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTBCONFIG2", Err, objCTBConfig.sCodigo, objCTBConfig.iFilialEmpresa)

        Case 44665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CTBCONFIG", Err, objCTBConfig.sCodigo, objCTBConfig.iFilialEmpresa)

        Case 44666
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CTBCONFIG", Err, objCTBConfig.sCodigo, objCTBConfig.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150557)

    End Select

    Exit Function

End Function

Function CTBConfig_Grava_Configs(ByVal colConfigs As Collection) As Long
'Grava na tabela CRConfig uma coleção de registros de configuração

Dim lErro As Long, iIndice As Integer
Dim lTransacao As Long

On Error GoTo Erro_CTBConfig_Grava_Configs

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 44660

    lErro = CF("CTBConfig_Grava_Configs_Trans", colConfigs)
    If lErro <> SUCESSO Then Error 44670

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 44662

    CTBConfig_Grava_Configs = SUCESSO

    Exit Function

Erro_CTBConfig_Grava_Configs:

    CTBConfig_Grava_Configs = Err

    Select Case Err

        Case 44660
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 44662
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 44670

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150558)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function CTBConfig_Grava_Configs_Trans(ByVal colConfigs As Collection) As Long
'Grava na tabela CRConfig uma coleção de registros de configuração
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 3) As Long
Dim objCTBConfig As ClassCTBConfig

On Error GoTo Erro_CTBConfig_Grava_Configs_Trans

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 44659
    Next

   'Para cada elemento da coleção
    For Each objCTBConfig In colConfigs

        lErro = CTBConfig_Grava_Trans(objCTBConfig, alComando())
        If lErro <> SUCESSO Then Error 44661

    Next

   'Fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CTBConfig_Grava_Configs_Trans = SUCESSO

    Exit Function

Erro_CTBConfig_Grava_Configs_Trans:

    CTBConfig_Grava_Configs_Trans = Err

    Select Case Err

        Case 44659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44661

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150559)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Function Lancamento_Exclui(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
    
Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iExercicio As Integer
Dim iPeriodo As Integer
Dim iStatus As Integer
Dim iIndice As Integer
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Lancamento_Exclui

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 207107

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 207108
        
    Next
    
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
    
    Do While objLancamento_Cabecalho.iFilialEmpresa > 0 And objLancamento_Cabecalho.iFilialEmpresa < 100
    
        'verifica se o exercicio não está fechado
        lErro = Comando_ExecutarLockado(alComando(1), "SELECT Exercicio, Status FROM Exercicios WHERE DataInicio <= ? AND DataFim >= ?", iExercicio, iStatus, objLancamento_Cabecalho.dtData, objLancamento_Cabecalho.dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 207109
        
        'le o exercicio, se estiver no BD
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 207110
        
        objLancamento_Cabecalho.iExercicio = iExercicio
        
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 207111
        
        'se o exercicio estiver fechado, não pode editar os lançamentos
        If giDesconsideraFechamentoPeriodo <> MARCADO Then
            If iStatus = EXERCICIO_FECHADO Then gError 207112
        End If
        
        'verifica qual o periodo em questão
        lErro = Comando_ExecutarLockado(alComando(2), "SELECT Periodo FROM Periodo WHERE Exercicio=? AND DataInicio <= ? AND DataFim >= ?", iPeriodo, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.dtData, objLancamento_Cabecalho.dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 207113
        
        'le o periodo, se estiver no BD
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 207114
        
        objLancamento_Cabecalho.iPeriodoLote = iPeriodo
        objLancamento_Cabecalho.iPeriodoLan = iPeriodo
        
        lErro = Comando_LockShared(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 207115
        
        
        objPeriodosFilial.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
        objPeriodosFilial.iExercicio = objLancamento_Cabecalho.iExercicio
        objPeriodosFilial.iPeriodo = objLancamento_Cabecalho.iPeriodoLan
        objPeriodosFilial.sOrigem = objLancamento_Cabecalho.sOrigem

        lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
        If lErro <> SUCESSO Then gError 207116
        
        'se o periodo estiver fechado, não pode alterar os lançamentos do periodo
        If giDesconsideraFechamentoPeriodo <> MARCADO Then
            If objPeriodosFilial.iFechado = PERIODO_FECHADO Then gError 207117
        End If
        
        lErro = Lancamento_Exclui1(objLancamento_Cabecalho)
        If lErro <> SUCESSO Then gError 207118
  
        If giContabGerencial = 0 Then Exit Do
        
        objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa - giFilialAuxiliar
            
    Loop
                  
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 207119
    
    Lancamento_Exclui = SUCESSO

    Exit Function

Erro_Lancamento_Exclui:
    
    Lancamento_Exclui = gErr
    
    Select Case gErr
    
        Case 207107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 207108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 207109, 207110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOS", gErr)
                        
        Case 207111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", gErr, objLancamento_Cabecalho.iExercicio)
                        
        Case 207112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", gErr, objLancamento_Cabecalho.iExercicio)
           
        Case 207113, 207114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO1", gErr)
                        
        Case 207115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODO", gErr, objLancamento_Cabecalho.iExercicio, iPeriodo)
                        
        Case 207116, 207118
                        
        Case 207117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", gErr, iPeriodo, objLancamento_Cabecalho.iExercicio)
                        
        Case 207119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150560)
            
    End Select
   
    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Function Lancamento_Exclui0(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
'Exclui os lançamentos passados como parametro. Continuação  da rotina Lancamento_Exclui
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
    
    objLancamento_Cabecalho.iFilialEmpresa = giFilialEmpresa
        
    Lancamento_Exclui0 = CF("Lancamento_Exclui2", objLancamento_Cabecalho)

End Function

Function Lancamento_Exclui2(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
'Exclui os lançamentos passados como parametro. Continuação  da rotina Lancamento_Exclui
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
    
Dim alComando(1 To 3) As Long
Dim lErro As Long
Dim iExercicio As Integer
Dim iPeriodo As Integer
Dim iStatus As Integer
Dim iIndice As Integer
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Lancamento_Exclui2

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 5750
        
    Next
    
    'verifica se o exercicio não está fechado
    lErro = Comando_ExecutarLockado(alComando(1), "SELECT Exercicio, Status FROM Exercicios WHERE DataInicio <= ? AND DataFim >= ?", iExercicio, iStatus, objLancamento_Cabecalho.dtData, objLancamento_Cabecalho.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 5751
    
    'le o exercicio, se estiver no BD
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 5752
    
    objLancamento_Cabecalho.iExercicio = iExercicio
    
    lErro = Comando_LockShared(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 5753
    
    'se o exercicio estiver fechado, não pode editar os lançamentos
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If iStatus = EXERCICIO_FECHADO Then Error 5754
    End If
    
    'verifica qual o periodo em questão
    lErro = Comando_ExecutarLockado(alComando(2), "SELECT Periodo FROM Periodo WHERE Exercicio=? AND DataInicio <= ? AND DataFim >= ?", iPeriodo, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.dtData, objLancamento_Cabecalho.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 5781
    
    'le o periodo, se estiver no BD
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 5782
    
    objLancamento_Cabecalho.iPeriodoLote = iPeriodo
    objLancamento_Cabecalho.iPeriodoLan = iPeriodo
    
    'obs.: o lock é exclusivo pois pode ser chamado numa transacao que já utiliza este tipo de lock
    lErro = Comando_LockExclusive(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 5783
    
    
    objPeriodosFilial.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
    objPeriodosFilial.iExercicio = objLancamento_Cabecalho.iExercicio
    objPeriodosFilial.iPeriodo = objLancamento_Cabecalho.iPeriodoLan
    objPeriodosFilial.sOrigem = objLancamento_Cabecalho.sOrigem

    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then gError 20849
    
    'se o periodo estiver fechado, não pode alterar os lançamentos do periodo
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 5784
    End If
  
    lErro = Lancamento_Exclui1(objLancamento_Cabecalho)
    If lErro <> SUCESSO Then Error 10156
  
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Lancamento_Exclui2 = SUCESSO

    Exit Function

Erro_Lancamento_Exclui2:
    
    Lancamento_Exclui2 = Err
    
    Select Case Err
    
        Case 5750
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5751, 5752
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOS", Err)
                        
        Case 5753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, objLancamento_Cabecalho.iExercicio)
                        
        Case 5754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objLancamento_Cabecalho.iExercicio)
           
        Case 5781, 5782
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO1", Err)
                        
        Case 5783
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODO", Err, objLancamento_Cabecalho.iExercicio, iPeriodo)
                        
        Case 5784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, iPeriodo, objLancamento_Cabecalho.iExercicio)
                        
        Case 10151, 10152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iExercicio, iPeriodo)
            
        Case 10153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODOSFILIAL", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iExercicio, iPeriodo)
            
        Case 10156, 20849
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150561)
            
    End Select
   
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Private Function Lancamento_Exclui1(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
'Exclui os lançamentos passados como parametro. Continuação  da rotina Lancamento_Exclui0
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim dCredito As Double
Dim dDebito As Double
Dim iNumLancAtual As Integer
Dim iNumDocAtual As Integer
Dim objLote As New ClassLote
Dim iLoteAtualizado As Integer
Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long

On Error GoTo Erro_Lancamento_Exclui1

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10154

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10155

    objLote.sOrigem = objLancamento_Cabecalho.sOrigem
    objLote.iExercicio = objLancamento_Cabecalho.iExercicio
    objLote.iPeriodo = objLancamento_Cabecalho.iPeriodoLote
    objLote.iLote = objLancamento_Cabecalho.iLote
    objLote.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
    
    'verifica se o lote  está atualizado
    lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
    If lErro <> SUCESSO Then Error 5995
    
    If iLoteAtualizado = LOTE_ATUALIZADO Then Error 5785
  
    'Pesquisa Lote associado aos Lançamentos
    lErro = Comando_ExecutarPos(lComando1, "SELECT TotCre, TotDeb, NumLancAtual, NumDocAtual FROM LotePendente WHERE FilialEmpresa = ? AND Origem=? AND Exercicio=? AND Periodo=? AND Lote=?", 0, dCredito, dDebito, iNumLancAtual, iNumDocAtual, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 6078

    'Lê o Lote
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6079
    
    If lErro = AD_SQL_SEM_DADOS Then Error 6080

    'Lock do Lote dos Lançamentos
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 6081
  
    'exclui os lançamentos deste voucher
    lErro = CF("Lancamento_Exclui_BD", objLancamento_Cabecalho, dDebito, dCredito, iNumLancAtual, iNumDocAtual)
    If lErro Then Error 5800
  
    lErro = Comando_ExecutarPos(lComando2, "UPDATE LotePendente SET TotCre = ?, TotDeb = ?, NumLancAtual = ?, NumDocAtual = ?", lComando1, dCredito, dDebito, iNumLancAtual, iNumDocAtual)
    If lErro <> AD_SQL_SUCESSO Then Error 5802

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Lancamento_Exclui1 = SUCESSO

    Exit Function

Erro_Lancamento_Exclui1:
    
    Lancamento_Exclui1 = Err
    
    Select Case Err
    
        Case 10154, 10155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iLote, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.sOrigem)
            
        Case 5800, 5995
                   
        Case 5802
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LOTEPENDENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
                   
        Case 6078, 6079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
                      
        Case 6080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTEPENDENTE_INEXISTENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
            
        Case 6081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LOTEPENDENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150562)
            
    End Select
   
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function Lancamento_Grava(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection, Optional ByVal iExcluiLctos As Integer = 0) As Long
'Grava os lançamentos associados a um lote passados como parametro

Dim lTransacao As Long
Dim lErro As Long
Dim colLancamento_Detalhe_Aux As Collection
Dim objLancamento_Detalhe As ClassLancamento_Detalhe

On Error GoTo Erro_Lancamento_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 6013
    
    Do While objLancamento_Cabecalho.iFilialEmpresa > 0 And objLancamento_Cabecalho.iFilialEmpresa < 100
       
        Set colLancamento_Detalhe_Aux = New Collection
       
        For Each objLancamento_Detalhe In colLancamento_Detalhe
        
            If giContabGerencial = 1 Then
                If (objLancamento_Detalhe.iGerencial = 0 Or objLancamento_Detalhe.iGerencial = 2) And ((objLancamento_Cabecalho.iFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (objLancamento_Cabecalho.iFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                    colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
                End If
                    
                If (objLancamento_Detalhe.iGerencial = 1 Or objLancamento_Detalhe.iGerencial = 2) And ((objLancamento_Cabecalho.iFilialEmpresa < giFilialAuxiliar And giFilialAuxiliar > 0) Or (objLancamento_Cabecalho.iFilialEmpresa > Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
                End If
                
            Else
                    
                'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
            End If
                
        Next
       
        lErro = CF("Lancamento_Grava0", objLancamento_Cabecalho, colLancamento_Detalhe_Aux, iExcluiLctos)
        If lErro <> SUCESSO Then Error 20474
    
        If giContabGerencial = 0 Then Exit Do

        objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa - giFilialAuxiliar
        
    Loop
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 6015

    Lancamento_Grava = SUCESSO

    Exit Function

Erro_Lancamento_Grava:
    
    Lancamento_Grava = Err
    
    Select Case Err
    
        Case 6013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
                       
        Case 6015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
                      
        Case 20474
                       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150563)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function LanPendente_Altera_BD(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection) As Long
'Altera os dados do lançamento pendente passado como parametro
    
Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim dValor As Double
Dim iSeq As Integer
Dim iTransacao As Integer
Dim lNumIntDoc As Long
Dim objLanc As ClassLancamento_Detalhe
Dim iIndice As Integer
Dim dValorLivroAux As Double

On Error GoTo Erro_LanPendente_Altera_BD

    'Abre comando para pesquisar Lancamentos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 83510
    
    'Abre comando para transação de exclusão
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 83511
    
    'Pesquisa a existencia dos Lancamentos do Documento no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT Seq, ValorLivroAuxiliar FROM LanPendente WHERE FilialEmpresa=? AND Origem=? AND Exercicio=? AND PeriodoLan=? AND Doc=? And Lote = ? ORDER BY Seq", 0, iSeq, dValorLivroAux, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc, objLancamento_Cabecalho.iLote)
    If lErro <> AD_SQL_SUCESSO Then gError 83512
     
    'Lê o primeiro lançamento
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83513

    iIndice = 0

    'Inicia Loop de Alteração dos dados com Lançamentos do GRID
    Do While lErro = AD_SQL_SUCESSO
        
        iIndice = iIndice + 1
        
        Set objLanc = colLancamento_Detalhe.Item(iIndice)
        
        'se o numero de elementos da tela exceder o numero de elementos do bd ==> erro
        If iIndice > colLancamento_Detalhe.Count Then gError 83516
        
        If dValorLivroAux > 0 Then
            dValorLivroAux = Abs(objLanc.dValor)
        ElseIf dValorLivroAux < 0 Then
            dValorLivroAux = -Abs(objLanc.dValor)
        End If
        
        'Alteração Daniel em 30/08/2002
        If giHistoricoObrigatorio = MARCADO Then
            If Len(Trim(objLanc.sHistorico)) = 0 Then gError 32280
        End If
        'Fim da Alteração
        
        'Exclui Lançamento da Tabela de Lançamentos Pendentes
        lErro = Comando_ExecutarPos(lComando1, "UPDATE LanPendente SET Conta =?, Ccl=?, Valor=?, Historico=?, ValorLivroAuxiliar=?", lComando, objLanc.sConta, objLanc.sCcl, objLanc.dValor, objLanc.sHistorico, dValorLivroAux)
        If lErro <> AD_SQL_SUCESSO Then gError 83514
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83515
    
    Loop
    
    'se o numero do bd exceder o numero de elementos da tela  ==> erro
    If iIndice < colLancamento_Detalhe.Count Then gError 83517
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    LanPendente_Altera_BD = SUCESSO
    
    Exit Function
    
Erro_LanPendente_Altera_BD:

    LanPendente_Altera_BD = gErr
    
    Select Case gErr
    
    'Alteracao Daniel em 30/08/2002
        Case 32280
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_SEM_HISTORICO", gErr)
    'Fim da Alteração Daniel em 30/08/02
        
        Case 83510, 83511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 83512, 83513, 83515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS_PENDENTES1", gErr, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc, objLancamento_Cabecalho.iLote)
            
        Case 83514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANPENDENTE", gErr, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc, iSeq)
            
        Case 83516
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_LANC_TELA_MAIOR_BD", gErr)
            
        Case 83517
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_LANC_TELA_MENOR_BD", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150564)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
                   
End Function

Function Lancamento_Grava0(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection, Optional iExcluiLctos As Integer = 0) As Long
'Grava os dados do lançamento. Utilizado por Lancamento_Grava
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
    
Dim alComando(1 To 3) As Long
Dim lErro As Long
Dim iExercicio As Integer
Dim iPeriodo As Integer
Dim iStatus As Integer
Dim iIndice As Integer
Dim iLote As Integer

On Error GoTo Erro_Lancamento_Grava0

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 5755
        
    Next
    
    'faz os testes de duplicidade, documento já contabilizado, periodo fechado, etc.
    lErro = Lancamento_Grava2(alComando, objLancamento_Cabecalho, colLancamento_Detalhe)
    If lErro <> SUCESSO Then Error 20814
    
    lErro = Lancamento_Grava1(objLancamento_Cabecalho, colLancamento_Detalhe, iExcluiLctos)
    If lErro <> SUCESSO Then Error 10148
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Lancamento_Grava0 = SUCESSO

    Exit Function

Erro_Lancamento_Grava0:
    
    Lancamento_Grava0 = Err
    
    Select Case Err
    
        Case 5755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                   
        Case 10148, 20814
                       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150565)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Lancamento_Grava1(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection, iExcluiLctos As Integer) As Long
'Grava os dados do lançamento. Utilizado por Lancamento_Grava0
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lErro As Long
Dim iLoteAtualizado As Integer
Dim objLote As New ClassLote
Dim dDebito As Double
Dim dCredito As Double
Dim iNumLancAtual As Integer
Dim iNumDocAtual As Integer
Dim lComando1 As Long
Dim lComando2 As Long

On Error GoTo Erro_Lancamento_Grava1

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10149
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10150

    objLote.sOrigem = objLancamento_Cabecalho.sOrigem
    objLote.iExercicio = objLancamento_Cabecalho.iExercicio
    objLote.iPeriodo = objLancamento_Cabecalho.iPeriodoLote
    objLote.iLote = objLancamento_Cabecalho.iLote
    objLote.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
    
    'verifica se o lote  está atualizado
    lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
    If lErro <> SUCESSO Then Error 5996
    
    If iLoteAtualizado = LOTE_ATUALIZADO Then Error 5803
    
    'Pesquisa Lote associado aos Lançamentos
    lErro = Comando_ExecutarPos(lComando1, "SELECT TotCre, TotDeb, NumLancAtual, NumDocAtual FROM LotePendente WHERE FilialEmpresa=? AND Origem=? AND Exercicio=? AND Periodo=? AND Lote=?", 0, dCredito, dDebito, iNumLancAtual, iNumDocAtual, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 6009

    'Lê o Lote
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6010
        
    If lErro = AD_SQL_SEM_DADOS Then Error 6070

    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 6011
    
    If iExcluiLctos = EXCLUI_LANCAMENTOS_CONTABEIS Then
    
        'exclui os lançamentos deste voucher
        lErro = CF("Lancamento_Exclui_BD", objLancamento_Cabecalho, dDebito, dCredito, iNumLancAtual, iNumDocAtual)
        If lErro <> SUCESSO Then Error 5786

    End If

    lErro = CF("Lancamento_Grava_BD", objLancamento_Cabecalho, colLancamento_Detalhe, dDebito, dCredito, iNumLancAtual)
    If lErro <> SUCESSO Then Error 5797
    
    lErro = Comando_ExecutarPos(lComando2, "UPDATE LotePendente SET TotCre = ?, TotDeb = ?, NumLancAtual = ?, NumDocAtual = ?", lComando1, dCredito, dDebito, iNumLancAtual, iNumDocAtual + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 5798

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Lancamento_Grava1 = SUCESSO

    Exit Function

Erro_Lancamento_Grava1:
    
    Lancamento_Grava1 = Err
    
    Select Case Err
    
        Case 5786, 5797, 5996
                       
        Case 5798
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LOTEPENDENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
                       
        Case 5803
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iLote, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.sOrigem)
               
        Case 6009, 6010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
            
        Case 6011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LOTEPENDENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
                        
        Case 6070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTEPENDENTE_INEXISTENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
                    
        Case 10149, 10150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150566)
            
    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function LanPendente_Grava(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim colLancamento_Detalhe_Aux As Collection
Dim objLancamento_Detalhe As ClassLancamento_Detalhe

On Error GoTo Erro_LanPendente_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 83521
    
    Do While objLancamento_Cabecalho.iFilialEmpresa > 0 And objLancamento_Cabecalho.iFilialEmpresa < 100
       
        Set colLancamento_Detalhe_Aux = New Collection
       
        For Each objLancamento_Detalhe In colLancamento_Detalhe
        
            If giContabGerencial = 1 Then
                If (objLancamento_Detalhe.iGerencial = 0 Or objLancamento_Detalhe.iGerencial = 2) And ((objLancamento_Cabecalho.iFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (objLancamento_Cabecalho.iFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                    colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
                End If
                    
                If (objLancamento_Detalhe.iGerencial = 1 Or objLancamento_Detalhe.iGerencial = 2) And ((objLancamento_Cabecalho.iFilialEmpresa < giFilialAuxiliar And giFilialAuxiliar > 0) Or (objLancamento_Cabecalho.iFilialEmpresa > Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
                End If
                
            Else
                    
                'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
            End If
                
        Next
       
        lErro = CF("LanPendente_Grava_EmTrans", objLancamento_Cabecalho, colLancamento_Detalhe_Aux)
        If lErro <> SUCESSO Then Error 83519
    
        If giContabGerencial = 0 Then Exit Do

        objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa - giFilialAuxiliar
        
    Loop
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 83522
    
    LanPendente_Grava = SUCESSO

    Exit Function

Erro_LanPendente_Grava:
    
    LanPendente_Grava = gErr
    
    Select Case gErr
    
        Case 83519
                       
        Case 83521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
                       
        Case 83522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
                       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150567)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Function LanPendente_Grava_EmTrans(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection) As Long

Dim alComando(1 To 3) As Long
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_LanPendente_Grava_EmTrans
    
    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 83518
        
    Next
    
    'faz os testes de duplicidade, documento já contabilizado, periodo fechado, etc.
    lErro = Lancamento_Grava2(alComando, objLancamento_Cabecalho, colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError 83519
    
    lErro = LanPendente_Altera_BD(objLancamento_Cabecalho, colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError 83520
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    LanPendente_Grava_EmTrans = SUCESSO

    Exit Function

Erro_LanPendente_Grava_EmTrans:
    
    LanPendente_Grava_EmTrans = gErr
    
    Select Case gErr
    
        Case 83518
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                   
        Case 83519, 83520
                       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150568)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
                      
    Exit Function

End Function


Private Function Lancamento_Grava2(alComando() As Long, ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection) As Long
'faz os testes de duplicidade, documento já contabilizado, periodo fechado, etc.
    
Dim lErro As Long
Dim iExercicio As Integer
Dim iPeriodo As Integer
Dim iStatus As Integer
Dim iIndice As Integer
Dim iLote As Integer
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Lancamento_Grava2

    'verifica se o exercicio não está fechado
    lErro = Comando_ExecutarLockado(alComando(1), "SELECT Exercicio, Status FROM Exercicios WHERE DataInicio <= ? AND DataFim >= ?", iExercicio, iStatus, objLancamento_Cabecalho.dtData, objLancamento_Cabecalho.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 5756
    
    'le o exercicio, se estiver no BD
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 5757
    
    objLancamento_Cabecalho.iExercicio = iExercicio
    
    lErro = Comando_LockShared(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 5758
    
    'se o exercicio estiver fechado, não pode editar os lançamentos
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If iStatus = EXERCICIO_FECHADO Then Error 5759
    End If
    
    'verifica qual o periodo do lançamento
    lErro = Comando_ExecutarLockado(alComando(2), "SELECT Periodo FROM Periodo WHERE Exercicio=? AND DataInicio <= ? AND DataFim >= ?", iPeriodo, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.dtData, objLancamento_Cabecalho.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 5776
    
    'le o periodo, se estiver no BD
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 5777
    
    objLancamento_Cabecalho.iPeriodoLote = iPeriodo
    objLancamento_Cabecalho.iPeriodoLan = iPeriodo
    
    lErro = Comando_LockShared(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 5778
    
    objPeriodosFilial.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
    objPeriodosFilial.iExercicio = objLancamento_Cabecalho.iExercicio
    objPeriodosFilial.iPeriodo = objLancamento_Cabecalho.iPeriodoLan
    objPeriodosFilial.sOrigem = objLancamento_Cabecalho.sOrigem

    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 20850
    
    'se o periodo estiver fechado, não pode alterar os lançamentos do periodo
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 5779
    End If
    
    'pesquisa se o número já foi utilizado em lançamentos contabilizados
    lErro = CF("Lancamento_Le_Doc", objLancamento_Cabecalho)
    If lErro <> SUCESSO And lErro <> 5726 Then Error 9585
    
    'Se este voucher já foi usado ==> erro
    If lErro = SUCESSO Then Error 9586
        
    'guarda o lote que vai ser alterado pela funcao LanPendente_Le_Doc1
    iLote = objLancamento_Cabecalho.iLote
        
    'pesquisa se o número já foi utilizado em lançamentos pendentes em outro periodo
    lErro = CF("LanPendente_Le_Doc1", objLancamento_Cabecalho)
    If lErro <> SUCESSO And lErro <> 9590 Then Error 20776
    
    'testa se o lote foi trocado. Se o lote foi trocado ==> erro. Neste caso o doc tem que ser excluido e incluido em outro lote.
    If iLote <> objLancamento_Cabecalho.iLote Then Error 20777
    
    Lancamento_Grava2 = SUCESSO

    Exit Function

Erro_Lancamento_Grava2:
    
    Lancamento_Grava2 = Err
    
    Select Case Err
    
        Case 5756, 5757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOS", Err)
           
        Case 5758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, objLancamento_Cabecalho.iExercicio)
                       
        Case 5759
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objLancamento_Cabecalho.iExercicio)
                    
        Case 5776, 5777
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO1", Err)
            
        Case 5778
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODO", Err, objLancamento_Cabecalho.iExercicio, iPeriodo)
            
        Case 5779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, iPeriodo, objLancamento_Cabecalho.iExercicio)
                   
        Case 9585, 20776, 20850
                       
        Case 9586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_ATUALIZADO", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
            
        Case 10145, 10146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iExercicio, iPeriodo)
            
        Case 10147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODOSFILIAL", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iExercicio, iPeriodo)
            
        Case 20777
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_PENDENTE_OUTRO_LOTE", Err, objLancamento_Cabecalho.iLote)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150569)
            
    End Select
    
    Exit Function

End Function

Function Lancamento_Grava_BD(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection, dCredito As Double, dDebito As Double, iNumLancAtual As Integer) As Long
'grava na tabela de lançamentos pendentes os lançamentos do voucher em questão
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim alComando(0 To 15) As Long
Dim lErro As Long
Dim iTipoConta As Integer
Dim iStatus As Integer
Dim sCcl As String
Dim sConta As String
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim iPossuiCcl As Integer
Dim objHistPadrao As New ClassHistPadrao
Dim iIndice As Integer
Dim iApropriacao As Integer

On Error GoTo Erro_Lancamento_Grava_BD


    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 6012
    Next
    
    'Inicia Loop de GRAVACAO com Lançamentos do GRID
    For Each objLancamento_Detalhe In colLancamento_Detalhe
            
        'Alteração Daniel em 30/08/2002
        If giHistoricoObrigatorio = MARCADO Then
            If Len(Trim(objLancamento_Detalhe.sHistorico)) = 0 Then Error 32280
        End If
        'Fim da Alteração

        'Pesquisa Conta associada a um Lançamento
        lErro = Comando_ExecutarLockado(alComando(0), "SELECT TipoConta, Ativo FROM PlanoConta WHERE Conta=?", iTipoConta, iStatus, objLancamento_Detalhe.sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 6073
        
        'Lê a Conta
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6075
        
        If lErro = AD_SQL_SEM_DADOS Then Error 6074
        
        lErro = Comando_LockShared(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 6076
            
        'verifica se a conta está ativa
        If iStatus <> CONTA_ATIVA Then Error 5788
        
        'verifica se a conta é analítica
        If iTipoConta <> CONTA_ANALITICA Then Error 5789
        
        iPossuiCcl = 0
            
        'Se o lançamento envolver centro de custo extra contabil ou contabil
        'Verificar se a associação conta x centro de custo está cadastrada
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL And Len(objLancamento_Detalhe.sCcl) > 0 Then
        
            sCcl = objLancamento_Detalhe.sCcl
            
            sConta = String(STRING_CONTA, 0)
            
            'Pesquisa Ccl associado ao Lançamento
            lErro = Comando_ExecutarLockado(alComando(2), "SELECT Conta FROM ContaCcl WHERE Conta=? AND Ccl=?", sConta, objLancamento_Detalhe.sConta, sCcl)
            If lErro <> AD_SQL_SUCESSO Then Error 5791
        
           'Lê a associação Conta x Ccl
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5792
        
            If lErro = AD_SQL_SEM_DADOS Then Error 5793
        
            lErro = Comando_LockShared(alComando(2))
            If lErro <> AD_SQL_SUCESSO Then Error 5794
            
            lErro = CF("Ccl_Critica_Analitica", sCcl)
            If lErro <> SUCESSO Then Error 32311
            
        End If
        
        If giSetupUsoCcl = CCL_USA_CONTABIL Then
        
            sCcl = String(STRING_CCL, 0)
            
            lErro = Mascara_RetornaCcl(objLancamento_Detalhe.sConta, sCcl)
            If lErro <> AD_SQL_SUCESSO Then Error 9409
        
            objLancamento_Detalhe.sCcl = sCcl
        
            lErro = CF("Ccl_Critica_Analitica", sCcl)
            If lErro <> SUCESSO Then Error 32312
            
        End If
            
        objLancamento_Detalhe.dtData = objLancamento_Cabecalho.dtData
        objLancamento_Detalhe.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
            
        'trata a parte de custo que pode ser tratado neste momento (Custo Medio e Custo Standard)
        lErro = CF("LanPendente_Trata_Produto", alComando(), objLancamento_Detalhe, iApropriacao)
        If lErro <> SUCESSO Then Error 20530
        
        If objLancamento_Detalhe.dtDataEstoque = 0 Then objLancamento_Detalhe.dtDataEstoque = DATA_NULA
        
        If gobjCTB.giValidaCtaCcl = MARCADO And objLancamento_Cabecalho.sOrigem <> "APE" Then
            lErro = CF("Lancamento_Valida_ContaCcl", objLancamento_Detalhe.sConta, objLancamento_Detalhe.sCcl)
            If lErro <> SUCESSO Then Error 20530
        End If
            
        'Faz inserção do Lançamento na Tabela de Lançamentos Pendentes
        'reprocessamento (incluido Quantidade e DataEstoque )
        lErro = Comando_Executar(alComando(1), "INSERT INTO LanPendente (FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Lote,PeriodoLote,Data,Seq,Conta,Ccl,Valor,Historico,NumIntDoc,FilialCliForn,CliForn,Transacao,Aglutina,Produto, ApropriaCRProd, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Quantidade, DataEstoque, Status, Modelo, Gerencial, SubTipo, Usuario, DataRegistro, HoraRegistro) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc, objLancamento_Cabecalho.iLote, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.dtData, objLancamento_Detalhe.iSeq, objLancamento_Detalhe.sConta, objLancamento_Detalhe.sCcl, objLancamento_Detalhe.dValor, objLancamento_Detalhe.sHistorico, objLancamento_Detalhe.lNumIntDoc, objLancamento_Detalhe.iFilialCliForn, objLancamento_Detalhe.lCliForn, objLancamento_Detalhe.iTransacao, objLancamento_Detalhe.iAglutina, objLancamento_Detalhe.sProduto, objLancamento_Cabecalho.iCustoRealProducao, _
                                 objLancamento_Detalhe.lContaSimples, objLancamento_Detalhe.iSeqContraPartida, objLancamento_Detalhe.iEscaninho_Custo, objLancamento_Detalhe.dValorLivroAuxiliar, objLancamento_Detalhe.iClienteFornecedor, objLancamento_Detalhe.sDocOrigem, objLancamento_Detalhe.dQuantidade, objLancamento_Detalhe.dtDataEstoque, objLancamento_Detalhe.iStatus, objLancamento_Detalhe.sModelo, objLancamento_Detalhe.iGerencial, objLancamento_Cabecalho.iSubTipo, gsUsuario, Date, CDbl(Time))
        If lErro <> AD_SQL_SUCESSO Then Error 6014
        
        lErro = Comando_Unlock(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 5787
        
        If iPossuiCcl = 1 Then
            lErro = Comando_Unlock(alComando(2))
            If lErro <> AD_SQL_SUCESSO Then Error 5796
        End If
        
        
        If objLancamento_Detalhe.dValor > 0 Then
            dCredito = dCredito + objLancamento_Detalhe.dValor
        Else
            dDebito = dDebito - objLancamento_Detalhe.dValor
        End If
        
        iNumLancAtual = iNumLancAtual + 1
        
    Next


    dCredito = Format(dCredito, "Fixed")
    dDebito = Format(dDebito, "Fixed")
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Lancamento_Grava_BD = SUCESSO
    
    Exit Function
    
Erro_Lancamento_Grava_BD:

    Lancamento_Grava_BD = Err
    
    Select Case Err
    
        Case 5787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_PLANOCONTA", Err, objLancamento_Detalhe.sConta)
    
        Case 5788
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", Err, objLancamento_Detalhe.sConta)
    
        Case 5789
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, objLancamento_Detalhe.sConta)
    
        Case 5790, 20530, 32311, 32112
    
        Case 5791, 5792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, objLancamento_Detalhe.sConta, sCcl)
    
        Case 5793
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACCL_NAO_CADASTRADO", Err, objLancamento_Detalhe.sConta, sCcl)
    
        Case 5794
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTACCL", Err, objLancamento_Detalhe.sConta, sCcl)
    
        Case 5796
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_CONTACCL", Err, objLancamento_Detalhe.sConta, sCcl)
            
        Case 6012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 6073, 6075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, objLancamento_Detalhe.sConta)
    
        Case 6074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, objLancamento_Detalhe.sConta)

        Case 6076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, objLancamento_Detalhe.sConta)
    
        Case 6014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTOS", Err)
    
        Case 9409
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCcl", Err, objLancamento_Detalhe.sConta)
    
    'Alteracao Daniel em 30/08/2002
        Case 32280
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_SEM_HISTORICO", Err)
    'Fim da Alteração Daniel em 30/08/02
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150570)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function LanPendente_Trata_Produto(alComando() As Long, ByVal objLancamento_Detalhe As ClassLancamento_Detalhe, iApropriacao As Integer) As Long
'trata a parte de custo que pode ser tratado neste momento (Custo Medio e Custo Standard)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim dCustoUnitario As Double

On Error GoTo Erro_LanPendente_Trata_Produto

    If Len(objLancamento_Detalhe.sProduto) > 0 Then
    
        objLancamento_Detalhe.dQuantidade = objLancamento_Detalhe.dValor
        
        lErro = CF("Produto_Critica1", alComando(3), objLancamento_Detalhe.sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 36532

        iApropriacao = objProduto.iApropriacaoCusto

'        'le o custo medio, standard ou custo dos escaninhos
'        lErro = CF("CustoMedio_Le", alComando(), objLancamento_Detalhe, objProduto.iApropriacaoCusto)
'        If lErro <> SUCESSO Then gError 83523
        
        'o codigo acima foi substituido em 26/11/02 pois agora o custo é pego do movimento de estoque associado ao documento de origem
        lErro = CF("Retorna_CustoUnitario", objLancamento_Detalhe.iTransacao, objLancamento_Detalhe.lNumIntDoc, objLancamento_Detalhe.sProduto, dCustoUnitario)
        If lErro <> SUCESSO Then gError 110048
        
        objLancamento_Detalhe.dValor = objLancamento_Detalhe.dValor * dCustoUnitario
        
    End If

    LanPendente_Trata_Produto = SUCESSO
    
    Exit Function
    
Erro_LanPendente_Trata_Produto:

    LanPendente_Trata_Produto = gErr
    
    Select Case gErr
                   
        Case 36532, 83523, 110048
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150571)
            
    End Select
    
    Exit Function
    
End Function

Function LanPendente_Reprocessa(ByVal iOrigemLcto As Integer, ByVal lNumIntDocOrigem As Long, ByVal iFilialEmpresa As Integer) As Long
'atualiza o custo dos lançamentos pendentes que tem o documento origem passado como parametro.
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim alComando(0 To 15) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim iSeq As Integer
Dim objLancamento_Detalhe As New ClassLancamento_Detalhe
Dim tLancamento As typeLancamento
Dim iApropriacao As Integer

On Error GoTo Erro_LanPendente_Reprocessa

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 83811
    Next
    
    tLancamento.sProduto = String(STRING_PRODUTO, 0)
    tLancamento.sOrigem = String(STRING_ORIGEM, 0)
    tLancamento.iFilialEmpresa = iFilialEmpresa
    
    'Pesquisa a existencia dos Lancamentos Pendentes
    lErro = Comando_Executar(alComando(1), "SELECT Origem, Exercicio, PeriodoLan, Doc, Seq, Quantidade, EscaninhoCusto, DataEstoque, Produto, Transacao, NumIntDoc FROM LanPendente, TransacaoCTBCodigo WHERE FilialEmpresa = ? AND NumIntDoc = ? AND LanPendente.Transacao = TransacaoCTBCodigo.Codigo AND TransacaoCTBCodigo.OrigemLcto = ? AND Produto <> '' ", tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq, tLancamento.dQuantidade, tLancamento.iEscaninho_Custo, tLancamento.dtDataEstoque, tLancamento.sProduto, tLancamento.iTransacao, tLancamento.lNumIntDoc, iFilialEmpresa, lNumIntDocOrigem, iOrigemLcto)
    If lErro <> AD_SQL_SUCESSO Then gError 83812
     
    'Lê o primeiro lançamento
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83813

    'Inicia Loop de Alteração do custo
    Do While lErro = AD_SQL_SUCESSO
        
        objLancamento_Detalhe.iFilialEmpresa = iFilialEmpresa
        objLancamento_Detalhe.dQuantidade = tLancamento.dQuantidade
        objLancamento_Detalhe.iEscaninho_Custo = tLancamento.iEscaninho_Custo
        objLancamento_Detalhe.dtDataEstoque = tLancamento.dtDataEstoque
        objLancamento_Detalhe.sProduto = tLancamento.sProduto
        'coloca a quantidade no local de valor pois as rotinas localizadas em LanPenente_Trata_Produto tratam como se em valor estivesse a quantidade. Lembrar que antigamente não existia o campo quantidade em LanPendente.
        objLancamento_Detalhe.dValor = tLancamento.dQuantidade
        objLancamento_Detalhe.iTransacao = tLancamento.iTransacao
        objLancamento_Detalhe.lNumIntDoc = tLancamento.lNumIntDoc
        
        'descobre o  custo que pode ser tratado neste momento (Custo Medio e Custo Standard)
        lErro = CF("LanPendente_Trata_Produto", alComando(), objLancamento_Detalhe, iApropriacao)
        If lErro <> SUCESSO Then gError 83814
        
        If iApropriacao = APROPR_CUSTO_MEDIO Or iApropriacao = APROPR_CUSTO_STANDARD Then
        
            'Pesquisa a existencia dos Lancamentos do Documento no BD
            lErro = Comando_ExecutarPos(alComando(0), "SELECT Seq FROM LanPendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ? AND Seq = ?", 0, iSeq, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
            If lErro <> AD_SQL_SUCESSO Then gError 83815
            
            'Lê o primeiro lançamento
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO Then gError 83816
            
            'Altera o valor do lançamento
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE LanPendente SET Valor=?", alComando(0), objLancamento_Detalhe.dValor)
            If lErro <> AD_SQL_SUCESSO Then gError 83817
        
        End If
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83818
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    LanPendente_Reprocessa = SUCESSO
    
    Exit Function
    
Erro_LanPendente_Reprocessa:

    LanPendente_Reprocessa = gErr
    
    Select Case gErr
    
        Case 83811
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 83812, 83813, 83818
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE7", gErr, lNumIntDocOrigem, iOrigemLcto)
            
        Case 83814
            
        Case 83815, 83816
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE6", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
        Case 83817
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANPENDENTE", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150572)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
                   
End Function

Function Lancamento_Atualiza(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection) As Long
    
Dim lTransacao As Long
Dim lErro As Long
Dim colLancamento_Detalhe_Aux As Collection
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
    
On Error GoTo Erro_Lancamento_Atualiza

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 11548
       
    Do While objLancamento_Cabecalho.iFilialEmpresa > 0 And objLancamento_Cabecalho.iFilialEmpresa < 100
       
        Set colLancamento_Detalhe_Aux = New Collection
       
        For Each objLancamento_Detalhe In colLancamento_Detalhe
        
            If giContabGerencial = 1 Then
                If (objLancamento_Detalhe.iGerencial = 0 Or objLancamento_Detalhe.iGerencial = 2) And ((objLancamento_Cabecalho.iFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (objLancamento_Cabecalho.iFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                    colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
                End If
                    
                If (objLancamento_Detalhe.iGerencial = 1 Or objLancamento_Detalhe.iGerencial = 2) And ((objLancamento_Cabecalho.iFilialEmpresa < giFilialAuxiliar And giFilialAuxiliar > 0) Or (objLancamento_Cabecalho.iFilialEmpresa > Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
                End If
                
            Else
                    
                'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
            End If
                
        Next
       
        lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe_Aux)
        If lErro <> SUCESSO Then Error 20475
    
        If giContabGerencial = 0 Then Exit Do

        objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa - giFilialAuxiliar
        
    Loop
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 11551

    Lancamento_Atualiza = SUCESSO

    Exit Function

Erro_Lancamento_Atualiza:
    
    Lancamento_Atualiza = Err
    
    Select Case Err
                   
        Case 11548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
                       
        Case 11551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
                      
        Case 20475
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150573)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Function Lancamento_Atualiza0(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection, Optional ByVal bExclusaoFisica As Boolean = False) As Long
'Atualiza os lançamentos passados como parametro. Continuação de Lancamento_Atualiza.
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
    
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim iExercicio As Integer
Dim iPeriodo As Integer
Dim iStatus As Integer
Dim objPeriodosFilial As New ClassPeriodosFilial
    
On Error GoTo Erro_Lancamento_Atualiza0

    For iIndice = LBound(alComando) To UBound(alComando)
        
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 11538
        
    Next

    'verifica se o exercicio não está fechado
    lErro = Comando_ExecutarLockado(alComando(1), "SELECT Exercicio, Status FROM Exercicios WHERE DataInicio <= ? AND DataFim >= ?", iExercicio, iStatus, objLancamento_Cabecalho.dtData, objLancamento_Cabecalho.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 11540
    
    'le o exercicio, se estiver no BD
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 11541
    
    objLancamento_Cabecalho.iExercicio = iExercicio
    
    lErro = Comando_LockShared(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 11542
    
    'se o exercicio estiver fechado, não pode editar os lançamentos
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If iStatus = EXERCICIO_FECHADO Then Error 11543
    End If
    
    'seleciona o periodo em questão
    lErro = Comando_ExecutarLockado(alComando(2), "SELECT Periodo FROM Periodo WHERE Exercicio=? AND DataInicio <= ? AND DataFim >= ?", iPeriodo, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.dtData, objLancamento_Cabecalho.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 11544
    
    'le o periodo, se estiver no BD
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 11545
    
    objLancamento_Cabecalho.iPeriodoLan = iPeriodo
    objLancamento_Cabecalho.iPeriodoLote = iPeriodo
    
    'lock do periodo
    lErro = Comando_LockExclusive(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 11546
    
    objPeriodosFilial.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
    objPeriodosFilial.iExercicio = objLancamento_Cabecalho.iExercicio
    objPeriodosFilial.iPeriodo = objLancamento_Cabecalho.iPeriodoLan
    objPeriodosFilial.sOrigem = objLancamento_Cabecalho.sOrigem

    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 20851
    
    'se o periodo estiver fechado, não pode alterar os lançamentos do periodo
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 11547
    End If
    
    
    'se o periodo estiver fechado, não pode alterar os lançamentos do periodo
    lErro = Lancamento_Atualiza1(objLancamento_Cabecalho, colLancamento_Detalhe, bExclusaoFisica)
    If lErro <> SUCESSO Then Error 10165

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Lancamento_Atualiza0 = SUCESSO

    Exit Function

Erro_Lancamento_Atualiza0:
    
    Lancamento_Atualiza0 = Err
    
    Select Case Err
    
        Case 10165, 20851
            
        Case 10166, 10167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iExercicio, iPeriodo)
            
        Case 10168
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODOSFILIAL", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iExercicio, iPeriodo)
            
        Case 11538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 11540, 11541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOS", Err)
           
        Case 11542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, objLancamento_Cabecalho.iExercicio)
                       
        Case 11543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objLancamento_Cabecalho.iExercicio)
                    
        Case 11544, 11545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO1", Err)
            
        Case 11546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODO", Err, objLancamento_Cabecalho.iExercicio, iPeriodo)
            
        Case 11547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, iPeriodo, objLancamento_Cabecalho.iExercicio)
                   
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150574)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Lancamento_Atualiza1(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection, Optional ByVal bExclusaoFisica As Boolean = False) As Long
'Atualiza os lançamentos passados como parametro. Continuação de Lancamento_Atualiza0.
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lErro As Long
Dim dDebito As Double
Dim dCredito As Double
Dim iNumLancAtual As Integer

On Error GoTo Erro_Lancamento_Atualiza1

    'pesquisa se o número já foi utilizado em lançamentos contabilizados
    lErro = CF("Lancamento_Le_Doc", objLancamento_Cabecalho)
    If lErro <> SUCESSO And lErro <> 5726 Then Error 9587
    
    'Se este voucher já foi usado ==> erro
    If lErro = SUCESSO Then Error 9588
    
    'pesquisa se o número já foi utilizado em lançamentos pendentes
    lErro = CF("LanPendente_Le_Doc1", objLancamento_Cabecalho)
    If lErro <> SUCESSO And lErro <> 9590 Then Error 9595
    
    'Se este voucher já foi usado ==> erro
    If lErro = SUCESSO Then Error 9596
    
    lErro = CF("Lancamento_Atualiza_BD", objLancamento_Cabecalho, colLancamento_Detalhe, dDebito, dCredito, iNumLancAtual)
    If lErro <> SUCESSO Then Error 11549
            
    lErro = CF("Rotina_Atualiza_Lancamentos", objLancamento_Cabecalho, bExclusaoFisica)
    If lErro <> SUCESSO Then Error 11550
        
    Lancamento_Atualiza1 = SUCESSO

    Exit Function

Erro_Lancamento_Atualiza1:
    
    Lancamento_Atualiza1 = Err
    
    Select Case Err
    
        Case 9587, 9595
    
        Case 9588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_ATUALIZADO", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
    
        Case 9596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_PENDENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
                   
        Case 11549, 11550
                                                                                  
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150575)
            
    End Select
    
    Exit Function
    
End Function

Function Lancamento_Atualiza_BD(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection, dCredito As Double, dDebito As Double, iNumLancAtual As Integer) As Long
'grava na tabela de lançamentos pendentes os lançamentos do voucher em questão
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim iTipoConta As Integer
Dim iStatus As Integer
Dim sCcl As String
Dim sConta As String
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim iPossuiCcl As Integer
Dim objHistPadrao As New ClassHistPadrao

On Error GoTo Erro_Lancamento_Atualiza_BD

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11552
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 11553
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 11554
    
    'Inicia Loop de GRAVACAO com Lançamentos do GRID
    For Each objLancamento_Detalhe In colLancamento_Detalhe
            
        'Alteração Daniel em 30/08/2006
        If giHistoricoObrigatorio = MARCADO Then
            If Len(Trim(objLancamento_Detalhe.sHistorico)) = 0 Then Error 32280
        End If
        'Fim da Alteração
        
        'Pesquisa Conta associada a um Lançamento
        lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta, Ativo FROM PlanoConta WHERE Conta=?", iTipoConta, iStatus, objLancamento_Detalhe.sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 11555
        
       'Lê a Conta
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11556
        
        If lErro = AD_SQL_SEM_DADOS Then Error 11557
        
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 11558
            
        'verifica se a conta está ativa
        If iStatus <> CONTA_ATIVA Then Error 11559
        
        'verifica se a conta é analítica
        If iTipoConta <> CONTA_ANALITICA Then Error 11560
                    
        iPossuiCcl = 0
            
        'Se o lançamento envolver centro de custo extra contabil ou contabil
        'Verificar se a associação conta x centro de custo está cadastrada
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL And Len(objLancamento_Detalhe.sCcl) > 0 Then
        
            sCcl = objLancamento_Detalhe.sCcl
            
            sConta = String(STRING_CONTA, 0)
            
            'Pesquisa Ccl associado ao Lançamento
            lErro = Comando_ExecutarLockado(lComando2, "SELECT Conta FROM ContaCcl WHERE Conta=? AND Ccl=?", sConta, objLancamento_Detalhe.sConta, sCcl)
            If lErro <> AD_SQL_SUCESSO Then Error 11561
        
           'Lê a associação Conta x Ccl
            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11562
        
            If lErro = AD_SQL_SEM_DADOS Then Error 11563
        
            lErro = Comando_LockShared(lComando2)
            If lErro <> AD_SQL_SUCESSO Then Error 11564
            
            lErro = CF("Ccl_Critica_Analitica", sCcl)
            If lErro <> SUCESSO Then Error 32313
        
        End If
        
        If giSetupUsoCcl = CCL_USA_CONTABIL Then
        
            sCcl = String(STRING_CCL, 0)
            
            lErro = Mascara_RetornaCcl(objLancamento_Detalhe.sConta, sCcl)
            If lErro <> SUCESSO Then Error 11565
            
            objLancamento_Detalhe.sCcl = sCcl
            
            lErro = CF("Ccl_Critica_Analitica", sCcl)
            If lErro <> SUCESSO Then Error 32314
        
        End If
            
        If objLancamento_Detalhe.dtDataEstoque = 0 Then objLancamento_Detalhe.dtDataEstoque = DATA_NULA

        If gobjCTB.giValidaCtaCcl = MARCADO And objLancamento_Cabecalho.sOrigem <> "APE" Then
            lErro = CF("Lancamento_Valida_ContaCcl", objLancamento_Detalhe.sConta, objLancamento_Detalhe.sCcl)
            If lErro <> SUCESSO Then Error 32314
        End If
        
        'Faz inserção do Lançamento na Tabela de Lançamentos Pendentes
        lErro = Comando_Executar(lComando1, "INSERT INTO LanPendente (FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Data,Seq,PeriodoLote,Lote,Conta,Ccl,Valor,Historico,NumIntDoc,FilialCliForn,CliForn,Transacao,Aglutina, Produto, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Quantidade, DataEstoque, ApropriaCRProd, Status, Modelo, Gerencial, SubTipo, Usuario, DataRegistro, HoraRegistro) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc, objLancamento_Cabecalho.dtData, objLancamento_Detalhe.iSeq, objLancamento_Cabecalho.iPeriodoLan, 0, objLancamento_Detalhe.sConta, objLancamento_Detalhe.sCcl, objLancamento_Detalhe.dValor, objLancamento_Detalhe.sHistorico, objLancamento_Detalhe.lNumIntDoc, objLancamento_Detalhe.iFilialCliForn, objLancamento_Detalhe.lCliForn, objLancamento_Detalhe.iTransacao, objLancamento_Detalhe.iAglutina, objLancamento_Detalhe.sProduto, objLancamento_Detalhe.lContaSimples, objLancamento_Detalhe.iSeqContraPartida, objLancamento_Detalhe.iEscaninho_Custo, objLancamento_Detalhe.dValorLivroAuxiliar, objLancamento_Detalhe.iClienteFornecedor, objLancamento_Detalhe.sDocOrigem, objLancamento_Detalhe.dQuantidade, objLancamento_Detalhe.dtDataEstoque, objLancamento_Detalhe.iApropriaCRProd, objLancamento_Detalhe.iStatus, objLancamento_Detalhe.sModelo, _
        objLancamento_Detalhe.iGerencial, objLancamento_Cabecalho.iSubTipo, gsUsuario, Date, CDbl(Time))
        If lErro <> AD_SQL_SUCESSO Then Error 11566
        
        lErro = Comando_Unlock(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 11567
        
        If iPossuiCcl = 1 Then
            lErro = Comando_Unlock(lComando2)
            If lErro <> AD_SQL_SUCESSO Then Error 11568
        End If
        
        
        If objLancamento_Detalhe.dValor > 0 Then
            dCredito = dCredito + objLancamento_Detalhe.dValor
        Else
            dDebito = dDebito - objLancamento_Detalhe.dValor
        End If
        
        iNumLancAtual = iNumLancAtual + 1
        
    Next

    dCredito = Format(dCredito, "Fixed")
    dDebito = Format(dDebito, "Fixed")
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Lancamento_Atualiza_BD = SUCESSO
    
    Exit Function
    
Erro_Lancamento_Atualiza_BD:

    Lancamento_Atualiza_BD = Err
    
    Select Case Err
    
        Case 32313, 32314
    
        Case 11552, 11553, 11554
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 11555, 11556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, objLancamento_Detalhe.sConta)
    
        Case 11557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, objLancamento_Detalhe.sConta)

        Case 11558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, objLancamento_Detalhe.sConta)
    
        Case 11559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", Err, objLancamento_Detalhe.sConta)
    
        Case 11560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, objLancamento_Detalhe.sConta)
    
        Case 11561, 11562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, objLancamento_Detalhe.sConta, sCcl)
    
        Case 11563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACCL_NAO_CADASTRADO", Err, objLancamento_Detalhe.sConta, sCcl)
    
        Case 11564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTACCL", Err, objLancamento_Detalhe.sConta, sCcl)
    
        Case 11565
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCcl", Err, objLancamento_Detalhe.sConta)
    
        Case 11566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTOS", Err)
    
        Case 11567
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_PLANOCONTA", Err, objLancamento_Detalhe.sConta)
    
        Case 11568
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_CONTACCL", Err, objLancamento_Detalhe.sConta, sCcl)
            
    'Alteracao Daniel em 30/08/2002
        Case 32280
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_SEM_HISTORICO", Err)
    'Fim da Alteração Daniel em 30/08/02
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150576)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function Lancamento_Exclui_BD(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, dCredito As Double, dDebito As Double, iNumLancAtual As Integer, iNumDocAtual As Integer) As Long
'exclui da tabela de lançamentos pendentes os lançamentos do voucher em questão
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim dValor As Double
Dim iSeq As Integer
Dim iTransacao As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_Lancamento_Exclui_BD

    'Abre comando para pesquisar Lancamentos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6082
    
    'Abre comando para transação de exclusão
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 6083
    
    'Pesquisa a existencia dos Lancamentos do Documento no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT Transacao, NumIntDoc, Valor FROM LanPendente WHERE FilialEmpresa=? AND Origem=? AND Exercicio=? AND PeriodoLan=? AND Doc=?", 0, iTransacao, lNumIntDoc, dValor, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 6085
     
    'Lê o primeiro lançamento
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6086

    If objLancamento_Cabecalho.iTransacao <> 0 And objLancamento_Cabecalho.lNumIntDoc <> 0 And lErro = AD_SQL_SUCESSO Then
        If objLancamento_Cabecalho.iTransacao <> iTransacao Or objLancamento_Cabecalho.lNumIntDoc <> lNumIntDoc Then Error 20727
    End If

    If lErro = SUCESSO Then iNumDocAtual = iNumDocAtual - 1

    'Inicia Loop de EXCLUSAO com Lançamentos do GRID
    Do While lErro <> AD_SQL_SEM_DADOS  'Refere-se ao BuscarProximo
        
        'Exclui Lançamento da Tabela de Lançamentos Pendentes
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM LanPendente", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 6087
        
        If dValor > 0 Then
            dCredito = dCredito - dValor
        Else
            dDebito = dDebito + dValor
        End If
        
        iNumLancAtual = iNumLancAtual - 1
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6088
    
    Loop
    
    dCredito = Format(dCredito, "Fixed")
    dDebito = Format(dDebito, "Fixed")
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Lancamento_Exclui_BD = SUCESSO
    
    Exit Function
    
Erro_Lancamento_Exclui_BD:

    Lancamento_Exclui_BD = Err
    
    Select Case Err
    
        Case 6082, 6083
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 6085, 6086, 6088
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS_PENDENTES", Err)
            
        Case 6087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LANCAMENTO", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc, iSeq)
            
        Case 20727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANPENDENTE_TRANSACAO_NUMINTDOC", Err, iTransacao, lNumIntDoc, objLancamento_Cabecalho.iTransacao, objLancamento_Cabecalho.lNumIntDoc)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150577)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
                   
End Function

Function LotePendente_Grava(ByVal objLote As ClassLote) As Long
'grava o lote no banco de dados
    
Dim lTransacao As Long
Dim lErro As Long
    
On Error GoTo Erro_LotePendente_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 5416
        
    Do While objLote.iFilialEmpresa > 0 And objLote.iFilialEmpresa < 100
       
        lErro = CF("LotePendente_Grava_Trans", objLote)
        If lErro <> SUCESSO Then Error 32166
       
        If giContabGerencial = 0 Then Exit Do

        objLote.iFilialEmpresa = objLote.iFilialEmpresa - giFilialAuxiliar
        
    Loop
        
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 5420
    
    LotePendente_Grava = SUCESSO

    Exit Function

Erro_LotePendente_Grava:
    
    LotePendente_Grava = Err
    
    Select Case Err
    
        Case 32166
        
        Case 5416
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 5420
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150578)
        
    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Function LotePendente_Grava_Totais(ByVal objLote As ClassLote) As Long

Dim lTransacao As Long
Dim lErro As Long
    
On Error GoTo Erro_LotePendente_Grava_Totais

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 8039
    
    'faz a atualizacao do bd
    lErro = CF("LotePendente_Grava_Totais_Trans", objLote)
    If lErro <> SUCESSO Then Error 32165
            
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 8043

    LotePendente_Grava_Totais = SUCESSO

    Exit Function

Erro_LotePendente_Grava_Totais:
    
    LotePendente_Grava_Totais = Err
    
    Select Case Err
    
        Case 32165
            
        Case 8039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 8043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150579)
        
    End Select
    
     Call Transacao_Rollback
    
    Exit Function

End Function

Function LotePendente_Exclui(ByVal objLote As ClassLote) As Long
'exclui o lote do banco de dados
    
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iLote As Integer
Dim iLoteAtualizado As Integer
Dim objPeriodosFilial As New ClassPeriodosFilial
Dim iStatus As Integer

On Error GoTo Erro_LotePendente_Exclui

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5421

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5422

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5760

    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 5765

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 5426
            
    Do While objLote.iFilialEmpresa > 0 And objLote.iFilialEmpresa < 100
       
        'verifica se o exercicio não está fechado
        lErro = Comando_ExecutarLockado(lComando3, "SELECT Status FROM Exercicios WHERE Exercicio=?", iStatus, objLote.iExercicio)
        If lErro <> AD_SQL_SUCESSO Then Error 5761
        
        'le o exercicio, se estiver no BD
        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 5762
        
        lErro = Comando_LockShared(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 5763
        
        If giDesconsideraFechamentoPeriodo <> MARCADO Then
            If iStatus = EXERCICIO_FECHADO Then Error 5764
        End If
                
        objPeriodosFilial.iFilialEmpresa = objLote.iFilialEmpresa
        objPeriodosFilial.iExercicio = objLote.iExercicio
        objPeriodosFilial.iPeriodo = objLote.iPeriodo
        objPeriodosFilial.sOrigem = objLote.sOrigem

        lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
        If lErro <> SUCESSO Then Error 20852
        
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 5769
                
        'verifica se o lote  está atualizado
        lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
        If lErro <> SUCESSO Then Error 5993
        
        If iLoteAtualizado = LOTE_ATUALIZADO Then Error 5425
                
        'Pesquisa a existencia do lote no BD
        lErro = Comando_ExecutarPos(lComando1, "SELECT Lote FROM LotePendente WHERE FilialEmpresa = ? AND Origem=? AND Exercicio=? AND Periodo=? AND Lote=?", 0, iLote, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 5423
    
        'le o lote, se existir
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 5424
        
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 5517
        
        'verificar se o lote possui lançamentos pendentes. Se tiver, não pode excluir
        lErro = CF("LanPendente_Le_Lote", objLote)
        If lErro <> SUCESSO And lErro <> 5524 Then Error 5527
            
        If lErro = SUCESSO Then Error 5526
        
        'exclui o lote do BD
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM LotePendente", lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 5427
            
        If giContabGerencial = 0 Then Exit Do

        objLote.iFilialEmpresa = objLote.iFilialEmpresa - giFilialAuxiliar
        
    Loop
            
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 5428

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    LotePendente_Exclui = SUCESSO

    Exit Function

Erro_LotePendente_Exclui:
    
    LotePendente_Exclui = Err
    
    Select Case Err
    
        Case 5421, 5422
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5423
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 5424
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTEPENDENTE_NAO_CADASTRADO", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 5425
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_EXCLUIR", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 5426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 5427
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            
        Case 5428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case 5526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_COM_LANC_PEND_NAO_EXCLUIR", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 5517
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 5527, 5993, 20852
        
        Case 5760
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5761, 5762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, objLote.iExercicio)
            
        Case 5763
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, objLote.iExercicio)
        
        Case 5764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_EXERCICIO_FECHADO", Err, objLote.iExercicio)
        
        Case 5765
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5766, 5767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo)
            
        Case 5768
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODOSFILIAL", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo)
        
        Case 5769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_PERIODO_FECHADO", Err, objLote.iExercicio, objLote.iPeriodo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150580)
        
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    Exit Function

End Function

Function Lote_Automatico(ByVal objLote As ClassLote) As Long
'retorna o número do proximo lote disponivel

Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lErro As Long
Dim lErro1 As Integer
Dim iFim_de_Arquivo As Integer
Dim iLote As Integer
Dim iLote2 As Integer
Dim lTransacao As Long
Dim iPeriodo As Integer

On Error GoTo Erro_Lote_Automatico

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5390
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5392

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 9001

    'se a numeração do lote não é reinicializada a cada periodo ==> não leva em consideração o periodo passado como parametro
    If giSetupLotePorPeriodo = 0 Then
        iPeriodo = 0
    Else
        iPeriodo = objLote.iPeriodo
    End If
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 5397

    'Pesquisa o número do lote gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT Lote FROM ExercicioOrigem WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ? AND Origem = ?", 0, iLote, objLote.iFilialEmpresa, objLote.iExercicio, iPeriodo, objLote.sOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 5393

    'le o registro com o número do lote automático
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5394

    'loca o registro que guarda o numero do lote gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5398

    'verifica se este número ainda não foi utilizado
    'Pesquisa as numerações de lote maiores ou iguais a numeração gerada automaticamente
    If giSetupLotePorPeriodo Then
    
        lErro = Comando_Executar(lComando3, "SELECT Lote FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote >= ? ORDER BY Lote", iLote2, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 9002
        
    Else
    
        lErro = Comando_Executar(lComando3, "SELECT Lote FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Lote >= ? ORDER BY Lote", iLote2, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 9006
        
    End If
        
    'le o primeiro lote pendente >= ao lote automatico
    lErro1 = Comando_BuscarPrimeiro(lComando3)
    If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then Error 9003

    'pesquisa o proximo número de lote disponível
    Do While lErro1 <> AD_SQL_SEM_DADOS
        
        If iLote < iLote2 Then
        
            objLote.iLote = iLote
        
            'pesquisa se o número já foi utilizado em um lote contabilizado
            lErro = CF("Lote_Le", objLote)
            If lErro <> SUCESSO And lErro <> 9293 Then Error 5730
            
            'Se ainda não foi usado ==> este é o número a ser usado
            If lErro = 9293 Then Exit Do
        
            'Se o número já foi usado ==> pega o próximo
            iLote = iLote + 1
        
        
        Else
        
            iLote = iLote + 1
            
            lErro1 = Comando_BuscarProximo(lComando3)
            If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then Error 9005
        
        End If
        
    Loop
        
    'atualiza o numero do lote gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE ExercicioOrigem SET Lote = ?", lComando, iLote + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 5399
        
    objLote.iLote = iLote
        
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 5400
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Lote_Automatico = SUCESSO
        
    Exit Function
    
    
Erro_Lote_Automatico:

    Lote_Automatico = Err

    Select Case Err
    
    Case 5390, 5392, 9001
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
    Case 5393, 5394
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOORIGEM", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo, objLote.sOrigem)
        
    Case 9002, 9003, 9005, 9006
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE1", Err)
        
    Case 5397
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
    Case 5398
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOORIGEM", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo, objLote.sOrigem)
        
    Case 5399
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOORIGEM", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo, objLote.sOrigem)
        
    Case 5400
        lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
    Case 5730
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150581)
        
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function
    
End Function

Function Lote_Automatico1(ByVal objLote As ClassLote) As Long
'retorna o número do proximo lote disponivel
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lErro As Long
Dim lErro1 As Long
Dim iLote As Integer
Dim iLote2 As Integer
Dim iPeriodo As Integer

On Error GoTo Erro_Lote_Automatico1

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 39710
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 39711

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 39712

    'se a numeração do lote não é reinicializada a cada periodo ==> não leva em consideração o periodo passado como parametro
    If giSetupLotePorPeriodo = 0 Then
        iPeriodo = 0
    Else
        iPeriodo = objLote.iPeriodo
    End If
    
    'Pesquisa o número do lote gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT Lote FROM ExercicioOrigem WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ? AND Origem = ?", 0, iLote, objLote.iFilialEmpresa, objLote.iExercicio, iPeriodo, objLote.sOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 39713

    'le o registro com o número do lote automático
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 39714

    'loca o registro que guarda o numero do lote gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 39715

    'verifica se este número ainda não foi utilizado
    'Pesquisa as numerações de lote maiores ou iguais a numeração gerada automaticamente
    If giSetupLotePorPeriodo Then
    
        lErro = Comando_Executar(lComando3, "SELECT Lote FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote >= ? ORDER BY Lote", iLote2, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 39716
        
    Else
    
        lErro = Comando_Executar(lComando3, "SELECT Lote FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Lote >= ? ORDER BY Lote", iLote2, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 39717
        
    End If
        
    'le o primeiro lote pendente >= ao lote automatico
    lErro1 = Comando_BuscarPrimeiro(lComando3)
    If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then Error 39718

    'pesquisa o proximo número de lote disponível
    Do While lErro1 <> AD_SQL_SEM_DADOS
        
        If iLote < iLote2 Then
        
            objLote.iLote = iLote
        
            'pesquisa se o número já foi utilizado em um lote contabilizado
            lErro = CF("Lote_Le", objLote)
            If lErro <> SUCESSO And lErro <> 9293 Then Error 39719
            
            'Se ainda não foi usado ==> este é o número a ser usado
            If lErro = 9293 Then Exit Do
        
            'Se o número já foi usado ==> pega o próximo
            iLote = iLote + 1
        
        
        Else
            iLote = iLote + 1
            lErro1 = Comando_BuscarProximo(lComando3)
            If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then Error 39720
        
        End If
        
    Loop
        
    'atualiza o numero do lote gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE ExercicioOrigem SET Lote = ?", lComando, iLote + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 39721
        
    objLote.iLote = iLote
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Lote_Automatico1 = SUCESSO
        
    Exit Function
    
    
Erro_Lote_Automatico1:

    Lote_Automatico1 = Err

    Select Case Err
    
    Case 39710, 39711, 39712
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
    Case 39713, 39714
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOORIGEM", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo, objLote.sOrigem)
        
    Case 39715
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOORIGEM", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo, objLote.sOrigem)
        
    Case 39716, 39717, 39718, 39720
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE1", Err)
        
    Case 39719
    
    Case 39721
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOORIGEM", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo, objLote.sOrigem)
        
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150582)
        
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function
    
End Function

Function Voucher_Automatico(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal sOrigem As String, lDoc As Long) As Long
'retorna o número do proximo voucher (documento) disponivel

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim iFim_de_Arquivo2 As Integer
Dim lDoc1 As Long
Dim lDoc3 As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim iFilialEmpresaAux As Integer
Dim lDoc2 As Long

On Error GoTo Erro_Voucher_Automatico

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 64349
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 64350
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 64351

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 64352

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 64353
    
    'se a numeração do voucher(documento) não é reinicializada a cada periodo ==> não leva em consideração o periodo passado como parametro
    If giSetupDocPorPeriodo = 0 Then
        iPeriodo = 0
    End If
    
    'Pesquisa o ultimo número de voucher gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT Doc FROM ExercicioOrigem WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ? AND Origem = ?", 0, lDoc, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 64354

    'le o registro com o número do voucher
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 64355

    'loca o registro que guarda o numero do voucher gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 64356

    'verifica se este número ainda não foi utilizado por LanPedente
    'Pesquisa as numerações de voucher maiores ou iguais a numeração gerada automaticamente
    If giSetupDocPorPeriodo Then
    
        lErro = Comando_Executar(lComando1, "SELECT DISTINCT Doc FROM LanPendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc >= ? ORDER BY Doc", lDoc1, iFilialEmpresa, sOrigem, iExercicio, iPeriodo, lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 64357
        
    Else
    
        lErro = Comando_Executar(lComando1, "SELECT DISTINCT Doc FROM LanPendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Doc >= ? ORDER BY Doc", lDoc1, iFilialEmpresa, sOrigem, iExercicio, lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 64358
        
    End If
        
    'le o primeiro voucher >= ao voucher automatatico (LanPedente)
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64359
        
    'verifica se este número ainda não foi utilizado Lancamento
    'Pesquisa as numerações de voucher maiores ou iguais a numeração gerada automaticamente
    If giSetupDocPorPeriodo Then
    
        lErro = Comando_Executar(lComando3, "SELECT DISTINCT Doc FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc >= ? ORDER BY Doc", lDoc3, iFilialEmpresa, sOrigem, iExercicio, iPeriodo, lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 64360
        
    Else
    
        lErro = Comando_Executar(lComando3, "SELECT DISTINCT Doc FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Doc >= ? ORDER BY Doc", lDoc3, iFilialEmpresa, sOrigem, iExercicio, lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 64361
        
    End If
    
    'le o primeiro voucher >= ao voucher automatico (Lancamentos)
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64362
    
    Do While lDoc = lDoc1 Or lDoc = lDoc3
            
        If lDoc1 = lDoc Then
            iFim_de_Arquivo = Comando_BuscarProximo(lComando1)
            If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then gError 64363
        ElseIf lDoc3 = lDoc Then
            iFim_de_Arquivo2 = Comando_BuscarProximo(lComando3)
            If iFim_de_Arquivo2 <> AD_SQL_SUCESSO And iFim_de_Arquivo2 <> AD_SQL_SEM_DADOS Then gError 64364
        End If
            
        lDoc = lDoc + 1
            
    Loop
    
    'atualiza o número do voucher gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE ExercicioOrigem SET Doc = ?", lComando, lDoc + 1)
    If lErro <> AD_SQL_SUCESSO Then gError 64365
        
        
    If giContabGerencial = 1 Then
        
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 193323
        Next
        
        If iFilialEmpresa > Abs(giFilialAuxiliar) Then
            iFilialEmpresaAux = iFilialEmpresa - Abs(giFilialAuxiliar)
        Else
            iFilialEmpresaAux = iFilialEmpresa + Abs(giFilialAuxiliar)
        End If
            
        'Pesquisa o ultimo número de voucher gerado automaticamente pelo sistema
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Doc FROM ExercicioOrigem WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ? AND Origem = ?", 0, lDoc2, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 193324
    
        'le o registro com o número do voucher
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 193325
    
        'loca o registro que guarda o numero do voucher gerado automaticamente
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 193326
        
        'atualiza o numero do voucher gerado automaticamente
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ExercicioOrigem SET Doc = ?", alComando(1), lDoc + 1)
        If lErro <> AD_SQL_SUCESSO Then gError 193327
        
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
        
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 64366
    
    Voucher_Automatico = SUCESSO
        
    Exit Function
        
Erro_Voucher_Automatico:

    Voucher_Automatico = gErr

    Select Case gErr
    
        Case 64349, 64350, 64351, 64352, 193323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 64353
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 64354, 64355
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOORIGEM", gErr, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
        
        Case 64356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOORIGEM", gErr, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
        
        Case 64357, 64358, 64359, 64363
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS_PENDENTES", gErr)
        
        Case 64360, 64361, 64362, 64364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)
        
        Case 64365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOORIGEM", gErr, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
        
        Case 64366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 193324, 193325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOORIGEM", gErr, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        
        Case 193326
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOORIGEM", gErr, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        
        Case 193327
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOORIGEM", gErr, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150583)
        
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function Voucher_Automatico_Trans(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal sOrigem As String, lDoc As Long) As Long
'retorna o número do proximo voucher (documento) disponivel
'Tem que ser chamado em transacao

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim iFim_de_Arquivo2 As Integer
Dim lDoc1 As Long
Dim lDoc3 As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim iFilialEmpresaAux As Integer
Dim lDoc2 As Long

On Error GoTo Erro_Voucher_Automatico_Trans

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 64333
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 64334
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 64335

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 64336

    'se a numeração do voucher(documento) não é reinicializada a cada periodo ==> não leva em consideração o periodo passado como parametro
    If giSetupDocPorPeriodo = 0 Then
        iPeriodo = 0
    End If
    
    'Pesquisa o ultimo número de voucher gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT Doc FROM ExercicioOrigem WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ? AND Origem = ?", 0, lDoc, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 64337

    'le o registro com o número do voucher
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 64338

    'loca o registro que guarda o numero do voucher gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 64339

    'verifica se este número ainda não foi utilizado por LanPedente
    'Pesquisa as numerações de voucher maiores ou iguais a numeração gerada automaticamente
    If giSetupDocPorPeriodo Then
    
        lErro = Comando_Executar(lComando1, "SELECT DISTINCT Doc FROM LanPendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc >= ? ORDER BY Doc", lDoc1, iFilialEmpresa, sOrigem, iExercicio, iPeriodo, lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 64340
        
    Else
    
        lErro = Comando_Executar(lComando1, "SELECT DISTINCT Doc FROM LanPendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Doc >= ? ORDER BY Doc", lDoc1, iFilialEmpresa, sOrigem, iExercicio, lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 64341
        
    End If
        
    'le o primeiro voucher >= ao voucher automatatico (LanPedente)
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64342
        
    'verifica se este número ainda não foi utilizado Lancamento
    'Pesquisa as numerações de voucher maiores ou iguais a numeração gerada automaticamente
    If giSetupDocPorPeriodo Then
    
        lErro = Comando_Executar(lComando3, "SELECT DISTINCT Doc FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc >= ? ORDER BY Doc", lDoc3, iFilialEmpresa, sOrigem, iExercicio, iPeriodo, lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 64343
        
    Else
    
        lErro = Comando_Executar(lComando3, "SELECT DISTINCT Doc FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Doc >= ? ORDER BY Doc", lDoc3, iFilialEmpresa, sOrigem, iExercicio, lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 64344
        
    End If
    
    'le o primeiro voucher >= ao voucher automatico (Lancamentos)
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64345
    
    Do While lDoc = lDoc1 Or lDoc = lDoc3
            
        If lDoc1 = lDoc Then
            iFim_de_Arquivo = Comando_BuscarProximo(lComando1)
            If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then gError 64346
        ElseIf lDoc3 = lDoc Then
            iFim_de_Arquivo2 = Comando_BuscarProximo(lComando3)
            If iFim_de_Arquivo2 <> AD_SQL_SUCESSO And iFim_de_Arquivo2 <> AD_SQL_SEM_DADOS Then gError 64347
        End If
            
        lDoc = lDoc + 1
            
    Loop
    
    'atualiza o número do voucher gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE ExercicioOrigem SET Doc = ?", lComando, lDoc + 1)
    If lErro <> AD_SQL_SUCESSO Then gError 64348
        
    If giContabGerencial = 1 Then
        
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 193328
        Next
        
        If iFilialEmpresa > Abs(giFilialAuxiliar) Then
            iFilialEmpresaAux = iFilialEmpresa - Abs(giFilialAuxiliar)
        Else
            iFilialEmpresaAux = iFilialEmpresa + Abs(giFilialAuxiliar)
        End If
            
        'Pesquisa o ultimo número de voucher gerado automaticamente pelo sistema
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Doc FROM ExercicioOrigem WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ? AND Origem = ?", 0, lDoc2, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 193329
    
        'le o registro com o número do voucher
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 193330
    
        'loca o registro que guarda o numero do voucher gerado automaticamente
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 193331
        
        'atualiza o numero do voucher gerado automaticamente
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ExercicioOrigem SET Doc = ?", alComando(1), lDoc + 1)
        If lErro <> AD_SQL_SUCESSO Then gError 193332
        
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Voucher_Automatico_Trans = SUCESSO
        
    Exit Function
        
Erro_Voucher_Automatico_Trans:

    Voucher_Automatico_Trans = gErr

    Select Case gErr
    
        Case 64333, 64334, 64335, 64336, 193328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 64337, 64338
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOORIGEM", gErr, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
        
        Case 64339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOORIGEM", gErr, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
        
        Case 64340, 64341, 64342, 64346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS_PENDENTES", gErr)
        
        Case 64343, 64344, 64345, 64347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)
        
        Case 64348
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOORIGEM", gErr, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
        
        Case 193329, 193330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOORIGEM", gErr, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        
        Case 193331
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOORIGEM", gErr, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        
        Case 193332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOORIGEM", gErr, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150584)
        
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function DocAuto_Automatico(lDoc As Long) As Long
'retorna o número do proximo voucher (documento) disponivel

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim lDoc1 As Long
Dim lTransacao As Long
Dim lDoc2 As Long

On Error GoTo Erro_DocAuto_Automatico

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11021
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 11022
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 11035
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 11023

    'Pesquisa o ultimo número de voucher gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT DocAuto FROM Configuracao", 0, lDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 11024

    'le o registro com o número do voucher
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 11032

    'loca o registro que guarda o numero do voucher gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 11033

    'verifica se este número ainda não foi utilizado
    'Pesquisa as numerações de voucher maiores ou iguais a numeração gerada automaticamente
    lErro = Comando_Executar(lComando1, "SELECT Doc FROM DocAuto WHERE  Doc >= ? ORDER BY Doc", lDoc1, lDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 11034
        
    'le o primeiro voucher >= ao voucher automatico
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36793

    'pesquisa o proximo número de voucher disponível
    Do While lErro <> AD_SQL_SEM_DADOS
        
        If lDoc < lDoc1 Then
              
            'Se ainda não foi usado ==> este é o número a ser usado
             Exit Do
        
        Else
            'se o número for igual ==> passa para o proximo
            lDoc = lDoc + 1
        
            lDoc2 = lDoc1
        
            'pesquisa o proximo voucher
            Do While lDoc2 = lDoc1 And lErro <> AD_SQL_SEM_DADOS
        
                lErro = Comando_BuscarProximo(lComando1)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36794

            Loop
            
        End If
        
    Loop
    
    'atualiza o numero do voucher gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Configuracao SET DocAuto = ?", lComando, lDoc + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 11025
        
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 11026
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    DocAuto_Automatico = SUCESSO
        
    Exit Function
        
Erro_DocAuto_Automatico:

    DocAuto_Automatico = Err

    Select Case Err
    
        Case 11021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 11022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 11023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 11024, 11032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
        
        Case 11025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONFIGURACAO", Err)
        
        Case 11026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 11033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)
        
        Case 11034, 36793, 36794
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCAUTO", Err)
     
        Case 11035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150585)
        
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function
    
End Function

Function DocAuto_Exclui(ByVal lDoc As Long) As Long
    
Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_DocAuto_Exclui

    lTransacao = 0
   
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 11092
  
    'exclui os lançamentos deste voucher
    lErro = CF("DocAuto_Exclui_BD", lDoc)
    If lErro <> SUCESSO And lErro <> 20598 Then Error 11093
  
    If lErro = 20598 Then Error 20599
  
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 11094

    DocAuto_Exclui = SUCESSO

    Exit Function

Erro_DocAuto_Exclui:
    
    DocAuto_Exclui = Err
    
    Select Case Err
    
         
        Case 11092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
           
        Case 11093
          
        Case 11094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 20599
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCAUTO_NAO_CADASTRADO2", Err, lDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150586)
            
   End Select
   
   Call Transacao_Rollback
       
   Exit Function

End Function

Function DocAuto_Grava(ByVal lDoc As Long, ByVal sOrigem As String, ByVal sDescricao As String, ByVal colDocAutos As Collection) As Long
                                
Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_DocAuto_Grava

    lTransacao = 0
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 11095
    
    'exclui os lançamentos deste voucher
    lErro = CF("DocAuto_Exclui_BD", lDoc)
    If lErro <> SUCESSO And lErro <> 20598 Then Error 11096

    lErro = CF("DocAuto_Grava_BD", lDoc, sOrigem, sDescricao, colDocAutos)
    If lErro <> SUCESSO Then Error 11097
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 11098
   
    DocAuto_Grava = SUCESSO

    Exit Function

Erro_DocAuto_Grava:
    
    DocAuto_Grava = Err
    
    Select Case Err
    
        Case 11095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 11096, 11097

        Case 11098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150587)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Function DocAuto_Grava_BD(ByVal lDoc As Long, ByVal sOrigem As String, ByVal sDescricao As String, ByVal colDocAutos As Collection) As Long
'grava na tabela de Documentos Automaticos (DocAuto)lançamentos do voucher em questão

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim iTipoConta As Integer
Dim sCcl As String
Dim sConta As String
Dim iPossuiCcl As Integer
Dim iStatus As Integer
Dim objDocAuto As ClassDocAuto

On Error GoTo Erro_DocAuto_Grava_BD

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11099
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 11100
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 11101
    
    'Inicia Loop de GRAVACAO com Lançamentos do GRID
    For Each objDocAuto In colDocAutos
            
        'Pesquisa Conta associada a um Lançamento
        lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta, Ativo FROM PlanoConta WHERE Conta=?", iTipoConta, iStatus, objDocAuto.sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 11102
        
       'Lê a Conta
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11103
        
        If lErro = AD_SQL_SEM_DADOS Then Error 11104
        
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 11105
            
        'verifica se a conta é analítica
        If iTipoConta <> CONTA_ANALITICA Then Error 11106
        
        sCcl = String(STRING_CCL, 0)

        iPossuiCcl = 0
           
        'Se o lançamento envolver centro de custo extra contabil
        'Verificar se a associação conta x centro de custo está cadastrada
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL And Len(objDocAuto.sCcl) > 0 Then
        
            sCcl = objDocAuto.sCcl
            
            sConta = String(STRING_CONTA, 0)
            
            'Pesquisa Ccl associado ao Lançamento
            lErro = Comando_ExecutarLockado(lComando2, "SELECT Conta FROM ContaCcl WHERE Conta=? AND Ccl=?", sConta, objDocAuto.sConta, sCcl)
            If lErro <> AD_SQL_SUCESSO Then Error 11107
        
           'Lê a associação Conta x Ccl
            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11108
        
            If lErro = AD_SQL_SEM_DADOS Then Error 11109
        
            lErro = Comando_LockShared(lComando2)
            If lErro <> AD_SQL_SUCESSO Then Error 11110
            
        End If
        
        'Faz inserção do Lançamento na Tabela de Documentos Automáticos
        lErro = Comando_Executar(lComando1, "INSERT INTO DocAuto (Doc, Seq, Origem, Descricao, Conta,  Ccl,  Historico, Valor, SeqContraPartida ) VALUES (?,?,?,?,?,?,?,?,?)", lDoc, objDocAuto.iSeq, sOrigem, sDescricao, objDocAuto.sConta, objDocAuto.sCcl, objDocAuto.sHistorico, objDocAuto.dValor, objDocAuto.iSeqContraPartida)
        If lErro <> AD_SQL_SUCESSO Then Error 11111
        
        lErro = Comando_Unlock(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 11112
        
        If iPossuiCcl = 1 Then
            lErro = Comando_Unlock(lComando2)
            If lErro <> AD_SQL_SUCESSO Then Error 11113
        End If
           
    Next

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    DocAuto_Grava_BD = SUCESSO
    
    Exit Function
    
Erro_DocAuto_Grava_BD:

    DocAuto_Grava_BD = Err
    
    Select Case Err
    
        Case 11099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 11100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 11101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 11102, 11103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, objDocAuto.sConta)
    
        Case 11104
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, objDocAuto.sConta)

        Case 11105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, objDocAuto.sConta)
    
        Case 11106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, objDocAuto.sConta)
    
        Case 11107, 11108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, objDocAuto.sConta, sCcl)
    
        Case 11109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACCL_NAO_CADASTRADO", Err, objDocAuto.sConta, sCcl)
    
        Case 11110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTACCL", Err, objDocAuto.sConta, sCcl)
    
        Case 11111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_DOCAUTO", Err)
    
        Case 11112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_PLANOCONTA", Err, objDocAuto.sConta)
    
        Case 11113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_CONTACCL", Err, objDocAuto.sConta, sCcl)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150588)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function DocAuto_Exclui_BD(ByVal lDoc As Long)
'exclui da tabela DocAuto os Documentos Automáticos do voucher em questão

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iSeq As Integer
Dim tDocAuto As typeDocAuto

On Error GoTo Erro_DocAuto_Exclui_BD

    'Abre comando para pesquisar Lancamentos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11114
    
    'Abre comando para transação de exclusão
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 11115
    
    tDocAuto.lDoc = lDoc
    tDocAuto.sOrigem = String(STRING_ORIGEM, 0)
    tDocAuto.sConta = String(STRING_CONTA, 0)
    tDocAuto.sCcl = String(STRING_CCL, 0)
    tDocAuto.sHistorico = String(STRING_HISTORICO, 0)
    tDocAuto.sDescricao = String(STRING_DOCAUTO_DESCRICAO, 0)
   
    'Pesquisa a existencia dos Documentos no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT Seq, Origem, Descricao, Conta,  Ccl,  Historico, Valor FROM DocAuto WHERE Doc=?", 0, tDocAuto.iSeq, tDocAuto.sOrigem, tDocAuto.sDescricao, tDocAuto.sConta, tDocAuto.sCcl, tDocAuto.sHistorico, tDocAuto.dValor, tDocAuto.lDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 11116
        
    'Lê o primeiro Documento
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11117
    
    If lErro = AD_SQL_SEM_DADOS Then Error 20598
    
    'Inicia Loop de EXCLUSAO com Documentos do GRID
    Do While lErro <> AD_SQL_SEM_DADOS  'Refere-se ao BuscarProximo
        
        'Exclui os documentoa Automaticos da tabela DocAuto
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM DocAuto", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 11118
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11119
    
    Loop
   
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    DocAuto_Exclui_BD = SUCESSO
    
    Exit Function
    
Erro_DocAuto_Exclui_BD:

    DocAuto_Exclui_BD = Err
    
    Select Case Err
    
        Case 11114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 11115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 11116, 11117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCAUTO", Err)
            
        Case 11118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DOCAUTO", Err)
            
        Case 11119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCAUTO", Err)
            
        Case 20598
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150589)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
                   
End Function

Function PlanoConta_Insere(ByVal objPlanoConta As ClassPlanoConta) As Long
'insere uma conta no banco de dados
    
'??? comentar

Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iExercicio As Integer
Dim iDisponivel As Integer
Dim iNivelConta As Integer
Dim sCcl As String
Dim sConta As String
Dim dSldIni As Double
Dim dSldIni1 As Double
Dim objFiliais As AdmFiliais
Dim objPlanoContaHistorico As New ClassPlanoContaHistorico
Dim sContaPai As String

On Error GoTo Erro_PlanoConta_Insere

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 5494

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 5495

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 5496

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 5497
    
    'verifica se o saldo inicial pode ser alterado
    lErro = CF("Saldo_Inicial_Critica", iDisponivel)
    If lErro <> SUCESSO Then gError 5973
    
    'se o saldo não estiver disponivel ==> não pode colocar valor
    If iDisponivel = SLDINI_NAO_DISPONIVEL Then objPlanoConta.dSldIni = 0
    
    If objPlanoConta.iTipoConta <> CONTA_ANALITICA Then objPlanoConta.dSldIni = 0
        
    'Guarda os Valores no obj para futura Gravacao de Historico
    With objPlanoContaHistorico
    
        'Busca a ContaPai imediatamente acima da conta em questao
        lErro = Mascara_RetornaContaPai(objPlanoConta.sConta, sContaPai)
        If lErro <> SUCESSO Then gError 97753
        
        .sContaPai = sContaPai
        .dtDataAtualizacao = gdtDataHoje
        .iTipo = objPlanoConta.iTipoConta
        .sCodigo = objPlanoConta.sConta
        .sDescricao = objPlanoConta.sDescConta
                
    End With
    
    'FAZ A GRAVACAO DO HISTORICO
    lErro = PlanoContaHistorico_Grava(objPlanoContaHistorico)
    If lErro <> SUCESSO Then gError 97754
    
    'insere a conta na tabela PlanoConta
    lErro = Comando_Executar(lComando1, "INSERT INTO PlanoConta (Conta,NivelConta,DescConta,TipoConta,ContaSimples,Natureza,DigitoVerif,DigitoVerifSimples, HistPadrao, Ativo, UsaContaSimples, Categoria, CP, CR, TES, FAT, EST, COM, FIS, LJ, FluxoCaixa, NaturezaSped) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objPlanoConta.sConta, objPlanoConta.iNivelConta, objPlanoConta.sDescConta, objPlanoConta.iTipoConta, objPlanoConta.lContaSimples, objPlanoConta.iNatureza, objPlanoConta.iDigitoVerif, objPlanoConta.iDigitoVerifSimples, objPlanoConta.iHistPadrao, objPlanoConta.iAtivo, objPlanoConta.iUsaContaSimples, objPlanoConta.iCategoria, CInt(objPlanoConta.UsaModulo(MODULO_CONTASAPAGAR)), CInt(objPlanoConta.UsaModulo(MODULO_CONTASARECEBER)), CInt(objPlanoConta.UsaModulo(MODULO_TESOURARIA)), CInt(objPlanoConta.UsaModulo(MODULO_FATURAMENTO)), CInt(objPlanoConta.UsaModulo(MODULO_ESTOQUE)), CInt(objPlanoConta.UsaModulo(MODULO_COMPRAS)), CInt(objPlanoConta.UsaModulo(MODULO_LIVROSFISCAIS)), _
    CInt(objPlanoConta.UsaModulo(MODULO_LOJA)), objPlanoConta.iFluxoCaixa, objPlanoConta.iNaturezaSped)
    If lErro <> AD_SQL_SUCESSO Then gError 5498
        
    'insere os saldos iniciais de conta (zerados) para cada filial
    lErro = SaldoInicialConta_Insere(objPlanoConta.sConta)
    If lErro <> SUCESSO Then gError 10337
    
    'se o saldo inicial for diferente de zero alterar as contas "pai"
    If objPlanoConta.dSldIni <> 0 And gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        lErro = SaldoInicialConta_Altera(objPlanoConta)
        If lErro <> SUCESSO Then gError 9763
        
    End If
    
    dSldIni1 = objPlanoConta.dSldIni
    
    lErro = Comando_Executar(lComando2, "SELECT Exercicio FROM Exercicios ORDER BY Exercicio", iExercicio)
    If lErro <> AD_SQL_SUCESSO Then gError 5499
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9573
    
    Do While lErro = AD_SQL_SUCESSO
    
        For Each objFiliais In gcolFiliais
        
            If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then
    
                If objFiliais.iCodFilial = objPlanoConta.iFilialEmpresa Or objFiliais.iCodFilial = EMPRESA_TODA Then
                    dSldIni = dSldIni1
                Else
                    dSldIni = 0
                End If
                    
                'insere a conta na tabela MvPerCta
                lErro = Comando_Executar(lComando3, "INSERT INTO MvPerCta (FilialEmpresa, Exercicio, Conta, SldIni) VALUES (?,?,?,?)", objFiliais.iCodFilial, iExercicio, objPlanoConta.sConta, dSldIni)
                If lErro <> AD_SQL_SUCESSO Then gError 5500
                
            End If
            
        Next
        
        dSldIni1 = 0
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9574
    
    Loop
    
    'se a conta envolver um centro de custo contabil, cadastra a associacao da conta com o centro de custo
    lErro = ContaCcl_Insere_Conta(objPlanoConta)
    If lErro <> SUCESSO Then gError 9579
    
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 5501

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    PlanoConta_Insere = SUCESSO

    Exit Function

Erro_PlanoConta_Insere:
    
    PlanoConta_Insere = gErr
    
    Select Case gErr
    
        Case 5494, 5495, 5496
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 5497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 5498
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOCONTA", gErr, objPlanoConta.sConta)
        
        Case 5499, 9573, 9574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOS", gErr)
                    
        Case 5500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCTA", gErr, iExercicio, objPlanoConta.sConta)
            
        Case 5501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 5973, 9579, 9763, 10337, 97753, 97754
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150590)
        
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function ContaCcl_Insere_Conta(ByVal objPlanoConta As ClassPlanoConta) As Long
'insere uma associacao de conta com centro de custo a partir de uma conta com centro de custo contabil

Dim sCcl As String
Dim sConta As String
Dim lErro As Long
Dim iNivelConta As Integer
Dim objContaCcl As New ClassContaCcl

On Error GoTo Erro_ContaCcl_Insere_Conta

    'Se está trabalhando com centro de custo contábil
    If giSetupUsoCcl = CCL_USA_CONTABIL Then
    
        sCcl = String(STRING_CCL, 0)
        
        'verifica se a conta contem um centro de custo
        lErro = Mascara_RetornaCcl(objPlanoConta.sConta, sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 9577
        
        'se contem um centro de custo
        If Len(sCcl) > 0 Then
    
            'descobre o nivel da conta
            lErro = Mascara_Conta_ObterNivel(objPlanoConta.sConta, iNivelConta)
            If lErro <> SUCESSO Then Error 9572
    
            sConta = String(STRING_CONTA, 0)
    
            'pega o pai da conta em questao
            lErro = Mascara_RetornaContaNoNivel(iNivelConta - 1, objPlanoConta.sConta, sConta)
            If lErro <> SUCESSO Then Error 9576
    
            objContaCcl.iFilialEmpresa = objPlanoConta.iFilialEmpresa
            objContaCcl.sConta = sConta
            objContaCcl.sCcl = sCcl
            objContaCcl.dSldIni = objPlanoConta.dSldIni
            
            'inclui a associacao da conta com o centro de custo
            lErro = ContaCcl_Inclui(objContaCcl, objPlanoConta.sConta)
            If lErro <> SUCESSO Then Error 9578
            
        End If
    
    End If

    ContaCcl_Insere_Conta = SUCESSO

    Exit Function

Erro_ContaCcl_Insere_Conta:
    
    ContaCcl_Insere_Conta = Err
    
    Select Case Err
    
        Case 9572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, objPlanoConta.sConta)
            
        Case 9576
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, objPlanoConta.sConta, iNivelConta - 1)
            
        Case 9577
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCcl", Err, objPlanoConta.sConta)
            
        Case 9578
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150591)
        
    End Select
    
    Exit Function

End Function

'Function ContaCcl_Altera_Conta(objPlanoConta As ClassPlanoConta, iExercicio As Integer) As Long
''altera o saldo inicial do centro de custo/lucro se a conta tiver um centro de custo contábil
'
'Dim sCcl As String
'Dim sConta As String
'Dim lErro As Long
'Dim iNivelConta As Integer
'
'On Error GoTo Erro_ContaCcl_Altera_Conta
'
'    'Se está trabalhando com centro de custo contábil
'    If giSetupUsoCcl = CCL_USA_CONTABIL Then
'
'        sCcl = String(STRING_CCL, 0)
'
'        'verifica se a conta contem um centro de custo
'        lErro = Mascara_RetornaCcl(objPlanoConta.sConta, sCcl)
'        If lErro <> AD_SQL_SUCESSO Then gError 9974
'
'        'se contem um centro de custo
'        If Len(sCcl) > 0 Then
'
'            'descobre o nivel da conta
'            lErro = Mascara_Conta_ObterNivel(objPlanoConta.sConta, iNivelConta)
'            If lErro <> SUCESSO Then gError 9975
'
'            sConta = String(STRING_CONTA, 0)
'
'            'pega o pai da conta em questao
'            lErro = Mascara_RetornaContaNoNivel(iNivelConta - 1, objPlanoConta.sConta, sConta)
'            If lErro <> SUCESSO Then gError 9976
'
'            'inclui a associacao da conta com o centro de custo
'            lErro = ContaCcl_Inclui(sConta, sCcl, objPlanoConta.dSldIni, objPlanoConta.sConta)
'            If lErro <> SUCESSO Then gError 9578
'
'        End If
'
'    End If
'
'    ContaCcl_Altera_Conta = SUCESSO
'
'    Exit Function
'
'Erro_ContaCcl_Altera_Conta:
'
'    ContaCcl_Altera_Conta = gErr
'
'    Select Case gErr
'
'        Case 9974
'            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCcl", gErr, objPlanoConta.sConta)
'
'        Case 9975
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", gErr, objPlanoConta.sConta)
'
'        Case 9976
'            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", gErr, objPlanoConta.sConta, iNivelConta - 1)
'
'        Case 9578
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150592)
'
'    End Select
'
'    Exit Function
'
'End Function


Function PlanoConta_Altera(ByVal objPlanoConta As ClassPlanoConta) As Long
'altera uma conta no banco de dados
    
'??? comentar

Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iDisponivel As Integer
Dim iExercicio As Integer
Dim objContaCcl As New ClassContaCcl
Dim iAtivo As Integer
Dim dSldIni As Double
Dim objPlanoContaHistorico As New ClassPlanoContaHistorico
Dim sContaPai As String

On Error GoTo Erro_PlanoConta_Altera

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 5503

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 5504

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 9830

    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then gError 9831

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 5505
    
    'seleciona a conta a ser alterada
    lErro = Comando_ExecutarPos(lComando1, "SELECT Ativo FROM PlanoConta WHERE Conta = ?", 0, iAtivo, objPlanoConta.sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 5506
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 5596
    
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 10582
    
    'critica o tipo da conta
    lErro = CF("Conta_Critica_Tipo", objPlanoConta.sConta, objPlanoConta.iTipoConta)
    If lErro <> SUCESSO Then gError 10583
    
    iDisponivel = SLDINI_NAO_DISPONIVEL
    
    'se a conta for analitica
    If objPlanoConta.iTipoConta = CONTA_ANALITICA Then
    
        'verifica se o saldo inicial pode ser alterado
        lErro = CF("Saldo_Inicial_Critica", iDisponivel)
        If lErro <> SUCESSO Then gError 5974
        
    End If
    
    'Guarda os Valores no obj para futura Gravacao de Historico
    With objPlanoContaHistorico
    
        'Busca a ContaPai imediatamente acima da conta em questao
        lErro = Mascara_RetornaContaPai(objPlanoConta.sConta, sContaPai)
        If lErro <> SUCESSO Then gError 97755
        
        .sContaPai = sContaPai
        .dtDataAtualizacao = gdtDataHoje
        .iTipo = objPlanoConta.iTipoConta
        .sCodigo = objPlanoConta.sConta
        .sDescricao = objPlanoConta.sDescConta
                
    End With
    
    'FAZ A GRAVACAO DO HISTORICO
    lErro = PlanoContaHistorico_Grava(objPlanoContaHistorico)
    If lErro <> SUCESSO Then gError 97756
            
    'altera a conta na tabela PlanoConta
    lErro = Comando_ExecutarPos(lComando2, "UPDATE PlanoConta SET NivelConta=?,DescConta =?,TipoConta=?,ContaSimples=?,Natureza=?,DigitoVerif=?,DigitoVerifSimples=?, HistPadrao=?, Ativo=?, UsaContaSimples=?, Categoria=?, CP=?, CR=?, TES=?, FAT=?, EST=?, COM=?, FIS=?, LJ=?, FluxoCaixa=?, NaturezaSped = ?", lComando1, objPlanoConta.iNivelConta, objPlanoConta.sDescConta, objPlanoConta.iTipoConta, objPlanoConta.lContaSimples, objPlanoConta.iNatureza, objPlanoConta.iDigitoVerif, objPlanoConta.iDigitoVerifSimples, objPlanoConta.iHistPadrao, objPlanoConta.iAtivo, objPlanoConta.iUsaContaSimples, objPlanoConta.iCategoria, CInt(objPlanoConta.UsaModulo(MODULO_CONTASAPAGAR)), CInt(objPlanoConta.UsaModulo(MODULO_CONTASARECEBER)), CInt(objPlanoConta.UsaModulo(MODULO_TESOURARIA)), CInt(objPlanoConta.UsaModulo(MODULO_FATURAMENTO)), CInt(objPlanoConta.UsaModulo(MODULO_ESTOQUE)), CInt(objPlanoConta.UsaModulo(MODULO_COMPRAS)), CInt(objPlanoConta.UsaModulo(MODULO_LIVROSFISCAIS)), CInt(objPlanoConta.UsaModulo(MODULO_LOJA)), _
    objPlanoConta.iFluxoCaixa, objPlanoConta.iNaturezaSped)
    If lErro <> AD_SQL_SUCESSO Then gError 5507
        
    'se o saldo estiver disponivel ==> pode alterar seu valor
    If iDisponivel = SLDINI_DISPONIVEL And gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        'seleciona o saldo da conta a ser alterado
        lErro = Comando_Executar(lComando3, "SELECT SldIni FROM SaldoInicialConta WHERE FilialEmpresa=? AND Conta = ?", dSldIni, objPlanoConta.iFilialEmpresa, objPlanoConta.sConta)
        If lErro <> AD_SQL_SUCESSO Then gError 9832

        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO Then gError 9833

        objContaCcl.iFilialEmpresa = objPlanoConta.iFilialEmpresa
        objContaCcl.sConta = objPlanoConta.sConta
        objContaCcl.dSldIni = objPlanoConta.dSldIni
        
        'altera a associação da conta com centro de custo/lucro
        lErro = ContaCcl_Altera_Conta_Contabil(objContaCcl)
        If lErro <> SUCESSO Then gError 9984
        
        'Altera os saldos iniciais das contas pais de uma conta analitica. Se vIncluiConta estiver preenchido ==> inclui também a conta que está sendo passada como parametro. Altera os saldos da filial em questão e da empresa toda.
        lErro = SaldoInicialConta_Altera_Pais1(objPlanoConta.iFilialEmpresa, objPlanoConta.sConta, objPlanoConta.dSldIni - dSldIni, INCLUI_CONTA_SALDO_INICIAL)
        If lErro <> SUCESSO Then gError 9772
        
    End If
    
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 5508

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    PlanoConta_Altera = SUCESSO

    Exit Function

Erro_PlanoConta_Altera:
    
    PlanoConta_Altera = gErr
    
    Select Case gErr
    
        Case 5503, 5504, 9830, 9831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 5505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 5506, 5596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", gErr, objPlanoConta.sConta)
        
        Case 5507, 10047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PLANOCONTA", gErr, objPlanoConta.sConta)
         
        Case 5508
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 5974, 9772, 9984, 10583, 97755, 97756
        
        Case 9832, 9833
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA", gErr, objPlanoConta.iFilialEmpresa, objPlanoConta.sConta)
        
        Case 10582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", objPlanoConta.sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150593)
        
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    Exit Function

End Function

Function Conta_Exclui(ByVal sConta As String) As Long
'exclui a conta (rotina de mais alto nível)

Dim lErro As Long
Dim iTipoConta As Integer
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim iContaTemMovimento As Integer
Dim dSldIni As Double

On Error GoTo Erro_Conta_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 5529
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 5535
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 5552
    
    'seleciona a conta em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT TipoConta FROM PlanoConta WHERE Conta=?", 0, iTipoConta, sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 5530
    
    'le a conta
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 5532
    
    If lErro = AD_SQL_SEM_DADOS Then gError 5531
    
    'lock da conta
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 5534
    
    'testa se a conta possui movimento. Se tiver, não pode excluir
    lErro = CF("Conta_Critica_Possui_Movimento", sConta, iContaTemMovimento)
    If lErro <> SUCESSO Then gError 5645
    
    If iContaTemMovimento = 1 Then gError 5646
    
    '################################################################
    'Inserido por Wagner 26/05/2006
    lErro = CF("Exclusao_Valida", VALIDAEXCLUSAO_CODIGO_CONTACONTABIL, sConta)
    If lErro <> SUCESSO Then gError 177945
    '################################################################
    
    'se estiver excluindo uma conta analitica
    If iTipoConta = CONTA_ANALITICA Then
    
        'verifica se a conta possui lancamentos associados, rateios, etc.
        lErro = CF("ContaAnalitica_Critica_Exclusao", sConta)
        If lErro <> SUCESSO Then gError 5533

        If giSetupUsoCcl = CCL_USA_CONTABIL Then
        
            'exclui as associações de centro de custo com a conta em questão
            lErro = CF("ContaCcl_Exclui_Conta_Contabil", sConta)
            If lErro <> SUCESSO Then gError 5559
            
        ElseIf giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            'exclui as associações da conta com centros de custo extra contábil
            lErro = CF("ContaCcl_Exclui_Conta_ExtraContabil", sConta)
            If lErro <> SUCESSO Then gError 10319
            
        End If

        'exclui os documentos automaticos associados a conta
        lErro = CF("DocAuto_Exclui_Conta", sConta)
        If lErro <> SUCESSO Then gError 5588

        'exclui os saldos de conta associados a conta
        lErro = CF("MvPerCta_Exclui_Conta", sConta)
        If lErro <> SUCESSO Then gError 5651

        'exclui os saldos de ccl associados a conta
        lErro = CF("MvPerCcl_Exclui_Conta", sConta)
        If lErro <> SUCESSO Then gError 5652
        
        'exclui os orcamentos associados a conta
        lErro = Orcamento_Exclui_Conta(sConta)
        If lErro <> SUCESSO Then gError 16083
        
        'exclui os saldos iniciais de conta
        lErro = SaldoInicialConta_Exclui_Conta(sConta)
        If lErro <> SUCESSO Then gError 10297
        
        'exclui a conta
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM PlanoConta", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 5537
        
        
    Else
    
        'se estiver excluindo uma conta sintetica
        lErro = ContaSintetica_Exclui(sConta)
        If lErro <> SUCESSO Then gError 5551
        
    End If

    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 5553

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Conta_Exclui = SUCESSO
    
    Exit Function
    
Erro_Conta_Exclui:

    Conta_Exclui = gErr
    
    Select Case gErr
    
        Case 5529, 5535
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 5530, 5532
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", gErr, sConta)
    
        Case 5531
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", gErr, sConta)
            
        Case 5533
        
        Case 5534
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", gErr, sConta)
        
        Case 5537
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOCONTA", gErr, sConta)
            
        Case 5551, 177945
            
        Case 5552
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 5553
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 5559, 5588, 5645, 5651, 5652, 9774, 10297, 10319, 16083
        
        Case 5646
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTA_COM_MOVIMENTO1", gErr, sConta)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150594)
    
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function ContaSintetica_Exclui(ByVal sConta As String) As Long
'exclui a conta sintética passada como parametro

Dim lErro As Long
Dim iTipoConta As Integer
Dim lComando1 As Long
Dim sConta1 As String
Dim iNivel As Integer
Dim iNivel1 As Integer
Dim lErro1 As Integer

On Error GoTo Erro_ContaSintetica_Exclui

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5547
    
    sConta1 = String(STRING_CONTA, 0)
        
    'verifica se tem contas filhas
    lErro = Comando_ExecutarPos(lComando1, "SELECT Conta, TipoConta FROM PlanoConta WHERE Conta > ? ORDER BY Conta", 0, sConta1, iTipoConta, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5538
        
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10399
        
    If lErro = AD_SQL_SUCESSO Then
        
        lErro1 = Mascara_Conta_ObterNivel(sConta, iNivel)
        If lErro1 <> SUCESSO Then Error 5539
        
        lErro1 = Mascara_Conta_ObterNivel(sConta1, iNivel1)
        If lErro1 <> SUCESSO Then Error 5540
            
    End If
        
    'Para cada conta filha, se ela for analitica verifica se tem
    'lançamentos associados.
    Do While lErro = AD_SQL_SUCESSO And iNivel < iNivel1

        'lock da conta e suas filhas
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 5541

        'se a conta for analitica
        If iTipoConta = CONTA_ANALITICA Then

            'verifica se a conta possui lancamentos associados
            lErro = CF("ContaAnalitica_Critica_Exclusao", sConta1)
            If lErro <> SUCESSO Then Error 5542

        End If

        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10400

        If lErro = SUCESSO Then
            lErro1 = Mascara_Conta_ObterNivel(sConta1, iNivel1)
            If lErro1 <> SUCESSO Then Error 5948
        End If

    Loop
    
    'exclui a conta sintetica e suas filhas, se houverem.
    lErro = Conta_Exclui_BD(sConta)
    If lErro <> SUCESSO Then Error 5550
    
    Call Comando_Fechar(lComando1)

    ContaSintetica_Exclui = SUCESSO
    
    Exit Function
    
Erro_ContaSintetica_Exclui:

    ContaSintetica_Exclui = Err
    
    Select Case Err
    
        Case 5547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5538, 10399, 10400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTA", Err)
    
        Case 5539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)
    
        Case 5540, 5948
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta1)
    
        Case 5541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, sConta1)
    
        Case 5542, 5550
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150595)
    
    End Select
    
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function Conta_Exclui_BD(ByVal sConta As String) As Long
'exclui a conta e seus descententes

Dim lErro As Long
Dim lErro1 As Integer
Dim lComando2 As Long
Dim lComando3 As Long
Dim iNivel As Integer
Dim iNivel1 As Integer
Dim sConta1 As String
Dim iTipoConta As Integer

On Error GoTo Erro_Conta_Exclui_BD

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5548
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5549

    sConta1 = String(STRING_CONTA, 0)

    'pesquisa a conta e suas filhas visando sua exclusão
    lErro = Comando_ExecutarPos(lComando2, "SELECT Conta, TipoConta FROM PlanoConta WHERE Conta >= ? ORDER BY Conta", 0, sConta1, iTipoConta, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5543
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10401
     
    If lErro = AD_SQL_SUCESSO Then
        
        lErro1 = Mascara_Conta_ObterNivel(sConta, iNivel)
        If lErro1 <> SUCESSO Then Error 5544
        
        iNivel1 = iNivel + 1
            
    End If
        
    'exclui as contas da tabela PlanoConta
    Do While lErro = AD_SQL_SUCESSO And iNivel < iNivel1
        
        If iTipoConta = CONTA_ANALITICA Then
        
            If giSetupUsoCcl = CCL_USA_EXTRACONTABIL Then
            
                'exclui as associações de centro de custo com a conta em questão, se houverem
                lErro = CF("ContaCcl_Exclui_Conta_ExtraContabil", sConta1)
                If lErro <> SUCESSO Then Error 5558
    
    
            ElseIf giSetupUsoCcl = CCL_USA_CONTABIL Then
                
                'exclui a associação do centro de custo com a conta em questão, se houver
                lErro = CF("ContaCcl_Exclui_Conta_Contabil", sConta1)
                If lErro <> SUCESSO Then Error 9994
                
            End If
        
        End If
        
        'exclui os documentos automaticos associados a conta
        lErro = CF("DocAuto_Exclui_Conta", sConta1)
        If lErro <> SUCESSO Then Error 5587

        'exclui os saldos de conta associados a conta
        lErro = CF("MvPerCta_Exclui_Conta", sConta1)
        If lErro <> SUCESSO Then Error 5649

        'exclui os saldos de ccl associados a conta
        lErro = CF("MvPerCcl_Exclui_Conta", sConta1)
        If lErro <> SUCESSO Then Error 5650

        'exclui os orcamentos associados a conta
        lErro = Orcamento_Exclui_Conta(sConta1)
        If lErro <> SUCESSO Then Error 16084
        
        'exclui a conta
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM PlanoConta", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 5546
        
        'se for a conta sintetica que está sendo excluida ==> exclui o saldo inicial da conta sintetica das contas ancestrais
        If sConta = sConta1 Then
        
            lErro = SaldoInicialConta_Exclui_Conta(sConta1)
            If lErro <> SUCESSO Then Error 9773
            
        Else
        
            lErro = SaldoInicialConta_Exclui_Conta1(sConta1)
            If lErro <> SUCESSO Then Error 10326
        
        End If
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10402
        
        If lErro = AD_SQL_SUCESSO Then
            lErro1 = Mascara_Conta_ObterNivel(sConta1, iNivel1)
            If lErro1 <> SUCESSO Then Error 5640
        End If
        
            
    Loop

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Conta_Exclui_BD = SUCESSO
    
    Exit Function
    
Erro_Conta_Exclui_BD:

    Conta_Exclui_BD = Err
    
    Select Case Err
    
        Case 5543, 10401, 10402
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTA", Err)
    
        Case 5544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)
    
        Case 5546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOCONTA", Err, sConta1)
            
        Case 5548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5549
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5558, 5587, 5649, 5650, 9773, 9994, 16084
        
        Case 5640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta1)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150596)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function ContaCcl_Exclui_Conta_ExtraContabil(ByVal sConta As String) As Long
'exclui as associações da conta com centros de custo /m

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim sCcl As String
Dim objContaCcl As New ClassContaCcl

On Error GoTo Erro_ContaCcl_Exclui_Conta_ExtraContabil

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5554
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5555

    sCcl = String(STRING_CCL, 0)

    'pesquisa as associações da conta com centros de custo, visando sua exclusão
    lErro = Comando_ExecutarPos(lComando2, "SELECT Ccl FROM ContaCcl WHERE Conta = ?", 0, sCcl, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5556
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9911
     
    'exclui as associações com centros de custo
    Do While lErro = AD_SQL_SUCESSO
        
        objContaCcl.sCcl = sCcl
        objContaCcl.sConta = sConta
        
        lErro = SaldoInicialContaCcl_Exclui(objContaCcl)
        If lErro <> SUCESSO Then Error 10318
        
        'exclui a associação
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM ContaCcl", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 5557
            
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9912
        
    Loop

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    ContaCcl_Exclui_Conta_ExtraContabil = SUCESSO
    
    Exit Function
    
Erro_ContaCcl_Exclui_Conta_ExtraContabil:

    ContaCcl_Exclui_Conta_ExtraContabil = Err
    
    Select Case Err
    
        Case 5554, 5555
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5556, 9911, 9912
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL1", Err, sConta)
    
        Case 5557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACCL", Err, sConta, sCcl)
            
        Case 10318
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150597)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function PlanoConta_Exclui_Conta(ByVal sConta As String) As Long
'Exclui a conta de PlanoContas. /m
'O PROGRAMA DEVE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sConta1 As String

On Error GoTo Erro_PlanoConta_Exclui_Conta

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8384
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8385
    
    sConta1 = String(STRING_CONTA, 0)
    
    'Pesquisa a conta em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Conta FROM PlanoConta WHERE Conta=?", 0, sConta1, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 8386
    
    'Le a conta em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8388
        
    If lErro = AD_SQL_SEM_DADOS Then Error 8387
        
    'faz o lock da conta em questão
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9927
        
    'exclui a conta
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM PlanoConta", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 8389
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    PlanoConta_Exclui_Conta = SUCESSO
    
    Exit Function
    
Erro_PlanoConta_Exclui_Conta:

    PlanoConta_Exclui_Conta = Err
    
    Select Case Err
    
        Case 8384, 8385
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8386, 8388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, sConta)
    
        Case 8387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
    
        Case 8389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOCONTA", Err, sConta)
            
        Case 9927
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150598)
    
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function DocAuto_Exclui_Conta(ByVal sConta As String) As Long
'exclui os lançamentos de documento automático associados a conta em questão.

Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim lComando2 As Long
Dim lComando3 As Long
Dim lDoc As Long
Dim iSeq As Integer

On Error GoTo Erro_DocAuto_Exclui_Conta

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5560
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5561

    'pesquisa os documentos automáticos que envolvem a conta em questão.
    lErro = Comando_ExecutarPos(lComando2, "SELECT Doc, Seq FROM DocAuto WHERE Conta = ?", 0, lDoc, iSeq, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5562
        
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando2)
     
    'exclui os documentos que envolvem a conta questão.
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO
        
        'exclui o documento
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM DocAuto", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 5563
            
        iFim_de_Arquivo = Comando_BuscarProximo(lComando2)
            
    Loop

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    DocAuto_Exclui_Conta = SUCESSO
    
    Exit Function
    
Erro_DocAuto_Exclui_Conta:

    DocAuto_Exclui_Conta = Err
    
    Select Case Err
    
        Case 5560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCAUTO", Err)
    
        Case 5563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DOCAUTO", Err, lDoc, iSeq)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150599)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function ContaCcl_Exclui_Ccl(ByVal sCcl As String) As Long
'exclui as associações do centro de custo/lucro com contas /m

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim sConta As String

On Error GoTo Erro_ContaCcl_Exclui_Ccl

    lErro = ContaCcl_Exclui_Ccl1(sCcl)
    If lErro <> SUCESSO Then Error 10438

    lErro = SaldoInicialContaCcl_Exclui_Ccl(sCcl)
    If lErro <> SUCESSO Then Error 10387
                    
    'Exclui os saldos de Centro de custo/lucro (MvPerCcl) associados
    lErro = CF("MvPerCcl_Exclui_Ccl", sCcl)
    If lErro <> SUCESSO Then Error 8056
            
    ContaCcl_Exclui_Ccl = SUCESSO
    
    Exit Function
    
Erro_ContaCcl_Exclui_Ccl:

    ContaCcl_Exclui_Ccl = Err
    
    Select Case Err
    
        Case 8056, 10387, 10438
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150600)
    
    End Select
    
    Exit Function

End Function

Function ContaCcl_Exclui_Ccl1(ByVal sCcl As String) As Long
'exclui as associações do centro de custo/lucro com contas /m

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim sConta As String

On Error GoTo Erro_ContaCcl_Exclui_Ccl1

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5629
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5630

    'verifica se o centro de custo/lucro possui lancamentos associados
    lErro = CF("CclExtraContabil_Critica_Exclusao", sCcl)
    If lErro <> SUCESSO Then Error 10448
    
    'Exclui os documentos automaticos associados ao Centro de custo/lucro
    lErro = CF("DocAuto_Exclui_Ccl", sCcl)
    If lErro <> SUCESSO Then Error 8055
    
    'exclui os orcamentos associados ao centro de custo em questão
    lErro = Orcamento_Exclui_Ccl(sCcl)
    If lErro <> SUCESSO Then Error 10416

    
    sConta = String(STRING_CONTA, 0)

    'pesquisa as associações do ccl com contas
    lErro = Comando_ExecutarPos(lComando2, "SELECT Conta FROM ContaCcl WHERE Ccl = ?", 0, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 5631
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9873
     
    'exclui as associações com contas
    Do While lErro = AD_SQL_SUCESSO
        
        'exclui a associação
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM ContaCcl", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 5632
            
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9874
            
    Loop

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    ContaCcl_Exclui_Ccl1 = SUCESSO
    
    Exit Function
    
Erro_ContaCcl_Exclui_Ccl1:

    ContaCcl_Exclui_Ccl1 = Err
    
    Select Case Err
    
        Case 5629, 5630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5631, 9873, 9874
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL", Err)
    
        Case 5632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACCL", Err, sConta, sCcl)
        
        Case 8055, 10416, 10448
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150601)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function Ccl_Grava(ByVal objCcl As ClassCcl) As Long
'grava o centro de custo/lucro no banco de dados

Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iStatus As Integer
Dim sCcl As String


On Error GoTo Erro_Ccl_Grava

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5633

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5634

    sCcl = String(STRING_CCL, 0)

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 5635

    'pesquisa no BD o ccl em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Ccl FROM Ccl WHERE Ccl=?", 0, sCcl, objCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 5636

    'le o ccl, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)

    'encontrou o ccl ==> isto é uma alteração
    If lErro = AD_SQL_SUCESSO Then

        'atualiza o ccl
        lErro = Comando_ExecutarPos(lComando2, "UPDATE Ccl SET DescCcl=?", lComando1, objCcl.sDescCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 5637

    Else
        'não encontrou o ccl ==>  é uma inserção de ccl
        'insere o ccl
        lErro = Comando_Executar(lComando2, "INSERT INTO Ccl (Ccl,DescCcl) VALUES (?,?)", objCcl.sCcl, objCcl.sDescCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 5638

    End If

    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 5639

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando1)

    Ccl_Grava = SUCESSO

    Exit Function

Erro_Ccl_Grava:

    Ccl_Grava = Err

    Select Case Err

        Case 5633
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 5636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err, objCcl.sCcl)

        Case 5637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CCL", Err, objCcl.sCcl)

        Case 5638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CCL", Err, objCcl.sCcl)

        Case 5639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150602)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Ccl_Insere(ByVal objCcl As ClassCcl) As Long
'insere o centro de custo/lucro no banco de dados /m

Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim objContaCcl As New ClassContaCcl
Dim iExercicio As Integer
Dim objFiliais As AdmFiliais
Dim iCcl_Total As Integer
Dim objCclHistorico As New ClassCclHistorico

On Error GoTo Erro_Ccl_Insere

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 9597

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 10569

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 10570

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 9598

    'Se for Versão Light
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        'Lê o Número Total de Transportadoras
'        lErro = CF("Ccl_Le_Total", iCcl_Total)
'        If lErro <> SUCESSO Then gError 64084
'
'        'Se ultrapassou o número máximo de transportadoras ---> ERRO
'        If iCcl_Total >= LIMITE_CCL_VLIGHT Then gError 64085
'
'    End If

    'Alteracao Daniel
    'Atualiza os dados para futura gravacao de histórico
    With objCclHistorico

        .dtDataAtualizacao = gdtDataHoje
        .sCodigoCcl = objCcl.sCcl
        .iTipoCcl = objCcl.iTipoCcl
        .sDescCcl = objCcl.sDescCcl

    End With

    'FAZ A GRAVACAO DO HISTORICO
    lErro = CclHistorico_Grava(objCclHistorico)
    If lErro <> SUCESSO Then gError 97660
    
    'insere o ccl
    lErro = Comando_Executar(lComando1, "INSERT INTO Ccl (Ccl,DescCcl,TipoCcl,AtivoCcl) VALUES (?,?,?,?)", objCcl.sCcl, objCcl.sDescCcl, objCcl.iTipoCcl, objCcl.iAtivo)
    If lErro <> AD_SQL_SUCESSO Then gError 9599

    objContaCcl.sCcl = objCcl.sCcl
    objContaCcl.sConta = ""

    'insere os saldos iniciais de centro de custo (zerados) para cada filial
    lErro = SaldoInicialContaCcl_Insere(objContaCcl)
    If lErro <> SUCESSO Then gError 10568

    lErro = Comando_Executar(lComando2, "SELECT Exercicio FROM Exercicios ORDER BY Exercicio", iExercicio)
    If lErro <> AD_SQL_SUCESSO Then gError 10571
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 10572
    
    Do While lErro = AD_SQL_SUCESSO
    
        For Each objFiliais In gcolFiliais
    
            If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then
    
                'insere a conta na tabela MvPerCcl
                lErro = Comando_Executar(lComando3, "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Ccl, Conta, SldIni) VALUES (?,?,?,?,?)", objFiliais.iCodFilial, iExercicio, objCcl.sCcl, "", 0)
                If lErro <> AD_SQL_SUCESSO Then gError 10573
                
            End If
            
        Next
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 10574
    
    Loop

    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 9600

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Ccl_Insere = SUCESSO

    Exit Function

Erro_Ccl_Insere:

    Ccl_Insere = gErr

    Select Case gErr

        Case 9597, 10569, 10570
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 9598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 9599
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CCL", gErr, objCcl.sCcl)

        Case 9600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 10568, 64084, 97660

        Case 10571, 10572, 10574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", gErr)
        
        Case 10573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", gErr, objFiliais.iCodFilial, iExercicio, objCcl.sCcl, "")
        
        Case 64085
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_CCL_VLIGHT", gErr, LIMITE_CCL_VLIGHT)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150603)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Ccl_Altera(ByVal objCcl As ClassCcl) As Long
'altera o centro de custo/lucro no banco de dados /m

Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iStatus As Integer
Dim iTipoCcl As Integer
Dim iTemFilho As Integer
Dim objCclHistorico As New ClassCclHistorico

On Error GoTo Erro_Ccl_Altera

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 9601

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 9602

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 9603

    'pesquisa no BD o ccl em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT TipoCcl FROM Ccl WHERE Ccl=?", 0, iTipoCcl, objCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 9604

    'le o ccl, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 9605

    'lock o ccl
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 9606

    If objCcl.iTipoCcl <> iTipoCcl Then

        'critica o tipo do centro de custo/lucro
        lErro = CF("Ccl_Critica_Tipo", objCcl)
        If lErro <> SUCESSO Then gError 10579
    
    End If
    
    'Alteracao Daniel
    'Atualiza os dados para futura gravacao de histórico
    With objCclHistorico
    
        .dtDataAtualizacao = gdtDataHoje
        .sCodigoCcl = objCcl.sCcl
        .iTipoCcl = objCcl.iTipoCcl
        .sDescCcl = objCcl.sDescCcl
        
    End With
    
    'FAZ A GRAVACAO DO HISTORICO
    lErro = CclHistorico_Grava(objCclHistorico)
    If lErro <> SUCESSO Then gError 97628

    'atualiza o ccl
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Ccl SET DescCcl=?, TipoCcl=?, AtivoCcl=?", lComando1, objCcl.sDescCcl, objCcl.iTipoCcl, objCcl.iAtivo)
    If lErro <> AD_SQL_SUCESSO Then gError 9607

    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 9608

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando1)

    Ccl_Altera = SUCESSO

    Exit Function

Erro_Ccl_Altera:

    Ccl_Altera = gErr

    Select Case gErr

        Case 9601, 9602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 9603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 9604, 9605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", gErr, objCcl.sCcl)
            
        Case 9606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", gErr, objCcl.sCcl)

        Case 9607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CCL", gErr, objCcl.sCcl)

        Case 9608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 10579, 97628

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150604)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function MvPerCta_Exclui_Conta(ByVal sConta As String) As Long
'exclui os saldos de conta associados a conta em questão /m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim lComando2 As Long
Dim lComando3 As Long
Dim iExercicio As Integer

On Error GoTo Erro_MvPerCta_Exclui_Conta

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5641
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5642

    'pesquisa os saldos de conta para a conta em questão
    lErro = Comando_ExecutarPos(lComando2, "SELECT Exercicio FROM MvPerCta WHERE Conta = ?", 0, iExercicio, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5643
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9870
     
    'exclui as associações com centros de custo
    Do While lErro = AD_SQL_SUCESSO
        
        'exclui o saldo
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM MvPerCta", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 5644
            
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9871
            
    Loop

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    MvPerCta_Exclui_Conta = SUCESSO
    
    Exit Function
    
Erro_MvPerCta_Exclui_Conta:

    MvPerCta_Exclui_Conta = Err
    
    Select Case Err
    
        Case 5641, 5642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5643, 9870, 9871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA2", Err, sConta)
    
        Case 5644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MVPERCTA", Err, iExercicio, sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150605)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function


Function MvPerCcl_Exclui_Conta(ByVal sConta As String) As Long
'exclui os saldos de ccl associados a conta em questão
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim sCcl As String
Dim iExercicio As Integer

On Error GoTo Erro_MvPerCcl_Exclui_Conta

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5645
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5646

    sCcl = String(STRING_CCL, 0)

    'pesquisa os saldos de ccl para a conta em questão
    lErro = Comando_ExecutarPos(lComando2, "SELECT Exercicio, Ccl FROM MvPerCcl WHERE Conta = ?", 0, iExercicio, sCcl, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5647
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10172
     
    'exclui as associações com centros de custo
    Do While lErro = AD_SQL_SUCESSO
        
        'exclui o saldo
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM MvPerCcl", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 5648
            
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10173
            
    Loop

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    MvPerCcl_Exclui_Conta = SUCESSO
    
    Exit Function
    
Erro_MvPerCcl_Exclui_Conta:

    MvPerCcl_Exclui_Conta = Err
    
    Select Case Err
    
        Case 5645, 5646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 5647, 10172, 10173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL2", Err, sConta)
    
        Case 5648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MVPERCCL", Err, iExercicio, sCcl, sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150606)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function Ccl_Exclui(ByVal sCcl As String) As Long
'exclui o Centro de custo/lucro /m

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim iTipoCcl As Integer
Dim iCclTemMovimento As Integer

On Error GoTo Erro_Ccl_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 8044
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 8045
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 8046
    
    'pesquisa o Centro de custo/lucro em questão no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT TipoCcl FROM Ccl WHERE Ccl=?", 0, iTipoCcl, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 8047
    
    'le o centro de custo/lucro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 8049
        
    If lErro = AD_SQL_SEM_DADOS Then gError 8048
    
    'lock do Centro de custo/lucro
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 8050
    
    'testa se o centro de custo/lucro possui movimento. Se tiver, não pode excluir
    lErro = CF("Ccl_Critica_Possui_Movimento", sCcl, iCclTemMovimento)
    If lErro <> SUCESSO Then gError 10377
    
    If iCclTemMovimento = 1 Then gError 10378
    
    'se estiver excluindo um centro de custo analitico
    If iTipoCcl = CCL_ANALITICA Then
    
    
        'verifica se Ccl tem vinculos com outras tabelas antes da sua exclusão
        lErro = Ccl_Critica_Vinculos_Compras(sCcl)
        If lErro <> SUCESSO Then gError 85240

        'Verifica se o Centro de custo/lucro é contabil
        If giSetupUsoCcl = CCL_USA_CONTABIL Then
        
            'Exclui o Centro de custo/lucro Contabil
            lErro = CF("Ccl_Exclui_Conta", sCcl)
            If lErro <> SUCESSO Then gError 8058
            
        Else
        
            'Exclui o Centro de custo/lucro Extra-Contabil
            lErro = ContaCcl_Exclui_Ccl(sCcl)
            If lErro <> SUCESSO Then gError 8129
                
        End If
                    
        
        'exclui o Centro do custo/lucro do BD
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM Ccl", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 8059
        
    Else
    
        'se estiver excluindo um centro de custo sintetico
        lErro = Ccl_Exclui_BD(sCcl)
        If lErro <> SUCESSO Then gError 10440
    
    End If
    
        
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 8060

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Ccl_Exclui = SUCESSO
    
    Exit Function
    
Erro_Ccl_Exclui:

    Ccl_Exclui = gErr
    
    Select Case gErr
    
        Case 8044, 8045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 8046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 8047, 8049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", gErr, sCcl)
            
        Case 8048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, sCcl)
        
        Case 8050
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", gErr, sCcl)
        
        Case 8058, 10377, 10388, 10440, 85240
        
        Case 8059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CCL", gErr, sCcl)
        
        Case 8060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 8129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACCL", gErr, sCcl)
           
        Case 10378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CCL_COM_MOVIMENTO", gErr, sCcl)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150607)
    
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function



Function Ccl_Exclui_BD(ByVal sCcl As String) As Long
'exclui o centro de custo/lucro e seus descententes

Dim lErro As Long
Dim lErro1 As Integer
Dim lComando2 As Long
Dim lComando3 As Long
Dim iNivel As Integer
Dim iNivel1 As Integer
Dim sCcl1 As String
Dim iTipoCcl As Integer

On Error GoTo Erro_Ccl_Exclui_BD

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10403
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10404

    sCcl1 = String(STRING_CCL, 0)

    'pesquisa o centro de custo/lucro e seus filhos visando sua exclusão
    lErro = Comando_ExecutarPos(lComando2, "SELECT Ccl, TipoCcl FROM Ccl WHERE Ccl >= ? ORDER BY Ccl", 0, sCcl1, iTipoCcl, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 10405
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10406
     
    If lErro = AD_SQL_SUCESSO Then
        
        lErro1 = Mascara_Ccl_ObterNivel(sCcl, iNivel)
        If lErro1 <> SUCESSO Then Error 10407
        
        iNivel1 = iNivel + 1
            
    End If
        
    'exclui os centros de custo
    Do While lErro = AD_SQL_SUCESSO And iNivel < iNivel1
        
        'lock do centro de custo e seus filhos
        lErro = Comando_LockExclusive(lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 10439
     
        If iTipoCcl = CONTA_ANALITICA Then
        
            'Verifica se o Centro de custo/lucro é contabil
            If giSetupUsoCcl = CCL_USA_CONTABIL Then
            
                'Exclui o Centro de custo/lucro Contabil
                lErro = Ccl_Exclui_Conta1(sCcl1)
                If lErro <> SUCESSO Then Error 10449
                
            Else
            
                'Exclui o Centro de custo/lucro Extra-Contabil
                lErro = ContaCcl_Exclui_Ccl1(sCcl1)
                If lErro <> SUCESSO Then Error 10450
                    
            End If
        
        End If
        
        'exclui os saldos de ccl associados ao centro de custo/lucro em questão
        lErro = CF("MvPerCcl_Exclui_Ccl", sCcl1)
        If lErro <> SUCESSO Then Error 10451

        'exclui os orcamentos associados ao centro de custo/lucro
        lErro = Orcamento_Exclui_Ccl(sCcl1)
        If lErro <> SUCESSO Then Error 10452
        
        'exclui o centro de custo/lucro
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM Ccl", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 10453
        
        If sCcl = sCcl1 Then
        
            lErro = SaldoInicialContaCcl_Exclui_Ccl(sCcl1)
            If lErro <> SUCESSO Then Error 10454
            
        Else
        
            lErro = SaldoInicialContaCcl_Exclui_Ccl1(sCcl1)
            If lErro <> SUCESSO Then Error 10455
            
        End If
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10429
        
        If lErro = AD_SQL_SUCESSO Then
            lErro1 = Mascara_Ccl_ObterNivel(sCcl1, iNivel1)
            If lErro1 <> SUCESSO Then Error 10456
        End If
        
    Loop

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Ccl_Exclui_BD = SUCESSO
    
    Exit Function
    
Erro_Ccl_Exclui_BD:

    Ccl_Exclui_BD = Err
    
    Select Case Err
    
        Case 10403, 10404
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10405, 10406, 10429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL1", Err)
    
        Case 10407
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CCL_OBTERNIVEL", Err, sCcl)
        
        Case 10439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", Err, sCcl1)
        
        Case 10453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CCL", Err, sCcl1)
            
        Case 10456
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CCL_OBTERNIVEL", Err, sCcl1)
    
        Case 10449, 10450, 10451, 10452, 10454, 10455
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150608)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function


Function DocAuto_Exclui_Ccl(ByVal sCcl As String) As Long
'exclui os lançamentos de documento automático associados o Ccl em questão./m

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lDoc As Long
Dim iSeq As Integer

On Error GoTo Erro_DocAuto_Exclui_Ccl

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 8061
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 8062

    'pesquisa os documentos automáticos que envolvem o Ccl em questão.
    lErro = Comando_ExecutarPos(lComando2, "SELECT Doc, Seq FROM DocAuto WHERE Ccl = ?", 0, lDoc, iSeq, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8063
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9864
     
    'exclui os documentos que envolvem o Ccl em questão.
    Do While lErro = AD_SQL_SUCESSO
        
        'exclui o documento
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM DocAuto", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 8064
            
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9865
            
    Loop

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    DocAuto_Exclui_Ccl = SUCESSO
    
    Exit Function
    
Erro_DocAuto_Exclui_Ccl:

    DocAuto_Exclui_Ccl = Err
    
    Select Case Err
    
        Case 8061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8063, 9864, 9865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCAUTO", Err)
    
        Case 8064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DOCAUTO", Err, lDoc, iSeq)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150609)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function MvPerCcl_Exclui_Ccl(ByVal sCcl As String) As Long
'exclui os saldos de ccl associados ao centro de custo em questao/m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim iExercicio As Integer
Dim sConta As String

On Error GoTo Erro_MvPerCcl_Exclui_Ccl

    sConta = String(STRING_CONTA, 0)
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 8065
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 8066

    'pesquisa os saldos para o Ccl em questão
    lErro = Comando_ExecutarPos(lComando2, "SELECT Exercicio, Conta FROM MvPerCcl WHERE Ccl = ?", 0, iExercicio, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8067
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9866
     
    'exclui as associações com centros de custo
    Do While lErro = AD_SQL_SUCESSO
        
        'exclui o saldo
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM MvPerCcl", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 8068
            
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9867
            
    Loop

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    MvPerCcl_Exclui_Ccl = SUCESSO
    
    Exit Function
    
Erro_MvPerCcl_Exclui_Ccl:

    MvPerCcl_Exclui_Ccl = Err
    
    Select Case Err
    
        Case 8065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8067, 9866, 9867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL3", Err, sCcl)
    
        Case 8068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MVPERCCL", Err, iExercicio, sCcl, sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150610)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function Ccl_Exclui_Conta(ByVal sCcl As String) As Long
'Exclui as contas associadas ao centro de custo e os respectivos relacionamentos. /m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lComando5 As Long
Dim sConta As String
Dim sConta1 As String

On Error GoTo Erro_Ccl_Exclui_Conta

    'Exclui as contas associadas ao centro de custo
    lErro = Ccl_Exclui_Conta1(sCcl)
    If lErro <> SUCESSO Then Error 10431

    'Exclui os saldos de Centro de custo/lucro (MvPerCcl) associados
    lErro = CF("MvPerCcl_Exclui_Ccl", sCcl)
    If lErro <> SUCESSO Then Error 10419
            
    'exclui os saldos iniciais associados ao centro de custo em questão
    lErro = SaldoInicialContaCcl_Exclui_Ccl(sCcl)
    If lErro <> SUCESSO Then Error 10420

    Ccl_Exclui_Conta = SUCESSO
    
    Exit Function
    
Erro_Ccl_Exclui_Conta:

    Ccl_Exclui_Conta = Err
    
    Select Case Err
    
        Case 10419, 10420, 10431
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150611)
    
    End Select
    
    Exit Function

End Function

Function Ccl_Exclui_Conta1(ByVal sCcl As String) As Long
'Exclui as contas associadas ao centro de custo e seus relacionamentos  /m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lComando5 As Long
Dim sConta As String
Dim sConta1 As String

On Error GoTo Erro_Ccl_Exclui_Conta1

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 8069
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 8070

    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 8093

    lComando5 = Comando_Abrir()
    If lComando5 = 0 Then Error 8094

    sConta = String(STRING_CONTA, 0)

    'pesquisa as associações da conta com centros de custo
    lErro = Comando_ExecutarPos(lComando2, "SELECT Conta FROM ContaCcl WHERE Ccl = ?", 0, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8071
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9868
        
    'exclui as referidas contas e associacoes
    Do While lErro = AD_SQL_SUCESSO
                        
        lErro = Comando_LockExclusive(lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 10430
                        
        'Verifica se existe lançamentos pendentes e rateios para esta conta. Se existem, abortar.
        lErro = CF("ContaAnalitica_Critica_Exclusao", sConta)
        If lErro <> SUCESSO Then Error 10417
                        
        'exclui o saldo inicial associado a conta em questão
        lErro = SaldoInicialConta_Exclui_Conta(sConta)
        If lErro <> SUCESSO Then Error 10379
                        
        'exclui os saldos associado a conta em questão
        lErro = CF("MvPerCta_Exclui_Conta", sConta)
        If lErro <> SUCESSO Then Error 9872
                        
        'exclui os orçamentos associados a conta em questão
        lErro = Orcamento_Exclui_Conta(sConta)
        If lErro <> SUCESSO Then Error 10415
                        
        'Exclui os documentos automaticos associados ao Centro de custo/lucro
        lErro = CF("DocAuto_Exclui_Conta", sConta)
        If lErro <> SUCESSO Then Error 10418
                        
        sConta1 = String(STRING_CONTA, 0)
                        
        'posiciona conta a ser excluida
        lErro = Comando_ExecutarPos(lComando3, "SELECT Conta FROM PlanoConta WHERE Conta = ?", 0, sConta1, sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 8091
           
        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 8095
                        
        'exclui a conta
        lErro = Comando_ExecutarPos(lComando4, "DELETE FROM PlanoConta", lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 8072
        
        'exclui associacao Ccl/Conta
        lErro = Comando_ExecutarPos(lComando5, "DELETE FROM ContaCcl", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 8092
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9869
            
    Loop
                    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    Call Comando_Fechar(lComando5)

    Ccl_Exclui_Conta1 = SUCESSO
    
    Exit Function
    
Erro_Ccl_Exclui_Conta1:

    Ccl_Exclui_Conta1 = Err
    
    Select Case Err
    
        Case 8069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 8071, 9868, 9869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL", Err)
            
        Case 8072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTA", Err, sConta)
                
        Case 8091, 8095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, sConta)
            
        Case 8092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACCL", Err, sConta, sCcl)
                
        Case 8093
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 9872, 10379, 10415, 10417, 10418
        
        Case 10430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTACCL", Err, sConta, sCcl)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150612)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    Call Comando_Fechar(lComando5)
    
    Exit Function

End Function


Function HistPadrao_Exclui(ByVal iHistPadrao As Integer) As Long
'exclui o Histórico Padrão

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim iHistPadrao1 As Integer

On Error GoTo Erro_HistPadrao_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6152
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 6153
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 6154
    
    'pesquisa o HistPadrao em questão no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT HistPadrao FROM HistPadrao WHERE HistPadrao=?", 0, iHistPadrao1, iHistPadrao)
    If lErro <> AD_SQL_SUCESSO Then Error 6155
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6157
        
    If lErro = AD_SQL_SEM_DADOS Then Error 6156
    
    'lock do histórico padrão
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 6158

    'Verifica se o Histórico Padrão é usado no Plano de Contas
    lErro = CF("PlanoContas_Le_HistPadrao", iHistPadrao)
    If lErro <> SUCESSO And lErro <> 6165 Then Error 6162
    
    'Existe o Histórico Padrão no Plano de Contas
    If lErro = SUCESSO Then Error 6161

    'exclui o histórico padrão do BD
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM HistPadrao", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 6159
    
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 6160

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    HistPadrao_Exclui = SUCESSO
    
    Exit Function
    
Erro_HistPadrao_Exclui:

    HistPadrao_Exclui = Err
    
    Select Case Err
    
        Case 6152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 6153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 6154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 6155, 6157
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_HISTPADRAO", Err, iHistPadrao)
            
        Case 6156
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_NAO_CADASTRADO", Err, iHistPadrao)
        
        Case 6158
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_HISTPADRAO", Err, iHistPadrao)
        
        Case 6159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_HISTPADRAO", Err, iHistPadrao)
        
        Case 6160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case 6161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTPADRAO_PRESENTE_PLANO_CONTAS", Err, iHistPadrao)
            
        Case 6162
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150613)
    
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function HistPadrao_Grava(ByVal objHistPadrao As ClassHistPadrao, iOperacao As Integer) As Long
'grava o Histórico Padrão no banco de dados

Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iStatus As Integer
Dim iHistPadrao As Integer

On Error GoTo Erro_HistPadrao_Grava

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 6144

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 6145

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 6146

    'pesquisa no BD o HistPadrao em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT HistPadrao FROM HistPadrao WHERE HistPadrao=?", 0, iHistPadrao, objHistPadrao.iHistPadrao)
    If lErro <> AD_SQL_SUCESSO Then Error 6147

    'le o HistPadrao, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6148

    'encontrou o HistPadrao ==> isto é uma alteração
    If lErro = AD_SQL_SUCESSO Then
    
        iOperacao = MODIFICACAO
    
        'atualiza o HistPadrao
        lErro = Comando_ExecutarPos(lComando2, "UPDATE HistPadrao SET DescHistPadrao=?", lComando1, objHistPadrao.sDescHistPadrao)
        If lErro <> AD_SQL_SUCESSO Then Error 6149

    Else
        'não encontrou o HistPadrao ==>  é uma inserção de HistPadrao
        
        iOperacao = GRAVACAO
        
        'insere o HistPadrao
        lErro = Comando_Executar(lComando2, "INSERT INTO HistPadrao (HistPadrao,DescHistPadrao) VALUES (?,?)", objHistPadrao.iHistPadrao, objHistPadrao.sDescHistPadrao)
        
        If lErro <> AD_SQL_SUCESSO Then Error 6150

    End If

    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 6151

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando1)

    HistPadrao_Grava = SUCESSO

    Exit Function

Erro_HistPadrao_Grava:

    HistPadrao_Grava = Err

    Select Case Err

        Case 6144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 6147, 6148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_HISTPADRAO", Err, objHistPadrao.iHistPadrao)

        Case 6149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_HISTPADRAO", Err, objHistPadrao.iHistPadrao)

        Case 6150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_HISTPADRAO", Err, objHistPadrao.iHistPadrao)

        Case 6151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150614)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function HistPadrao_Automatico(iHistPadrao As Integer) As Long
'retorna o número do proximo HistPadrao disponivel

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim iHistPadrao1 As Integer
Dim lTransacao As Long

On Error GoTo Erro_HistPadrao_Automatico

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6127
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 6128
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 6129
   
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 6130

    'Pesquisa o número do HistPadrao gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT HistPadrao FROM Configuracao", 0, iHistPadrao)
    If lErro <> AD_SQL_SUCESSO Then Error 6131

    'le o registro com o número do HistPadrao automático
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 6132
    
    'loca o registro que guarda o numero do HistPadrao gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 6134

    'verifica se este número ainda não foi utilizado
    'Pesquisa as numerações de HistPadrao maiores ou iguais a numeração gerada automaticamente
    lErro = Comando_Executar(lComando1, "SELECT HistPadrao FROM HistPadrao WHERE HistPadrao >= ? ORDER BY HistPadrao", iHistPadrao1, iHistPadrao)
    If lErro <> AD_SQL_SUCESSO Then Error 6133
        
    'le o primeiro HistPadrao >= ao HistPadrao automatico
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9715

    'pesquisa o proximo número de HistPadrao disponível
    Do While lErro <> AD_SQL_SEM_DADOS
        
        If iHistPadrao < iHistPadrao1 Then Exit Do
        
        If iHistPadrao = iHistPadrao1 Then
            iHistPadrao = iHistPadrao + 1
        End If
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9716
        
    Loop

    'atualiza o numero do HistPadrao gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Configuracao SET HistPadrao = ?", lComando, iHistPadrao + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 6135
        
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 6136
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    HistPadrao_Automatico = SUCESSO
        
    Exit Function
    
    
Erro_HistPadrao_Automatico:

    HistPadrao_Automatico = Err

    Select Case Err
    
        Case 6127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 6128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 6129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 6131, 6132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
            
        Case 6133, 9715, 9716
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_HISTPADRAO1", Err)
            
        Case 6130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 6134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)
            
        Case 6135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONFIGURACAO", Err)
            
        Case 6136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150615)
        
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function
    
End Function

Public Function ContaCcl_Inclui_Associacao_ExtraCont(ByVal sConta As String, ByVal sCcl As String) As Long
'Inclui no BD as associacoes de sConta com cada Ccl marcado /m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim sConta1 As String
Dim iExercicio As Integer
Dim alComando(1 To 6) As Long
Dim iNivelConta As Integer
Dim sContaFinal As String
Dim iFilialEmpresa As Integer
Dim iTipoCcl As Integer
Dim objContaCcl As New ClassContaCcl
Dim iIndice As Integer, dSldIni As Double

On Error GoTo Erro_ContaCcl_Inclui_Associacao_ExtraCont

    For iIndice = LBound(alComando) To UBound(alComando)
        
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 8204
        
    Next
        
    sConta1 = String(STRING_CONTA, 0)
        
    'pesquisa a associacao em questão no BD
    lErro = Comando_Executar(alComando(1), "SELECT Conta FROM ContaCcl WHERE Conta=? AND Ccl=?", sConta1, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8140
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8141
        
    'Se a associação conta x centro de custo/lucro ainda não estiver cadastrada
    If lErro = AD_SQL_SEM_DADOS Then
        
        'pesquisa o centro de custo em questão no BD
        lErro = Comando_Executar(alComando(3), "SELECT TipoCcl FROM Ccl WHERE Ccl=?", iTipoCcl, sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 10483
    
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10484
        
        'verifica se o centro de custo é analitico
        If iTipoCcl <> CCL_ANALITICA Then Error 10485
        
        'Incluir associacao no BD
        lErro = Comando_Executar(alComando(2), "INSERT INTO ContaCcl (Conta, Ccl) VALUES (?,?)", sConta, sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 8143
        
        objContaCcl.sConta = sConta
        objContaCcl.sCcl = sCcl
        
        'insere os saldos iniciais de centro de custo (zerados) para cada filial
        lErro = SaldoInicialContaCcl_Insere(objContaCcl)
        If lErro <> SUCESSO Then Error 10584
        
        'incluir a assoc em MvPerCcl para os exercicios existentes
        'Seleciona exercicios
        lErro = Comando_Executar(alComando(4), "SELECT FilialEmpresa, Exercicio  FROM ExerciciosFilial", iFilialEmpresa, iExercicio)
        If lErro <> AD_SQL_SUCESSO Then Error 8148
        
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9371
               
        'loop dos exercicios
        Do While lErro = AD_SQL_SUCESSO
                            
'''''            '??? este teste nao deveria existir mas foi colocado p/contornar problema com bd da INPAL em 05/11/03
'''''            lErro = Comando_Executar(alComando(6), "SELECT SldIni FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ? AND Ccl = ?", dSldIni, iFilialEmpresa, iExercicio, sConta, sCcl)
'''''            If lErro <> AD_SQL_SUCESSO Then Error 8149
'''''
'''''            lErro = Comando_BuscarProximo(alComando(6))
'''''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8149
'''''
'''''            If lErro = AD_SQL_SEM_DADOS Then
'''''
'''''                lErro = Comando_Executar(alComando(5), "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Conta, Ccl) VALUES (?,?,?,?)", iFilialEmpresa, iExercicio, sConta, sCcl)
'''''                If lErro <> AD_SQL_SUCESSO Then Error 8149
'''''
'''''            End If
            
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9372
                       
        Loop
      
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ContaCcl_Inclui_Associacao_ExtraCont = SUCESSO
    
    Exit Function

Erro_ContaCcl_Inclui_Associacao_ExtraCont:

    ContaCcl_Inclui_Associacao_ExtraCont = Err
    
    Select Case Err
            
        Case 8140, 8141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, sConta, sCcl)
               
        Case 8204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                                    
        Case 8143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONTACCL", Err, sConta, sCcl)
            
        Case 8149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", Err, iFilialEmpresa, iExercicio, sConta, sCcl)
            
        Case 8148, 9371, 9372
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", Err)
            
        Case 10483, 10484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err, sCcl)
        
        Case 10485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA", Err, sCcl)
        
        Case 10584
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 150616)
        
    End Select
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function ContaCcl_Inclui_Associacao_Contabil(tPlanoConta As typePlanoConta, ByVal sCcl As String) As Long
'Inclui no BD as associacoes de sConta com cada Ccl marcado /m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim sConta1 As String
Dim sCcl1 As String
Dim sConta3 As String
Dim iExercicio As Integer
Dim alComando(1 To 7) As Long
Dim iNivelConta As Integer
Dim sContaFinal As String
Dim iIndice As Integer
Dim iFilialEmpresa As Integer
Dim iTipoCcl As Integer
Dim objContaCcl As New ClassContaCcl

On Error GoTo Erro_ContaCcl_Inclui_Associacao_Contabil

    'abre os comandos de acesso ao bd
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 8369
    Next
    
    sConta3 = String(STRING_CONTA, 0)
    sCcl1 = String(STRING_CCL, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT Ccl FROM ContaCcl WHERE Conta=? AND Ccl=?", sCcl1, tPlanoConta.sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8375
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8376
        
    If lErro = AD_SQL_SEM_DADOS Then
        
        'pesquisa o centro de custo em questão no BD
        lErro = Comando_Executar(alComando(7), "SELECT TipoCcl FROM Ccl WHERE Ccl=?", iTipoCcl, sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 10486
    
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10487
        
        'verifica se o centro de custo é analitico
        If iTipoCcl <> CCL_ANALITICA Then Error 10488
        
        sContaFinal = left(tPlanoConta.sConta, Len(tPlanoConta.sConta) - Len(sCcl)) & sCcl
        
        lErro = Mascara_Conta_ObterNivel(sContaFinal, iNivelConta)
        If lErro <> SUCESSO Then Error 8379
        
        lErro = Comando_Executar(alComando(3), "INSERT INTO PlanoConta (Conta, NivelConta, TipoConta, CP, CR, TES, FAT, EST) VALUES (?,?,?,?,?,?,?,?)", sContaFinal, iNivelConta, CONTA_ANALITICA, tPlanoConta.iCP, tPlanoConta.iCR, tPlanoConta.iTES, tPlanoConta.iFAT, tPlanoConta.iEST)
        If lErro <> AD_SQL_SUCESSO Then Error 8380
        
        lErro = SaldoInicialConta_Insere(sContaFinal)
        If lErro <> SUCESSO Then Error 10588
        
        'Incluir associacao no BD
        lErro = Comando_Executar(alComando(2), "INSERT INTO ContaCcl (Conta, Ccl) VALUES (?,?)", tPlanoConta.sConta, sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 8377
        
        'incluir a conta em MvPerCcl e em MvPerCta para os exercicios existentes
        'Seleciona exercicios
        lErro = Comando_Executar(alComando(4), "SELECT FilialEmpresa, Exercicio FROM ExerciciosFilial", iFilialEmpresa, iExercicio)
        If lErro <> AD_SQL_SUCESSO Then Error 8381
        
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9373
               
        'loop dos exercicios
        Do While lErro = AD_SQL_SUCESSO
                            
'''''            lErro = Comando_Executar(alComando(5), "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Conta, Ccl) VALUES (?,?,?,?)", iFilialEmpresa, iExercicio, sContaFinal, sCcl)
'''''            If lErro <> AD_SQL_SUCESSO Then Error 8382
        
            lErro = Comando_Executar(alComando(6), "INSERT INTO MvPerCta (FilialEmpresa, Exercicio, Conta) VALUES (?,?,?)", iFilialEmpresa, iExercicio, sContaFinal)
            If lErro <> AD_SQL_SUCESSO Then Error 8383
                    
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9374
                       
        Loop
           
        objContaCcl.sCcl = sCcl
        objContaCcl.sConta = tPlanoConta.sConta
           
        lErro = SaldoInicialContaCcl_Insere(objContaCcl)
        If lErro <> SUCESSO Then Error 10589
        
    End If
       
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ContaCcl_Inclui_Associacao_Contabil = SUCESSO
    
    Exit Function

Erro_ContaCcl_Inclui_Associacao_Contabil:

    ContaCcl_Inclui_Associacao_Contabil = Err
    
    Select Case Err
            
        Case 8369
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                        
        Case 8375, 8376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, tPlanoConta.sConta, sCcl)
               
        Case 8377, 8149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONTACCL", Err, tPlanoConta.sConta, sCcl)
            
        Case 8379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sContaFinal)
            
        Case 8380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOCONTA", Err, sContaFinal)
            
        Case 8381, 9373, 9374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", Err)
            
        Case 8382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", Err, iFilialEmpresa, iExercicio, sCcl, sContaFinal)
        
        Case 8383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCTA", Err, iFilialEmpresa, iExercicio, sContaFinal)
            
        Case 10486, 10487
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err, sCcl)
        
        Case 10488
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA", Err, sCcl)
        
        Case 10588, 10589
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 150617)
        
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
        
End Function

Public Function ContaCcl_Inclui(ByVal objContaCcl As ClassContaCcl, ByVal sConta1 As String) As Long
'Inclui no BD a associacao de sConta com sCcl. sConta contem a conta pai da conta com centro de custo. sConta1 contem a conta com centro de custo contabil
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim iExercicio As Integer
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim iNivelConta As Integer
Dim objFiliais As AdmFiliais
Dim dSldIni As Double
Dim dSldIni1 As Double
Dim objSaldoInicialContaCcl As New ClassSaldoInicialContaCcl
Dim iTipoCcl As Integer

On Error GoTo Erro_ContaCcl_Inclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9559
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9560
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 9566
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 9567
        
    'acessa o centro de custo em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT TipoCcl FROM Ccl WHERE Ccl=? ", 0, iTipoCcl, objContaCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9561
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9562
    
    If lErro <> AD_SQL_SUCESSO Then Error 9563
    
    'loca o registro que guarda o centro de custo
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9564
    
    If iTipoCcl <> CCL_ANALITICA Then Error 10567
    
    'Incluir associacao no BD
    lErro = Comando_Executar(lComando1, "INSERT INTO ContaCcl (Conta, Ccl) VALUES (?,?)", objContaCcl.sConta, objContaCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9565
        
    lErro = SaldoInicialContaCcl_Insere(objContaCcl)
    If lErro <> SUCESSO Then Error 10340
        
    If objContaCcl.dSldIni <> 0 Then
    
        objSaldoInicialContaCcl.iFilialEmpresa = objContaCcl.iFilialEmpresa
        objSaldoInicialContaCcl.sConta = objContaCcl.sConta
        objSaldoInicialContaCcl.sCcl = objContaCcl.sCcl
        objSaldoInicialContaCcl.dSldIni = objContaCcl.dSldIni
    
        lErro = SaldoInicialContaCcl_Altera(objSaldoInicialContaCcl)
        If lErro <> SUCESSO Then Error 10341
        
    End If
        
    dSldIni1 = objContaCcl.dSldIni
        
    'incluir a associacao conta x centro de custo em MvPerCcl para os exercicios existentes
    'Seleciona exercicios
    lErro = Comando_Executar(lComando2, "SELECT Exercicio FROM Exercicios ORDER BY Exercicio", iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 9568
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9569
           
    'loop dos exercicios
    Do While lErro = AD_SQL_SUCESSO
                        
        For Each objFiliais In gcolFiliais
                        
            If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then
                        
                If objFiliais.iCodFilial = objContaCcl.iFilialEmpresa Or objFiliais.iCodFilial = EMPRESA_TODA Then
                    dSldIni = dSldIni1
                Else
                    dSldIni = 0
                End If
                            
'''''                lErro = Comando_Executar(lComando3, "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Conta, Ccl, SldIni) VALUES (?,?,?,?,?)", objFiliais.iCodFilial, iExercicio, sConta1, objContaCcl.sCcl, dSldIni)
'''''                If lErro <> AD_SQL_SUCESSO Then Error 9570
        
            End If
        
        Next
        
        dSldIni1 = 0
    
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9571
                   
    Loop
           
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    ContaCcl_Inclui = SUCESSO
    
    Exit Function

Erro_ContaCcl_Inclui:

    ContaCcl_Inclui = Err
    
    Select Case Err
            
        Case 9559, 9560, 9566, 9567
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 9561, 9562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err, objContaCcl.sCcl)
            
        Case 9563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objContaCcl.sCcl)
            
        Case 9564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", Err, objContaCcl.sCcl)
                        
        Case 9565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONTACCL", Err, objContaCcl.sConta, objContaCcl.sCcl)
            
        Case 9568, 9569, 9571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", Err)
            
        Case 9570
             lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", Err, iExercicio, objContaCcl.sCcl, sConta1)
             
        Case 10340, 10341
       
        Case 10567
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA1", Err, objContaCcl.sCcl)
        
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150618)
        
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function
        
End Function

Public Function Ccl_Exclui_Sem_Assoc_ExtraContabil(ByVal sConta As String, ByVal sCcl As String) As Long
'Exclui do BD a associação da conta sConta com o centro de custo/lucro sCcl /m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando1 As Long
Dim iTemMovimento As Integer
Dim sCcl1 As String
Dim sConta1 As String
Dim iTemLanPendente As Integer
Dim objContaCcl As New ClassContaCcl

On Error GoTo Erro_Ccl_Exclui_Sem_Assoc_ExtraContabil

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8208
    
    sCcl1 = String(STRING_CCL, 0)
    sConta1 = String(STRING_CONTA, 0)
    
    'pesquisa a associacao em questão no BD
    lErro = Comando_ExecutarPos(lComando1, "SELECT Conta FROM ContaCcl WHERE Conta=? AND Ccl=?", 0, sConta1, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8154
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8155
        
    'associação conta x centro de custo/lucro não cadastrada
    If lErro = AD_SQL_SEM_DADOS Then Error 8156
    
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 9913
    
    'verifica se a associação da conta com centro de custo/lucro está apta para ser excluida
    'Isto é, não possui movimento, lançamentos pendentes ou rateioon.
    lErro = CF("ContaCcl_Critica_Exclusao", sConta, sCcl)
    If lErro <> SUCESSO Then Error 9909
    
    'Exclui os documentos automaticos
    lErro = CF("DocAuto_Exclui_ContaCcl", sConta, sCcl)
    If lErro <> SUCESSO And lErro <> 8187 Then Error 8163
    
    objContaCcl.sConta = sConta
    objContaCcl.sCcl = sCcl
    
    'exclui o saldo inicial da associação conta x ccl
    lErro = SaldoInicialContaCcl_Exclui(objContaCcl)
    If lErro <> SUCESSO Then Error 10585
    
    'Excluir os saldos de Ccl associados
    lErro = CF("MvPerCcl_Exclui_ContaCcl", sConta, sCcl)
    If lErro <> SUCESSO And lErro <> 8193 Then Error 8164
    
    'Exclui a associacao
    lErro = CF("ContaCcl_Exclui_Assoc", sConta, sCcl)
    If lErro <> SUCESSO Then Error 8210
    
    Call Comando_Fechar(lComando1)
    
    Ccl_Exclui_Sem_Assoc_ExtraContabil = SUCESSO
    
    Exit Function

Erro_Ccl_Exclui_Sem_Assoc_ExtraContabil:

    Ccl_Exclui_Sem_Assoc_ExtraContabil = Err
    
    Select Case Err
                
        Case 8208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                    
        Case 8154, 8155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL", Err, sConta, sCcl)
            
        Case 8156
            
        Case 8161, 8163, 8164, 8165, 8166, 8210, 8211, 10585
                   
        Case 8209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, sCcl)
            
        Case 9909
            
        Case 9913
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTACCL", Err, sConta, sCcl)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 150619)
        
    End Select
    
    Call Comando_Fechar(lComando1)
    
    Exit Function
    
End Function

Public Function Ccl_Exclui_Sem_Assoc_Contabil(ByVal sConta As String, ByVal sCcl As String) As Long
'Exclui do BD a associação da conta sConta com o centro de custo/lucro sCcl /m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando1 As Long
Dim iTemMovimento As Integer
Dim sCcl1 As String
Dim sConta2 As String
Dim sConta3 As String
Dim objContaCcl As New ClassContaCcl

On Error GoTo Erro_Ccl_Exclui_Sem_Assoc_Contabil

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8346
    
    sConta2 = String(STRING_CONTA, 0)
    sCcl1 = String(STRING_CCL, 0)
    
    'pesquisa a associacao da conta com centro de custo/lucro
    lErro = Comando_ExecutarLockado(lComando1, "SELECT Conta FROM ContaCcl WHERE Conta=? AND Ccl=?", sConta2, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8352
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8353
        
    If lErro = AD_SQL_SEM_DADOS Then Error 8354
    
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 9928
    
    sConta3 = left(sConta, Len(sConta) - Len(sCcl)) & sCcl
    
    'Verifica se existem movimentações para a conta Assoc. Se existem, abortar.
    lErro = CF("Conta_Critica_Possui_Movimento", sConta3, iTemMovimento)
    If lErro <> SUCESSO Then Error 8357
            
    If iTemMovimento = 1 Then Error 8358

    'verifica se a conta possui lancamentos associados, rateios, etc.
    lErro = CF("ContaAnalitica_Critica_Exclusao", sConta3)
    If lErro <> SUCESSO Then Error 9926

    'Exclui os documentos automaticos
    lErro = CF("DocAuto_Exclui_Conta", sConta3)
    If lErro <> SUCESSO Then Error 8363
    
    'Excluir os saldos de Ccl associados
    lErro = CF("MvPerCcl_Exclui_Conta", sConta3)
    If lErro <> SUCESSO Then Error 8364
    
    'Excluir os saldos de Conta associados
    lErro = CF("MvPerCta_Exclui_Conta", sConta3)
    If lErro <> SUCESSO Then Error 8365
        
    lErro = SaldoInicialConta_Exclui_Conta(sConta3)
    If lErro <> SUCESSO Then Error 10586
        
    'exclui a associação da conta com o centro de custo
    lErro = CF("ContaCcl_Exclui_Assoc", sConta, sCcl)
    If lErro <> SUCESSO Then Error 8367
    
    objContaCcl.sConta = sConta
    objContaCcl.sCcl = sCcl
    
    lErro = SaldoInicialContaCcl_Exclui(objContaCcl)
    If lErro <> SUCESSO Then Error 10587
        
    'Exclui a conta
    lErro = PlanoConta_Exclui_Conta(sConta3)
    If lErro <> SUCESSO Then Error 8368
    
    Call Comando_Fechar(lComando1)
    
    Ccl_Exclui_Sem_Assoc_Contabil = SUCESSO
    
    Exit Function

Erro_Ccl_Exclui_Sem_Assoc_Contabil:

    Ccl_Exclui_Sem_Assoc_Contabil = Err
    
    Select Case Err
                
        Case 8346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                    
        Case 8352, 8353
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, sConta, sCcl)
            
        Case 8354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACCL_NAO_CADASTRADO", Err, sConta, sCcl)
            
        Case 8358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_COM_MOVIMENTO", Err, sConta3, sCcl)
        
        Case 8357, 8363, 8364, 8365, 8367, 8368, 9926, 10586, 10587
        
        Case 9928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTACCL", Err, sConta, sCcl)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 150620)
        
    End Select
    
    Call Comando_Fechar(lComando1)
    
    Exit Function
    
End Function


Public Function ContaCcl_Exclui_RateioOff_Contabil(ByVal sConta As String) As Long
'Exclui do BD os rateios off associados a sConta
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim lSeq As Long


On Error GoTo Erro_ContaCcl_Exclui_RateioOff_Contabil

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8333
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8334
        
    lErro = Comando_ExecutarPos(lComando, "SELECT Seq FROM RateioOff WHERE ContaOrigem=? OR ContaCre=? OR Conta=?", 0, lSeq, sConta, sConta, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 8335
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8337
        
    Do While lErro = AD_SQL_SUCESSO
        
        'exclui o rateio
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM RateioOff", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 8338
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9376
        
    Loop
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    ContaCcl_Exclui_RateioOff_Contabil = SUCESSO
    
    Exit Function
 
Erro_ContaCcl_Exclui_RateioOff_Contabil:

    ContaCcl_Exclui_RateioOff_Contabil = Err

    Select Case Err
    
        Case 8333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8335, 8337, 9376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFF", Err, sConta, "")
    
        Case 8338
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACCL_COM_RATEIOOFF", Err, sConta, "")
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150621)
    
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Public Function ContaCcl_Exclui_RateioOn_ExtraContab(ByVal sConta As String, ByVal sCcl As String) As Long
'Exclui do BD os rateios on associados a sConta+sCcl
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iCodigo As Integer
Dim iSeq As Integer

On Error GoTo Erro_ContaCcl_Exclui_RateioOn_ExtraContab

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8175
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8212
        
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, Seq FROM RateioOn WHERE Conta=? AND Ccl=?", 0, iCodigo, iSeq, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8176
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8178
        
    Do While lErro = AD_SQL_SUCESSO
        
        'exclui o rateio
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM RateioOn", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 8179
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9377
        
    Loop
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    ContaCcl_Exclui_RateioOn_ExtraContab = SUCESSO
    
    Exit Function
 
Erro_ContaCcl_Exclui_RateioOn_ExtraContab:

    ContaCcl_Exclui_RateioOn_ExtraContab = Err

    Select Case Err
    
        Case 8175
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                    
        Case 8176, 8178, 9377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOON", Err, sConta, sCcl)
        
        Case 8179
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACCL_COM_RATEIOON", Err, sConta, sCcl)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150622)
    
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Public Function ContaCcl_Exclui_RateioOn_Contabil(ByVal sConta As String) As Long
'Exclui do BD os rateios on associados a sConta
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iCodigo As Integer
Dim iSeq As Integer

On Error GoTo Erro_ContaCcl_Exclui_RateioOn_Contabil

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8339
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8340
        
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, Seq FROM RateioOn WHERE Conta=?", 0, iCodigo, iSeq, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 8341
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8343
        
    Do While lErro = AD_SQL_SUCESSO
    
        'exclui o rateio
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM RateioOn", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 8344
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9378
        
    Loop
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    ContaCcl_Exclui_RateioOn_Contabil = SUCESSO
    
    Exit Function
 
Erro_ContaCcl_Exclui_RateioOn_Contabil:

    ContaCcl_Exclui_RateioOn_Contabil = Err

    Select Case Err
    
        Case 8339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                    
        Case 8341, 8343, 9378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOON", Err, sConta, "")
        
        Case 8344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACCL_COM_RATEIOON", Err, sConta, "")
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150623)
    
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Public Function DocAuto_Exclui_ContaCcl(ByVal sConta As String, ByVal sCcl As String) As Long
'Exclui do BD os documentos automaticos associados a sConta+sCcl
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim lDoc As Long
Dim iSeq As Integer

On Error GoTo Erro_DocAuto_Exclui_ContaCcl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8185
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8214
    
    lErro = Comando_ExecutarPos(lComando, "SELECT Doc, Seq FROM DocAuto WHERE Conta=? AND Ccl=?", 0, lDoc, iSeq, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8186
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8188
        
    Do While lErro = AD_SQL_SUCESSO
    
        'exclui o DocAuto
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM DocAuto", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 8189
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9379
        
    Loop
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    DocAuto_Exclui_ContaCcl = SUCESSO
    
    Exit Function
 
Erro_DocAuto_Exclui_ContaCcl:

    DocAuto_Exclui_ContaCcl = Err

    Select Case Err
    
        Case 8185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8186, 8188, 9379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DOCAUTO", Err, sConta, sCcl)
    
        Case 8189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DOCAUTO", Err, sConta, sCcl)
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150624)
    
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Public Function MvPerCcl_Exclui_ContaCcl(ByVal sConta As String, ByVal sCcl As String) As Long
'Exclui do BD os saldos de centro de custo associados a sConta+sCcl em questão /m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iExercicio As Integer
Dim sConta1 As String

On Error GoTo Erro_MvPerCcl_Exclui_ContaCcl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8190
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8191

    sConta1 = String(STRING_CONTA, 0)

    'pesquisa os saldos para a assoc em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Exercicio, Conta FROM MvPerCcl WHERE Conta=? AND Ccl=?", 0, iExercicio, sConta1, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8192
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8194
        
    Do While lErro = AD_SQL_SUCESSO
    
        'exclui a associação
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM MvPerCcl", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 8195
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9380
        
    Loop
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    MvPerCcl_Exclui_ContaCcl = SUCESSO
    
    Exit Function
    
Erro_MvPerCcl_Exclui_ContaCcl:

    MvPerCcl_Exclui_ContaCcl = Err
    
    Select Case Err
    
        Case 8190, 8191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8192, 8194, 9380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL3", Err, sCcl, sConta)
    
        Case 8193
    
        Case 8195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MVPERCCL", Err, iExercicio, sCcl, sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150625)
    
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Public Function ContaCcl_Exclui_Assoc(ByVal sConta As String, ByVal sCcl As String) As Long
'Exclui do BD a associação de Conta com Centro de Custo/Lucro /m
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sCcl1 As String

On Error GoTo Erro_ContaCcl_Exclui_Assoc

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8196
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8197

    sCcl1 = String(STRING_CCL, 0)

    'posiciona a associaçao da conta com o centro de custo, visando sua exclusão
    lErro = Comando_ExecutarPos(lComando, "SELECT Ccl FROM ContaCcl WHERE Conta = ? AND Ccl=?", 0, sCcl1, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 8198
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8200
        
    If lErro = AD_SQL_SEM_DADOS Then Error 8199
     
    'exclui a associaçao
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ContaCcl", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 8201
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ContaCcl_Exclui_Assoc = SUCESSO
    
    Exit Function
    
Erro_ContaCcl_Exclui_Assoc:

    ContaCcl_Exclui_Assoc = Err
    
    Select Case Err
    
        Case 8196, 8197
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8198, 8200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, sConta, sCcl)
    
        Case 8199
            
        Case 8201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACCL", Err, sConta, sCcl)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150626)
    
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Public Function ContaCcl_Atualizacao_ExtraContabil(ByVal colContas As Collection, ByVal colCcl As Collection) As Long
'Inicia as atualizacoes correspondentes às colecoes colContas e colCcl  /m

Dim lErro As Long
Dim lTransacao As Long
Dim vConta As Variant
Dim vCcl As Variant
Dim sConta1 As String
Dim lComando As Long
Dim sConta As String
Dim sCcl As String

On Error GoTo Erro_ContaCcl_Atualizacao_ExtraContabil

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 8130
   
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8131
    
    'Para cada conta selecionada
    For Each vConta In colContas
              
        sConta1 = String(STRING_CONTA, 0)
              
        sConta = vConta
        
        'pesquisa sua existencia no plano de contas
        lErro = Comando_ExecutarPos(lComando, "SELECT Conta FROM PlanoConta WHERE Conta=?", 0, sConta1, sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 8133
        
        'Verifica se a conta esta no BD
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 8135
        
        If lErro = AD_SQL_SEM_DADOS Then Error 8134
        
        'lock da conta
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 8136
    
        lErro = CF("ContaCcl_Critica_Colecao_ExtraContab", sConta, colCcl)
        If lErro <> SUCESSO Then Error 8202
                
        For Each vCcl In colCcl
                    
            sCcl = vCcl
            
            lErro = CF("ContaCcl_Inclui_Associacao_ExtraCont", sConta, sCcl)
            If lErro <> SUCESSO Then Error 8203
                                
        Next
            
    
    Next

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 8139
    
    Call Comando_Fechar(lComando)
    
    ContaCcl_Atualizacao_ExtraContabil = SUCESSO
    
    Exit Function
    
Erro_ContaCcl_Atualizacao_ExtraContabil:

    ContaCcl_Atualizacao_ExtraContabil = Err
    
    Select Case Err
                
        Case 8130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 8131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8133, 8135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, sConta)
               
        Case 8134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, sConta)
        
        Case 8136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, sConta)
        
        Case 8202, 8203
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 150627)
        
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function ContaCcl_Atualizacao_Contabil(ByVal colContas As Collection, ByVal colCcl As Collection) As Long
'Inicia as atualizacoes correspondentes às coleções colContas e colCcl /m

Dim lErro As Long
Dim lTransacao As Long
Dim vConta As Variant
Dim sConta As String
Dim vCcl As Variant
Dim sCcl As String
Dim sConta2 As String
Dim lComando As Long
Dim iContaPreenchida As Integer
Dim tPlanoConta As typePlanoConta

On Error GoTo Erro_ContaCcl_Atualizacao_Contabil

    sConta2 = String(STRING_CONTA, 0)
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 8322
   
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 8323
    
    'Para cada conta selecionada
    For Each vConta In colContas
                              
        sConta = vConta
        
        lErro = CF("ContaCcl_Critica_Colecao_Contabil", sConta, colCcl)
        If lErro <> SUCESSO Then Error 8328
        
        tPlanoConta.sConta = String(STRING_CONTA, 0)
                
        lErro = Comando_Executar(lComando, "SELECT Conta, NivelConta, CP, CR, TES, FAT, EST FROM PlanoConta WHERE Conta=?", tPlanoConta.sConta, tPlanoConta.iNivelConta, tPlanoConta.iCP, tPlanoConta.iCR, tPlanoConta.iTES, tPlanoConta.iFAT, tPlanoConta.iEST, sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 6029
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6031
                
        For Each vCcl In colCcl
                    
            sCcl = vCcl
                    
            lErro = ContaCcl_Inclui_Associacao_Contabil(tPlanoConta, sCcl)
            If lErro <> SUCESSO Then Error 8329
                                
        Next
            
    Next

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 8331
    
    Call Comando_Fechar(lComando)
    
    ContaCcl_Atualizacao_Contabil = SUCESSO
    
    Exit Function
    
Erro_ContaCcl_Atualizacao_Contabil:

    ContaCcl_Atualizacao_Contabil = Err
    
    Select Case Err
                
        Case 8322
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 8323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 8328, 8329
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 150628)
        
    End Select
        
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RateioOn_Exclui(ByVal iCodigo As Integer) As Long
    
Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_RateioOn_Exclui

    lTransacao = 0
       
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 11180
  
    'exclui os lançamentos do rateio com codigo = icodigo
    lErro = RateioOn_Exclui_BD(iCodigo, CTB_RATEIOON_TESTA_EXISTENCIA)
    If lErro <> SUCESSO Then Error 11181
  
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 11182

    RateioOn_Exclui = SUCESSO

    Exit Function

Erro_RateioOn_Exclui:
    
    RateioOn_Exclui = Err
    
    Select Case Err
            
        Case 11180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
           
        Case 11181
          
        Case 11182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150629)
            
   End Select
   
   Call Transacao_Rollback
   
   Exit Function

End Function

Function RateioOn_Grava(ByVal iCodigo As Integer, ByVal sDescricao As String, ByVal colRateioOns As Collection) As Long
                                
Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_RateioOn_Grava

    lTransacao = 0
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 11176
    
    'exclui os lançamentos do Rateio com código = icodigo
    lErro = RateioOn_Exclui_BD(iCodigo, CTB_RATEIOON_NAO_TESTA_EXISTENCIA)
    If lErro <> SUCESSO Then Error 11177

    'inclui os lançamentos do Rateio
    lErro = CF("RateioOn_Grava_BD", iCodigo, sDescricao, colRateioOns)
    If lErro <> SUCESSO Then Error 11178
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 11179

   
    RateioOn_Grava = SUCESSO

    Exit Function

Erro_RateioOn_Grava:
    
    RateioOn_Grava = Err
    
    Select Case Err
    
        Case 11176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 11177, 11178

        Case 11179
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150630)
            
    End Select
    
    Call Transacao_Rollback

    Exit Function

End Function

Function RateioOn_Grava_BD(ByVal iCodigo As Integer, ByVal sDescricao As String, ByVal colRateioOns As Collection) As Long
'grava na tabela de Rateios Automaticos (RateioOn)lançamentos do voucher em questão

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim iTipoConta As Integer
Dim sCcl As String
Dim sConta As String
Dim iPossuiCcl As Integer
Dim iStatus As Integer
Dim objRateioOn As ClassRateioOn

On Error GoTo Erro_RateioOn_Grava_BD

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11155
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 11156
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 11157
    
    'Inicia Loop de GRAVACAO com Lançamentos do GRID
    For Each objRateioOn In colRateioOns
            
        'Pesquisa Conta associada a um Lançamento
        lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta, Ativo FROM PlanoConta WHERE Conta=?", iTipoConta, iStatus, objRateioOn.sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 11158
        
       'Lê a Conta
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11159
        
        If lErro = AD_SQL_SEM_DADOS Then Error 11160
        
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 11161
            
        'verifica se a conta é analítica
        If iTipoConta <> CONTA_ANALITICA Then Error 11162
        
        sCcl = String(STRING_CCL, 0)

        iPossuiCcl = 0
           
        'Se o lançamento envolver centro de custo extra contabil
        'Verificar se a associação conta x centro de custo está cadastrada
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL And Len(objRateioOn.sCcl) > 0 Then
        
            sCcl = objRateioOn.sCcl
            
            sConta = String(STRING_CONTA, 0)
            
            'Pesquisa Ccl associado ao Lançamento
            lErro = Comando_ExecutarLockado(lComando2, "SELECT Conta FROM ContaCcl WHERE Conta=? AND Ccl=?", sConta, objRateioOn.sConta, sCcl)
            If lErro <> AD_SQL_SUCESSO Then Error 11163
        
           'Lê a associação Conta x Ccl
            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11164
        
            If lErro = AD_SQL_SEM_DADOS Then Error 11165
        
            lErro = Comando_LockShared(lComando2)
            If lErro <> AD_SQL_SUCESSO Then Error 11166
            
        End If
        
        'Faz inserção do Lançamento na Tabela de Rateio OnLine
        lErro = Comando_Executar(lComando1, "INSERT INTO RateioOn (Codigo, Seq, Descricao, Conta,  Ccl,  Historico, Percentual ) VALUES (?,?,?,?,?,?,?)", iCodigo, objRateioOn.iSeq, sDescricao, objRateioOn.sConta, objRateioOn.sCcl, objRateioOn.sHistorico, objRateioOn.dPercentual)
        If lErro <> AD_SQL_SUCESSO Then Error 11167
        
        lErro = Comando_Unlock(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 11168
        
        If iPossuiCcl = 1 Then
            lErro = Comando_Unlock(lComando2)
            If lErro <> AD_SQL_SUCESSO Then Error 11169
        End If
           
    Next

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    RateioOn_Grava_BD = SUCESSO
    
    Exit Function
    
Erro_RateioOn_Grava_BD:

    RateioOn_Grava_BD = Err
    
    Select Case Err
    
        Case 11155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 11156
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 11157
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 11158, 11159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, objRateioOn.sConta)
    
        Case 11160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, objRateioOn.sConta)

        Case 11161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, objRateioOn.sConta)
    
        Case 11162
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, objRateioOn.sConta)
    
        Case 11163, 11164
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, objRateioOn.sConta, sCcl)
    
        Case 11165
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACCL_NAO_CADASTRADO", Err, objRateioOn.sConta, sCcl)
    
        Case 11166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTACCL", Err, objRateioOn.sConta, sCcl)
    
        Case 11167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RATEIOON", Err)
    
        Case 11168
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_PLANOCONTA", Err, objRateioOn.sConta)
    
        Case 11169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_CONTACCL", Err, objRateioOn.sConta, sCcl)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150631)
                
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Private Function RateioOn_Exclui_BD(ByVal iCodigo As Integer, ByVal iTestaExistencia As Integer) As Long
'exclui da tabela RateioOn os Rateios Automáticos do voucher em questão

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iSeq As Integer
Dim tRateioOn As typeRateioOn

On Error GoTo Erro_RateioOn_Exclui_BD

    'Abre comando para pesquisar Lancamentos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11170
    
    'Abre comando para transação de exclusão
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 11171
    
    tRateioOn.iCodigo = iCodigo
    tRateioOn.sConta = String(STRING_CONTA, 0)
    tRateioOn.sCcl = String(STRING_CCL, 0)
    tRateioOn.sHistorico = String(STRING_HISTORICO, 0)
    tRateioOn.sDescricao = String(STRING_RATEIO_DESCRICAO, 0)
   
    'Pesquisa a existencia dos Rateios no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT Seq, Descricao, Conta,  Ccl,  Historico, Percentual FROM RateioOn WHERE Codigo = ?", 0, tRateioOn.iSeq, tRateioOn.sDescricao, tRateioOn.sConta, tRateioOn.sCcl, tRateioOn.sHistorico, tRateioOn.dPercentual, tRateioOn.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 11172
        
    'Lê o primeiro Rateio
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11173

    If iTestaExistencia = CTB_RATEIOON_TESTA_EXISTENCIA Then
    
        'se o rateio não estiver cadastrado ==> erro
        If lErro = AD_SQL_SEM_DADOS Then Error 55701
        
    End If
    
    'Inicia Loop de EXCLUSAO com Rateios do GRID
    Do While lErro <> AD_SQL_SEM_DADOS  'Refere-se ao BuscarProximo
        
        'Exclui os Rateios com icodigo da tabela RateioOn
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM RateioOn", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 11174
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11175
    
    Loop
   
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    RateioOn_Exclui_BD = SUCESSO
    
    Exit Function
    
Erro_RateioOn_Exclui_BD:

    RateioOn_Exclui_BD = Err
    
    Select Case Err
    
        Case 11170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 11171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 11172, 11173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOON", Err)
            
        Case 11174
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RATEIOON", Err, tRateioOn.iCodigo)
            
        Case 11175
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOON", Err)
            
        Case 55701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RATEIOON_NAO_CADASTRADO", Err, tRateioOn.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150632)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
                   
End Function

Function RateioOn_Automatico(iCodigo As Integer) As Long
'retorna o número do proximo código de Rateio disponivel

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iCodigo1 As Integer
Dim iCodigo2 As Integer

On Error GoTo Erro_RateioOn_Automatico

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11145
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 11146
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 11147
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 11148

    'Pesquisa o ultimo codigo gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT CodRateioOn FROM Configuracao", 0, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 11149

    'le o registro com o número do codigo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 11150

    'loca o registro que guarda o codigo gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 11151

    'verifica se este número ainda não foi utilizado
    'Pesquisa os códigos maiores ou iguais ao codigo gerado automaticamente
            
    lErro = Comando_Executar(lComando1, "SELECT Codigo FROM RateioOn WHERE  Codigo >= ? ORDER BY Codigo", iCodigo1, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 11152
        
    'le o primeiro codigo >= ao codigo automatico
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36791
    
    'pesquisa o proximo código disponível
    Do While lErro <> AD_SQL_SEM_DADOS
        
        If iCodigo < iCodigo1 Then
              
            'Se ainda não foi usado ==> este é o número a ser usado
             Exit Do
        
        Else
            'se o número for igual ==> passa para o proximo
            iCodigo = iCodigo + 1
        
            iCodigo2 = iCodigo1
        
            'pesquisa o proximo codigo
            Do While iCodigo2 = iCodigo1 And lErro <> AD_SQL_SEM_DADOS
        
                lErro = Comando_BuscarProximo(lComando1)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36792
            
            Loop
            
        End If
        
    Loop
    
    'atualiza o codigo gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Configuracao SET CodRateioOn = ?", lComando, iCodigo + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 11153
        
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 11154
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    RateioOn_Automatico = SUCESSO
        
    Exit Function
    
    
Erro_RateioOn_Automatico:

    RateioOn_Automatico = Err

    Select Case Err
    
        Case 11145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 11146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 11147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 11148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 11149, 11150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
        
        Case 11151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)
        
        Case 11152, 36791, 36792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOON2", Err)
        
        Case 11153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONFIGURACAO", Err)
        
        Case 11154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150633)
        
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function
        
End Function

Function Segmento_Grava_Conta(ByVal colSegmentos As Collection) As Long
'grava os registros da colecao na tabela Segmento

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Segmento_Grava_Conta

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 12132

    lErro = CF("Segmento_Grava_Conta_Trans", colSegmentos)
    If lErro <> SUCESSO Then Error 20719
     
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 12144
     
    Segmento_Grava_Conta = SUCESSO
    
    Exit Function

Erro_Segmento_Grava_Conta:

    Segmento_Grava_Conta = Err

    Select Case Err

        Case 12132
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 12144
             lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)

        Case 20719

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150634)

    End Select
    
    Call Transacao_Rollback

    Exit Function

End Function

Function Segmento_Grava_Conta_Trans(ByVal colSegmentos As Collection) As Long
'grava os registros da colecao na tabela Segmento de Conta
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim iPossui_Conta As Integer
Dim iExercicio As Integer
Dim iIndice As Integer
Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim tSegmento As typeSegmento
Dim sCodigo As String
Dim sCodigo2 As String
Dim objSegmento As New ClassSegmento
Dim iCcl As Integer

On Error GoTo Erro_Segmento_Grava_Conta_Trans

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12129
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 12130
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 12131
  
    sCodigo = String(STRING_SEGMENTO_CODIGO, 0)
    sCodigo2 = String(STRING_SEGMENTO_CODIGO, 0)
    tSegmento.sDelimitador = String(STRING_SEGMENTO_DELIMITADOR, 0)
    
    'pega um objeto da colSegmentos para saber qual o formato( = campo codigo na tabela Segmento)
    For Each objSegmento In colSegmentos
        'guarda o formato em sCodigo para depois fazer o Select na tabela Segmentos
        sCodigo = objSegmento.sCodigo
        Exit For
    Next
       
    If sCodigo <> SEGMENTO_CONTA Then Error 20724
       
    'verifica existencia de pelo menos uma conta em PlanoConta
    lErro = CF("PlanoConta_ExisteConta", iPossui_Conta)
    If lErro <> SUCESSO Then Error 12136

    If iPossui_Conta = POSSUI_CONTA Then
    'so faz modificacao no delimitador e no preenchimento da tabela Segmento
        
         lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, Nivel, Delimitador, Preenchimento FROM Segmento WHERE Codigo = ? ORDER BY Nivel", 0, sCodigo2, tSegmento.iNivel, tSegmento.sDelimitador, tSegmento.iPreenchimento, sCodigo)
         If lErro <> AD_SQL_SUCESSO Then Error 12137
     
         'percorre cada registro em colSegmentos para fazer a modificacao nos delimitadores e no preenchimento
         For Each objSegmento In colSegmentos
                     
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO Then Error 12138
            
            lErro = Comando_ExecutarPos(lComando2, "UPDATE Segmento SET Codigo =? , Nivel =? , Delimitador =? , Preenchimento =? ", lComando, sCodigo, objSegmento.iNivel, objSegmento.sDelimitador, objSegmento.iPreenchimento)
            If lErro <> AD_SQL_SUCESSO Then Error 12139
            
         Next
    
    Else
        'apaga toda tabela Segmento e insere os dados novos em Segmento
        
        'prepara o lock na tabela Configuracao
        lErro = Comando_ExecutarPos(lComando3, "SELECT UsoCcl FROM Configuracao", 0, iCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 10366
    
        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10367
    
        'Lock da Configuracao
        lErro = Comando_LockExclusive(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 10368
    
        'Seleciona registros na tabela de Segmento
        lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, Nivel FROM Segmento WHERE Codigo = ?", 0, sCodigo2, tSegmento.iNivel, sCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 12140
        
        'le o primeiro segmento
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12141
        
        'verifica se ainda existe registro a ser deletado
        Do While lErro <> AD_SQL_SEM_DADOS
           
            'exclui o segmento
            lErro = Comando_ExecutarPos(lComando2, "DELETE FROM Segmento", lComando)
            If lErro <> AD_SQL_SUCESSO Then Error 12142
        
            'le o proximo segmento
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10365
        
        Loop
        
        For Each objSegmento In colSegmentos
           
           lErro = Comando_Executar(lComando2, "INSERT INTO Segmento (Codigo,Nivel,Tipo,Tamanho,Delimitador,Preenchimento) VALUES (?,?,?,?,?,?)", sCodigo, objSegmento.iNivel, objSegmento.iTipo, objSegmento.iTamanho, objSegmento.sDelimitador, objSegmento.iPreenchimento)
           If lErro <> AD_SQL_SUCESSO Then Error 12143
           
        Next
        
    End If
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
     
    Segmento_Grava_Conta_Trans = SUCESSO
    
    Exit Function

Erro_Segmento_Grava_Conta_Trans:

    Segmento_Grava_Conta_Trans = Err

    Select Case Err

        Case 10365, 12140, 12141, 12137, 12138
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SEGMENTO1", Err, sCodigo)
             
        Case 10366, 10367
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
        
        Case 10368
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)

        Case 12129, 12130, 12131
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12136

        Case 12139
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SEGMENTO", Err, sCodigo, tSegmento.iNivel)

        Case 12142
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SEGMENTO", Err, sCodigo, tSegmento.iNivel)

        Case 12143
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SEGMENTO", Err, sCodigo, objSegmento.iNivel)

        Case 20724
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_CONTA_INVALIDO", Err, sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150635)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function ConfiguracaoSetup_Grava(ByVal objConfiguracao As ClassConfiguracao) As Long
'Insere os seguintes campos da tabela Configuracao : LotePorPeriodo, DocPorPeriodo, UsoCcl

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ConfiguracaoSetup_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 12178
    
    lErro = CF("ConfiguracaoSetup_Grava_Trans", objConfiguracao)
    If lErro <> SUCESSO Then Error 20721
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 12183
        
    ConfiguracaoSetup_Grava = SUCESSO
    
    Exit Function

Erro_ConfiguracaoSetup_Grava:

    ConfiguracaoSetup_Grava = Err

    Select Case Err

        Case 12178
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 12183
             lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)
        
        Case 20721
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150636)

    End Select
    
    Call Transacao_Rollback

    Exit Function

End Function

Function ConfiguracaoSetup_Grava_Trans(ByVal objConfiguracao As ClassConfiguracao) As Long
'Insere os seguintes campos da tabela Configuracao : LotePorPeriodo, DocPorPeriodo, UsoCcl
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim iCcl As Integer
Dim iHistPadrao As Integer
Dim iIndice As Integer
Dim lComando2 As Long

On Error GoTo Erro_ConfiguracaoSetup_Grava_Trans

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 12177
    
    'faz as alteracoes em TipoConta , Origem e Natureza na tabela de Configurarcao
    lErro = Comando_Executar(lComando2, "INSERT INTO Configuracao (IdAtualizacao, UsoCcl, LotePorPeriodo, DocPorPeriodo, TipoContaDefault, NaturezaDefault, HistPadrao, DataImplantacao, DocAuto, CodRateioOn, Categoria, CodRateioOff) VALUES (1,?,?,?,?,?,1,?,1,1,1,1)", objConfiguracao.iUsoCcl, objConfiguracao.iLotePorPeriodo, objConfiguracao.iDocPorPeriodo, objConfiguracao.iTipoContaDefault, objConfiguracao.iNaturezaDefault, Date)
    If lErro <> AD_SQL_SUCESSO Then Error 12182
                     
    Call Comando_Fechar(lComando2)
    
    ConfiguracaoSetup_Grava_Trans = SUCESSO
    
    Exit Function

Erro_ConfiguracaoSetup_Grava_Trans:

    ConfiguracaoSetup_Grava_Trans = Err

    Select Case Err

        Case 12176
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12177
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12179, 12180
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)

        Case 12181
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)

        Case 12182
             lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CONFIGURACAO", Err, objConfiguracao.iTipoContaDefault, "", objConfiguracao.iNaturezaDefault)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150637)

    End Select
    
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Configuracao_Altera(ByVal objConfiguracao As ClassConfiguracao) As Long
'grava os registros do objeto na tabela Configuracao

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Configuracao_Altera

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 12164
    
    lErro = CF("Configuracao_Altera_Trans", objConfiguracao)
    If lErro <> SUCESSO Then Error 20732
                     
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 12169
        
    Configuracao_Altera = SUCESSO
    
    Exit Function

Erro_Configuracao_Altera:

    Configuracao_Altera = Err

    Select Case Err

        Case 12164
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 12169
             lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)
             
        Case 20732
        
        Case Else
            
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150638)

    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Function Configuracao_Altera_Trans(ByVal objConfiguracao As ClassConfiguracao) As Long
'grava os registros do objeto na tabela Configuracao

Dim lErro As Long
Dim iCcl As Integer
Dim iHistPadrao As Integer
Dim iIndice As Integer
Dim lComando As Long
Dim lComando2 As Long

On Error GoTo Erro_Configuracao_Altera_Trans

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12162
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 12163
    
    'apenas prepara o lock na tabela Configuracao
    lErro = Comando_ExecutarPos(lComando, "SELECT UsoCcl , HistPadrao FROM Configuracao", 0, iCcl, iHistPadrao)
    If lErro <> AD_SQL_SUCESSO Then Error 12165
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12166
    
    'Lock da Configuracao
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 12167
           
    'faz as alteracoes em TipoConta e Natureza na tabela de Configurarcao
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Configuracao SET TipoContaDefault =? , NaturezaDefault =? ", lComando, objConfiguracao.iTipoContaDefault, objConfiguracao.iNaturezaDefault)
    If lErro <> AD_SQL_SUCESSO Then Error 12168
                     
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Configuracao_Altera_Trans = SUCESSO
    
    Exit Function

Erro_Configuracao_Altera_Trans:

    Configuracao_Altera_Trans = Err

    Select Case Err

        Case 12162
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12163
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12165, 12166
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)

        Case 12167
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)

        Case 12168
             lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CONFIGURACAO", Err, objConfiguracao.iTipoContaDefault, "", objConfiguracao.iNaturezaDefault)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150639)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function



Function MvPerCcl_Exclui_Exercicio(ByVal iExercicio As Integer) As Long
'exclui todos os saldos de centro de custo/lucro associados com o Exercicio passado como parametro
'ATENÇÃO: Tem que estar em transação o programa que chamou esta rotina

Dim lErro As Long
Dim sCcl As String
Dim sConta As String
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_MvPerCcl_Exclui_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9323

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9324

    sCcl = String(STRING_CCL, 0)
    sConta = String(STRING_CONTA, 0)

    'Pesquisa todos os saldos para o exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Ccl, Conta FROM MvPerCcl WHERE Exercicio = ?", 0, sCcl, sConta, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 9325

    'le o primeiro saldo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9326
    
    Do While lErro = AD_SQL_SUCESSO

        'Remove o registro de Saldo de Ccl
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM MvPerCcl", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 9327
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9328
        
    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    MvPerCcl_Exclui_Exercicio = SUCESSO

    Exit Function

Erro_MvPerCcl_Exclui_Exercicio:

    MvPerCcl_Exclui_Exercicio = Err

    Select Case Err

        Case 9323, 9324
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9325, 9326, 9328
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", Err)

        Case 9327
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MVPERCCL", Err, iExercicio, sCcl, sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150640)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function MvPerCta_Exclui_Exercicio(ByVal iExercicio As Integer) As Long
'exclui todos os saldos de conta associados com o Exercicio passado como parametro
'ATENÇÃO: Tem que estar em transação o programa que chamou esta rotina

Dim lErro As Long
Dim sConta As String
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_MvPerCta_Exclui_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9329

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9330

    sConta = String(STRING_CONTA, 0)

    'Pesquisa todos os saldos para o exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Conta FROM MvPerCta WHERE Exercicio = ?", 0, sConta, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 9331

    'le o primeiro saldo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9332
    
    Do While lErro = AD_SQL_SUCESSO

        'Remove o registro de Saldo de Conta
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM MvPerCta", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 9333
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9334
        
    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    MvPerCta_Exclui_Exercicio = SUCESSO

    Exit Function

Erro_MvPerCta_Exclui_Exercicio:

    MvPerCta_Exclui_Exercicio = Err

    Select Case Err

        Case 9329, 9330
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9331, 9332, 9334
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", Err)

        Case 9333
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MVPERCTA", Err, iExercicio, sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150641)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Periodo_Exclui_Exercicio(ByVal iExercicio As Integer) As Long
'exclui todos os periodos associados com o Exercicio passado como parametro
'ATENÇÃO: Tem que estar em transação o programa que chamou esta rotina

Dim lErro As Long
Dim iPeriodo As Integer
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_Periodo_Exclui_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9335

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9336

    'Pesquisa todos os periodos para o exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Periodo FROM Periodo WHERE Exercicio = ?", 0, iPeriodo, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 9337

    'le o primeiro periodo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9338
    
    Do While lErro = AD_SQL_SUCESSO

        'Remove o registro de Periodo
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM Periodo", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 9339
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9340
        
    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Periodo_Exclui_Exercicio = SUCESSO

    Exit Function

Erro_Periodo_Exclui_Exercicio:

    Periodo_Exclui_Exercicio = Err

    Select Case Err

        Case 9335, 9336
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9337, 9338, 9340
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO1", Err)

        Case 9339
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PERIODO", Err, iPeriodo, iExercicio)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150642)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function PeriodosFilial_Exclui_Exercicio(ByVal iExercicio As Integer) As Long
'exclui todos os periodos associados com o Exercicio passado como parametro
'ATENÇÃO: Tem que estar em transação o programa que chamou esta rotina

Dim lErro As Long
Dim iPeriodo As Integer
Dim iFilialEmpresa As Integer
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_PeriodosFilial_Exclui_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10180

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10181

    'Pesquisa todos os periodos para o exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa, Periodo FROM PeriodosFilial WHERE Exercicio = ?", 0, iFilialEmpresa, iPeriodo, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 10182

    'le o primeiro periodo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10183
    
    Do While lErro = AD_SQL_SUCESSO

        'Remove o registro de Periodo
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM PeriodosFilial", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 10184
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10185
        
    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    PeriodosFilial_Exclui_Exercicio = SUCESSO

    Exit Function

Erro_PeriodosFilial_Exclui_Exercicio:

    PeriodosFilial_Exclui_Exercicio = Err

    Select Case Err

        Case 10180, 10181
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10182, 10183, 10185
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL1", Err)

        Case 10184
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PERIODOSFILIAL", Err, iFilialEmpresa, iExercicio, iPeriodo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150643)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function ExerciciosFilial_Exclui_Exercicio(ByVal iExercicio As Integer) As Long
'exclui todos os exerciciosFilial associados com o Exercicio passado como parametro
'ATENÇÃO: Tem que estar em transação o programa que chamou esta rotina

Dim lErro As Long
Dim iFilialEmpresa As Integer
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_ExerciciosFilial_Exclui_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10727

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10728

    'Pesquisa todos os registros da tabela exerciciosFilial para o exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa FROM ExerciciosFilial WHERE Exercicio = ?", 0, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 10729

    'le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10730
    
    Do While lErro = AD_SQL_SUCESSO

        'Remove o registro de Periodo
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ExerciciosFilial", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 10731
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10732
        
    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ExerciciosFilial_Exclui_Exercicio = SUCESSO

    Exit Function

Erro_ExerciciosFilial_Exclui_Exercicio:

    ExerciciosFilial_Exclui_Exercicio = Err

    Select Case Err

        Case 10727, 10728
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10729, 10730, 10732
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL2", Err)

        Case 10731
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EXERCICIOSFILIAL", Err, iFilialEmpresa, iExercicio)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150644)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function


Function ExercicioOrigem_Exclui_Exercicio(ByVal iExercicio As Integer) As Long
'exclui todos os registros da tabela ExercicioOrigem associados com o Exercicio passado como parametro
'ATENÇÃO: Tem que estar em transação o programa que chamou esta rotina

Dim lErro As Long
Dim sOrigem As String
Dim iPeriodo As Integer
Dim iFilialEmpresa As Integer
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_ExercicioOrigem_Exclui_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9341

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9342

    sOrigem = String(STRING_ORIGEM, 0)

    'Pesquisa todos os registros de ExercicioOrigem para o exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa, Periodo, Origem FROM ExercicioOrigem WHERE Exercicio = ?", 0, iFilialEmpresa, iPeriodo, sOrigem, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 9343

    'le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9344
    
    Do While lErro = AD_SQL_SUCESSO

        'Remove o registro de ExercicioOrigem
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ExercicioOrigem", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 9345
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9346
        
    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ExercicioOrigem_Exclui_Exercicio = SUCESSO

    Exit Function

Erro_ExercicioOrigem_Exclui_Exercicio:

    ExercicioOrigem_Exclui_Exercicio = Err

    Select Case Err

        Case 9341, 9342
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9343, 9344, 9346
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOORIGEM1", Err)

        Case 9345
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EXERCICIOORIGEM", Err, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150645)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Function MvPerCcl_Insere_Exercicio(ByVal iExercicio As Integer, ByVal iIncluiSaldo As Integer, ByVal colFiliais As Collection) As Long
'Pesquisa todos os centros de custo e as associações conta x centro de custo/lucro e insere registros de saldo na tabela MvPerCcl.
'ATENÇÃO: Esta rotina tem que ser chamada por um programa que esteja em transação

Dim lErro As Long
Dim sCcl As String
Dim sConta As String
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim dSldIni As Double
Dim sConta1 As String
Dim objFiliais As AdmFiliais

On Error GoTo Erro_MvPerCcl_Insere_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9347

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9348

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10206

    sConta = String(STRING_CONTA, 0)
    sCcl = String(STRING_CCL, 0)

    'insere os saldos na tabela MvPerCcl de todos os centros de custo
    lErro = MvPerCcl_Insere_Exercicio_Ccl(iExercicio, iIncluiSaldo, colFiliais)
    If lErro <> SUCESSO Then Error 10950

    'Pesquisa todas as associações de conta com ccl
    lErro = Comando_Executar(lComando, "SELECT Ccl, Conta FROM ContaCcl", sCcl, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 9349

    'le a primeira associação Conta x Ccl
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9350
    
    Do While lErro = AD_SQL_SUCESSO

        'Se está trabalhando com centro de custo contábil
        If giSetupUsoCcl = CCL_USA_CONTABIL Then
            'conta acoplada ao centro de custo
            sConta1 = left(sConta, Len(sConta) - Len(sCcl)) & sCcl
        Else
            sConta1 = sConta
        End If
        
'''''        For Each objFiliais In colFiliais
'''''
'''''            If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then
'''''
'''''                'se não deve incluir saldo inicial ==> zera o saldo inicial
'''''                If iIncluiSaldo = 0 Then
'''''                    dSldIni = 0
'''''                Else
'''''
'''''                    'Pesquisa o saldo inicial da conta x centro de custo/lucro em questão
'''''                    lErro = Comando_Executar(lComando2, "SELECT SldIni FROM SaldoInicialContaCcl WHERE FilialEmpresa=? AND Conta=? AND Ccl=?", dSldIni, objFiliais.iCodFilial, sConta, sCcl)
'''''                    If lErro <> AD_SQL_SUCESSO Then Error 10207
'''''
'''''                    'le o saldo
'''''                    lErro = Comando_BuscarPrimeiro(lComando2)
'''''                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10208
'''''
'''''                    If lErro = AD_SQL_SEM_DADOS Then dSldIni = 0
'''''
'''''                End If
'''''
'''''                'Insere o registro de Saldo de Ccl
'''''                lErro = Comando_Executar(lComando1, "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Ccl, Conta, SldIni) VALUES (?, ?, ?, ?, ?)", objFiliais.iCodFilial, iExercicio, sCcl, sConta1, dSldIni)
'''''                If lErro <> AD_SQL_SUCESSO Then Error 9351
'''''
'''''            End If
'''''
'''''        Next
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9352
            
    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    MvPerCcl_Insere_Exercicio = SUCESSO

    Exit Function

Erro_MvPerCcl_Insere_Exercicio:

    MvPerCcl_Insere_Exercicio = Err

    Select Case Err

        Case 9347, 9348, 10206
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9349, 9350, 9352
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL", Err)

        Case 9351
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", Err, objFiliais.iCodFilial, iExercicio, sCcl, sConta)

        Case 10207, 10208
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL3", Err, objFiliais.iCodFilial, sConta, sCcl)
            
        Case 10950
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150646)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Private Function MvPerCcl_Insere_Exercicio_Ccl(ByVal iExercicio As Integer, ByVal iIncluiSaldo As Integer, ByVal colFiliais As Collection) As Long
'Pesquisa todos os centros de custo  e insere registros de saldo na tabela MvPerCcl.
'ATENÇÃO: Esta rotina tem que ser chamada por um programa que esteja em transação

Dim lErro As Long
Dim sCcl As String
Dim sConta As String
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim dSldIni As Double
Dim sConta1 As String
Dim objFiliais As AdmFiliais

On Error GoTo Erro_MvPerCcl_Insere_Exercicio_Ccl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10951

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10952

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10953

    sCcl = String(STRING_CCL, 0)

    'Pesquisa todos os centros de custo/lucro
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM Ccl", sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 10954

    'le o primeiro centro de custo/lucro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10955

    Do While lErro = AD_SQL_SUCESSO

        For Each objFiliais In colFiliais

            If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then

                'se não deve incluir saldo inicial ==> zera o saldo inicial
                If iIncluiSaldo = 0 Then
                    dSldIni = 0
                Else

                    'Pesquisa o saldo inicial do centro de custo/lucro em questão
                    lErro = Comando_Executar(lComando2, "SELECT SldIni FROM SaldoInicialContaCcl WHERE FilialEmpresa=? AND Ccl=? AND Conta=? ", dSldIni, objFiliais.iCodFilial, sCcl, "")
                    If lErro <> AD_SQL_SUCESSO Then Error 10956

                    'le o saldo
                    lErro = Comando_BuscarPrimeiro(lComando2)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10957

                    If lErro = AD_SQL_SEM_DADOS Then dSldIni = 0

                End If

                'Insere o registro de Saldo de Ccl
                lErro = Comando_Executar(lComando1, "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Ccl, Conta, SldIni) VALUES (?, ?, ?, ?, ?)", objFiliais.iCodFilial, iExercicio, sCcl, "", dSldIni)
                If lErro <> AD_SQL_SUCESSO Then Error 10958

            End If

        Next

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10959

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    MvPerCcl_Insere_Exercicio_Ccl = SUCESSO

    Exit Function

Erro_MvPerCcl_Insere_Exercicio_Ccl:

    MvPerCcl_Insere_Exercicio_Ccl = Err

    Select Case Err

        Case 10951, 10952, 10953
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10954, 10955, 10959
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL1", Err)

        Case 10956, 10957
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL3", Err, objFiliais.iCodFilial, "", sCcl)
            
        Case 10958
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", Err, objFiliais.iCodFilial, iExercicio, sCcl, "")

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150647)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Private Function MvPerCta_Insere_Exercicio(ByVal iExercicio As Integer, ByVal iIncluiSaldo As Integer, ByVal colFiliais As Collection) As Long
'Pesquisa todas as contas do plano de contas e insere registros de saldo na tabela MvPerCta para cada conta
'ATENÇÃO: Esta rotina tem que ser chamada por um programa que esteja em transação

Dim lErro As Long
Dim sConta As String
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim objExercicio As New ClassExercicio
Dim dSldIni As Double
Dim objFiliais As AdmFiliais

On Error GoTo Erro_MvPerCta_Insere_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9353

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9354

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10203

    sConta = String(STRING_CONTA, 0)

    'Pesquisa todas as contas do Plano de Contas
    lErro = Comando_Executar(lComando, "SELECT Conta FROM PlanoConta", sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 9355

    'le a primeira conta
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9356
    
    Do While lErro = AD_SQL_SUCESSO

        For Each objFiliais In colFiliais

            If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then

                'se não deve incluir saldo inicial ==> zera o saldo inicial
                If iIncluiSaldo = 0 Then
                    dSldIni = 0
                Else
                    
                    'Pesquisa o saldo inicial da conta em questão
                    lErro = Comando_Executar(lComando2, "SELECT SldIni FROM SaldoInicialConta WHERE FilialEmpresa=? AND Conta=?", dSldIni, objFiliais.iCodFilial, sConta)
                    If lErro <> AD_SQL_SUCESSO Then Error 10204
    
                    'le a primeira conta
                    lErro = Comando_BuscarPrimeiro(lComando2)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10205
                    
                    If lErro = AD_SQL_SEM_DADOS Then dSldIni = 0
                    
                End If
                
    
                'Insere o registro de Saldo de Conta da Filial
                lErro = Comando_Executar(lComando1, "INSERT INTO MvPerCta (FilialEmpresa, Exercicio, Conta, SldIni) VALUES (?, ?, ?, ?)", objFiliais.iCodFilial, iExercicio, sConta, dSldIni)
                If lErro <> AD_SQL_SUCESSO Then Error 9357
                
            End If
            
        Next
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9358
        
    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    MvPerCta_Insere_Exercicio = SUCESSO

    Exit Function

Erro_MvPerCta_Insere_Exercicio:

    MvPerCta_Insere_Exercicio = Err

    Select Case Err

        Case 9353, 9354, 10203
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9355, 9356, 9358
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)
            
        Case 9357
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCTA", Err, objFiliais.iCodFilial, iExercicio, sConta)

        Case 10204, 10205
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA", Err, objFiliais.iCodFilial, sConta)

        Case 10224
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCTA", Err, 0, iExercicio, sConta)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150648)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Periodo_Insere_Exercicio(ByVal iExercicio As Integer, ByVal iNumPeriodos As Integer, adtDataInicio() As Date, adtDataFim() As Date, asNomeExterno() As String, aiFechado() As Integer) As Long
'Insere os periodos relativos a um exercicio
'ATENÇÃO: Esta rotina tem que ser chamada por um programa que esteja em transação

Dim lErro As Long
Dim lComando As Long
Dim iPeriodo As Integer

On Error GoTo Erro_Periodo_Insere_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5266

    For iPeriodo = 1 To iNumPeriodos

        'Insere o Periodo em questão
        lErro = Comando_Executar(lComando, "INSERT INTO Periodo (Exercicio, Periodo, NomeExterno, DataInicio, DataFim ) VALUES (?,?,?,?,?)", iExercicio, iPeriodo, asNomeExterno(iPeriodo - 1), adtDataInicio(iPeriodo - 1), adtDataFim(iPeriodo - 1))
        If lErro <> AD_SQL_SUCESSO Then Error 5267

    Next

    Call Comando_Fechar(lComando)

    Periodo_Insere_Exercicio = SUCESSO

    Exit Function

Erro_Periodo_Insere_Exercicio:

    Periodo_Insere_Exercicio = Err

    Select Case Err

        Case 5266
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5267
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PERIODO", Err, iExercicio, iPeriodo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150649)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PeriodosFilial_Insere_Exercicio(ByVal iExercicio As Integer, ByVal iNumPeriodos As Integer, ByVal colFiliais As Collection) As Long
'Insere os PeriodosFilial relativos a um exercicio
'ATENÇÃO: Esta rotina tem que ser chamada por um programa que esteja em transação

Dim lErro As Long
Dim lComando As Long
Dim iPeriodo As Integer
Dim objFiliais As AdmFiliais

On Error GoTo Erro_PeriodosFilial_Insere_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10209

    For iPeriodo = 1 To iNumPeriodos

        For Each objFiliais In colFiliais
    
            If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then
    
                'Insere o Periodo em questão
                lErro = Comando_Executar(lComando, "INSERT INTO PeriodosFilial (FilialEmpresa, Exercicio, Periodo, Lancamentos, Apurado, Lote, Fechado, DocApuracao) VALUES (?,?,?,?,?,?,?,?)", objFiliais.iCodFilial, iExercicio, iPeriodo, LANCAMENTOS_NO_BD, PERIODO_NAO_APURADO, 0, PERIODO_ABERTO, 0)
                If lErro <> AD_SQL_SUCESSO Then Error 10210
            
            End If
            
        Next

    Next

    Call Comando_Fechar(lComando)

    PeriodosFilial_Insere_Exercicio = SUCESSO

    Exit Function

Erro_PeriodosFilial_Insere_Exercicio:

    PeriodosFilial_Insere_Exercicio = Err

    Select Case Err

        Case 10209
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10210
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PERIODOSFILIAL", Err, objFiliais.iCodFilial, iExercicio, iPeriodo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150650)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function ExercicioOrigem_Insere_Exercicio(ByVal iExercicio As Integer, ByVal iNumPeriodos As Integer, ByVal colFiliais As Collection) As Long
'Pesquisa a tabela de origens e insere na tabela de ExercicioOrigem um registro para cada origem
'ATENÇÃO: Esta rotina tem que ser chamada por um programa que esteja em transação

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim sOrigem As String
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo
Dim iPeriodo As Integer
Dim objFiliais As AdmFiliais

On Error GoTo Erro_ExercicioOrigem_Insere_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9303

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9304

    sOrigem = String(STRING_ORIGEM, 0)

    'Pesquisa todas as origens
    lErro = Comando_Executar(lComando, "SELECT Origem FROM Origem", sOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 9305
    
    'le a primeira origem
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9306
    
    Do While lErro = AD_SQL_SUCESSO

        For iPeriodo = 0 To iNumPeriodos
        
            For Each objFiliais In gcolFiliais
        
                If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then
        
                    'insere um registro na tabela ExercicioOrigem
                    lErro = Comando_Executar(lComando1, "INSERT INTO ExercicioOrigem (FilialEmpresa, Exercicio, Periodo, Origem, Doc, Lote) VALUES (?,?,?,?,1,1)", objFiliais.iCodFilial, iExercicio, iPeriodo, sOrigem)
                    If lErro <> AD_SQL_SUCESSO Then Error 9307
                
                End If
                
            Next
            
        Next
        
        'le a proxima origem
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9308
            
    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ExercicioOrigem_Insere_Exercicio = SUCESSO

    Exit Function

Erro_ExercicioOrigem_Insere_Exercicio:

    ExercicioOrigem_Insere_Exercicio = Err

    Select Case Err

        Case 9303, 9304
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9305, 9306, 9308
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORIGEM", Err)

        Case 9307
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_EXERCICIOORIGEM", Err, iExercicio, iPeriodo, sOrigem)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150651)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Exercicio_Inicializacao(ByVal objExercicio As ClassExercicio, ByVal iPeriodo_Inicial As Integer, adtDataInicio() As Date, adtDataFim() As Date, asNomeExterno() As String, aiFechado() As Integer) As Long
'AVISO: ESSA ROTINA SO PODE SER CHAMADA COM A TRANSACAO ABERTA

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iStatus As Integer
Dim iCcl As Integer
Dim lNumReg As Long
Dim iIncluiSaldo As Integer

On Error GoTo Erro_Exercicio_Inicializacao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5142

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9835

    'apenas prepara o lock na tabela Configuracao
    lErro = Comando_ExecutarPos(lComando1, "SELECT UsoCcl FROM Configuracao", 0, iCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9836
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 9837
    
    'Lock da Configuracao, para não permitir a criacao simultanea de exercicios
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 9838

    lErro = CF("Exercicio_Le_NumRegistros", lNumReg)
    If lErro <> SUCESSO Then Error 9843

    If lNumReg = 0 Then
        iIncluiSaldo = 1
    Else
        iIncluiSaldo = 0
    End If

    'Pesquisa o Exercicio em questão
    lErro = Comando_Executar(lComando, "SELECT Status FROM Exercicios WHERE Exercicio = ?", iStatus, objExercicio.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5143

    'le o Exercicio
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9301

    'se o Exercicio já existir ==> erro
    If lErro = AD_SQL_SUCESSO Then Error 5144

    'Insere o Exercicio em questão
    lErro = Comando_Executar(lComando, "INSERT INTO Exercicios (Exercicio, NumPeriodos, Status, NomeExterno, DataInicio, DataFim) VALUES (?, ?, ?, ?, ?, ?)", objExercicio.iExercicio, objExercicio.iNumPeriodos, EXERCICIO_ABERTO, objExercicio.sNomeExterno, objExercicio.dtDataInicio, objExercicio.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then Error 5147
    
    lErro = ExerciciosFilial_Insere_Exercicio(objExercicio.iExercicio, gcolFiliais)
    If lErro <> SUCESSO Then Error 10221

    'inicializa os saldos de conta
    lErro = MvPerCta_Insere_Exercicio(objExercicio.iExercicio, iIncluiSaldo, gcolFiliais)
    If lErro <> SUCESSO Then Error 5146

    'inicializa os saldos de ccl
    lErro = MvPerCcl_Insere_Exercicio(objExercicio.iExercicio, iIncluiSaldo, gcolFiliais)
    If lErro <> SUCESSO Then Error 9359

    'Insere os periodos
    lErro = Periodo_Insere_Exercicio(objExercicio.iExercicio, objExercicio.iNumPeriodos, adtDataInicio(), adtDataFim(), asNomeExterno(), aiFechado())
    If lErro <> SUCESSO Then Error 5148
    
    'Insere os periodos das filiais
    lErro = CF("PeriodosFilial_Insere_Exercicio", objExercicio.iExercicio, objExercicio.iNumPeriodos, gcolFiliais)
    If lErro <> SUCESSO Then Error 10222
    
    'Insere exercicio origem
    lErro = ExercicioOrigem_Insere_Exercicio(objExercicio.iExercicio, objExercicio.iNumPeriodos, gcolFiliais)
    If lErro <> SUCESSO Then Error 9302

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exercicio_Inicializacao = SUCESSO

    Exit Function

Erro_Exercicio_Inicializacao:

    Exercicio_Inicializacao = Err

    Select Case Err

        Case 5142, 9835
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5143, 9301
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, objExercicio.iExercicio)

        Case 5144
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_JA_EXISTE", Err, objExercicio.iExercicio)

        Case 5146, 5148, 9302, 9359, 10221, 10222

        Case 5147
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_EXERCICIO", Err, objExercicio.iExercicio)

        Case 9836, 9837
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)

        Case 9838
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150652)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Exercicio_Reinicializacao(ByVal objExercicio1 As ClassExercicio, ByVal iPeriodo_Inicial As Integer, adtDataInicio() As Date, adtDataFim() As Date, asNomeExterno() As String, aiFechado() As Integer) As Long
'exclui os dados do exercicio e recria-o de acordo com os parametros
'AVISO: ESSA ROTINA DEVE SER CHAMADA COM A TRANSACAO JA ABERTA

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iStatus As Integer
Dim objExercicio2 As New ClassExercicio
Dim iIncluiSaldo As Integer

On Error GoTo Erro_Exercicio_Reinicializacao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9309

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9316

    'Pesquisa o Exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Status FROM Exercicios WHERE Exercicio = ?", 0, iStatus, objExercicio1.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 9310

    'le o Exercicio
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9311

    lErro = CF("Exercicio_Le_Primeiro", objExercicio2)
    If lErro <> SUCESSO Then Error 9849

    'se o exercicio em questão é o primeiro exercicio do sistema ==> atualiza os saldos iniciais de conta e centro de custo
    If objExercicio2.iExercicio = objExercicio1.iExercicio Then
        iIncluiSaldo = 1
    Else
        iIncluiSaldo = 0
    End If

    'Exclui os saldos de conta
    lErro = MvPerCta_Exclui_Exercicio(objExercicio1.iExercicio)
    If lErro <> SUCESSO Then Error 9313

    'Exclui todos os saldos de centro de custo/lucro associados com o Exercicio passado como parametro
    lErro = MvPerCcl_Exclui_Exercicio(objExercicio1.iExercicio)
    If lErro <> SUCESSO Then Error 11385

    'Exclui exercicio origem
    lErro = ExercicioOrigem_Exclui_Exercicio(objExercicio1.iExercicio)
    If lErro <> SUCESSO Then Error 9314

    'Exclui os periodos
    lErro = Periodo_Exclui_Exercicio(objExercicio1.iExercicio)
    If lErro <> SUCESSO Then Error 9315

    'Exclui os periodosFilial
    lErro = PeriodosFilial_Exclui_Exercicio(objExercicio1.iExercicio)
    If lErro <> SUCESSO Then Error 10196

    'Atualiza o Exercicio em questão
    lErro = Comando_ExecutarPos(lComando1, "UPDATE Exercicios Set NumPeriodos=?, Status=?, DataInicio = ? , DataFim = ?", lComando, objExercicio1.iNumPeriodos, EXERCICIO_ABERTO, objExercicio1.dtDataInicio, objExercicio1.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then Error 9317

    lErro = ExerciciosFilial_Reinicia_Exerc(objExercicio1.iExercicio)
    If lErro <> SUCESSO Then Error 10212

    'inicializa os saldos de conta
    lErro = MvPerCta_Insere_Exercicio(objExercicio1.iExercicio, iIncluiSaldo, gcolFiliais)
    If lErro <> SUCESSO Then Error 9318

    'inicializa os saldos de centro de custo/lucro
    lErro = MvPerCcl_Insere_Exercicio(objExercicio1.iExercicio, iIncluiSaldo, gcolFiliais)
    If lErro <> SUCESSO Then Error 11386
    
    'Insere os periodos
    lErro = Periodo_Insere_Exercicio(objExercicio1.iExercicio, objExercicio1.iNumPeriodos, adtDataInicio(), adtDataFim(), asNomeExterno(), aiFechado())
    If lErro <> SUCESSO Then Error 9320

    lErro = CF("PeriodosFilial_Insere_Exercicio", objExercicio1.iExercicio, objExercicio1.iNumPeriodos, gcolFiliais)
    If lErro <> SUCESSO Then Error 10211

    'Insere exercicio origem
    lErro = ExercicioOrigem_Insere_Exercicio(objExercicio1.iExercicio, objExercicio1.iNumPeriodos, gcolFiliais)
    If lErro <> SUCESSO Then Error 9321

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exercicio_Reinicializacao = SUCESSO

    Exit Function

Erro_Exercicio_Reinicializacao:

    Exercicio_Reinicializacao = Err

    Select Case Err

        Case 9309, 9316
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9310, 9311
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, objExercicio1.iExercicio)

        Case 9313, 9314, 9315, 9318, 9320, 9321, 10196, 10211, 10212, 11385, 11386
        
        Case 9317
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIO", Err, objExercicio1.iExercicio)

        Case 9849

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150653)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function ExerciciosFilial_Reinicia_Exerc(ByVal iExercicio As Integer) As Long
'Reinicializa os registros de exercicio de cada filial.
'ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_ExerciciosFilial_Reinicia_Exerc

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10214
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10215
    
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa FROM ExerciciosFilial WHERE Exercicio = ?", 0, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 10216

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10217

    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_ExecutarPos(lComando1, "UPDATE ExerciciosFilial SET LoteApuracao=?, DataApuracao=?, DocApuracao=?", lComando, 0, 0, 0)
        If lErro <> AD_SQL_SUCESSO Then Error 10218

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10219
        
    Loop
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    ExerciciosFilial_Reinicia_Exerc = SUCESSO
    
    Exit Function

Erro_ExerciciosFilial_Reinicia_Exerc:

    ExerciciosFilial_Reinicia_Exerc = Err
    
    Select Case Err
    
        Case 10214, 10215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10216, 10217, 10219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL1", Err, iExercicio)
    
        Case 10218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOSFILIAL", Err, iFilialEmpresa, iExercicio)
    
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
    
End Function

Function Orcamento_Exclui_Conta(ByVal sConta As String) As Long
'exclui os orcamentos associados a conta em questão
'ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim iPeriodo As Integer
Dim iExercicio As Integer

On Error GoTo Erro_Orcamento_Exclui_Conta

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 16010
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 16011

    'pesquisa os orcamentos para a conta em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Exercicio, Periodo FROM Orcamento WHERE Conta = ?", 0, iExercicio, iPeriodo, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 16012
        
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10213
     
    'exclui os orcamentos encontrados na pesquisa
    Do While lErro = AD_SQL_SUCESSO
        
        'exclui o orcamento
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM Orcamento", lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 16082
            
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10214
            
    Loop

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Orcamento_Exclui_Conta = SUCESSO
    
    Exit Function
    
Erro_Orcamento_Exclui_Conta:

    Orcamento_Exclui_Conta = Err
    
    Select Case Err
    
        Case 16010, 16011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10213, 10214, 16012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTO2", Err, sConta)
    
        Case 16082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTO1", Err, iExercicio, iPeriodo, sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150654)
    
    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function Orcamento_Exclui_Ccl(ByVal sCcl As String) As Long
'exclui os orcamentos associados ao centro de custo em questão
'ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim iPeriodo As Integer
Dim iExercicio As Integer
Dim sConta As String

On Error GoTo Erro_Orcamento_Exclui_Ccl

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10409
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10410

    sConta = String(STRING_CONTA, 0)

    'pesquisa os orcamentos para a conta em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Exercicio, Periodo, Conta FROM Orcamento WHERE Ccl = ?", 0, iExercicio, iPeriodo, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 10411
        
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10412
     
    'exclui os orcamentos encontrados na pesquisa
    Do While lErro = AD_SQL_SUCESSO
        
        'exclui o orcamento
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM Orcamento", lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 10413
            
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10414
            
    Loop

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Orcamento_Exclui_Ccl = SUCESSO
    
    Exit Function
    
Erro_Orcamento_Exclui_Ccl:

    Orcamento_Exclui_Ccl = Err
    
    Select Case Err
    
        Case 10409, 10410
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10411, 10412, 10414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTO3", Err, sCcl)
    
        Case 10413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTO1", Err, iExercicio, iPeriodo, sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150655)
    
    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function


Function Orcamento_Exclui_BD(ByVal objOrcamento As ClassOrcamento) As Long
'exclui os registros da tabela de Orcamento para a conta x exercicio em questão
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iPeriodo As Integer

On Error GoTo Erro_Orcamento_Exclui_BD

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16050
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 16051
    
    'selecao dos registros a serem excluidos
    lErro = Comando_ExecutarPos(lComando, "SELECT Periodo FROM Orcamento WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ? AND Ccl = ?", 0, iPeriodo, objOrcamento.iFilialEmpresa, objOrcamento.iExercicio, objOrcamento.sConta, objOrcamento.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 16072
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16073
    
    If lErro = AD_SQL_SEM_DADOS Then Error 64453
    
    Do While lErro = AD_SQL_SUCESSO
        
        'exclusao dos registros
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM Orcamento", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 16074
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16075
        
    Loop
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Orcamento_Exclui_BD = SUCESSO
    
    Exit Function
    
Erro_Orcamento_Exclui_BD:

    Orcamento_Exclui_BD = Err
    
    Select Case Err
    
        Case 16050, 16051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                    
        Case 16072, 16073, 16075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTO", Err)
                        
        Case 16074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTO_CTB", Err)
        
        Case 64453 'Orcamento não encontrado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150656)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
        
End Function

Function Orcamento_Insere_BD(ByVal colOrcamento As Collection) As Long
'grava na tabela de Orcamento os novos registros do orcamento em questão
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lErro As Long
Dim lComando As Long
Dim objOrcamento As ClassOrcamento

On Error GoTo Erro_Orcamento_Insere_BD

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16052
    
    For Each objOrcamento In colOrcamento
    
        'insere novos registros na tabela Orcamento
        lErro = Comando_Executar(lComando, "INSERT INTO Orcamento (FilialEmpresa, Conta, Exercicio, Ccl, Periodo, Valor) VALUES (?,?,?,?,?,?)", objOrcamento.iFilialEmpresa, objOrcamento.sConta, objOrcamento.iExercicio, objOrcamento.sCcl, objOrcamento.iPeriodo, objOrcamento.dValor)
        If lErro <> AD_SQL_SUCESSO Then Error 16053
    
    Next
    
    Call Comando_Fechar(lComando)
    
    Orcamento_Insere_BD = SUCESSO
    
    Exit Function
    
Erro_Orcamento_Insere_BD:

    Orcamento_Insere_BD = Err
    
    Select Case Err
    
        Case 16052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                    
        Case 16053
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ORCAMENTO", Err)
          
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150657)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Orcamento_Exclui(ByVal objOrcamento As ClassOrcamento) As Long
'Chama a rotina de exclusao de registros da tabela Orcamento, após obter a conta e o exercicio da tela

Dim lErro As Long
Dim iExercicio As Integer
Dim lComando As Long
Dim sConta As String
Dim lTransacao As Long
Dim iStatus As Integer

On Error GoTo Erro_Orcamento_Exclui

    iExercicio = objOrcamento.iExercicio
    sConta = objOrcamento.sConta

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16055

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 16056

    'Pesquisa registros da tabela Exercicios
    lErro = Comando_ExecutarLockado(lComando, "SELECT Status FROM Exercicios WHERE Exercicio = ?", iStatus, objOrcamento.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 16057

    'le o exercicio selecionado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 16058

    'loca o exercicio selecionado
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 16059

    'verifica se o exercicio não está fechado
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If iStatus = EXERCICIO_FECHADO Then Error 16060
    End If

    'exclui registros da tabela Orcamento
    lErro = CF("Orcamento_Exclui_BD", objOrcamento)
    If lErro <> SUCESSO And lErro <> 64453 Then Error 16063

    'Se não encontrou o Orcamento
    If lErro = 64453 Then Error 64454
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 16064

    Call Comando_Fechar(lComando)

    Orcamento_Exclui = SUCESSO

    Exit Function

Erro_Orcamento_Exclui:

    Orcamento_Exclui = Err

    Select Case Err

        Case 16055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 16057, 16058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, iExercicio)

        Case 16059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, iExercicio)

        Case 16060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_ABERTO", Err)

        Case 16063

        Case 16064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 64454
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTO_NAO_CADASTRADO", Err, objOrcamento.iExercicio, objOrcamento.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150658)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Orcamento_Grava(ByVal colOrcamento As Collection) As Long
'Grava os novos registros da tabela Orcamento, obtidos na tela de Orcamento

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim sConta As String
Dim iExercicio As Integer
Dim iStatus As Integer
Dim sConta1 As String
Dim objOrcamento1 As New ClassOrcamento

On Error GoTo Erro_Orcamento_Grava
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16034
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 16035
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10273
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 16036
    
    'recupera conta e exercicio passados como parametros
    Set objOrcamento1 = colOrcamento.Item(1)
    sConta = objOrcamento1.sConta
    iExercicio = objOrcamento1.iExercicio
    
    'Pesquisa registro do exercicio passado como parametro
    lErro = Comando_ExecutarLockado(lComando, "SELECT Status FROM Exercicios WHERE Exercicio = ?", iStatus, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 16043
    
    'le o exercicio selecionado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16037
    
    'Exercicio não cadastrado.
    If lErro = AD_SQL_SEM_DADOS Then Error 10274
    
    'loca o exercicio selecionado
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 16038
    
    'verifica se o exercicio não está fechado
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If iStatus = EXERCICIO_FECHADO Then Error 16039
    End If
    
    sConta1 = String(STRING_CONTA, 0)
    
    'Pesquisa registro da conta passada como parametro
    lErro = Comando_ExecutarLockado(lComando1, "SELECT Conta FROM PlanoConta WHERE Conta = ?", sConta1, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 16040
    
    'le a conta selecionada
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16041
    
    If lErro = AD_SQL_SEM_DADOS Then Error 10275
    
    'loca a conta selecionada
    lErro = Comando_LockShared(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 16042
    
    If giSetupUsoCcl = CCL_USA_EXTRACONTABIL And Len(objOrcamento1.sCcl) > 0 Then
        
        sConta1 = String(STRING_CONTA, 0)
        
        'Pesquisa registro da associacao conta x centro de cussto/lucro passada como parametro
        lErro = Comando_ExecutarLockado(lComando2, "SELECT Conta FROM ContaCcl WHERE Conta = ? AND Ccl=?", sConta1, objOrcamento1.sConta, objOrcamento1.sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 10276
    
        'le a associacao conta x centro de custo/lucro selecionada
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10277
    
        If lErro = AD_SQL_SEM_DADOS Then Error 10278
    
        'loca a associacao selecionada
        lErro = Comando_LockShared(lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 10279

    
    End If
    
    'exclui registros da tabela Orcamento
    lErro = CF("Orcamento_Exclui_BD", objOrcamento1)
    If lErro <> SUCESSO And lErro <> 64453 Then Error 16045
    
    'inclui registros na tabela Orcamento
    lErro = CF("Orcamento_Insere_BD", colOrcamento)
    If lErro <> SUCESSO Then Error 16046
    
    'finaliza a transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 16049
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Orcamento_Grava = SUCESSO
    
    Exit Function
    
Erro_Orcamento_Grava:

    Orcamento_Grava = Err
    
    
    Select Case Err
    
        Case 10273, 16034, 16035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 10274
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, objOrcamento1.iExercicio)
        
        Case 10275
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", Err, objOrcamento1.sConta)
        
        Case 10276, 10277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, objOrcamento1.sConta, objOrcamento1.sCcl)
        
        Case 10278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACCL_NAO_CADASTRADO", Err, objOrcamento1.sConta, objOrcamento1.sCcl)
        
        Case 10279
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTACCL", Err, objOrcamento1.sConta, objOrcamento1.sCcl)
        
        Case 16036
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
                       
        Case 16037, 16043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, iExercicio)
                        
        Case 16038
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, iExercicio)
                   
        Case 16039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_ABERTO", Err)
                    
        Case 16040, 16041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, sConta)
                       
        Case 16042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, sConta)
                       
        Case 16045, 16046
           
        Case 16049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
                       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150659)

    End Select
      
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
      
    Exit Function
    
End Function

Function RateioOff_Exclui(ByVal lCodigo As Long) As Long
'/m

Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_RateioOff_Exclui
              
    lTransacao = 0
                  
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 11305
   
    'exclui os lançamentos do rateio com codigo = icodigo
    lErro = CF("RateioOff_Exclui_BD", lCodigo)
    If lErro <> SUCESSO Then Error 11309
  
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 11310

    RateioOff_Exclui = SUCESSO

    Exit Function

Erro_RateioOff_Exclui:
    
    RateioOff_Exclui = Err
    
    Select Case Err
            
        Case 11305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 11310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150660)
   End Select
   
   Call Transacao_Rollback
   
   Exit Function

End Function

Function RateioOff_Grava(ByVal objRateioOff As ClassRateioOff, ByVal colRateioOff As Collection, ByVal colContas As Collection) As Long
'grava um rateio off-line
'/m

Dim lTransacao As Long
Dim lComando As Long
Dim iTipoConta As Integer
Dim lErro As Long
Dim iStatus As Integer
Dim iTipoCcl As Integer

On Error GoTo Erro_RateioOff_Grava
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11311
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 11312
    
    'pesquisa o Ccl em questão no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT TipoCcl FROM Ccl WHERE Ccl=?", 0, iTipoCcl, objRateioOff.sCclOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 11313

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11314
    
    If lErro = AD_SQL_SEM_DADOS Then Error 11315
    
    'lock do ccl
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 11316
    
    If iTipoCcl <> CCL_ANALITICA Then Error 10504
        
    lErro = Comando_ExecutarPos(lComando, "SELECT TipoConta, Ativo FROM PlanoConta WHERE Conta=?", 0, iTipoConta, iStatus, objRateioOff.sContaCre)
    If lErro <> AD_SQL_SUCESSO Then Error 10505

   'Lê a Conta Credito
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10506
    
    If lErro = AD_SQL_SEM_DADOS Then Error 10507
    
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10508
        
    'verifica se a conta credito está ativa
    If iStatus <> CONTA_ATIVA Then Error 10509
    
    'verifica se a conta credito é analítica
    If iTipoConta <> CONTA_ANALITICA Then Error 10510
    
    'Exclui os rateios com essas carcteristicas do BD
    lErro = CF("RateioOff_Exclui_BD", objRateioOff.lCodigo)
    If lErro <> SUCESSO Then Error 11323

    'Grava todos os rateio contidos em colRateioOff
    lErro = CF("RateioOff_Grava_BD", objRateioOff, colRateioOff)
    If lErro <> SUCESSO Then Error 11324
    
    'Grava todos as contas origem contidos em colRateioOffContas
    lErro = RateioOff_Grava_BD1(objRateioOff, colContas)
    If lErro <> SUCESSO Then Error 55802
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 11325

    Call Comando_Fechar(lComando)
   
    RateioOff_Grava = SUCESSO

    Exit Function

Erro_RateioOff_Grava:
    
    RateioOff_Grava = Err
    
    Select Case Err
    
        Case 10504
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_ANALITICA1", Err, objRateioOff.sCclOrigem)
    
        Case 10505, 10506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, objRateioOff.sContaCre)
    
        Case 10507
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, objRateioOff.sContaCre)
    
        Case 10508
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, objRateioOff.sContaCre)
        
        Case 10509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", Err, objRateioOff.sContaCre)
        
        Case 10510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, objRateioOff.sContaCre)
        
        Case 11311
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 11312
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 11313, 11314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err, objRateioOff.sCclOrigem)
                    
        Case 11315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objRateioOff.sCclOrigem)
            
        Case 11316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", Err, objRateioOff.sCclOrigem)
            
        Case 11323, 11324, 55802

        Case 11325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150661)
            
    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RateioOff_Grava_BD(ByVal objRateioOff1 As ClassRateioOff, ByVal colRateioOff As Collection) As Long
'grava na tabela de Rateios Automaticos (RateioOff)lançamentos do voucher em questão /m
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim iTipoConta As Integer
'Dim sCcl As String
Dim sConta As String
Dim iPossuiCcl As Integer
Dim iStatus As Integer
Dim objRateioOff As ClassRateioOff
Dim iTipoCcl As Integer

On Error GoTo Erro_RateioOff_Grava_BD

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11326
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 11327
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 20548
    
    'Inicia Loop de GRAVACAO com Lançamentos do GRID
    For Each objRateioOff In colRateioOff
                
        'Pesquisa Conta associada a um Lançamento
        lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta, Ativo FROM PlanoConta WHERE Conta=?", iTipoConta, iStatus, objRateioOff.sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 11328
        
       'Lê a Conta
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11329
        
        If lErro = AD_SQL_SEM_DADOS Then Error 11330
        
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 11331
            
        'verifica se a conta está ativa
        If iStatus <> CONTA_ATIVA Then Error 11332
               
        'verifica se a conta é analítica
        If iTipoConta <> CONTA_ANALITICA Then Error 11333
        
        If giSetupUsoCcl = CCL_USA_EXTRACONTABIL And Len(objRateioOff.sCcl) > 0 Then
        
            sConta = String(STRING_CONTA, 0)
        
            'Pesquisa a associação Conta x Ccl
            lErro = Comando_ExecutarLockado(lComando2, "SELECT Conta FROM ContaCcl WHERE Conta=? AND Ccl=?", sConta, objRateioOff.sConta, objRateioOff.sCcl)
            If lErro <> AD_SQL_SUCESSO Then Error 20571
        
            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20572
            
            If lErro = AD_SQL_SEM_DADOS Then Error 20573
            
            'lock da associacao conta x ccl
            lErro = Comando_LockExclusive(lComando2)
            If lErro <> AD_SQL_SUCESSO Then Error 20574
            
            lErro = CF("Ccl_Critica_Analitica", objRateioOff.sCcl)
            If lErro <> SUCESSO Then Error 32315
        
        End If
        
        'Faz inserção do Lançamento na Tabela de RateioOff
        lErro = Comando_Executar(lComando1, "INSERT INTO RateioOff (Codigo, CclOrigem, Descricao, Seq, ContaCre, Tipo, Conta, Ccl, Percentual ) VALUES (?,?,?,?,?,?,?,?,?)", objRateioOff1.lCodigo, objRateioOff1.sCclOrigem, objRateioOff1.sDescricao, objRateioOff.lSeq, objRateioOff1.sContaCre, objRateioOff1.iTipo, objRateioOff.sConta, objRateioOff.sCcl, objRateioOff.dPercentual)
        If lErro <> AD_SQL_SUCESSO Then Error 11334
        
    Next

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    RateioOff_Grava_BD = SUCESSO
    
    Exit Function
    
Erro_RateioOff_Grava_BD:

    RateioOff_Grava_BD = Err
    
    Select Case Err
    
        Case 32315
        
        Case 11326, 11327, 20548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 11328, 11329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, objRateioOff.sConta)
    
        Case 11330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, objRateioOff.sConta)

        Case 11331
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, objRateioOff.sConta)
    
        Case 11332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", Err)
    
        Case 11333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, objRateioOff.sConta)
        
        Case 11334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RATEIOOFF", Err)
    
        Case 20571, 20572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, objRateioOff.sConta, objRateioOff.sCcl)
    
        Case 20573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACCL_NAO_CADASTRADO", Err, objRateioOff.sConta, objRateioOff.sCcl)
    
        Case 20574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTACCL", Err, objRateioOff.sConta, objRateioOff.sCcl)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150662)
                
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Private Function RateioOff_Grava_BD1(ByVal objRateioOff1 As ClassRateioOff, ByVal colContas As Collection) As Long
'grava na tabela de Rateios Off line a faixa de contas origem do rateio
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iTipoConta As Integer
Dim sCcl As String
Dim sConta As String
Dim objRateioOffContas As ClassRateioOffContas
Dim iStatus As Integer
Dim iIndice As Integer

On Error GoTo Erro_RateioOff_Grava_BD1

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55787
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 55788
    
    'Inicia Loop de GRAVACAO com Lançamentos do GRID
    For Each objRateioOffContas In colContas
                
        iIndice = iIndice + 1
                
        'Pesquisa Conta Inicio
        lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta, Ativo FROM PlanoConta WHERE Conta=?", iTipoConta, iStatus, objRateioOffContas.sContaInicio)
        If lErro <> AD_SQL_SUCESSO Then Error 55789
        
       'Lê a Conta
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55790
        
        If lErro = AD_SQL_SEM_DADOS Then Error 55791
        
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 55792
            
        'verifica se a conta está ativa
        If iStatus <> CONTA_ATIVA Then Error 55793
               
        'verifica se a conta é analítica
        If iTipoConta <> CONTA_ANALITICA Then Error 55794
        
        'Pesquisa Conta Inicio
        lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta, Ativo FROM PlanoConta WHERE Conta=?", iTipoConta, iStatus, objRateioOffContas.sContaFim)
        If lErro <> AD_SQL_SUCESSO Then Error 55795
        
       'Lê a Conta
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55796
        
        If lErro = AD_SQL_SEM_DADOS Then Error 55797
        
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 55798
            
        'verifica se a conta está ativa
        If iStatus <> CONTA_ATIVA Then Error 55799
               
        'verifica se a conta é analítica
        If iTipoConta <> CONTA_ANALITICA Then Error 55800
        
        'Faz inserção do Lançamento na Tabela de RateioOff
        lErro = Comando_Executar(lComando1, "INSERT INTO RateioOffContas (Codigo, Item, ContaInicial, ContaFinal) VALUES (?,?,?,?)", objRateioOff1.lCodigo, iIndice, objRateioOffContas.sContaInicio, objRateioOffContas.sContaFim)
        If lErro <> AD_SQL_SUCESSO Then Error 55801
        
    Next

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    RateioOff_Grava_BD1 = SUCESSO
    
    Exit Function
    
Erro_RateioOff_Grava_BD1:

    RateioOff_Grava_BD1 = Err
    
    Select Case Err
    
        Case 55787, 55788
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 55789, 55790
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, objRateioOffContas.sContaInicio)
    
        Case 55791
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, objRateioOffContas.sContaInicio)

        Case 55792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, objRateioOffContas.sContaInicio)
    
        Case 55793
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", Err, objRateioOffContas.sContaInicio)
    
        Case 55794
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, objRateioOffContas.sContaInicio)
        
        Case 55795, 55796
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", Err, objRateioOffContas.sContaFim)
    
        Case 55797
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, objRateioOffContas.sContaFim)

        Case 55798
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, objRateioOffContas.sContaFim)
    
        Case 55799
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INATIVA", Err, objRateioOffContas.sContaFim)
    
        Case 55800
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA", Err, objRateioOffContas.sContaFim)
        
        Case 55801
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RATEIOOFFCONTAS", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150663)
                
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function RateioOff_Exclui_BD(ByVal lCodigo As Long) As Long
'Exclui da tabela RateioOff os Rateios com CclOrigem= sCclOrigem e ContaOrigem= sContaOrigem /m
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iSeq As Integer
Dim iTipo As Integer
Dim iItem As Integer

On Error GoTo Erro_RateioOff_Exclui_BD

    'Abre comando para pesquisar Lancamentos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11336
    
    'Abre comando para transação de exclusão
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 11337
      
    'Pesquisa a existencia dos Rateios no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT Tipo FROM RateioOff WHERE Codigo = ?", 0, iTipo, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 11338
        
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11339

    'Inicia Loop de EXCLUSAO com Rateios do GRID
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Exclui a conta origem
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM RateioOff", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 11340
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11341
    
    Loop
   
    'Pesquisa a existencia das Contas Origem no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT Item FROM RateioOffContas WHERE Codigo = ?", 0, iItem, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 55803
        
    'Lê o primeiro Rateio
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55804
   
    'Inicia Loop de EXCLUSAO das Contas Origem do Rateio
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Exclui os Rateios com icodigo da tabela RateioOff
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM RateioOffContas", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 55805
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55806
    
    Loop
   
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    RateioOff_Exclui_BD = SUCESSO
    
    Exit Function
    
Erro_RateioOff_Exclui_BD:

    RateioOff_Exclui_BD = Err
    
    Select Case Err
    
        Case 11336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 11337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 11338, 11339, 11341
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFF1", Err)
            
        Case 11340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RATEIOOFF", Err)
            
        Case 55803, 55804, 55806
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFFCONTAS", Err, lCodigo)
            
        Case 55805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RATEIOOFFCONTAS", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150664)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
                   
End Function

Function LoteApuracao_Automatico(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, iLote As Integer) As Long
'retorna o número do proximo lote disponivel

Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim iLote1 As Integer
Dim iLote2 As Integer
Dim lTransacao As Long
Dim sOrigem As String
Dim objLote As New ClassLote

On Error GoTo Erro_LoteApuracao_Automatico

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11624

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 11625

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 11626

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 11627

    sOrigem = "APE"

    'Pesquisa o número do lote gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT LoteApuracao FROM ExerciciosFilial WHERE FilialEmpresa=? AND Exercicio = ?", 0, iLote, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 11628

    'le o registro com o número do lote automático
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 11629

    'loca o registro que guarda o numero do lote gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 11630

    lErro = Comando_Executar(lComando3, "SELECT Lote FROM LotePendente WHERE FilialEmpresa=? AND Origem = ? AND Exercicio = ? AND Lote >= ? ORDER BY Lote", iLote2, iFilialEmpresa, sOrigem, iExercicio, iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 11631

    'le o primeiro lote pendente >= ao lote automatico
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11632

    'pesquisa o proximo número de lote disponível
    Do While lErro <> AD_SQL_SEM_DADOS

        If iLote < iLote2 Then

            objLote.iFilialEmpresa = iFilialEmpresa
            objLote.iExercicio = iExercicio
            objLote.sOrigem = sOrigem
            objLote.iLote = iLote

            'pesquisa se o número já foi utilizado em um lote contabilizado
            lErro = CF("Lote_Le", objLote)
            If lErro <> SUCESSO And lErro <> 9293 Then Error 11633 '********

            'Se ainda não foi usado ==> este é o número a ser usado
            If lErro = 9293 Then Exit Do

            'Se o número já foi usado ==> pega o próximo
            iLote = iLote + 1
        Else
            iLote = iLote + 1
            lErro = Comando_BuscarProximo(lComando3)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11634

        End If

    Loop

    'atualiza o numero do lote gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE ExerciciosFilial SET LoteApuracao = ?", lComando, iLote + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 11635

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 9925
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    LoteApuracao_Automatico = SUCESSO

    Exit Function

Erro_LoteApuracao_Automatico:

    LoteApuracao_Automatico = Err

    Select Case Err

        Case 9925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 11624, 11625, 11626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 11627
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 11628, 11629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL", Err, iFilialEmpresa, iExercicio)

        Case 11630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOSFILIAL", Err, iFilialEmpresa, iExercicio)

        Case 11631, 11632, 11634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE1", Err)

        Case 11635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOSFILIAL", Err, iFilialEmpresa, iExercicio)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150665)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function LotePendente_Atualiza(ByVal colLote As Collection, iIdAtualizacao As Integer) As Long
'atualiza o IdAtualizacao das tabelas Configuracao e LotePendente

Dim lErro As Long
Dim objLote As ClassLote
Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lTransacao As Long
Dim tLote As typeLote
Dim iTotaisIguais As Integer
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_LotePendente_Atualiza

    tLote.sOrigem = String(STRING_ORIGEM, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12216
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 12217
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 12218
    
    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 12219
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 12220
    
    'apenas prepara o lock na tabela Configuracao
    lErro = Comando_ExecutarPos(lComando, "SELECT IdAtualizacao FROM Configuracao", 0, iIdAtualizacao)
    If lErro <> AD_SQL_SUCESSO Then Error 12221
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12222
    
    'Lock do registro de configuracao
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 12223
           
    'faz a alteracao em IdAtualizacao na tabela de Configuracao
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Configuracao SET IdAtualizacao =? ", lComando, iIdAtualizacao + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 12224
                     
    For Each objLote In colLote
      
        Do While objLote.iFilialEmpresa > 0 And objLote.iFilialEmpresa < 100
      
            'apenas prepara o lock na tabela LotePendente
            lErro = Comando_ExecutarPos(lComando3, "SELECT IdAtualizacao, Status FROM LotePendente WHERE FilialEmpresa=? AND Origem = ? AND Exercicio = ? AND Periodo=? AND Lote=?", 0, tLote.iIdAtualizacao, tLote.iStatus, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            If lErro <> AD_SQL_SUCESSO Then Error 12225
          
            lErro = Comando_BuscarPrimeiro(lComando3)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12236
        
            If lErro = AD_SQL_SUCESSO Then
        
                'Lock do registro de LotePendente
                lErro = Comando_LockExclusive(lComando3)
                If lErro <> AD_SQL_SUCESSO Then Error 12226
            
                'se o lote está pronto para ser atualizado.
                If tLote.iStatus = LOTE_DESATUALIZADO Then
                
                    objLote.dTotCre = objLote.dTotInf
                    objLote.dTotDeb = objLote.dTotInf
                    objLote.iNumDocAtual = objLote.iNumDocInf
                
                    'Verifica a origem do lote
                    Select Case objLote.sOrigem
                    
                        'Se for um lote originado por um dos módulos (diferente de batch)
                        Case MODULO_CONTABILIDADE, MODULO_COMPRAS, MODULO_CONTASAPAGAR, MODULO_CONTASARECEBER, MODULO_ESTOQUE, MODULO_FATURAMENTO, MODULO_TESOURARIA
                            'Testas os totais do lote
                            lErro = CF("LanPendente_Critica_TotaisLote", objLote, iTotaisIguais)
                            If lErro <> SUCESSO Then Error 60791
                        
                    End Select
                
                    'Se os totais informados forem diferentes dos totais calculados
                    If iTotaisIguais = 1 Then
                        
                        'totais diferentes
                        vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_ATUALIZACAO_LOTE_TOTAIS_DIFERENTES", objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote, objLote.iNumDocInf, objLote.iNumDocAtual, objLote.dTotInf, objLote.dTotCre)
                        If vbMsgRet = vbYes Then
                        
                            lErro = LotePendente_Atualiza1(tLote, objLote, lComando3, lComando4, iIdAtualizacao)
                            If lErro <> SUCESSO Then Error 60793
                            
                        End If
                        
                    Else
                            
                        lErro = LotePendente_Atualiza1(tLote, objLote, lComando3, lComando4, iIdAtualizacao)
                        If lErro <> SUCESSO Then Error 60794
                            
                    End If
                    
                End If
            
            End If
        
            If giContabGerencial = 0 Then Exit Do
            
            objLote.iFilialEmpresa = objLote.iFilialEmpresa - giFilialAuxiliar
            
        Loop
            
        
    Next
      
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 12228
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    LotePendente_Atualiza = SUCESSO
    
    Exit Function

Erro_LotePendente_Atualiza:

    LotePendente_Atualiza = Err

    Select Case Err

        Case 12216, 12217, 12218, 12219
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
             
        Case 12220
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 12221, 12222
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
            
        Case 12223
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)
            
        Case 12224
             lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CONFIG", Err, iIdAtualizacao)
            
        Case 12225
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE_PENDENTE", Err, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            
        Case 12226
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LOTEPENDENTE", Err)
            
        Case 12228
             lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)

        Case 60791, 60793, 60794

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150666)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    Exit Function

End Function

Private Function LotePendente_Atualiza1(tLote As typeLote, ByVal objLote As ClassLote, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal iIdAtualizacao As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_LotePendente_Atualiza1

    'verifica se ja tem um numero de atualizacao
    If tLote.iIdAtualizacao <> 0 Then
    
        'se na tela o lote já estava marcardo "Atualizando"
        If objLote.iStatus = LOTE_ATUALIZANDO Then

            'faz o update na tabela de LotePendente
            lErro = Comando_ExecutarPos(lComando4, "UPDATE LotePendente SET IdAtualizacao =? ", lComando3, iIdAtualizacao)
            If lErro <> AD_SQL_SUCESSO Then Error 12227
            
        End If
           

    Else
    
        'se o lote ainda não possui um número de atualização (ou seja, não está em processo de atualização).

        'faz o update na tabela de LotePendente
        lErro = Comando_ExecutarPos(lComando4, "UPDATE LotePendente SET IdAtualizacao =? ", lComando3, iIdAtualizacao)
        If lErro <> AD_SQL_SUCESSO Then Error 60792
      
    End If
        
    LotePendente_Atualiza1 = SUCESSO
    
    Exit Function
    
Erro_LotePendente_Atualiza1:

    LotePendente_Atualiza1 = Err

    Select Case Err
    
        Case 12227, 60792
             lErro = Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_LOTEPENDENTE", Err, iIdAtualizacao)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150667)

    End Select
    
    Exit Function

End Function



Function ContaCategoria_Exclui(ByVal iCodigo As Integer) As Long
'exclui a Categoria

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim iApuracao As Integer

On Error GoTo Erro_ContaCategoria_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9656
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9657
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 9658
    
    'pesquisa a Categoria em questão no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT Apuracao FROM ContaCategoria WHERE Codigo=?", 0, iApuracao, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 9659
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9660
        
    If lErro = AD_SQL_SEM_DADOS Then Error 9661
    
    'lock da categoria
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9662

    'Verifica se a categoria é usada no Plano de Contas
    lErro = CF("PlanoContas_Le_Categoria", iCodigo)
    If lErro <> SUCESSO And lErro <> 9666 Then Error 9667
    
    'Se a categoria é usada no Plano de Contas
    If lErro = SUCESSO Then Error 9668

    'exclui a categoria do BD
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ContaCategoria", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9669
    
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 9670

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ContaCategoria_Exclui = SUCESSO
    
    Exit Function
    
Erro_ContaCategoria_Exclui:

    ContaCategoria_Exclui = Err
    
    Select Case Err
    
        Case 9656, 9657
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9658
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 9659, 9660
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACATEGORIA", Err, iCodigo)
            
        Case 9661
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", Err, iCodigo)
        
        Case 9662
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTACATEGORIA", Err, iCodigo)
        
        Case 9667
        
        Case 9668
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_PRESENTE_PLANO_CONTAS", Err, iCodigo)
            
        Case 9669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACATEGORIA", Err, iCodigo)
        
        Case 9670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150668)
    
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function ContaCategoria_Grava(ByVal objContaCategoria As ClassContaCategoria) As Long
'grava a Categoria no banco de dados

Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iApuracao As Integer

On Error GoTo Erro_ContaCategoria_Grava

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9688

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 9689

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 9690

    'pesquisa no BD a Categoria em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Apuracao FROM ContaCategoria WHERE Codigo=?", 0, iApuracao, objContaCategoria.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 9691

    'le a Categoria, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9692

    'encontrou a Categoria ==> isto é uma alteração
    If lErro = AD_SQL_SUCESSO Then
    
        'atualiza a Categoria
        lErro = Comando_ExecutarPos(lComando2, "UPDATE ContaCategoria SET Nome=?, Apuracao=?", lComando1, objContaCategoria.sNome, objContaCategoria.iApuracao)
        If lErro <> AD_SQL_SUCESSO Then Error 9693

    Else
        'não encontrou a Categoria ==>  é uma inserção de Categoria
        
        'insere a Categoria
        lErro = Comando_Executar(lComando2, "INSERT INTO ContaCategoria (Codigo, Nome, Apuracao) VALUES (?,?,?)", objContaCategoria.iCodigo, objContaCategoria.sNome, objContaCategoria.iApuracao)
        If lErro <> AD_SQL_SUCESSO Then Error 9694

    End If

    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 9695

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando1)

    ContaCategoria_Grava = SUCESSO

    Exit Function

Erro_ContaCategoria_Grava:

    ContaCategoria_Grava = Err

    Select Case Err

        Case 9688, 9689
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9690
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 9691, 9692
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACATEGORIA", Err, objContaCategoria.iCodigo)

        Case 9693
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONTACATEGORIA", Err, objContaCategoria.iCodigo)

        Case 9694
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONTACATEGORIA", Err, objContaCategoria.iCodigo)

        Case 9695
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150669)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function ContaCategoria_Automatico(iCodigo As Integer) As Long
'retorna o proximo codigo de Categoria disponivel

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim iCodigo1 As Integer
Dim lTransacao As Long

On Error GoTo Erro_ContaCategoria_Automatico

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9703
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9704
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 9705
   
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 9706

    'Pesquisa o proximo codigo da categoria gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT Categoria FROM Configuracao", 0, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 9707

    'le o registro com o codigo de categoria automático
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9708
    
    'loca o registro que guarda o codigo da categoria gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9709

    'verifica se este número ainda não foi utilizado
    'Pesquisa os codigos de categoria maiores ou iguais a numeração gerada automaticamente
    lErro = Comando_Executar(lComando1, "SELECT Codigo FROM ContaCategoria WHERE Codigo >= ? ORDER BY Codigo", iCodigo1, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 9710
        
    'le o primeiro codigo de categoria >= ao codigo de categoria automatico
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9711
    
    'pesquisa o proximo codigo de Categoria disponível
    Do While lErro <> AD_SQL_SEM_DADOS
        
        If iCodigo < iCodigo1 Then Exit Do
        
        If iCodigo = iCodigo1 Then
            iCodigo = iCodigo + 1
        End If
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9712
        
    Loop

    'atualiza o codigo da categoria gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Configuracao SET Categoria = ?", lComando, iCodigo + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 9713
        
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 9714
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    ContaCategoria_Automatico = SUCESSO
        
    Exit Function
    
    
Erro_ContaCategoria_Automatico:

    ContaCategoria_Automatico = Err

    Select Case Err
    
        Case 9703, 9704, 9705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 9706
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 9707, 9708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
            
        Case 9709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)
            
        Case 9710, 9711, 9712
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACATEGORIA1", Err)
            
        Case 9713
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONFIGURACAO", Err)
            
        Case 9714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150670)
        
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function
    
End Function

Function SaldoInicialConta_Altera_Pais(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal dSldIni As Double, Optional ByVal vIncluiConta As Variant) As Long
'Altera os saldos iniciais das contas pais de uma conta analitica. Se vIncluiConta estiver preenchido ==> inclui também a conta que está sendo passada como parametro. Altera somente os saldos da filial em questao.
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO


Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lErro As Long
Dim iNivel As Integer
Dim sConta1 As String
Dim sConta2 As String
Dim iExercicio As Integer

On Error GoTo Erro_SaldoInicialConta_Altera_Pais

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9764

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 9765

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 9825

    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 9826

    lErro = Mascara_Conta_ObterNivel(sConta, iNivel)
    If lErro <> SUCESSO Then Error 9766
    
    If Not IsMissing(vIncluiConta) Then iNivel = iNivel + 1
    
    Do While iNivel > 1
    
        iNivel = iNivel - 1
        
        sConta1 = String(STRING_CONTA, 0)
                
        'pega o pai da conta em questao
        lErro = Mascara_RetornaContaNoNivel(iNivel, sConta, sConta1)
        If lErro <> SUCESSO Then Error 9767
    
        sConta2 = String(STRING_CONTA, 0)
    
        'seleciona a conta a ser alterada
        lErro = Comando_ExecutarPos(lComando1, "SELECT Conta FROM SaldoInicialConta WHERE FilialEmpresa=? AND Conta = ?", 0, sConta2, iFilialEmpresa, sConta1)
        If lErro <> AD_SQL_SUCESSO Then Error 9768
        
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 9769
    
        'altera a conta na tabela SaldoInicialConta
        lErro = Comando_ExecutarPos(lComando2, "UPDATE SaldoInicialConta SET SldIni= SldIni + ?", lComando1, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 9770
        
        'seleciona o saldo da conta a ser alterado
        lErro = Comando_ExecutarPos(lComando3, "SELECT Exercicio FROM MvPerCta WHERE FilialEmpresa=? AND Conta = ? ORDER BY Exercicio", 0, iExercicio, iFilialEmpresa, sConta1)
        If lErro <> AD_SQL_SUCESSO Then Error 9827
        
        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 9828
    
        'altera a conta na tabela MvPerCta
        lErro = Comando_ExecutarPos(lComando4, "UPDATE MvPerCta SET SldIni= SldIni + ?", lComando3, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 9829
        
    Loop
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    SaldoInicialConta_Altera_Pais = SUCESSO

    Exit Function

Erro_SaldoInicialConta_Altera_Pais:
    
    SaldoInicialConta_Altera_Pais = Err
    
    Select Case Err
    
        Case 9764, 9765, 9825, 9826
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 9766
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)
            
        Case 9767
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, sConta, iNivel)
        
        Case 9768, 9769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA", Err, iFilialEmpresa, sConta1)
        
        Case 9770
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTA", Err, iFilialEmpresa, sConta1)
         
        Case 9827, 9828
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA2", Err, sConta1)
        
        Case 9829
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCTA", Err, iExercicio, sConta1)
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150671)
        
    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    Exit Function

End Function

Function SaldoInicialConta_Altera_Pais1(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal dSldIni As Double, Optional ByVal vIncluiConta As Variant) As Long
'Altera os saldos iniciais das contas pais de uma conta analitica. Se vIncluiConta estiver preenchido ==> inclui também a conta que está sendo passada como parametro. Altera os saldos da filial em questão e da empresa toda.
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO


Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lErro As Long
Dim iNivel As Integer
Dim sConta1 As String
Dim sConta2 As String
Dim iExercicio As Integer

On Error GoTo Erro_SaldoInicialConta_Altera_Pais1

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10542

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10543

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10544

    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 10545

    lErro = Mascara_Conta_ObterNivel(sConta, iNivel)
    If lErro <> SUCESSO Then Error 10546
    
    If Not IsMissing(vIncluiConta) Then iNivel = iNivel + 1
    
    Do While iNivel > 1
    
        iNivel = iNivel - 1
        
        sConta1 = String(STRING_CONTA, 0)
                
        'pega o pai da conta em questao
        lErro = Mascara_RetornaContaNoNivel(iNivel, sConta, sConta1)
        If lErro <> SUCESSO Then Error 10547
    
        sConta2 = String(STRING_CONTA, 0)
    
        'seleciona a conta a ser alterada
        lErro = Comando_ExecutarPos(lComando1, "SELECT Conta FROM SaldoInicialConta WHERE (FilialEmpresa=? OR FilialEmpresa = ?) AND Conta = ?", 0, sConta2, iFilialEmpresa, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta1)
        If lErro <> AD_SQL_SUCESSO Then Error 10548
        
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 10549
    
        'altera a conta na tabela SaldoInicialConta
        lErro = Comando_ExecutarPos(lComando2, "UPDATE SaldoInicialConta SET SldIni= SldIni + ?", lComando1, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10550
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 10551
    
        'altera a conta na tabela SaldoInicialConta
        lErro = Comando_ExecutarPos(lComando2, "UPDATE SaldoInicialConta SET SldIni= SldIni + ?", lComando1, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10552
        
        'seleciona o saldo da conta a ser alterado
        lErro = Comando_ExecutarPos(lComando3, "SELECT Exercicio FROM MvPerCta WHERE (FilialEmpresa=? OR FilialEmpresa=?) AND Conta = ? ORDER BY Exercicio", 0, iExercicio, iFilialEmpresa, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta1)
        If lErro <> AD_SQL_SUCESSO Then Error 10553
        
        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 10554
    
        'altera a conta na tabela MvPerCta
        lErro = Comando_ExecutarPos(lComando4, "UPDATE MvPerCta SET SldIni= SldIni + ?", lComando3, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10555
        
        lErro = Comando_BuscarProximo(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 10556
    
        'altera a conta na tabela MvPerCta
        lErro = Comando_ExecutarPos(lComando4, "UPDATE MvPerCta SET SldIni= SldIni + ?", lComando3, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10557
        
    Loop
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    SaldoInicialConta_Altera_Pais1 = SUCESSO

    Exit Function

Erro_SaldoInicialConta_Altera_Pais1:
    
    SaldoInicialConta_Altera_Pais1 = Err
    
    Select Case Err
    
        Case 10542, 10543, 10544, 10545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 10546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)
            
        Case 10547
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, sConta, iNivel)
        
        Case 10548, 10549, 10551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA", Err, iFilialEmpresa, sConta1)
        
        Case 10550, 10552
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTA", Err, iFilialEmpresa, sConta1)
         
        Case 10553, 10554, 10556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA2", Err, sConta1)
        
        Case 10555, 10557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCTA", Err, iExercicio, sConta1)
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150672)
        
    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    Exit Function

End Function

Function SaldoInicialContaCcl_Altera_Pais(ByVal iFilialEmpresa As Integer, ByVal sCcl As String, ByVal dSldIni As Double, Optional ByVal vIncluiCcl As Variant) As Long
'Altera os saldos iniciais dos centros de custo/lucro pais de um ccl. Altera somente a filial em questao
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO


Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lErro As Long
Dim iNivel As Integer
Dim sCcl1 As String
Dim sCcl2 As String
Dim iExercicio As Integer

On Error GoTo Erro_SaldoInicialContaCcl_Altera_Pais

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10305

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10306

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10307

    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 10308

    lErro = Mascara_Ccl_ObterNivel(sCcl, iNivel)
    If lErro <> SUCESSO Then Error 10309
    
    If Not IsMissing(vIncluiCcl) Then iNivel = iNivel + 1
    
    Do While iNivel > 1
    
        iNivel = iNivel - 1
        
        sCcl1 = String(STRING_CCL, 0)
                
        'pega o pai da conta em questao
        lErro = Mascara_RetornaCclNoNivel(iNivel, sCcl, sCcl1)
        If lErro <> SUCESSO Then Error 10310
    
        sCcl2 = String(STRING_CCL, 0)
    
        'seleciona o centro de custo a ser alterada
        lErro = Comando_ExecutarPos(lComando1, "SELECT Ccl FROM SaldoInicialContaCcl WHERE FilialEmpresa=? AND Ccl = ? AND Conta = ?", 0, sCcl2, iFilialEmpresa, sCcl1, "")
        If lErro <> AD_SQL_SUCESSO Then Error 10311
        
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 10312
    
        'altera o centro de custo
        lErro = Comando_ExecutarPos(lComando2, "UPDATE SaldoInicialContaCcl SET SldIni= SldIni + ?", lComando1, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10313
        
        'seleciona o saldo do centro de custo a ser alterado
        lErro = Comando_ExecutarPos(lComando3, "SELECT Exercicio FROM MvPerCcl WHERE FilialEmpresa=? AND Ccl = ? AND Conta = ? ORDER BY Exercicio", 0, iExercicio, iFilialEmpresa, sCcl1, "")
        If lErro <> AD_SQL_SUCESSO Then Error 10314
        
        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 10315
    
        'altera o saldo do centro de custo
        lErro = Comando_ExecutarPos(lComando4, "UPDATE MvPerCcl SET SldIni= SldIni + ?", lComando3, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10316
        
    Loop
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    SaldoInicialContaCcl_Altera_Pais = SUCESSO

    Exit Function

Erro_SaldoInicialContaCcl_Altera_Pais:
    
    SaldoInicialContaCcl_Altera_Pais = Err
    
    Select Case Err
    
        Case 10305, 10306, 10307, 10308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 10309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CCL_OBTERNIVEL", Err, sCcl)
            
        Case 10310
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclNoNivel", Err, sCcl, iNivel)
        
        Case 10311, 10312
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL2", Err, iFilialEmpresa, sCcl1)
        
        Case 10313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTACCL", Err, iFilialEmpresa, sCcl1)
         
        Case 10314, 10315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL3", Err, sCcl1)
        
        Case 10316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL3", Err, iFilialEmpresa, iExercicio, sCcl1)
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150673)
        
    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    Exit Function

End Function

Function SaldoInicialContaCcl_Altera_Pais1(alComando() As Long, ByVal iFilialEmpresa As Integer, ByVal sCcl As String, ByVal dSldIni As Double, Optional ByVal vIncluiCcl As Variant) As Long
'Altera os saldos iniciais dos centros de custo/lucro pais de um ccl.
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lErro As Long
Dim iNivel As Integer
Dim sCcl1 As String
Dim sCcl2 As String
Dim iExercicio As Integer
Dim iIndice As Integer

On Error GoTo Erro_SaldoInicialContaCcl_Altera_Pais1

    lErro = Mascara_Ccl_ObterNivel(sCcl, iNivel)
    If lErro <> SUCESSO Then Error 10354
    
    If Not IsMissing(vIncluiCcl) Then iNivel = iNivel + 1
    
    Do While iNivel > 1
    
        iNivel = iNivel - 1
        
        sCcl1 = String(STRING_CCL, 0)
                
        'pega o pai da conta em questao
        lErro = Mascara_RetornaCclNoNivel(iNivel, sCcl, sCcl1)
        If lErro <> SUCESSO Then Error 10355
    
        sCcl2 = String(STRING_CCL, 0)
    
        'seleciona o centro de custo a ser alterada
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Ccl FROM SaldoInicialContaCcl WHERE FilialEmpresa=? AND Ccl = ? AND Conta = ?", 0, sCcl2, iFilialEmpresa, sCcl1, "")
        If lErro <> AD_SQL_SUCESSO Then Error 10356
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 10357
    
        'altera o centro de custo
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE SaldoInicialContaCcl SET SldIni= SldIni + ?", alComando(1), dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10358
        
        sCcl2 = String(STRING_CCL, 0)
    
        'seleciona o centro de custo a ser alterada no ambito empresa
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Ccl FROM SaldoInicialContaCcl WHERE FilialEmpresa=? AND Ccl = ? AND Conta = ?", 0, sCcl2, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl1, "")
        If lErro <> AD_SQL_SUCESSO Then Error 10603
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 10604
    
        'altera o saldo do centro de custo no ambito empresa
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE SaldoInicialContaCcl SET SldIni= SldIni + ?", alComando(1), dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10605
        
        'seleciona o saldo do centro de custo a ser alterado
        lErro = Comando_ExecutarPos(alComando(3), "SELECT Exercicio FROM MvPerCcl WHERE FilialEmpresa=? AND Ccl = ? AND Conta = ? ORDER BY Exercicio", 0, iExercicio, iFilialEmpresa, sCcl1, "")
        If lErro <> AD_SQL_SUCESSO Then Error 10359
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO Then Error 10360
    
        'altera o saldo do centro de custo
        lErro = Comando_ExecutarPos(alComando(4), "UPDATE MvPerCcl SET SldIni= SldIni + ?", alComando(3), dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10361
        
        'seleciona o saldo do centro de custo a ser alterado no ambito empresa
        lErro = Comando_ExecutarPos(alComando(3), "SELECT Exercicio FROM MvPerCcl WHERE FilialEmpresa=? AND Ccl = ? AND Conta = ? ORDER BY Exercicio", 0, iExercicio, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl1, "")
        If lErro <> AD_SQL_SUCESSO Then Error 10606
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO Then Error 10607
    
        'altera o saldo do centro de custo
        lErro = Comando_ExecutarPos(alComando(4), "UPDATE MvPerCcl SET SldIni= SldIni + ?", alComando(3), dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10608
        
    Loop
    
    SaldoInicialContaCcl_Altera_Pais1 = SUCESSO

    Exit Function

Erro_SaldoInicialContaCcl_Altera_Pais1:
    
    SaldoInicialContaCcl_Altera_Pais1 = Err
    
    Select Case Err
    
        Case 10354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CCL_OBTERNIVEL", Err, sCcl)
            
        Case 10355
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclNoNivel", Err, sCcl, iNivel)
        
        Case 10356, 10357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL2", Err, iFilialEmpresa, sCcl1)
        
        Case 10358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTACCL", Err, iFilialEmpresa, sCcl1)
         
        Case 10359, 10360, 10606, 10607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL3", Err, sCcl1)
        
        Case 10361, 10608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL3", Err, iFilialEmpresa, iExercicio, sCcl1)
         
        Case 10603, 10604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL2", Err, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl1)
            
        Case 10605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTACCL", Err, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl1)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150674)
        
    End Select
    
    Exit Function

End Function

Function SaldoInicialContaCcl_Altera_Pais2(ByVal iFilialEmpresa As Integer, ByVal sCcl As String, ByVal dSldIni As Double, Optional ByVal vIncluiCcl As Variant) As Long
'Altera os saldos iniciais dos centros de custo/lucro pais de um ccl. Altera a filial em questao e a empresa.
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO


Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lErro As Long
Dim iNivel As Integer
Dim sCcl1 As String
Dim sCcl2 As String
Dim iExercicio As Integer

On Error GoTo Erro_SaldoInicialContaCcl_Altera_Pais2

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10305

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10306

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10307

    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 10308

    lErro = Mascara_Ccl_ObterNivel(sCcl, iNivel)
    If lErro <> SUCESSO Then Error 10309
    
    If Not IsMissing(vIncluiCcl) Then iNivel = iNivel + 1
    
    Do While iNivel > 1
    
        iNivel = iNivel - 1
        
        sCcl1 = String(STRING_CCL, 0)
                
        'pega o pai da conta em questao
        lErro = Mascara_RetornaCclNoNivel(iNivel, sCcl, sCcl1)
        If lErro <> SUCESSO Then Error 10310
    
        sCcl2 = String(STRING_CCL, 0)
    
        'seleciona o centro de custo a ser alterada
        lErro = Comando_ExecutarPos(lComando1, "SELECT Ccl FROM SaldoInicialContaCcl WHERE (FilialEmpresa=? OR FilialEmpresa=?) AND Ccl = ? AND Conta = ?", 0, sCcl2, iFilialEmpresa, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl1, "")
        If lErro <> AD_SQL_SUCESSO Then Error 10311
        
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 10312
    
        'altera o centro de custo
        lErro = Comando_ExecutarPos(lComando2, "UPDATE SaldoInicialContaCcl SET SldIni= SldIni + ?", lComando1, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10313
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 10559
    
        'altera o centro de custo
        lErro = Comando_ExecutarPos(lComando2, "UPDATE SaldoInicialContaCcl SET SldIni= SldIni + ?", lComando1, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10560
        
        'seleciona o saldo do centro de custo a ser alterado
        lErro = Comando_ExecutarPos(lComando3, "SELECT Exercicio FROM MvPerCcl WHERE (FilialEmpresa=? OR FilialEmpresa=?) AND Ccl = ? AND Conta = ? ORDER BY Exercicio", 0, iExercicio, iFilialEmpresa, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl1, "")
        If lErro <> AD_SQL_SUCESSO Then Error 10314
        
        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 10315
    
        'altera o saldo do centro de custo
        lErro = Comando_ExecutarPos(lComando4, "UPDATE MvPerCcl SET SldIni= SldIni + ?", lComando3, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10316
        
        lErro = Comando_BuscarProximo(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 10561
    
        'altera o saldo do centro de custo
        lErro = Comando_ExecutarPos(lComando4, "UPDATE MvPerCcl SET SldIni= SldIni + ?", lComando3, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then Error 10562
        
    Loop
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    SaldoInicialContaCcl_Altera_Pais2 = SUCESSO

    Exit Function

Erro_SaldoInicialContaCcl_Altera_Pais2:
    
    SaldoInicialContaCcl_Altera_Pais2 = Err
    
    Select Case Err
    
        Case 10305, 10306, 10307, 10308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 10309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CCL_OBTERNIVEL", Err, sCcl)
            
        Case 10310
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclNoNivel", Err, sCcl, iNivel)
        
        Case 10311, 10312, 10559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL2", Err, iFilialEmpresa, sCcl1)
        
        Case 10313, 10560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTACCL", Err, iFilialEmpresa, sCcl1)
         
        Case 10314, 10315, 10561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL3", Err, sCcl1)
        
        Case 10316, 10562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL3", Err, iFilialEmpresa, iExercicio, sCcl1)
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150675)
        
    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    Exit Function

End Function

Public Function CTBConfig_ObterProxLong(ByVal sChave As String, lValor As Long) As Long
'obtem o valor correspondente a sChave (do tipo long), retorna-o em lValor e incrementa o contador na tabela CTBconfig
Dim sConteudo As String
Dim lComando As Long, lComandoAlt As Long, lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CTBConfig_ObterProxLong

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 7106

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 7108
    
    lComandoAlt = Comando_Abrir()
    If lComandoAlt = 0 Then Error 7112
    
    sConteudo = String(255, 0)
    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM CTBConfig WHERE Codigo = ?", 0, sConteudo, sChave)
    If lErro <> AD_SQL_SUCESSO Then Error 7109
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7110
    
    If lErro = AD_SQL_SEM_DADOS Then Error 7111
    
    lValor = CLng(sConteudo)
    
    lErro = Comando_ExecutarPos(lComandoAlt, "UPDATE CTBConfig SET Conteudo=?", lComando, lValor + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 7113
    
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 7107
    
    Call Comando_Fechar(lComandoAlt)
    Call Comando_Fechar(lComando)
    
    CTBConfig_ObterProxLong = SUCESSO
    
    Exit Function
    
Erro_CTBConfig_ObterProxLong:
    
    CTBConfig_ObterProxLong = Err
    
    Select Case Err
    
        Case 7106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 7107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 7108, 7112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 7109, 7110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CTBCONFIG", Err)
        
        Case 7111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CTBCONFIG_CHV_NAO_ENC", Err, sChave)
        
        Case 7113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CTBCONFIG_ATUALIZAR_CHV", Err, sChave)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150676)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComandoAlt)
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function ContaCcl_Altera_Conta_Contabil(ByVal objContaCcl As ClassContaCcl) As Long
'altera uma associacao de conta com centro de custo a partir de uma conta com centro de custo contabil e seu respectivo saldo
'TEM QUE EXECUTAR DENTRO DE TRANSACAO

Dim sCcl As String
Dim sConta As String
Dim lErro As Long
Dim iNivelConta As Integer
Dim objSaldoInicialContaCcl As New ClassSaldoInicialContaCcl

On Error GoTo Erro_ContaCcl_Altera_Conta_Contabil

    'Se está trabalhando com centro de custo contábil
    If giSetupUsoCcl = CCL_USA_CONTABIL Then
    
        sCcl = String(STRING_CCL, 0)
        
        'verifica se a conta contem um centro de custo
        lErro = Mascara_RetornaCcl(objContaCcl.sConta, sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 9850
        
        'se contem um centro de custo
        If Len(sCcl) > 0 Then
    
            'descobre o nivel da conta
            lErro = Mascara_Conta_ObterNivel(objContaCcl.sConta, iNivelConta)
            If lErro <> SUCESSO Then Error 9851
    
            sConta = String(STRING_CONTA, 0)
    
            'pega o pai da conta em questao
            lErro = Mascara_RetornaContaNoNivel(iNivelConta - 1, objContaCcl.sConta, sConta)
            If lErro <> SUCESSO Then Error 9852
            
            objContaCcl.sCcl = sCcl
            
            objSaldoInicialContaCcl.iFilialEmpresa = objContaCcl.iFilialEmpresa
            objSaldoInicialContaCcl.sConta = sConta
            objSaldoInicialContaCcl.sCcl = sCcl
            objSaldoInicialContaCcl.dSldIni = objContaCcl.dSldIni
            
            'altera a tabela de conta x centro de custo/lucro
            lErro = SaldoInicialContaCcl_Altera(objSaldoInicialContaCcl)
            If lErro <> SUCESSO Then Error 9977
            
            'altera na tabela de saldos de centro de custo/lucro a associacao da conta com o centro de custo/lucro
            lErro = MvPerCcl_Altera(objContaCcl.iFilialEmpresa, objContaCcl.sConta, sCcl, objContaCcl.dSldIni)
            If lErro <> SUCESSO Then Error 9853
            
        End If
    
    End If

    ContaCcl_Altera_Conta_Contabil = SUCESSO

    Exit Function

Erro_ContaCcl_Altera_Conta_Contabil:
    
    ContaCcl_Altera_Conta_Contabil = Err
    
    Select Case Err
    
        Case 9850
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCcl", Err, objContaCcl.sConta)
            
        Case 9851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, objContaCcl.sConta)
            
        Case 9852
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, objContaCcl.sConta, iNivelConta - 1)
            
        Case 9853, 9977
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150677)
        
    End Select
    
    Exit Function

End Function

Public Function MvPerCcl_Altera(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal sCcl As String, ByVal dSldIni As Double) As Long
'Altera na tabela de saldos por centro de custo o saldo inicial da associacao de sConta com sCcl
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim dSldIni1 As Double
Dim dSldIni2 As Double
Dim iExercicio As Integer

On Error GoTo Erro_MvPerCcl_Altera

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 9854
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 9855
    
    'acessa o saldo da associação da conta com o centro de custo em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Exercicio, SldIni FROM MvPerCcl WHERE FilialEmpresa=? AND Conta=? AND Ccl=? ORDER BY Exercicio", 0, iExercicio, dSldIni1, iFilialEmpresa, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 9857
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9858
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Insere o registro de Saldo de Ccl
        lErro = Comando_Executar(lComando1, "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Ccl, Conta, SldIni) VALUES (?, ?, ?, ?, ?)", iFilialEmpresa, iExercicio, sCcl, sConta, dSldIni1)
        If lErro <> AD_SQL_SUCESSO Then gError 188299
    
    Else
        'loca o saldo da associação da conta com o centro de custo
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 9859
        
        'Alterar o saldo da  associacao no BD
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MvPerCcl SET SldIni=?", lComando, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then gError 9860
    End If
        
    'acessa o saldo da associação da conta com o centro de custo em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Exercicio, SldIni FROM MvPerCcl WHERE FilialEmpresa=? AND Conta=? AND Ccl=? ORDER BY Exercicio", 0, iExercicio, dSldIni2, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 10563
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 10564
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Insere o registro de Saldo de Ccl
        lErro = Comando_Executar(lComando1, "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Ccl, Conta, SldIni) VALUES (?, ?, ?, ?, ?)", IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio, sCcl, sConta, dSldIni2)
        If lErro <> AD_SQL_SUCESSO Then gError 188300
    
    Else
    
        'loca o saldo da associação da conta com o centro de custo
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 10565
        
        'Alterar o saldo da  associacao no BD
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MvPerCcl SET SldIni=?", lComando, dSldIni2 + (dSldIni - dSldIni1))
        If lErro <> AD_SQL_SUCESSO Then gError 10566
        
    End If
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    MvPerCcl_Altera = SUCESSO
    
    
    Exit Function

Erro_MvPerCcl_Altera:

    MvPerCcl_Altera = gErr
    
    Select Case gErr
            
        Case 9854, 9855
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 9856
        
        Case 9857, 9858, 10563, 10564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL4", gErr, sCcl, sConta)
            
        Case 9859, 10565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MVPERCCL1", gErr, sCcl, sConta)
                        
        Case 9860, 10566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL2", gErr, sCcl, sConta)
            
        Case 188299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", gErr, iFilialEmpresa, iExercicio, sCcl, sConta)

        Case 188300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", gErr, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio, sCcl, sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150678)
        
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
        
End Function

Public Function SaldoInicialContaCcl_Altera(ByVal objSaldoInicialContaCcl As ClassSaldoInicialContaCcl) As Long
'Altera na tabela SaldoInicialContaCcl o saldo inicial da associacao passada como parametro. Altera a filial e a empresa.
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim dSldIni1 As Double
Dim dSldIni2 As Double

On Error GoTo Erro_SaldoInicialContaCcl_Altera

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9978

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9979

    'acessa a associação da conta com o centro de custo em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT SldIni FROM SaldoInicialContaCcl WHERE FilialEmpresa=? AND Conta=? AND Ccl=?", 0, dSldIni1, objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9980

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9981

    'loca a associação da conta com o centro de custo
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9982

    'Alterar a  associacao no BD
    lErro = Comando_ExecutarPos(lComando1, "UPDATE SaldoInicialContaCcl SET SldIni=?", lComando, objSaldoInicialContaCcl.dSldIni)
    If lErro <> AD_SQL_SUCESSO Then Error 9983

    'acessa a associação da conta com o centro de custo em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT SldIni FROM SaldoInicialContaCcl WHERE FilialEmpresa=? AND Conta=? AND Ccl=?", 0, dSldIni2, IIf(objSaldoInicialContaCcl.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 10558

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10559

    'loca a associação da conta com o centro de custo
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10560

    'Alterar a  associacao no BD
    lErro = Comando_ExecutarPos(lComando1, "UPDATE SaldoInicialContaCcl SET SldIni=?", lComando, dSldIni2 + (objSaldoInicialContaCcl.dSldIni - dSldIni1))
    If lErro <> AD_SQL_SUCESSO Then Error 10561

    lErro = SaldoInicialContaCcl_Altera_Pais2(objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sCcl, objSaldoInicialContaCcl.dSldIni - dSldIni1, INCLUI_CCL_SALDO_INICIAL)
    If lErro <> SUCESSO Then Error 10327

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    SaldoInicialContaCcl_Altera = SUCESSO

    Exit Function

Erro_SaldoInicialContaCcl_Altera:

    SaldoInicialContaCcl_Altera = Err

    Select Case Err

        Case 9978, 9979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9980, 9981
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL", Err, objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case 9982
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SALDOINICIALCONTACCL", Err, objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case 9983
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTACCL1", Err, objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case 10327

        Case 10558, 10559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL", Err, IIf(objSaldoInicialContaCcl.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case 10560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SALDOINICIALCONTACCL", Err, IIf(objSaldoInicialContaCcl.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case 10561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTACCL1", Err, IIf(objSaldoInicialContaCcl.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150679)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Public Function SaldoInicialContaCcl_Altera1(alComando() As Long, ByVal objSaldoInicialContaCcl As ClassSaldoInicialContaCcl) As Long
'Altera na tabela SaldoInicialContaCcl o saldo inicial da associacao passada como parametro
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim dSldIni1 As Double
Dim dSldIni2 As Double
Dim alComando1(1 To 4) As Long
Dim iIndice As Integer

On Error GoTo Erro_SaldoInicialContaCcl_Altera1

    'acessa a associação da conta com o centro de custo em questão
    lErro = Comando_ExecutarPos(alComando(1), "SELECT SldIni FROM SaldoInicialContaCcl WHERE FilialEmpresa=? AND Conta=? AND Ccl=?", 0, dSldIni1, objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9980

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 9981

    'loca a associação da conta com o centro de custo
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 9982

    'Alterar a  associacao no BD
    lErro = Comando_ExecutarPos(alComando(2), "UPDATE SaldoInicialContaCcl SET SldIni=?", alComando(1), objSaldoInicialContaCcl.dSldIni)
    If lErro <> AD_SQL_SUCESSO Then Error 9983

    'acessa a associação da conta com o centro de custo em questão no ambito empresa
    lErro = Comando_ExecutarPos(alComando(1), "SELECT SldIni FROM SaldoInicialContaCcl WHERE FilialEmpresa=? AND Conta=? AND Ccl=?", 0, dSldIni2, IIf(objSaldoInicialContaCcl.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 10599

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 10600

    'loca a associação da conta com o centro de custo no ambito empresa
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 10601

    'Alterar a  associacao no BD no ambito empresa
    lErro = Comando_ExecutarPos(alComando(2), "UPDATE SaldoInicialContaCcl SET SldIni=?", alComando(1), dSldIni2 + (objSaldoInicialContaCcl.dSldIni - dSldIni1))
    If lErro <> AD_SQL_SUCESSO Then Error 10602

    For iIndice = 1 To 4
        alComando1(iIndice) = alComando(iIndice + 2)
    Next

    lErro = SaldoInicialContaCcl_Altera_Pais1(alComando1(), objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sCcl, objSaldoInicialContaCcl.dSldIni - dSldIni1, INCLUI_CCL_SALDO_INICIAL)
    If lErro <> SUCESSO Then Error 10327

    SaldoInicialContaCcl_Altera1 = SUCESSO

    Exit Function

Erro_SaldoInicialContaCcl_Altera1:

    SaldoInicialContaCcl_Altera1 = Err

    Select Case Err

        Case 9978, 9979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9980, 9981
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL", Err, objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case 9982
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SALDOINICIALCONTACCL", Err, objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case 9983
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTACCL1", Err, objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case 10327

        Case 10599, 10600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL", Err, IIf(objSaldoInicialContaCcl.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case 10601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SALDOINICIALCONTACCL", Err, IIf(objSaldoInicialContaCcl.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case 10602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTACCL1", Err, IIf(objSaldoInicialContaCcl.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150680)

    End Select

    Exit Function

End Function

Public Function SaldoInicialConta_Altera(ByVal objPlanoConta As ClassPlanoConta) As Long
'Altera na tabela SaldoInicialConta o saldo inicial da conta passada como parametro
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim dSldIni1 As Double
Dim dSldIni2 As Double

On Error GoTo Erro_SaldoInicialConta_Altera

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10328

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10329

    'acessa o saldo inicial da conta em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT SldIni FROM SaldoInicialConta WHERE FilialEmpresa=? AND Conta=? ", 0, dSldIni1, objPlanoConta.iFilialEmpresa, objPlanoConta.sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 10330

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10331

    'loca o saldo inicial da conta
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10332

    'Alterar a  associacao no BD
    lErro = Comando_ExecutarPos(lComando1, "UPDATE SaldoInicialConta SET SldIni=?", lComando, objPlanoConta.dSldIni)
    If lErro <> AD_SQL_SUCESSO Then Error 10333
        
    'acessa o saldo inicial da conta da EMPRESA_TODA
    lErro = Comando_ExecutarPos(lComando, "SELECT SldIni FROM SaldoInicialConta WHERE FilialEmpresa=? AND Conta=? ", 0, dSldIni2, IIf(objPlanoConta.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objPlanoConta.sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 10538
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10539

    'loca o saldo inicial da conta da EMPRESA_TODA
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10540

    'Alterar o saldo inicial da EMPRESA_TODA
    lErro = Comando_ExecutarPos(lComando1, "UPDATE SaldoInicialConta SET SldIni=?", lComando, dSldIni2 + (objPlanoConta.dSldIni - dSldIni1))
    If lErro <> AD_SQL_SUCESSO Then Error 10541

    'Altera os saldos iniciais das contas pais de uma conta analitica. Se vIncluiConta estiver preenchido ==> inclui também a conta que está sendo passada como parametro. Altera os saldos da filial em questão e da empresa toda.
    lErro = SaldoInicialConta_Altera_Pais1(objPlanoConta.iFilialEmpresa, objPlanoConta.sConta, objPlanoConta.dSldIni - dSldIni1)
    If lErro <> SUCESSO Then Error 10334

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    SaldoInicialConta_Altera = SUCESSO

    Exit Function

Erro_SaldoInicialConta_Altera:

    SaldoInicialConta_Altera = Err

    Select Case Err

        Case 10328, 10329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10330, 10331
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA", Err, objPlanoConta.iFilialEmpresa, objPlanoConta.sConta)

        Case 10332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SALDOINICIALCONTA", Err, objPlanoConta.iFilialEmpresa, objPlanoConta.sConta)

        Case 10333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTA", Err, objPlanoConta.iFilialEmpresa, objPlanoConta.sConta)

        Case 10334

        Case 10538, 10539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA", Err, IIf(objPlanoConta.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objPlanoConta.sConta)

        Case 10540
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SALDOINICIALCONTA", Err, IIf(objPlanoConta.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objPlanoConta.sConta)

        Case 10541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTA", Err, IIf(objPlanoConta.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150681)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Public Function SaldoInicialConta_Insere(ByVal sConta As String) As Long
'Insere na tabela SaldoInicialConta a conta passada como parametro
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_SaldoInicialConta_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10335

    For Each objFiliais In gcolFiliais
    
        If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then

            'Insere o saldo inicial da conta
            lErro = Comando_Executar(lComando, "INSERT INTO SaldoInicialConta (FilialEmpresa, Conta, SldIni) VALUES (?,?,?)", objFiliais.iCodFilial, sConta, 0)
            If lErro <> AD_SQL_SUCESSO Then Error 10336

        End If
        
    Next

    Call Comando_Fechar(lComando)

    SaldoInicialConta_Insere = SUCESSO

    Exit Function

Erro_SaldoInicialConta_Insere:

    SaldoInicialConta_Insere = Err

    Select Case Err

        Case 10335
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SALDOINICIALCONTA", Err, objFiliais.iCodFilial, sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150682)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function SaldoInicialConta_Insere_NovaFilial(ByVal iFilialEmpresa As Integer) As Long
'Insere na tabela SaldoInicialConta os registros sobre saldos iniciais de todas as contas para a nova filial, a partir dos dados da EMPRESA_TODA
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim sConta As String

On Error GoTo Erro_SaldoInicialConta_Insere_NovaFilial

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 55677
    Next

    sConta = String(STRING_CONTA, 0)

    'acessa o saldo inicial da conta em questão
    lErro = Comando_Executar(alComando(1), "SELECT Conta FROM SaldoInicialConta S1 WHERE FilialEmpresa=? AND NOT EXISTS (SELECT * FROM SaldoInicialConta S2 WHERE FilialEmpresa = ? AND S1.Conta = S2.Conta) ORDER BY Conta", sConta, EMPRESA_TODA, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 55678

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55679

    Do While lErro = AD_SQL_SUCESSO
    
        'Insere o saldo inicial da conta
        lErro = Comando_Executar(alComando(2), "INSERT INTO SaldoInicialConta (FilialEmpresa, Conta, SldIni) VALUES (?,?,0)", iFilialEmpresa, sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 55681

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55680

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    SaldoInicialConta_Insere_NovaFilial = SUCESSO

    Exit Function

Erro_SaldoInicialConta_Insere_NovaFilial:

    SaldoInicialConta_Insere_NovaFilial = Err

    Select Case Err

        Case 55677
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55678, 55679, 55680
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA2", Err)

        Case 55681
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SALDOINICIALCONTA", Err, iFilialEmpresa, sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150683)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function SaldoInicialContaCcl_Insere_NovaFilial(ByVal iFilialEmpresa As Integer) As Long
'Insere na tabela SaldoInicialContaCcl os registros sobre saldos iniciais de todas os pares conta/ccl para a nova filial, a partir dos dados da EMPRESA_TODA
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim sConta As String
Dim sCcl As String

On Error GoTo Erro_SaldoInicialContaCcl_Insere_NovaFilial

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 55682
    Next

    sConta = String(STRING_CONTA, 0)
    sCcl = String(STRING_CCL, 0)

    'acessa o saldo inicial da conta em questão
    lErro = Comando_Executar(alComando(1), "SELECT Conta, Ccl FROM SaldoInicialContaCcl S1 WHERE FilialEmpresa=? AND NOT EXISTS (SELECT * FROM SaldoInicialContaCcl S2 WHERE FilialEmpresa = ? AND S1.Conta = S2.Conta AND S1.Ccl = S2.Ccl) ORDER BY Conta, Ccl", sConta, sCcl, EMPRESA_TODA, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 55683

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55684

    Do While lErro = AD_SQL_SUCESSO
    
        'Insere o saldo inicial da conta
        lErro = Comando_Executar(alComando(2), "INSERT INTO SaldoInicialContaCcl (FilialEmpresa, Conta, Ccl, SldIni) VALUES (?,?,?,0)", iFilialEmpresa, sConta, sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 55686

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55685

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    SaldoInicialContaCcl_Insere_NovaFilial = SUCESSO

    Exit Function

Erro_SaldoInicialContaCcl_Insere_NovaFilial:

    SaldoInicialContaCcl_Insere_NovaFilial = Err

    Select Case Err

        Case 55682
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55683, 55684, 55685
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL5", Err)

        Case 55686
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SALDOINICIALCONTACCL", Err, iFilialEmpresa, sConta, sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150684)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function SaldoInicialContaCcl_Insere(ByVal objContaCcl As ClassContaCcl) As Long
'Insere na tabela SaldoInicialContaCcl a conta x ccl passada como parametro
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long, lComando1 As Long
Dim objFiliais As AdmFiliais, dSldIni As Double

On Error GoTo Erro_SaldoInicialContaCcl_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10338

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10338

    For Each objFiliais In gcolFiliais

        If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then

            '??? este teste nao deveria existir mas foi colocado p/contornar problema com bd da INPAL em 05/11/03
            lErro = Comando_Executar(lComando1, "SELECT SldIni FROM SaldoInicialContaCcl WHERE FilialEmpresa = ? AND Conta = ? AND Ccl = ?", dSldIni, objFiliais.iCodFilial, objContaCcl.sConta, objContaCcl.sCcl)
            If lErro <> AD_SQL_SUCESSO Then Error 10339
            
            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10339
            
            If lErro = AD_SQL_SEM_DADOS Then
                
                'Insere o saldo inicial da conta
                lErro = Comando_Executar(lComando, "INSERT INTO SaldoInicialContaCcl (FilialEmpresa, Conta, Ccl, SldIni) VALUES (?,?,?,?)", objFiliais.iCodFilial, objContaCcl.sConta, objContaCcl.sCcl, 0)
                If lErro <> AD_SQL_SUCESSO Then Error 10339
            
            End If
            
        End If

    Next

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    SaldoInicialContaCcl_Insere = SUCESSO

    Exit Function

Erro_SaldoInicialContaCcl_Insere:

    SaldoInicialContaCcl_Insere = Err

    Select Case Err

        Case 10338
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SALDOINICIALCONTACCL", Err, objFiliais.iCodFilial, objContaCcl.sConta, objContaCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150685)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function ContaCcl_Exclui_Conta_Contabil(ByVal sConta As String) As Long
'exclui a associacao de conta com centro de custo a partir de uma conta com centro de custo contabil
'TEM QUE EXECUTAR DENTRO DE TRANSACAO

Dim sCcl As String
Dim sConta1 As String
Dim lErro As Long
Dim iNivelConta As Integer
Dim objContaCcl As New ClassContaCcl

On Error GoTo Erro_ContaCcl_Exclui_Conta_Contabil

    sCcl = String(STRING_CCL, 0)
    
    'verifica se a conta contem um centro de custo
    lErro = Mascara_RetornaCcl(sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9985
    
    'se contem um centro de custo
    If Len(sCcl) > 0 Then

        'descobre o nivel da conta
        lErro = Mascara_Conta_ObterNivel(sConta, iNivelConta)
        If lErro <> SUCESSO Then Error 9986

        sConta1 = String(STRING_CONTA, 0)

        'pega o pai da conta em questao
        lErro = Mascara_RetornaContaNoNivel(iNivelConta - 1, sConta, sConta1)
        If lErro <> SUCESSO Then Error 9987
        
        objContaCcl.sConta = sConta1
        objContaCcl.sCcl = sCcl
        
        'exclui da tabela de conta x centro de custo/lucro
        lErro = ContaCcl_Exclui(objContaCcl)
        If lErro <> SUCESSO Then Error 9988
        
        lErro = SaldoInicialContaCcl_Exclui(objContaCcl)
        If lErro <> SUCESSO Then Error 10296
        
    End If

    ContaCcl_Exclui_Conta_Contabil = SUCESSO

    Exit Function

Erro_ContaCcl_Exclui_Conta_Contabil:
    
    ContaCcl_Exclui_Conta_Contabil = Err
    
    Select Case Err
    
        Case 9985
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCcl", Err, sConta)
            
        Case 9986
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_CONTA_OBTERNIVEL", Err, sConta)
            
        Case 9987
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, sConta, iNivelConta - 1)
            
        Case 9988, 10296
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150686)
        
    End Select
    
    Exit Function

End Function

Function ContaCcl_Exclui(ByVal objContaCcl As ClassContaCcl) As Long
'exclui a associação da conta sConta com o centro de custo sCcl/m
'TEM QUE EXECUTAR DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim sConta1 As String

On Error GoTo Erro_ContaCcl_Exclui

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 9989
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 9990

    sConta1 = String(STRING_CONTA, 0)

    'pesquisa a associações da conta com o centro de custo/lucro em questão
    lErro = Comando_ExecutarPos(lComando2, "SELECT Conta FROM ContaCcl WHERE Conta = ? And Ccl=?", 0, sConta1, objContaCcl.sConta, objContaCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9991
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9992
     
    'exclui a associação
    lErro = Comando_ExecutarPos(lComando3, "DELETE FROM ContaCcl", lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 9993
            
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    ContaCcl_Exclui = SUCESSO
    
    Exit Function
    
Erro_ContaCcl_Exclui:

    ContaCcl_Exclui = Err
    
    Select Case Err
    
        Case 9989, 9990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 9991, 9992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTACCL3", Err, objContaCcl.sConta, objContaCcl.sCcl)
    
        Case 9993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTACCL", Err, objContaCcl.sConta, objContaCcl.sCcl)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150687)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function SaldoInicialContaCcl_Exclui(ByVal objContaCcl As ClassContaCcl) As Long
'exclui a associação da conta sConta com o centro de custo sCcl/m
'TEM QUE EXECUTAR DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim iFilialEmpresa As Integer
Dim dSldIni As Double

On Error GoTo Erro_SaldoInicialContaCcl_Exclui

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10290
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10291

    'pesquisa a associações da conta com o centro de custo/lucro em questão
    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa, SldIni FROM SaldoInicialContaCcl WHERE Conta = ? And Ccl=?", 0, iFilialEmpresa, dSldIni, objContaCcl.sConta, objContaCcl.sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 10292
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10293
     
    Do While lErro = AD_SQL_SUCESSO
     
        If dSldIni <> 0 Then
        
            lErro = SaldoInicialContaCcl_Altera_Pais(iFilialEmpresa, objContaCcl.sCcl, -dSldIni, INCLUI_CCL_SALDO_INICIAL)
            If lErro <> SUCESSO Then Error 10317
            
        End If
     
        'exclui a associação
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM SaldoInicialContaCcl", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 10294
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10295
        
    Loop
            
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    SaldoInicialContaCcl_Exclui = SUCESSO
    
    Exit Function
    
Erro_SaldoInicialContaCcl_Exclui:

    SaldoInicialContaCcl_Exclui = Err
    
    Select Case Err
    
        Case 10290, 10291
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10292, 10293, 10295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL1", Err, objContaCcl.sConta, objContaCcl.sCcl)
    
        Case 10294
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SALDOINICIALCONTACCL", Err, iFilialEmpresa, objContaCcl.sConta, objContaCcl.sCcl)
        
        Case 10317
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150688)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function SaldoInicialContaCcl_Exclui_Ccl(ByVal sCcl As String) As Long
'exclui os saldos iniciais associados ao centro de custo em questão
'TEM QUE EXECUTAR DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim iFilialEmpresa As Integer
Dim dSldIni As Double
Dim sConta As String

On Error GoTo Erro_SaldoInicialContaCcl_Exclui_Ccl

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10380
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10381

    sConta = String(STRING_CONTA, 0)

    'pesquisa os saldos iniciais associados ao centro de custo/lucro em questão
    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa, SldIni, Conta FROM SaldoInicialContaCcl WHERE Ccl=? ", 0, iFilialEmpresa, dSldIni, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 10382
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10383
     
    Do While lErro = AD_SQL_SUCESSO
     
        'So precisa pegar as contas = "" pois contem a totalizacao de todas as contas associadas ao centro de custo.
        If dSldIni <> 0 And sConta = "" Then
        
            lErro = SaldoInicialContaCcl_Altera_Pais(iFilialEmpresa, sCcl, -dSldIni)
            If lErro <> SUCESSO Then Error 10386
            
        End If
        
        'exclui a associação
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM SaldoInicialContaCcl", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 10384
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10385
        
    Loop
            
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    SaldoInicialContaCcl_Exclui_Ccl = SUCESSO
    
    Exit Function
    
Erro_SaldoInicialContaCcl_Exclui_Ccl:

    SaldoInicialContaCcl_Exclui_Ccl = Err
    
    Select Case Err
    
        Case 10380, 10381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10382, 10383, 10385
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL4", Err, sCcl)
    
        Case 10384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SALDOINICIALCONTACCL", Err, iFilialEmpresa, sConta, sCcl)
        
        Case 10386
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150689)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function SaldoInicialContaCcl_Exclui_Ccl1(ByVal sCcl As String) As Long
'exclui os saldos iniciais associados ao centro de custo em questão
'TEM QUE EXECUTAR DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim iFilialEmpresa As Integer
Dim dSldIni As Double
Dim sConta As String

On Error GoTo Erro_SaldoInicialContaCcl_Exclui_Ccl1

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10432
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10433

    sConta = String(STRING_CONTA, 0)

    'pesquisa os saldos iniciais associados ao centro de custo/lucro em questão
    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa, SldIni, Conta FROM SaldoInicialContaCcl WHERE Ccl=?", 0, iFilialEmpresa, dSldIni, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 10434
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10435
     
    Do While lErro = AD_SQL_SUCESSO
     
        'exclui a associação
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM SaldoInicialContaCcl", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 10436
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10437
        
    Loop
            
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    SaldoInicialContaCcl_Exclui_Ccl1 = SUCESSO
    
    Exit Function
    
Erro_SaldoInicialContaCcl_Exclui_Ccl1:

    SaldoInicialContaCcl_Exclui_Ccl1 = Err
    
    Select Case Err
    
        Case 10432, 10433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10434, 10435, 10437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL4", Err, sCcl)
    
        Case 10436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SALDOINICIALCONTACCL", Err, iFilialEmpresa, sConta, sCcl)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150690)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function SaldoInicialConta_Exclui_Conta(ByVal sConta As String) As Long
'exclui os saldos iniciais da conta sConta.
'TEM QUE EXECUTAR DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim iFilialEmpresa As Integer
Dim dSldIni As Double

On Error GoTo Erro_SaldoInicialConta_Exclui_Conta

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10298
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10299

    'pesquisa os saldos iniciais da conta em questão
    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa, SldIni FROM SaldoInicialConta WHERE Conta = ?", 0, iFilialEmpresa, dSldIni, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 10300
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10301
     
    Do While lErro = AD_SQL_SUCESSO
    
        If dSldIni <> 0 Then
        
            'altera os saldos iniciais das contas pai
            lErro = SaldoInicialConta_Altera_Pais(iFilialEmpresa, sConta, -dSldIni)
            If lErro <> SUCESSO Then Error 10304
            
        End If
     
        'exclui o saldo inicial
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM SaldoInicialConta", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 10302
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10303
        
    Loop
            
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    SaldoInicialConta_Exclui_Conta = SUCESSO
    
    Exit Function
    
Erro_SaldoInicialConta_Exclui_Conta:

    SaldoInicialConta_Exclui_Conta = Err
    
    Select Case Err
    
        Case 10298, 10299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10300, 10301, 10303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA1", Err, sConta)
    
        Case 10302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SALDOINICIALCONTA", Err, iFilialEmpresa, sConta)
            
        Case 10304
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150691)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function SaldoInicialConta_Exclui_Conta1(ByVal sConta As String) As Long
'exclui os saldos iniciais da conta sConta.
'TEM QUE EXECUTAR DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim iFilialEmpresa As Integer
Dim dSldIni As Double

On Error GoTo Erro_SaldoInicialConta_Exclui_Conta1

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10320
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10321

    'pesquisa os saldos iniciais da conta em questão
    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa, SldIni FROM SaldoInicialConta WHERE Conta = ?", 0, iFilialEmpresa, dSldIni, sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 10322
        
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10323
     
    Do While lErro = AD_SQL_SUCESSO
    
        'exclui o saldo inicial
        lErro = Comando_ExecutarPos(lComando3, "DELETE FROM SaldoInicialConta", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 10324
        
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10325
        
    Loop
            
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    SaldoInicialConta_Exclui_Conta1 = SUCESSO
    
    Exit Function
    
Erro_SaldoInicialConta_Exclui_Conta1:

    SaldoInicialConta_Exclui_Conta1 = Err
    
    Select Case Err
    
        Case 10320, 10321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10322, 10323, 10325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA1", Err, sConta)
    
        Case 10324
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SALDOINICIALCONTA", Err, iFilialEmpresa, sConta)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150692)
    
    End Select
    
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function


Public Function ContaCcl_Atualiza_Saldo_Extra(ByVal colSaldoInicialContaCcl As Collection) As Long
'atualiza o saldo inicial de uma coleção de contas com centro de custo extra contábil

Dim lErro As Long
Dim sContaFinal As String
Dim objSaldoInicialContaCcl As ClassSaldoInicialContaCcl
Dim iTipoConta As Integer
Dim dSldIniAcum As Double
Dim dSldIni As Double
Dim alComando(0 To 8) As Long
Dim lTransacao As Long
Dim iDisponivel As Integer
Dim iIndice As Integer
Dim iNaturezaConta As Integer

On Error GoTo Erro_ContaCcl_Atualiza_Saldo_Extra

    For iIndice = LBound(alComando) To UBound(alComando)
        
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 10362
        
    Next

    lTransacao = 0
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 10020
    
    'verifica se o saldo inicial pode ser alterado
    lErro = CF("Saldo_Inicial_Critica", iDisponivel)
    If lErro <> SUCESSO Then Error 10021
    
    If iDisponivel = SLDINI_NAO_DISPONIVEL Then Error 10022
    
    dSldIniAcum = 0
        
    Set objSaldoInicialContaCcl = colSaldoInicialContaCcl.Item(1)
        
    'seleciona a conta que será alterada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Natureza FROM PlanoConta WHERE Conta = ?", 0, iNaturezaConta, objSaldoInicialContaCcl.sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 60732
    
    'le a conta a ser alterada
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 60733
    
    'loca a conta
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 60734

    For Each objSaldoInicialContaCcl In colSaldoInicialContaCcl

        If iNaturezaConta = CONTA_DEBITO Then objSaldoInicialContaCcl.dSldIni = -objSaldoInicialContaCcl.dSldIni
        
        'altera a tabela de conta x centro de custo/lucro
        lErro = SaldoInicialContaCcl_Altera1(alComando, objSaldoInicialContaCcl)
        If lErro <> SUCESSO Then Error 10023
        
        'altera na tabela de saldos de centro de custo/lucro a associacao da conta com o centro de custo/lucro
        lErro = MvPerCcl_Altera1(objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta, objSaldoInicialContaCcl.sCcl, objSaldoInicialContaCcl.dSldIni, alComando(7), alComando(8), dSldIni)
        If lErro <> SUCESSO Then Error 10024
        
        dSldIniAcum = dSldIniAcum + dSldIni
        
    Next

    If dSldIniAcum <> 0 Then
    
        Set objSaldoInicialContaCcl = colSaldoInicialContaCcl.Item(1)

        'Altera os saldos iniciais das contas pais de uma conta analitica. Se vIncluiConta estiver preenchido ==> inclui também a conta que está sendo passada como parametro. Altera os saldos da filial em questão e da empresa toda.
        lErro = SaldoInicialConta_Altera_Pais1(objSaldoInicialContaCcl.iFilialEmpresa, objSaldoInicialContaCcl.sConta, dSldIniAcum, INCLUI_CONTA_SALDO_INICIAL)
        If lErro <> SUCESSO Then Error 10025
    
    End If
            
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 10026

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ContaCcl_Atualiza_Saldo_Extra = SUCESSO

    Exit Function

Erro_ContaCcl_Atualiza_Saldo_Extra:
    
    ContaCcl_Atualiza_Saldo_Extra = Err
    
    Select Case Err
    
        Case 10362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 10020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 10021, 10023, 10024, 10025
    
        Case 10022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDOS_INICIAIS_NAO_ALTERAVEIS", Err)
        
        Case 10026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 60732, 60733
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, objSaldoInicialContaCcl.sConta)
        
        Case 60734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, objSaldoInicialContaCcl.sConta)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150693)
        
    End Select
    
    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function MvPerCcl_Altera1(ByVal iFilialEmpresa As Integer, ByVal sConta As String, ByVal sCcl As String, ByVal dSldIni As Double, ByVal lComando As Long, ByVal lComando1 As Long, dSldIni3 As Double) As Long
'Altera na tabela de saldos por centro de custo o saldo inicial da associacao de sConta com sCcl
'PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim dSldIni1 As Double
Dim dSldIni2 As Double
Dim iExercicio As Integer

On Error GoTo Erro_MvPerCcl_Altera1

    'acessa o saldo da associação da conta com o centro de custo em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Exercicio, SldIni FROM MvPerCcl WHERE FilialEmpresa=? AND Conta=? AND Ccl=? ORDER BY Exercicio", 0, iExercicio, dSldIni1, iFilialEmpresa, sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 9857
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9858
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Insere o registro de Saldo de Ccl
        lErro = Comando_Executar(lComando1, "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Ccl, Conta, SldIni) VALUES (?, ?, ?, ?, ?)", iFilialEmpresa, iExercicio, sCcl, sConta, dSldIni1)
        If lErro <> AD_SQL_SUCESSO Then gError 188301
    
    Else
    
        'loca o saldo da associação da conta com o centro de custo
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 9859
        
        'Alterar o saldo da  associacao no BD
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MvPerCcl SET SldIni=?", lComando, dSldIni)
        If lErro <> AD_SQL_SUCESSO Then gError 9860
        
    End If
        
    'acessa o saldo da associação da conta com o centro de custo em questão no ambito empresa
    lErro = Comando_ExecutarPos(lComando, "SELECT Exercicio, SldIni FROM MvPerCcl WHERE FilialEmpresa=? AND Conta=? AND Ccl=? ORDER BY Exercicio", 0, iExercicio, dSldIni2, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 10609
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 10610
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Insere o registro de Saldo de Ccl
        lErro = Comando_Executar(lComando1, "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Ccl, Conta, SldIni) VALUES (?, ?, ?, ?, ?)", IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio, sCcl, sConta, dSldIni2)
        If lErro <> AD_SQL_SUCESSO Then gError 188302
    
    Else
    
        
        'loca o saldo da associação da conta com o centro de custo no ambito empresa
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 10611
        
        'Alterar o saldo da  associacao no BD
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MvPerCcl SET SldIni=?", lComando, dSldIni2 + (dSldIni - dSldIni1))
        If lErro <> AD_SQL_SUCESSO Then gError 10612

    End If

    dSldIni3 = dSldIni - dSldIni1

    MvPerCcl_Altera1 = SUCESSO
    
    Exit Function

Erro_MvPerCcl_Altera1:

    MvPerCcl_Altera1 = gErr
    
    Select Case gErr
            
        Case 9856
        
        Case 9857, 9858, 10609, 10610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL4", gErr, sCcl, sConta)
            
        Case 9859, 10611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MVPERCCL1", gErr, sCcl, sConta)
                        
        Case 9860, 10612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL2", gErr, sCcl, sConta)
            
        Case 188301
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", gErr, iFilialEmpresa, iExercicio, sCcl, sConta)

        Case 188302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", gErr, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio, sCcl, sConta)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150694)
        
    End Select
    
    Exit Function
        
End Function

Public Function ContaCcl_Atualiza_Saldo_Contabil(ByVal colSaldoInicialContaCcl As Collection) As Long
'atualiza o saldo inicial de uma coleção de contas com centro de custo contábil

Dim lErro As Long
Dim iDisponivel As Integer
Dim alComando(0 To 16) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lComando As Integer

On Error GoTo Erro_ContaCcl_Atualiza_Saldo_Contabil
    
    For iIndice = LBound(alComando) To UBound(alComando)
        
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 10007
        
    Next
    
    lTransacao = 0
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 10011
    
    'verifica se o saldo inicial pode ser alterado
    lErro = CF("Saldo_Inicial_Critica", iDisponivel)
    If lErro <> SUCESSO Then Error 9997
    
    If iDisponivel = SLDINI_NAO_DISPONIVEL Then Error 9996
    
    lErro = ContaCcl_Atualiza_Saldo_Contabil1(colSaldoInicialContaCcl, alComando())
    If lErro <> SUCESSO Then Error 10029
    
    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 10012

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ContaCcl_Atualiza_Saldo_Contabil = SUCESSO

    Exit Function

Erro_ContaCcl_Atualiza_Saldo_Contabil:
    
    ContaCcl_Atualiza_Saldo_Contabil = Err
    
    Select Case Err
    
        Case 9996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDOS_INICIAIS_NAO_ALTERAVEIS", Err)
        
        Case 9997
        
        Case 10007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 10011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 10012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case 10029
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150695)
        
    End Select
    
    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function ContaCcl_Atualiza_Saldo_Contabil1(ByVal colSaldoInicialContaCcl As Collection, alComando() As Long) As Long
'atualiza o saldo inicial de uma coleção de contas com centro de custo contábil

Dim sContaFinal As String
Dim objSaldoInicialContaCcl As ClassSaldoInicialContaCcl
Dim iTipoConta As Integer
Dim dSldIniAcum As Double
Dim dSldIni1 As Double
Dim dSldIni As Double
Dim lErro As Long
Dim iExercicio As Integer
Dim alComando1(1 To 6) As Long
Dim iIndice As Integer
Dim dSldIni2 As Double
Dim iNaturezaConta As Integer


On Error GoTo Erro_ContaCcl_Atualiza_Saldo_Contabil1

    dSldIniAcum = 0
    
    For Each objSaldoInicialContaCcl In colSaldoInicialContaCcl
    
        sContaFinal = left(objSaldoInicialContaCcl.sConta, Len(objSaldoInicialContaCcl.sConta) - Len(objSaldoInicialContaCcl.sCcl)) & objSaldoInicialContaCcl.sCcl
    
        'seleciona a conta que será alterada
        lErro = Comando_ExecutarPos(alComando(0), "SELECT TipoConta, Natureza FROM PlanoConta WHERE Conta = ?", 0, iTipoConta, iNaturezaConta, sContaFinal)
        If lErro <> AD_SQL_SUCESSO Then Error 9998
        
        'le a conta a ser alterada
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 9999
        
        'loca a conta
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 10013
    
        'se a conta não for analítica ==> erro
        If iTipoConta <> CONTA_ANALITICA Then Error 10000
         
        If iNaturezaConta = CONTA_DEBITO Then objSaldoInicialContaCcl.dSldIni = -objSaldoInicialContaCcl.dSldIni
         
        'seleciona o registro que guarda o saldo inicial da conta a ser alterada
        lErro = Comando_ExecutarPos(alComando(1), "SELECT SldIni FROM SaldoInicialConta WHERE FilialEmpresa=? AND Conta = ?", 0, dSldIni, objSaldoInicialContaCcl.iFilialEmpresa, sContaFinal)
        If lErro <> AD_SQL_SUCESSO Then Error 10352
        
        'le o saldo inicial da conta a ser alterada
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 10353
        
        'loca o saldo inicial
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 10354
         
        'se o saldo cadastrado for diferente do saldo informado
        If dSldIni <> objSaldoInicialContaCcl.dSldIni Then
        
            'altera o saldo da conta na tabela SaldoInicialConta
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE SaldoInicialConta SET SldIni=?", alComando(1), objSaldoInicialContaCcl.dSldIni)
            If lErro <> AD_SQL_SUCESSO Then Error 10001
        
            lErro = ContaCcl_Atualiza_Saldo_Contabil2(alComando(), objSaldoInicialContaCcl, sContaFinal, dSldIni)
            If lErro <> SUCESSO Then Error 10590
            
            For iIndice = 1 To 6
                alComando1(iIndice) = alComando(iIndice + 4)
            Next
            
            'altera a associação conta x centro de custo/lucro
            lErro = SaldoInicialContaCcl_Altera1(alComando1(), objSaldoInicialContaCcl)
            If lErro <> SUCESSO Then Error 10030
            
            'altera o saldo da conta por centro de custo/lucro
            lErro = MvPerCcl_Altera1(objSaldoInicialContaCcl.iFilialEmpresa, sContaFinal, objSaldoInicialContaCcl.sCcl, objSaldoInicialContaCcl.dSldIni, alComando(11), alComando(12), dSldIni2)
            If lErro <> SUCESSO Then Error 10031
            
            dSldIniAcum = dSldIniAcum + (objSaldoInicialContaCcl.dSldIni - dSldIni)
            
        End If
        
    Next
    
    If dSldIniAcum <> 0 Then
    
        Set objSaldoInicialContaCcl = colSaldoInicialContaCcl.Item(1)
    
        'Altera os saldos iniciais das contas pais de uma conta analitica. Se vIncluiConta estiver preenchido ==> inclui também a conta que está sendo passada como parametro. Altera os saldos da filial em questão e da empresa toda.
        lErro = SaldoInicialConta_Altera_Pais1(objSaldoInicialContaCcl.iFilialEmpresa, sContaFinal, dSldIniAcum)
        If lErro <> SUCESSO Then Error 10006
        
    End If

    ContaCcl_Atualiza_Saldo_Contabil1 = SUCESSO

    Exit Function

Erro_ContaCcl_Atualiza_Saldo_Contabil1:
    
    ContaCcl_Atualiza_Saldo_Contabil1 = Err
    
    Select Case Err

        Case 9998, 9999
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA1", Err, sContaFinal)
        
        Case 10000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_ANALITICA_SALDO", Err, sContaFinal)
        
        Case 10001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTA", Err, objSaldoInicialContaCcl.iFilialEmpresa, sContaFinal)
            
        Case 10013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, sContaFinal)
            
        Case 10030, 10031, 10590
            
        Case 10352, 10353
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA", Err, objSaldoInicialContaCcl.iFilialEmpresa, sContaFinal)
        
        Case 10354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SALDOINICIALCONTA", Err, objSaldoInicialContaCcl.iFilialEmpresa, sContaFinal)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150696)
        
    End Select
    
    Exit Function

End Function

Private Function ContaCcl_Atualiza_Saldo_Contabil2(alComando() As Long, ByVal objSaldoInicialContaCcl As ClassSaldoInicialContaCcl, ByVal sContaFinal As String, ByVal dSldIni As Double) As Long
'atualiza o saldo inicial da conta sContaFinal

Dim lErro As Long
Dim dSldIni1 As Double
Dim iExercicio As Integer

On Error GoTo Erro_ContaCcl_Atualiza_Saldo_Contabil2

    'seleciona o registro que guarda o saldo inicial da conta a ser alterada, no ambito empresa.
    lErro = Comando_ExecutarPos(alComando(1), "SELECT SldIni FROM SaldoInicialConta WHERE FilialEmpresa=? AND Conta = ?", 0, dSldIni1, IIf(objSaldoInicialContaCcl.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sContaFinal)
    If lErro <> AD_SQL_SUCESSO Then Error 10591
    
    'le o saldo inicial da conta a ser alterada
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 10592
    
    'loca o saldo inicial
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 10593
     
    'altera o saldo da conta na tabela SaldoInicialConta
    lErro = Comando_ExecutarPos(alComando(2), "UPDATE SaldoInicialConta SET SldIni=?", alComando(1), dSldIni1 + (objSaldoInicialContaCcl.dSldIni - dSldIni))
    If lErro <> AD_SQL_SUCESSO Then Error 10594

    'seleciona o saldo da conta a ser alterado
    lErro = Comando_ExecutarPos(alComando(3), "SELECT Exercicio FROM MvPerCta WHERE FilialEmpresa=? AND Conta = ? ORDER BY Exercicio", 0, iExercicio, objSaldoInicialContaCcl.iFilialEmpresa, sContaFinal)
    If lErro <> AD_SQL_SUCESSO Then Error 10002
    
    'le o saldo da conta
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then Error 10003

    'loca o saldo da conta
    lErro = Comando_LockExclusive(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then Error 10014

    'altera a conta na tabela de saldos de conta
    lErro = Comando_ExecutarPos(alComando(4), "UPDATE MvPerCta SET SldIni = ?", alComando(3), objSaldoInicialContaCcl.dSldIni)
    If lErro <> AD_SQL_SUCESSO Then Error 10004

    'seleciona o saldo da conta a ser alterado
    lErro = Comando_ExecutarPos(alComando(3), "SELECT Exercicio FROM MvPerCta WHERE FilialEmpresa=? AND Conta = ? ORDER BY Exercicio", 0, iExercicio, IIf(objSaldoInicialContaCcl.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sContaFinal)
    If lErro <> AD_SQL_SUCESSO Then Error 10595
    
    'le o saldo da conta
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then Error 10596

    'loca o saldo da conta
    lErro = Comando_LockExclusive(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then Error 10597

    'altera a conta na tabela de saldos de conta
    lErro = Comando_ExecutarPos(alComando(4), "UPDATE MvPerCta SET SldIni = ?", alComando(3), dSldIni1 + (objSaldoInicialContaCcl.dSldIni - dSldIni))
    If lErro <> AD_SQL_SUCESSO Then Error 10598

    ContaCcl_Atualiza_Saldo_Contabil2 = SUCESSO

    Exit Function

Erro_ContaCcl_Atualiza_Saldo_Contabil2:
    
    ContaCcl_Atualiza_Saldo_Contabil2 = Err
    
    Select Case Err

        Case 10002, 10003, 10595, 10596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA2", Err, sContaFinal)
        
        Case 10004, 10598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCTA", Err, iExercicio, sContaFinal)
            
        Case 10014, 10597
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MVPERCTA", Err, iExercicio, sContaFinal)
            
        Case 10591, 10592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTA", Err, 0, sContaFinal)
        
        Case 10593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SALDOINICIALCONTA", Err, 0, sContaFinal)
            
        Case 10594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SALDOINICIALCONTA", Err, 0, sContaFinal)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150697)
            
        
    End Select
    
    Exit Function

End Function

Function Exercicio_Grava(ByVal objExercicio As ClassExercicio, ByVal colPeriodos As Collection) As Long
'grava os dados do exercicio objExercicio.

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Exercicio_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 13668

    lErro = CF("Exercicio_Grava_Trans", objExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 20717

    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 13676

    Exercicio_Grava = SUCESSO

    Exit Function

Erro_Exercicio_Grava:

    Exercicio_Grava = Err

    Select Case Err

        Case 13668
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 13676
             lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 20717

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150698)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Exercicio_Grava_Trans(ByVal objExercicio As ClassExercicio, ByVal colPeriodos As Collection) As Long
'grava os dados do exercicio objExercicio.
'IMPORTANTE: Tem que ser chamado em transação

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim tExercicio As typeExercicio
Dim iIdAtualizacao As Integer
Dim colExercicios As New Collection
Dim objExercicio1 As ClassExercicio

On Error GoTo Erro_Exercicio_Grava_Trans

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 13669

    lComando1 = Comando_Abrir()
    If lComando = 0 Then Error 13670

    'seleciona o registro de configuração
    lErro = Comando_ExecutarPos(lComando1, "SELECT IdAtualizacao FROM Configuracao", 0, iIdAtualizacao)
    If lErro <> AD_SQL_SUCESSO Then Error 10079

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 10080

    'lock da configuração
    'Força a alteração exclusiva da tabela de Exercicios
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 10081

    'le todos os exercícios existentes no BD
    lErro = CF("Exercicios_Le_Todos", colExercicios)
    If lErro <> SUCESSO Then Error 10078

    For Each objExercicio1 In colExercicios
    
        'critica se já existe um exercicio com o mesmo nome externo
        If objExercicio1.iExercicio <> objExercicio.iExercicio And objExercicio1.sNomeExterno = objExercicio.sNomeExterno Then Error 10077
        
    Next

    tExercicio.sNomeExterno = String(STRING_NOME_EXTERNO, 0)

    'verifica se o exercício está cadastrado
    lErro = Comando_ExecutarPos(lComando, "SELECT NumPeriodos, DataInicio, DataFim, NomeExterno FROM Exercicios WHERE Exercicio=?", 0, tExercicio.iNumPeriodos, tExercicio.dtDataInicio, tExercicio.dtDataFim, tExercicio.sNomeExterno, objExercicio.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 13671

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 13672

    If lErro = SUCESSO Then

        'exercício está cadastrado. Altera os dados do exercicio
        lErro = Exercicio_Grava1(lComando, objExercicio, tExercicio, colPeriodos, colExercicios)
        If lErro <> SUCESSO Then Error 10102


    Else
        
        'exercício não está cadastrado. Insere o exercicio.
        lErro = Exercicio_Insere(objExercicio, colExercicios, colPeriodos)
        If lErro <> SUCESSO Then Error 10119

    End If

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exercicio_Grava_Trans = SUCESSO
    
    Exit Function

Erro_Exercicio_Grava_Trans:

    Exercicio_Grava_Trans = Err

    Select Case Err

        Case 10077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_EXERCICIO_JA_USADO", Err, objExercicio.sNomeExterno, objExercicio1.iExercicio)
            
        Case 10078, 10102, 10119
            
        Case 10079, 10080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)

        Case 10081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)

        Case 13669, 13670
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 13671, 13672
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, objExercicio.iExercicio)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150699)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function


Function ExerciciosFilial_Grava(ByVal objExercicio As ClassExercicio, ByVal colPeriodos As Collection) As Long
'grava os dados do exercicio/filial contido em objExercicio

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iStatus As Integer

On Error GoTo Erro_ExerciciosFilial_Grava

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10239

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 10240

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 10241

    'seleciona o exercicio em questao
    lErro = Comando_ExecutarPos(lComando1, "SELECT Status FROM Exercicios WHERE Exercicio=?", 0, iStatus, objExercicio.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 10242

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 10243

    'lock do exercicio
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 10244

    'verifica se o exercicio está fechado
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If iStatus = EXERCICIO_FECHADO Then Error 10245
    End If

    'verifica se o exercíciofilial está cadastrado
    lErro = Comando_ExecutarPos(lComando, "SELECT Status FROM ExerciciosFilial WHERE FilialEmpresa = ? AND Exercicio=?", 0, iStatus, objExercicio.iFilialEmpresa, objExercicio.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 10246

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10247

    'lock do exerciciofilial
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10248

    lErro = CF("PeriodosFilial_Altera_Status", objExercicio.iFilialEmpresa, objExercicio.iExercicio, colPeriodos)
    If lErro <> SUCESSO Then Error 10249

    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 10250

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ExerciciosFilial_Grava = SUCESSO

    Exit Function

Erro_ExerciciosFilial_Grava:

    ExerciciosFilial_Grava = Err

    Select Case Err

        Case 10239, 10240
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10241
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 10242, 10243
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, objExercicio.iExercicio)

        Case 10244
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, objExercicio.iExercicio)

        Case 10245
             lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_FECHADO1", Err, objExercicio.iExercicio)

        Case 10246, 10247
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL", Err, objExercicio.iFilialEmpresa, objExercicio.iExercicio)

        Case 10248
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOSFILIAL", Err, objExercicio.iFilialEmpresa, objExercicio.iExercicio)

        Case 10249

        Case 10250
             lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150700)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Function Exercicio_Insere(ByVal objExercicio As ClassExercicio, ByVal colExercicios As Collection, ByVal colPeriodos As Collection) As Long
'insere um novo exercicio

Dim objExercicio1 As ClassExercicio
Dim iIndice As Integer
Dim adtDataInicio(NUM_MAX_PERIODOS) As Date
Dim adtDataFim(NUM_MAX_PERIODOS) As Date
Dim asNomeExterno(NUM_MAX_PERIODOS) As String
Dim aiFechado(NUM_MAX_PERIODOS) As Integer
Dim lErro As Long
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_Exercicio_Insere

    'se tem algum exercicio cadastrado
    If colExercicios.Count > 0 Then
    
        'pega o ultimo exercicio cadastrado
        Set objExercicio1 = colExercicios.Item(colExercicios.Count)
    
        'verifica se o exercicio que está sendo incluido tem data de inicio um dia maior que a data do ultimo exercicio
        If objExercicio1.dtDataFim + 1 <> objExercicio.dtDataInicio Then Error 11387

    End If
    
    iIndice = 0

    For Each objPeriodo In colPeriodos
        adtDataInicio(iIndice) = objPeriodo.dtDataInicio
        adtDataFim(iIndice) = objPeriodo.dtDataFim
        asNomeExterno(iIndice) = objPeriodo.sNomeExterno
        aiFechado(iIndice) = objPeriodo.iFechado
        iIndice = iIndice + 1
    Next

    lErro = CF("Exercicio_Inicializacao", objExercicio, 1, adtDataInicio(), adtDataFim(), asNomeExterno(), aiFechado())
    If lErro <> SUCESSO Then Error 11389
    
    Exercicio_Insere = SUCESSO
    
    Exit Function
    
Erro_Exercicio_Insere:

    Exercicio_Insere = Err
    
    Select Case Err
    
        Case 11387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIO_NOVO_EXERCICIO", Err, CStr(objExercicio1.dtDataFim + 1), CStr(objExercicio.dtDataInicio))
    
        Case 11389

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150701)

    End Select
    
    Exit Function

End Function

Private Function Exercicio_Grava1(ByVal lComando As Long, ByVal objExercicio As ClassExercicio, tExercicio As typeExercicio, ByVal colPeriodos As Collection, ByVal colExercicios As Collection) As Long
'altera os dados do exercicio passado como parametro
'TEM QUE SER EXECUTADO EM TRANSACAO

Dim lErro As Long

On Error GoTo Erro_Exercicio_Grava1

    'lock do exercício
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 13673

    'verifica se o exercício possui movimento, lote ou lançamento pendente
    lErro = CF("Exercicio_Critica_Movimento", objExercicio.iExercicio)
    If lErro <> SUCESSO And lErro <> 13663 Then Error 13674

    If lErro = 13663 Then
    
        'se há movimento para o exercicio em questão
        lErro = Exercicio_Grava_Com_Movimento(lComando, objExercicio, tExercicio, colPeriodos)
        If lErro <> SUCESSO Then Error 10103
    
    Else
    
        'se não há movimento para o exercicio em questão
        lErro = Exercicio_Grava_Sem_Movimento(lComando, objExercicio, tExercicio, colPeriodos, colExercicios)
        If lErro <> SUCESSO Then Error 10110
        
    End If

    Exercicio_Grava1 = SUCESSO

    Exit Function

Erro_Exercicio_Grava1:

    Exercicio_Grava1 = Err

    Select Case Err

        Case 13673
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, tExercicio.iExercicio)

        Case 13674, 10103, 10110

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150702)

    End Select
    
    Exit Function

End Function

Private Function Exercicio_Grava_Com_Movimento(ByVal lComando As Long, ByVal objExercicio As ClassExercicio, tExercicio As typeExercicio, ByVal colPeriodos As Collection) As Long
'grava as alteraçoes do exercicio passado como parametro. Este exercicio possui movimento.
'TEM QUE SER EXECUTADO EM TRANSACAO

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 5) As Long
Dim objPeriodo As ClassPeriodo
Dim iPeriodo As Integer

On Error GoTo Erro_Exercicio_Grava_Com_Movimento

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 10104
        
    Next

    'testa se as datas coincidem
    If objExercicio.dtDataInicio <> tExercicio.dtDataInicio Then Error 11399
    If objExercicio.dtDataFim <> tExercicio.dtDataFim Then Error 11401
    
    'testa se o numero de periodos coincide
    If objExercicio.iNumPeriodos <> tExercicio.iNumPeriodos Then Error 11397

    'Alterar nome do exercício
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE Exercicios SET NomeExterno=?", lComando, objExercicio.sNomeExterno)
    If lErro <> AD_SQL_SUCESSO Then Error 10118

    'Alterar nomes dos períodos
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Periodo FROM Periodo WHERE Exercicio = ?", 0, iPeriodo, objExercicio.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 11392

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11393

    Do While lErro = AD_SQL_SUCESSO

        Set objPeriodo = colPeriodos.Item(iPeriodo)

        lErro = Comando_ExecutarPos(alComando(3), "UPDATE Periodo SET NomeExterno=?", alComando(2), objPeriodo.sNomeExterno)
        If lErro <> AD_SQL_SUCESSO Then Error 11394

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11395
        
    Loop

'    ' Alterar status dos períodos da filial em questão
'    lErro = Comando_ExecutarPos(alComando(4), "SELECT Periodo FROM PeriodosFilial WHERE FilialEmpresa=? AND Exercicio = ?", 0, iPeriodo, objExercicio.iFilialEmpresa, objExercicio.iExercicio)
'    If lErro <> AD_SQL_SUCESSO Then Error 10188
'
'    lErro = Comando_BuscarPrimeiro(alComando(4))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10189
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        Set objPeriodo = colPeriodos.Item(iPeriodo)
'
'        lErro = Comando_ExecutarPos(alComando(5), "UPDATE PeriodosFilial SET Fechado=?", alComando(4), objPeriodo.iFechado)
'        If lErro <> AD_SQL_SUCESSO Then Error 10190
'
'        lErro = Comando_BuscarProximo(alComando(4))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10191
'
'    Loop

    Exercicio_Grava_Com_Movimento = SUCESSO

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

Erro_Exercicio_Grava_Com_Movimento:

    Exercicio_Grava_Com_Movimento = Err

    Select Case Err

        Case 10104
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10118
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOS", Err, tExercicio.iExercicio)

        Case 10188, 10189, 10191
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL2", Err, objExercicio.iFilialEmpresa, objExercicio.iExercicio)

        Case 10190
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PERIODOSFILIAL", Err, iPeriodo, objExercicio.iExercicio, objExercicio.iFilialEmpresa)

        Case 11392, 11393, 11395
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO2", Err, objExercicio.iExercicio)

        Case 11394
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PERIODO", Err, iPeriodo, objExercicio.iExercicio)

        Case 11397
             lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PERIODOS_ALTERADO", Err, tExercicio.iNumPeriodos, objExercicio.iNumPeriodos)
             
        Case 11399
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIO_EXERCICIO_ALTERADA", Err, CStr(tExercicio.dtDataInicio), CStr(objExercicio.dtDataInicio))

        Case 11401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFIM_EXERCICIO_ALTERADA", Err, CStr(tExercicio.dtDataFim), CStr(objExercicio.dtDataFim))

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150703)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Exercicio_Grava_Sem_Movimento(ByVal lComando As Long, ByVal objExercicio As ClassExercicio, tExercicio As typeExercicio, ByVal colPeriodos As Collection, ByVal colExercicios As Collection) As Long
'grava as alteraçoes do exercicio passado como parametro. Este exercicio está cadastrado e não possui movimento.
'TEM QUE SER EXECUTADO EM TRANSACAO

Dim lErro As Long
Dim iIndice As Integer
Dim objExercicio1 As ClassExercicio
Dim adtDataInicio(NUM_MAX_PERIODOS) As Date
Dim adtDataFim(NUM_MAX_PERIODOS) As Date
Dim asNomeExterno(NUM_MAX_PERIODOS) As String
Dim aiFechado(NUM_MAX_PERIODOS) As Integer
Dim objPeriodo As ClassPeriodo

On Error GoTo Erro_Exercicio_Grava_Sem_Movimento

    'pega o ultimo exercicio cadastrado
    Set objExercicio1 = colExercicios.Item(colExercicios.Count)
    
    'verifica se o exercicio em questão é o ultimo cadastrado
    If objExercicio.iExercicio = objExercicio1.iExercicio Then
    
        If colExercicios.Count > 1 Then
    
        'testa se a data inicio coincide. Permite alterar a data fim do exercicio
        If objExercicio.dtDataInicio <> tExercicio.dtDataInicio Then Error 10115
        
        End If
        
    Else
    
        'testa se as datas coincidem.
        If objExercicio.dtDataInicio <> tExercicio.dtDataInicio Then Error 10116
        If objExercicio.dtDataFim <> tExercicio.dtDataFim Then Error 10117
        
    End If

    'se o numero de periodos ou a data fim do exercicio foi alterada
    If objExercicio.iNumPeriodos <> tExercicio.iNumPeriodos Or objExercicio.dtDataFim <> tExercicio.dtDataFim Or objExercicio.dtDataInicio <> tExercicio.dtDataInicio Then

        iIndice = 0

        For Each objPeriodo In colPeriodos
            adtDataInicio(iIndice) = objPeriodo.dtDataInicio
            adtDataFim(iIndice) = objPeriodo.dtDataFim
            asNomeExterno(iIndice) = objPeriodo.sNomeExterno
            aiFechado(iIndice) = objPeriodo.iFechado
            iIndice = iIndice + 1
        Next

        lErro = CF("Exercicio_Reinicializacao", objExercicio, 1, adtDataInicio(), adtDataFim(), asNomeExterno(), aiFechado())
        If lErro <> SUCESSO Then Error 11400

    Else
            
        'se o numero de periodos e a data fim do exercicio ficaram inalteradas
        lErro = Exercicio_Grava_Sem_Movimento1(lComando, objExercicio, colPeriodos)
        If lErro <> SUCESSO Then Error 10111

    End If

    Exercicio_Grava_Sem_Movimento = SUCESSO

    Exit Function

Erro_Exercicio_Grava_Sem_Movimento:

    Exercicio_Grava_Sem_Movimento = Err

    Select Case Err

        Case 10111, 11400

        Case 10115, 10116
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAINICIO_EXERCICIO_ALTERADA", Err, CStr(tExercicio.dtDataInicio), CStr(objExercicio.dtDataInicio))

        Case 10117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAFIM_EXERCICIO_ALTERADA", Err, CStr(tExercicio.dtDataFim), CStr(objExercicio.dtDataFim))

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150704)

    End Select
    
    Exit Function
    
End Function

Private Function Exercicio_Grava_Sem_Movimento1(ByVal lComando As Long, ByVal objExercicio As ClassExercicio, ByVal colPeriodos As Collection) As Long
'altera os dados do exercicio passado como parametro. O Exercicio não possui movimento e não teve alterado o numero de periodos.
'TEM QUE SER EXECUTADO EM TRANSACAO

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iPeriodo As Integer
Dim objPeriodo As ClassPeriodo
Dim iIndice As Integer

On Error GoTo Erro_Exercicio_Grava_Sem_Movimento1

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 10107
        
    Next

    'Alterar nome do exercício
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE Exercicios SET NomeExterno=?", lComando, objExercicio.sNomeExterno)
    If lErro <> AD_SQL_SUCESSO Then Error 13675

    'Alterar os dados dos períodos
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Periodo FROM Periodo WHERE Exercicio = ? ", 0, iPeriodo, objExercicio.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 10112

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10113

    Do While lErro = SUCESSO

        Set objPeriodo = colPeriodos.Item(iPeriodo)

        lErro = Comando_ExecutarPos(alComando(3), "UPDATE Periodo SET NomeExterno=?, DataInicio=?, DataFim=?", alComando(2), objPeriodo.sNomeExterno, objPeriodo.dtDataInicio, objPeriodo.dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then Error 11418

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10114
        
    Loop
    
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exercicio_Grava_Sem_Movimento1 = SUCESSO

    Exit Function

Erro_Exercicio_Grava_Sem_Movimento1:

    Exercicio_Grava_Sem_Movimento1 = Err

    Select Case Err
    
        Case 10107
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10112, 10113, 10114
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO2", Err, objExercicio.iExercicio)

        Case 11418
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PERIODO", Err, iPeriodo, objExercicio.iExercicio)

        Case 13675
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOS", Err, objExercicio.iExercicio)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150705)
             
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function


Function PeriodosFilial_Altera_Status(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal colPeriodos As Collection) As Long
'altera o status dos periodos do exercicio passado como parametro. O Exercicio não possui movimento e não teve alterado o numero de periodos.
'TEM QUE SER EXECUTADO EM TRANSACAO

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iPeriodo As Integer
Dim objPeriodo As ClassPeriodo
Dim iIndice As Integer

On Error GoTo Erro_PeriodosFilial_Altera_Status

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 10220
        
    Next

    'Alterar status dos períodos da filial em questão
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Periodo FROM PeriodosFilial WHERE FilialEmpresa=? AND Exercicio = ?", 0, iPeriodo, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 10192

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10193

    Do While lErro = AD_SQL_SUCESSO

        Set objPeriodo = colPeriodos.Item(iPeriodo)

        lErro = Comando_ExecutarPos(alComando(2), "UPDATE PeriodosFilial SET Fechado=?", alComando(1), objPeriodo.iFechado)
        If lErro <> AD_SQL_SUCESSO Then Error 10194

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10195
        
    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PeriodosFilial_Altera_Status = SUCESSO

    Exit Function

Erro_PeriodosFilial_Altera_Status:

    PeriodosFilial_Altera_Status = Err

    Select Case Err
    
        Case 10220
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10192, 10193, 10195
             lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL2", Err, iFilialEmpresa, iExercicio)

        Case 10194
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PERIODOSFILIAL", Err, iPeriodo, iExercicio, iFilialEmpresa)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150706)
             
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function



Function Exercicio_Exclui(ByVal iExercicio As Integer) As Long
'exclui o exercicio em questão

Dim lTransacao As Long
Dim lErro As Long
Dim iUltimoExercicio As Integer
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim iStatus As Integer
Dim iIdAtualizacao As Integer
Dim colExercicios As New Collection
Dim objExercicio As ClassExercicio

On Error GoTo Erro_Exercicio_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11425

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 11426

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10096

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 11424

    'seleciona o registro de configuração
    lErro = Comando_ExecutarPos(lComando2, "SELECT IdAtualizacao FROM Configuracao", 0, iIdAtualizacao)
    If lErro <> AD_SQL_SUCESSO Then Error 10097

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 10098

    'lock da configuração.
    'Força a alteração exclusiva da tabela de Exercicios
    lErro = Comando_LockExclusive(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 10099

    'Pesquisa o Exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Status FROM Exercicios WHERE Exercicio = ?", 0, iStatus, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 11418

    'le o Exercicio
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 11419

    'se o exercicio não estiver cadastrado ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 10082

    lErro = Comando_LockExclusive(lComando)
    If lErro <> SUCESSO Then Error 11420
    
    lErro = CF("Exercicios_Le_Todos", colExercicios)
    If lErro <> SUCESSO Then Error 10129

    'pega o ultimo exercicio
    Set objExercicio = colExercicios(colExercicios.Count)

    'testa se o exercicio em questão é o ultimo. Se não for ==> não pode excluir.
    If iExercicio <> objExercicio.iExercicio Then Error 11427

    lErro = CF("Exercicio_Critica_Movimento", iExercicio)
    If lErro <> SUCESSO And lErro <> 13663 Then Error 11412

    If lErro = 13663 Then Error 11413

    lErro = Exercicio_Exclui_BD(iExercicio)
    If lErro <> SUCESSO Then Error 11414

    'Exclui o exercicio da Tabela de Exercicios
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM Exercicios", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 10127

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 10128

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exercicio_Exclui = SUCESSO


    Exit Function

Erro_Exercicio_Exclui:

    Exercicio_Exclui = Err

    Select Case Err
        
        Case 10082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", Err, iExercicio)

        Case 10096, 11425, 11426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10097, 10098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)

        Case 10099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)
            
        Case 10127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EXERCICIO", Err, iExercicio)

        Case 10128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 10129, 11412, 11414

        Case 11413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_COM_MOVIMENTO", Err, iExercicio)

        Case 11418, 11419
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, iExercicio)

        Case 11420
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, iExercicio)

        Case 11424
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 11427
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_ULTIMO", Err, objExercicio.iExercicio)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150707)

    End Select

    lErro = Transacao_Rollback()
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Private Function Exercicio_Exclui_BD(ByVal iExercicio As Integer) As Long
'AVISO: TEM QUE ESTAR EM TRANSACAO PARA CHAMAR ESSA FUNCAO

Dim lErro As Long
Dim objExercicio1 As New ClassExercicio

On Error GoTo Erro_Exercicio_Exclui_BD


    'Exclui os saldos de conta
    lErro = MvPerCta_Exclui_Exercicio(iExercicio)
    If lErro <> SUCESSO Then Error 11421

    'Exclui todos os saldos de centro de custo/lucro associados com o Exercicio passado como parametro
    lErro = MvPerCcl_Exclui_Exercicio(iExercicio)
    If lErro <> SUCESSO Then Error 11385

    'Exclui exercicio origem
    lErro = ExercicioOrigem_Exclui_Exercicio(iExercicio)
    If lErro <> SUCESSO Then Error 11422

    'Exclui os periodos
    lErro = Periodo_Exclui_Exercicio(iExercicio)
    If lErro <> SUCESSO Then Error 11423

    'Exclui os periodosFilial
    lErro = PeriodosFilial_Exclui_Exercicio(iExercicio)
    If lErro <> SUCESSO Then Error 10186

    'Exclui os ExerciciosFilial
    lErro = ExerciciosFilial_Exclui_Exercicio(iExercicio)
    If lErro <> SUCESSO Then Error 10733

    Exercicio_Exclui_BD = SUCESSO

    Exit Function

Erro_Exercicio_Exclui_BD:

    Exercicio_Exclui_BD = Err

    Select Case Err

        Case 10186, 10733, 11385, 11421, 11422, 11423

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150708)

    End Select

    Exit Function

End Function

Private Function ExerciciosFilial_Insere_Exercicio(ByVal iExercicio As Integer, ByVal colFiliais As Collection) As Long
'Insere os registros do exercicio para cada filial
'ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_ExerciciosFilial_Insere_Exercicio

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 10223
    
    For Each objFiliais In colFiliais
    
        If objFiliais.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Or objFiliais.iCodFilial = EMPRESA_TODA Then
    
            lErro = Comando_Executar(lComando, "INSERT INTO ExerciciosFilial (FilialEmpresa, Exercicio, Status, LoteApuracao, DocApuracao) VALUES (?,?,?,0,0)", objFiliais.iCodFilial, iExercicio, EXERCICIO_ABERTO)
            If lErro <> AD_SQL_SUCESSO Then Error 10224
            
        End If
        
    Next

    Call Comando_Fechar(lComando)
    
    ExerciciosFilial_Insere_Exercicio = SUCESSO
    
    Exit Function

Erro_ExerciciosFilial_Insere_Exercicio:

    ExerciciosFilial_Insere_Exercicio = Err
    
    Select Case Err
    
        Case 10223
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_EXERCICIOSFILIAL", Err, objFiliais.iCodFilial, iExercicio)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150709)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PadraoContab_Grava(ByVal objPadraoContab As ClassPadraoContab, ByVal colPadraoContabItem As ClassColPadraoContabItem) As Long
'grava o modelo padrão de contabilização em questão

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim iPadrao As Integer
Dim iObrigatoriedade As Integer

On Error GoTo Erro_PadraoContab_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 36168

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36169
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 36172
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 123000
     
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 123001
    
    'Seleciona a obrigatoriedade
    lErro = Comando_ExecutarPos(lComando2, "SELECT Obrigatoriedade FROM TransacaoCTB WHERE Sigla=? AND SubTipo=? AND Codigo=?", 0, iObrigatoriedade, objPadraoContab.sModulo, objPadraoContab.iSubTipo, objPadraoContab.iTransacao)
    If lErro <> AD_SQL_SUCESSO Then Error 123002
     
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 123003
    
    If lErro = AD_SQL_SUCESSO Then
    
        'se a Obrigatoriedade seleciona for diferente da obrigatoriedade
        'da tela realiza a atualizacao
        If iObrigatoriedade <> objPadraoContab.iObrigatoriedade Then
          
            lErro = Comando_LockExclusive(lComando2)
            If lErro <> AD_SQL_SUCESSO Then Error 123004
            
            'Atualiza a transacaoCTB
            lErro = Comando_ExecutarPos(lComando3, "UPDATE TransacaoCTB SET Obrigatoriedade=?", lComando2, objPadraoContab.iObrigatoriedade)
            If lErro <> AD_SQL_SUCESSO Then Error 123005
            
        End If
    
    End If
    
    'seleciona o modelo segundo os parametros passados
    lErro = Comando_ExecutarPos(lComando, "SELECT Padrao FROM PadraoContab WHERE Modulo=? AND Transacao=? AND SubTipo=? AND Modelo=?", 0, iPadrao, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.iSubTipo, objPadraoContab.sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 36170
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36171
    
    'se o modelo estiver cadastrado
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 36178
    
        'se o modelo foi colocado como padrao e nao era ==> verifica se existe algum modelo padrão e se houver desmarca-o
        If objPadraoContab.iPadrao = PADRAOCONTAB_MODELO_PADRAO And iPadrao <> objPadraoContab.iPadrao Then
            lErro = PadraoContab_Desmarca_Padrao(objPadraoContab)
            If lErro <> SUCESSO Then Error 36214
        End If
    
        'atualizar o padraocontab
        If iPadrao <> objPadraoContab.iPadrao Then
            lErro = Comando_ExecutarPos(lComando1, "UPDATE PadraoContab SET Padrao=?", lComando, objPadraoContab.iPadrao)
            If lErro <> AD_SQL_SUCESSO Then Error 36173
        End If
        
        'excluir os itens do padrao de contabilizacao
        lErro = CF("PadraoContabItem_Exclui", objPadraoContab)
        If lErro <> SUCESSO Then Error 36174

    Else
    
        'o modelo não está cadastrado ==> inserir o modelo
        lErro = Comando_Executar(lComando1, "INSERT INTO PadraoContab  (Modulo, Transacao, Modelo, Padrao, SubTipo) VALUES (?,?,?,?,?)", objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo, objPadraoContab.iPadrao, objPadraoContab.iSubTipo)
        If lErro <> AD_SQL_SUCESSO Then Error 36175

    End If
    
    'insere os lancamentos do padrao de contabilizacao
    lErro = CF("PadraoContabItem_Inserir", colPadraoContabItem)
    If lErro <> SUCESSO Then Error 36176

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 36177

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    PadraoContab_Grava = SUCESSO

    Exit Function

Erro_PadraoContab_Grava:

    PadraoContab_Grava = Err

    Select Case Err

        Case 36168
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 36169, 36172, 123000, 123001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 36170, 36171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOCONTAB1", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)

        Case 36173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PADRAOCONTAB", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)
            
        Case 36174, 36176
            
        Case 36175
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_PADRAOCONTAB", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)
            
        Case 36177
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case 36178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PADRAOCONTAB", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)
            
        Case 123002, 123003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSACAOCTB", Err, objPadraoContab.sModulo)
        
        Case 123004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRANSACAOCTB", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.iSubTipo)
            
        Case 123005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TRANSACAOCTB", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.iSubTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150710)

    End Select

    lErro = Transacao_Rollback()
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function PadraoContabItem_Exclui(ByVal objPadraoContab As ClassPadraoContab) As Long
'excluir os itens do padrao de contabilizacao
'ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iItem As Integer

On Error GoTo Erro_PadraoContabItem_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36179

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 36182

    'seleciona os itens do  modelo de geracao padrao passado como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT Item FROM PadraoContabItem WHERE Modulo=? AND Transacao=? AND SubTipo=? AND Modelo=?", 0, iItem, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.iSubTipo, objPadraoContab.sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 36180
    
    'le o primeiro item do modelo em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36181

    Do While lErro = SUCESSO
    
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM PadraoContabItem", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 36183
        
        'le o proximo item do modelo em questão
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36184
    
    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    PadraoContabItem_Exclui = SUCESSO

    Exit Function

Erro_PadraoContabItem_Exclui:

    PadraoContabItem_Exclui = Err

    Select Case Err

        Case 36179, 36182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 36180, 36181, 36184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOCONTABITEM", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)

        Case 36183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PADRAOCONTABITEM", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo, iItem)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150711)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function
        
Function PadraoContab_Exclui(ByVal objPadraoContab As ClassPadraoContab) As Long
'exclui o modelo padrão de contabilização em questão

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iPadrao As Integer

On Error GoTo Erro_PadraoContab_Exclui

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 36187

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36188
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 36189
    
    'seleciona o modelo em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Padrao FROM PadraoContab WHERE Modulo=? AND Transacao=? AND SubTipo=? AND Modelo=?", 0, iPadrao, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.iSubTipo, objPadraoContab.sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 36190
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36191
    
    'se o modelo não estiver cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 36192
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36193
    
    'exclui o padraocontab
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM PadraoContab", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36194
        
    'excluir os itens do padrao de contabilizacao
    lErro = CF("PadraoContabItem_Exclui", objPadraoContab)
    If lErro <> SUCESSO Then Error 36195

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 36196

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    PadraoContab_Exclui = SUCESSO

    Exit Function

Erro_PadraoContab_Exclui:

    PadraoContab_Exclui = Err

    Select Case Err

        Case 36187
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 36188, 36189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 36190, 36191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOCONTAB1", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)

        Case 36192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAOCONTAB_INEXISTENTE", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)

        Case 36193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PADRAOCONTAB", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)

        Case 36194
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PADRAOCONTAB", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)
            
        Case 36195
            
        Case 36196
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150712)

    End Select

    lErro = Transacao_Rollback()
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function PadraoContabItem_Inserir(ByVal colPadraoContabItem As ClassColPadraoContabItem) As Long
'insere os lancamentos do padrao de contabilizacao
'ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim objPadraoContabItem As ClassPadraoContabItem
Dim lComando As Long

On Error GoTo Erro_PadraoContabItem_Inserir

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36200

    For Each objPadraoContabItem In colPadraoContabItem

        'Insere o lancamento do modelo de contabilizacao em questao
        lErro = Comando_Executar(lComando, "INSERT INTO PadraoContabItem (Modulo, Transacao, SubTipo, Modelo, Item, Conta, Ccl, Credito, Debito, Historico, Custo, Produto, Aglutina, Gerencial) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objPadraoContabItem.sModulo, objPadraoContabItem.iTransacao, objPadraoContabItem.iSubTipo, objPadraoContabItem.sModelo, objPadraoContabItem.iItem, objPadraoContabItem.sConta, objPadraoContabItem.sCcl, objPadraoContabItem.sCredito, objPadraoContabItem.sDebito, objPadraoContabItem.sHistorico, objPadraoContabItem.iCusto, objPadraoContabItem.sProduto, objPadraoContabItem.iAglutina, objPadraoContabItem.iGerencial)
        If lErro <> AD_SQL_SUCESSO Then Error 36201

    Next

    Call Comando_Fechar(lComando)

    PadraoContabItem_Inserir = SUCESSO

    Exit Function

Erro_PadraoContabItem_Inserir:

    PadraoContabItem_Inserir = Err

    Select Case Err

        Case 36200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 36201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_PADRAOCONTABITEM", Err, objPadraoContabItem.sModulo, objPadraoContabItem.iTransacao, objPadraoContabItem.sModelo, objPadraoContabItem.iItem)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150713)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function


End Function

Private Function PadraoContab_Desmarca_Padrao(ByVal objPadraoContab As ClassPadraoContab) As Long
'verifica se existe algum modelo padrão diferente do passado como parametro. Se houver, desmarca-o.

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sModelo As String

On Error GoTo Erro_PadraoContab_Desmarca_Padrao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36208
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 36209
    
    sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    
    'seleciona o modelo de geracao padrao deste Modulo/Transacao
    lErro = Comando_ExecutarPos(lComando, "SELECT Modelo FROM PadraoContab WHERE Modulo=? AND Transacao=? And SubTipo=? AND Padrao=?", 0, sModelo, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.iSubTipo, PADRAOCONTAB_MODELO_PADRAO)
    If lErro <> AD_SQL_SUCESSO Then Error 36210
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36211
    
    'se existir um modelo padrao ==> desmarca-lo como padrao
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 36212
    
        'desmarcar o modelo como padrao
        lErro = Comando_ExecutarPos(lComando1, "UPDATE PadraoContab SET Padrao=?", lComando, PADRAOCONTAB_NAO_E_MODELO_PADRAO)
        If lErro <> AD_SQL_SUCESSO Then Error 36213
        

    End If
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    PadraoContab_Desmarca_Padrao = SUCESSO

    Exit Function

Erro_PadraoContab_Desmarca_Padrao:

    PadraoContab_Desmarca_Padrao = Err

    Select Case Err

        Case 36208, 36209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 36210, 36211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOCONTAB", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao)

        Case 36212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PADRAOCONTAB", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, sModelo)
            
        Case 36213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PADRAOCONTAB", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, sModelo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150714)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function LotePendente_Insere_LoteCusto(ByVal objLote As ClassLote) As Long
'grava um lote de custo na tabela de lotes pendentes
'ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO
    
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lErro As Long
Dim iStatus As Integer
Dim iIdAtualizacao As Integer
Dim objPeriodosFilial As New ClassPeriodosFilial
    
On Error GoTo Erro_LotePendente_Insere_LoteCusto

    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 36419

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 36420

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 36421

    'verifica se o exercicio não está fechado
    lErro = Comando_ExecutarLockado(lComando2, "SELECT Status FROM Exercicios WHERE Exercicio=?", iStatus, objLote.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 36422
    
    'le o exercicio, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 36423
    
    lErro = Comando_LockShared(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 36424
    
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If iStatus = EXERCICIO_FECHADO Then Error 36425
    End If
        
    objPeriodosFilial.iFilialEmpresa = objLote.iFilialEmpresa
    objPeriodosFilial.iExercicio = objLote.iExercicio
    objPeriodosFilial.iPeriodo = objLote.iPeriodo
    objPeriodosFilial.sOrigem = objLote.sOrigem

    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 20853
    
    If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 36429
        
        
    lErro = CF("Lote_Automatico1", objLote)
    If lErro <> SUCESSO Then Error 36430
        
    'insere o lote
    lErro = Comando_Executar(lComando1, "INSERT INTO LotePendente (FilialEmpresa,Origem,Exercicio,Periodo,Lote) VALUES (?,?,?,?,?)", objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 36431
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    LotePendente_Insere_LoteCusto = SUCESSO

    Exit Function

Erro_LotePendente_Insere_LoteCusto:
    
    LotePendente_Insere_LoteCusto = Err
    
    Select Case Err
    
        Case 20853, 36430
        
        Case 36419, 36420, 36421
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 36422, 36423
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, objLote.iExercicio)
            
        Case 36424
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, objLote.iExercicio)
        
        Case 36425
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_EXERCICIO_FECHADO", Err, objLote.iExercicio)
        
        Case 36426, 36427
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo)
            
        Case 36428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODOSFILIAL", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo)
        
        Case 36429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_PERIODO_FECHADO", Err, objLote.iExercicio, objLote.iPeriodo)
        
        Case 36431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150715)
        
    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function Voucher_Automatico1(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal sOrigem As String, lDoc As Long) As Long
'retorna o número do proximo voucher (documento) disponivel. Não verifica a tabela de Lancamentos ou Lancamentos Pendentes.
'ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim lDoc1 As Long
Dim lDoc2 As Long
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim iFilialEmpresaAux As Integer
Dim lDoc3 As Long

On Error GoTo Erro_Voucher_Automatico1

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36434
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 36435
    
    'se a numeração do voucher(documento) não é reinicializada a cada periodo ==> não leva em consideração o periodo passado como parametro
    If giSetupDocPorPeriodo = 0 Then
        iPeriodo = 0
    End If
    
    'Pesquisa o ultimo número de voucher gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT Doc FROM ExercicioOrigem WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ? AND Origem = ?", 0, lDoc, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 36436

    'le o registro com o número do voucher
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36437

    'loca o registro que guarda o numero do voucher gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36438
    
    'atualiza o numero do voucher gerado automaticamente
    lErro = Comando_ExecutarPos(lComando1, "UPDATE ExercicioOrigem SET Doc = ?", lComando, lDoc + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 36439
        
    If giContabGerencial = 1 Then
        
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 193333
        Next
        
        If iFilialEmpresa > Abs(giFilialAuxiliar) Then
            iFilialEmpresaAux = iFilialEmpresa - Abs(giFilialAuxiliar)
        Else
            iFilialEmpresaAux = iFilialEmpresa + Abs(giFilialAuxiliar)
        End If
            
        'Pesquisa o ultimo número de voucher gerado automaticamente pelo sistema
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Doc FROM ExercicioOrigem WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ? AND Origem = ?", 0, lDoc3, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 193334
    
        'le o registro com o número do voucher
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 193335
    
        'loca o registro que guarda o numero do voucher gerado automaticamente
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 193336
        
        'atualiza o numero do voucher gerado automaticamente
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ExercicioOrigem SET Doc = ?", alComando(1), lDoc + 1)
        If lErro <> AD_SQL_SUCESSO Then gError 193337
        
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Voucher_Automatico1 = SUCESSO
        
    Exit Function
    
    
Erro_Voucher_Automatico1:

    Voucher_Automatico1 = Err

    Select Case Err
    
        Case 36434, 36435, 193333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 36436, 36437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOORIGEM", Err, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
        
        Case 36438
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOORIGEM", Err, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
        
        Case 36439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOORIGEM", Err, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
        
        Case 193334, 193335
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOORIGEM", gErr, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        
        Case 193336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOORIGEM", gErr, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        
        Case 193337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOORIGEM", gErr, iFilialEmpresaAux, iExercicio, iPeriodo, sOrigem)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150716)
        
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function RateioOff_Automatico(lCodigo As Long) As Long
'retorna o número do proximo código de Rateio Offline disponivel

Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim lErro As Long
Dim lCodigo1 As Long
Dim lCodigo2 As Long
Dim iIndice As Integer

On Error GoTo Erro_RateioOff_Automatico

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 36782

    lErro = CF("RateioOff_Automatico_Trans", lCodigo)
    If lErro <> SUCESSO Then gError 197643

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 36790
        
    RateioOff_Automatico = SUCESSO
        
    Exit Function
    
    
Erro_RateioOff_Automatico:

    RateioOff_Automatico = gErr

    Select Case gErr
    
        Case 36782
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 36790
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 197643
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150717)
        
    End Select
    
    Call Transacao_Rollback

    Exit Function
        
End Function

Function Lancamento_Grava_Estorno(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal objLancamento_Cabecalho1 As ClassLancamento_Cabecalho) As Long
'objLancamento_Cabecalho contém algumas infos sobre os lançamentos de estorno a serem gravados. objLancamento_Cabecalho1 contem o lote ou doc a ser extornado

Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_Lancamento_Grava_Estorno

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 36884
    
    Do While objLancamento_Cabecalho.iFilialEmpresa > 0 And objLancamento_Cabecalho.iFilialEmpresa < 100
       
        lErro = Lancamento_Grava_Estorno0(objLancamento_Cabecalho, objLancamento_Cabecalho1)
        If lErro <> SUCESSO Then Error 36885
    
        If giContabGerencial = 0 Then Exit Do

        objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa - giFilialAuxiliar
        objLancamento_Cabecalho1.iFilialEmpresa = objLancamento_Cabecalho1.iFilialEmpresa - giFilialAuxiliar
        
    Loop
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 36886

    Lancamento_Grava_Estorno = SUCESSO

    Exit Function

Erro_Lancamento_Grava_Estorno:
    
    Lancamento_Grava_Estorno = Err
    
    Select Case Err
    
        Case 36884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
                       
        Case 36885
                      
        Case 36886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
                       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150718)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function Lancamento_Grava_Estorno0(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal objLancamento_Cabecalho1 As ClassLancamento_Cabecalho) As Long
'objLancamento_Cabecalho contém algumas infos sobre os lançamentos de estorno a serem gravados. objLancamento_Cabecalho1 contem o lote ou doc a ser extornado
'Grava os dados do lançamento. Utilizado por Lancamento_Grava_Estorno
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
    
Dim alComando(1 To 3) As Long
Dim lErro As Long
Dim iExercicio As Integer
Dim iPeriodo As Integer
Dim iStatus As Integer
Dim iIndice As Integer
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Lancamento_Grava_Estorno0

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 36887
        
    Next
    
    'verifica se o exercicio não está fechado
    lErro = Comando_ExecutarLockado(alComando(1), "SELECT Exercicio, Status FROM Exercicios WHERE DataInicio <= ? AND DataFim >= ?", iExercicio, iStatus, objLancamento_Cabecalho.dtData, objLancamento_Cabecalho.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 36888
    
    'le o exercicio, se estiver no BD
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 36889
    
    objLancamento_Cabecalho.iExercicio = iExercicio
    
    lErro = Comando_LockShared(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 36890
    
    'se o exercicio estiver fechado, não pode editar os lançamentos
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If iStatus = EXERCICIO_FECHADO Then Error 36891
    End If
    
    'verifica qual o periodo do lançamento
    lErro = Comando_ExecutarLockado(alComando(2), "SELECT Periodo FROM Periodo WHERE Exercicio=? AND DataInicio <= ? AND DataFim >= ?", iPeriodo, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.dtData, objLancamento_Cabecalho.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 36892
    
    'le o periodo, se estiver no BD
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 36893
    
    objLancamento_Cabecalho.iPeriodoLote = iPeriodo
    objLancamento_Cabecalho.iPeriodoLan = iPeriodo
    
    lErro = Comando_LockShared(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 36894
    
    objPeriodosFilial.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
    objPeriodosFilial.iExercicio = objLancamento_Cabecalho.iExercicio
    objPeriodosFilial.iPeriodo = objLancamento_Cabecalho.iPeriodoLan
    objPeriodosFilial.sOrigem = objLancamento_Cabecalho.sOrigem

    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 20854
    
    'se o periodo estiver fechado, não pode alterar os lançamentos do periodo
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 36898
    End If
    
    
    lErro = Lancamento_Grava_Estorno1(objLancamento_Cabecalho, objLancamento_Cabecalho1)
    If lErro <> SUCESSO Then Error 36899
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Lancamento_Grava_Estorno0 = SUCESSO

    Exit Function

Erro_Lancamento_Grava_Estorno0:
    
    Lancamento_Grava_Estorno0 = Err
    
    Select Case Err
    
        Case 36887
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 36888, 36889
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOS", Err)
           
        Case 36890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, objLancamento_Cabecalho.iExercicio)
                       
        Case 36891
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", Err, objLancamento_Cabecalho.iExercicio)
                    
        Case 36892, 36893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO1", Err)
            
        Case 36894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODO", Err, objLancamento_Cabecalho.iExercicio, iPeriodo)
            
        Case 36895, 36896
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iExercicio, iPeriodo)
            
        Case 36897
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODOSFILIAL", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iExercicio, iPeriodo)
        
        Case 36898
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", Err, iPeriodo, objLancamento_Cabecalho.iExercicio)
            
        Case 36899
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150719)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Lancamento_Grava_Estorno1(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal objLancamento_Cabecalho1 As ClassLancamento_Cabecalho) As Long
'objLancamento_Cabecalho contém algumas infos sobre os lançamentos de estorno a serem gravados. objLancamento_Cabecalho1 contem o lote ou doc a ser extornado
'Grava os dados do lançamento. Utilizado por Lancamento_Grava_Estorno0
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lErro As Long
Dim iLoteAtualizado As Integer
Dim objLote As New ClassLote
Dim objLote1 As New ClassLote
Dim dDebito As Double
Dim dCredito As Double
Dim iNumLancAtual As Integer
Dim iNumDocAtual As Integer
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_Lancamento_Grava_Estorno1

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 36900
        
    Next

'    objLote.sOrigem = objLancamento_Cabecalho.sOrigem
'    objLote.iExercicio = objLancamento_Cabecalho.iExercicio
'    objLote.iPeriodo = objLancamento_Cabecalho.iPeriodoLote
'    objLote.iLote = objLancamento_Cabecalho.iLote
'    objLote.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
'
'    'verifica se o lote  está atualizado
'    lErro = CF("Lote_Critica_Atualizado",objLote, iLoteAtualizado)
'    If lErro <> SUCESSO Then Error 36901
'
'    If iLoteAtualizado = LOTE_ATUALIZADO Then Error 36902
'
'    'Pesquisa Lote associado aos Lançamentos
'    lErro = Comando_ExecutarPos(alComando(1), "SELECT TotCre, TotDeb, NumLancAtual, NumDocAtual FROM LotePendente WHERE FilialEmpresa=? AND Origem=? AND Exercicio=? AND Periodo=? AND Lote=?", 0, dCredito, dDebito, iNumLancAtual, iNumDocAtual, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
'    If lErro <> AD_SQL_SUCESSO Then Error 36903
'
'    'Lê o Lote
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36904
'
'    If lErro = AD_SQL_SEM_DADOS Then Error 36905
'
'    lErro = Comando_LockExclusive(alComando(1))
'    If lErro <> AD_SQL_SUCESSO Then Error 36906
    
    'se for um estorno de lote
    If objLancamento_Cabecalho1.lDoc = 0 Then
    
        objLote1.sOrigem = objLancamento_Cabecalho1.sOrigem
        objLote1.iExercicio = objLancamento_Cabecalho1.iExercicio
        objLote1.iPeriodo = objLancamento_Cabecalho1.iPeriodoLote
        objLote1.iLote = objLancamento_Cabecalho1.iLote
        objLote1.iFilialEmpresa = objLancamento_Cabecalho1.iFilialEmpresa
    
        lErro = Lancamento_Grava_Estorno_Lote(objLancamento_Cabecalho, dDebito, dCredito, iNumLancAtual, iNumDocAtual, objLote1)
        If lErro <> SUCESSO Then Error 36907
        
    Else
    
        'se for um estorno de um voucher
        lErro = Lancamento_Grava_Estorno_Doc(objLancamento_Cabecalho, dDebito, dCredito, iNumLancAtual, iNumDocAtual, objLancamento_Cabecalho1)
        If lErro <> SUCESSO Then Error 36907
    
    End If
    
'    lErro = Comando_ExecutarPos(alComando(2), "UPDATE LotePendente SET TotCre = ?, TotDeb = ?, NumLancAtual = ?, NumDocAtual = ?", alComando(1), dCredito, dDebito, iNumLancAtual, iNumDocAtual)
'    If lErro <> AD_SQL_SUCESSO Then Error 36908

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Lancamento_Grava_Estorno1 = SUCESSO

    Exit Function

Erro_Lancamento_Grava_Estorno1:
    
    Lancamento_Grava_Estorno1 = Err
    
    Select Case Err
    
        Case 36900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 36901, 36907
                       
        Case 36902
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iLote, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.sOrigem)
               
        Case 36903, 36904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
            
        Case 36905
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTEPENDENTE_INEXISTENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
                    
        Case 36906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_LOTEPENDENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
                       
        Case 36908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LOTEPENDENTE", Err, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.iLote)
                       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150720)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Lancamento_Grava_Estorno_Lote(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, dCredito As Double, dDebito As Double, iNumLancAtual As Integer, iNumDocAtual As Integer, ByVal objLote As ClassLote) As Long
'grava na tabela de lançamentos pendentes os lançamentos do lote em questão com os valores invertidos (estorno)
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim alComando(0 To 1) As Long
Dim lErro As Long
Dim lErro1 As Long
Dim iIndice As Integer
Dim tLanc As typeLancamento
Dim lDoc As Long
Dim lDocEstorno As Long
Dim objLanc As ClassLancamento_Detalhe
Dim colLancamento_Detalhe As New Collection

On Error GoTo Erro_Lancamento_Grava_Estorno_Lote


    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 36909
    Next
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    'Pesquisa os Lançamentos do Lote em questão
    lErro1 = Comando_ExecutarPos(alComando(0), "SELECT Doc, Seq, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Produto, Quantidade, DataEstoque, ApropriaCRProd, Status, Gerencial FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLote = ? AND Lote = ? AND (Aglutinado = ? Or Aglutinado = ?) AND Status = 0 ORDER BY Doc", 0, _
    tLanc.lDoc, tLanc.iSeq, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iStatus, tLanc.iGerencial, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote, LANCAMENTO_AGLUTINA, LANCAMENTO_NAO_AGLUTINADO)
    If lErro1 <> AD_SQL_SUCESSO Then gError 36910
        
    'Lê o primeiro lançamento
    lErro1 = Comando_BuscarPrimeiro(alComando(0))
    If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 36911
        
    Do While lErro1 = AD_SQL_SUCESSO
    
        lDoc = tLanc.lDoc
    
        iNumDocAtual = iNumDocAtual + 1
    
        lErro = CF("Voucher_Automatico_Trans", objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.sOrigem, lDocEstorno)
        If lErro <> SUCESSO Then gError 36912
    
        objLancamento_Cabecalho.lDoc = lDocEstorno
        
        Set colLancamento_Detalhe = New Collection
    
        Do While lErro1 = AD_SQL_SUCESSO And lDoc = tLanc.lDoc
    
            'marca o lançamento como estornado
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE Lancamentos SET Status = ?", alComando(0), VOUCHER_ESTORNADO)
            If lErro <> SUCESSO Then gError 83881
    
            If objLancamento_Cabecalho.dtData < tLanc.dtData Then gError 92051
    
            Set objLanc = New ClassLancamento_Detalhe
        
            objLanc.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
            objLanc.sOrigem = objLancamento_Cabecalho.sOrigem
            objLanc.iExercicio = objLancamento_Cabecalho.iExercicio
            objLanc.iPeriodoLan = objLancamento_Cabecalho.iPeriodoLan
            objLanc.lDoc = lDocEstorno
            objLanc.iSeq = tLanc.iSeq
            objLanc.iLote = 0
            objLanc.iPeriodoLote = objLancamento_Cabecalho.iPeriodoLan
            objLanc.dtData = objLancamento_Cabecalho.dtData
            objLanc.sConta = tLanc.sConta
            objLanc.sCcl = tLanc.sCcl
            objLanc.sHistorico = tLanc.sHistorico
            objLanc.dValor = -tLanc.dValor
            objLanc.lNumIntDoc = tLanc.lNumIntDoc
            objLanc.iFilialCliForn = tLanc.iFilialCliForn
            objLanc.lCliForn = tLanc.lCliForn
            objLanc.iTransacao = tLanc.iTransacao
            objLanc.iAglutina = LANCAMENTO_NAO_AGLUTINADO
            objLanc.lContaSimples = tLanc.lContaSimples
            objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
            objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
            objLanc.dValorLivroAuxiliar = -tLanc.dValorLivroAuxiliar
            objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
            objLanc.sDocOrigem = tLanc.sDocOrigem
            objLanc.sProduto = tLanc.sProduto
            objLanc.dQuantidade = tLanc.dQuantidade
            objLanc.dtDataEstoque = tLanc.dtDataEstoque
            objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
            objLanc.iStatus = VOUCHER_ESTORNADOR
            objLanc.iGerencial = tLanc.iGerencial
        
            colLancamento_Detalhe.Add objLanc
    
            If -tLanc.dValor > 0 Then
                dCredito = dCredito + (-tLanc.dValor)
            Else
                dDebito = dDebito - (-tLanc.dValor)
            End If
        
            iNumLancAtual = iNumLancAtual + 1
        
            'Lê o proximo lançamento
            lErro1 = Comando_BuscarProximo(alComando(0))
            If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 36914
            
        Loop
        
        'contabiliza os lançamentos estornadores imediatamente
        lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe)
        If lErro <> SUCESSO Then gError 36913
        
    Loop

    dCredito = Format(dCredito, "Fixed")
    dDebito = Format(dDebito, "Fixed")
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Lancamento_Grava_Estorno_Lote = SUCESSO
    
    Exit Function
    
Erro_Lancamento_Grava_Estorno_Lote:

    Lancamento_Grava_Estorno_Lote = gErr
    
    Select Case gErr
        
        Case 36909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 36910, 36911, 36914
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS7", gErr, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
    
        Case 36912, 36913
    
        Case 83881
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANCAMENTO1", gErr, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, tLanc.lDoc, tLanc.iSeq)
    
        Case 92051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_ESTORNO_DATA_MENOR", gErr, objLancamento_Cabecalho.dtData, tLanc.dtData, objLote.iFilialEmpresa, objLote.sOrigem, tLanc.lDoc)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150721)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Lancamento_Grava_Estorno_Doc(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, dCredito As Double, dDebito As Double, iNumLancAtual As Integer, iNumDocAtual As Integer, ByVal objLancamento_Cabecalho1 As ClassLancamento_Cabecalho) As Long
'grava os lançamentos estornadores e marca os lançamentos estornados como tal
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim alComando(0 To 1) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim tLanc As typeLancamento
Dim lDoc As Long
Dim lDocEstorno As Long
Dim colLancamento_Detalhe As New Collection
Dim objLanc As ClassLancamento_Detalhe

On Error GoTo Erro_Lancamento_Grava_Estorno_Doc


    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 36915
    Next
    
    'pesquisa se o doc já foi utilizado em lançamentos contabilizados
    lErro = CF("Lancamento_Le_Doc", objLancamento_Cabecalho)
    If lErro <> SUCESSO And lErro <> 5726 Then gError 36916
    
    'Se este voucher já foi usado ==> erro
    If lErro = SUCESSO Then gError 36917
    
    'pesquisa se o doc já foi utilizado em lançamentos pendentes
    lErro = CF("LanPendente_Le_Doc1", objLancamento_Cabecalho)
    If lErro <> SUCESSO And lErro <> 9590 Then gError 41696
    
    'Se este voucher já foi usado ==> erro
    If lErro = SUCESSO Then gError 41697
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    'Pesquisa os Lançamentos do Voucher a ser extornado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Data, Seq, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Produto, Quantidade, DataEstoque, ApropriaCRProd, Status, Gerencial FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ? AND (Aglutinado = ? Or Aglutinado = ?) And Status = 0", 0, _
    tLanc.dtData, tLanc.iSeq, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iStatus, tLanc.iGerencial, objLancamento_Cabecalho1.iFilialEmpresa, objLancamento_Cabecalho1.sOrigem, objLancamento_Cabecalho1.iExercicio, objLancamento_Cabecalho1.iPeriodoLan, objLancamento_Cabecalho1.lDoc, LANCAMENTO_AGLUTINA, LANCAMENTO_NAO_AGLUTINADO)
    If lErro <> AD_SQL_SUCESSO Then gError 36918
        
    'Lê o primeiro lançamento
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 36919
        
    Do While lErro = AD_SQL_SUCESSO
    
        'marca o lançamento como estornado
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Lancamentos SET Status = ?", alComando(0), VOUCHER_ESTORNADO)
        If lErro <> SUCESSO Then gError 83879
    
        If objLancamento_Cabecalho.dtData < tLanc.dtData Then gError 92052
    
        Set objLanc = New ClassLancamento_Detalhe
    
        objLanc.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
        objLanc.sOrigem = objLancamento_Cabecalho.sOrigem
        objLanc.iExercicio = objLancamento_Cabecalho.iExercicio
        objLanc.iPeriodoLan = objLancamento_Cabecalho.iPeriodoLan
        objLanc.lDoc = objLancamento_Cabecalho.lDoc
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = 0
        objLanc.iPeriodoLote = objLancamento_Cabecalho.iPeriodoLan
        objLanc.dtData = objLancamento_Cabecalho.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = -tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        objLanc.iAglutina = LANCAMENTO_NAO_AGLUTINADO
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLanc.dValorLivroAuxiliar = -tLanc.dValorLivroAuxiliar
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.sDocOrigem = tLanc.sDocOrigem
        objLanc.sProduto = tLanc.sProduto
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
        objLanc.iStatus = VOUCHER_ESTORNADOR
        objLanc.iGerencial = tLanc.iGerencial
    
        colLancamento_Detalhe.Add objLanc
    
        If -tLanc.dValor > 0 Then
            dCredito = dCredito + (-tLanc.dValor)
        Else
            dDebito = dDebito - (-tLanc.dValor)
        End If
    
        iNumLancAtual = iNumLancAtual + 1
    
        'Lê o proximo lançamento
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 36921
            
    Loop

    'contabiliza os lançamentos estornadores imediatamente
    lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError 36920

    iNumDocAtual = iNumDocAtual + 1

    dCredito = Format(dCredito, "Fixed")
    dDebito = Format(dDebito, "Fixed")
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Lancamento_Grava_Estorno_Doc = SUCESSO
    
    Exit Function
    
Erro_Lancamento_Grava_Estorno_Doc:

    Lancamento_Grava_Estorno_Doc = gErr
    
    Select Case gErr
        
        Case 36915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 36916, 41696
        
        Case 36917
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_ATUALIZADO", gErr, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
    
        Case 36918, 36919, 36921, 188429, 188430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS8", gErr, objLancamento_Cabecalho1.iFilialEmpresa, objLancamento_Cabecalho1.sOrigem, objLancamento_Cabecalho1.iExercicio, objLancamento_Cabecalho1.iPeriodoLan, objLancamento_Cabecalho1.lDoc)
    
        Case 36920
        
        Case 41697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOC_PENDENTE", gErr, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
            
        Case 83879
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANCAMENTO1", gErr, objLancamento_Cabecalho1.iFilialEmpresa, objLancamento_Cabecalho1.sOrigem, objLancamento_Cabecalho1.iExercicio, objLancamento_Cabecalho1.iPeriodoLan, objLancamento_Cabecalho1.lDoc, tLanc.iSeq)
            
        Case 83882
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_ESTORNADO", gErr, objLancamento_Cabecalho1.iFilialEmpresa, objLancamento_Cabecalho1.sOrigem, objLancamento_Cabecalho1.iExercicio, objLancamento_Cabecalho1.iPeriodoLan, objLancamento_Cabecalho1.lDoc)
            
        Case 92052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_ESTORNO_DATA_MENOR", gErr, objLancamento_Cabecalho.dtData, tLanc.dtData, objLancamento_Cabecalho1.iFilialEmpresa, objLancamento_Cabecalho1.sOrigem, objLancamento_Cabecalho1.lDoc)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150722)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function RelDRE_Grava(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRelDRE As Collection, ByVal colRelDREConta As Collection, ByVal colRelDREFormula As Collection) As Long
'Grava no BD o modelo em questão

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RelDRE_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 44549

    'exclui os registros do modelo em questao da tabela  RelDRE
    lErro = CF("RelDRE_Exclui", sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 44550

    lErro = RelDREDetPer_Exclui(sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 44550

    'exclui os registros do modelo em questao da tabela  RelDREConta
    lErro = CF("RelDREConta_Exclui", sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 44551

    'exclui as formulas do modelo em questao da tabela  RelDREFormula
    lErro = CF("RelDREFormula_Exclui", sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 44552

    'inclui os elementos do modelo em questao da tabela  RelDRE
    lErro = CF("RelDRE_Insere", sRelatorio, sModelo, colRelDRE)
    If lErro <> SUCESSO Then Error 44553

    'inclui os elementos do tipo conta do modelo em questao da tabela  RelDREConta
    lErro = CF("RelDREConta_Insere", sRelatorio, sModelo, colRelDREConta)
    If lErro <> SUCESSO Then Error 44554

    'inclui os elementos do modelo em questao da tabela  RelDREFormula
    lErro = CF("RelDREFormula_Insere", sRelatorio, sModelo, colRelDREFormula)
    If lErro <> SUCESSO Then Error 44555

    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 44556

    RelDRE_Grava = SUCESSO

    Exit Function

Erro_RelDRE_Grava:

    RelDRE_Grava = Err

    Select Case Err

        Case 44549
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 44550, 44551, 44552, 44553, 44554, 44555

        Case 44556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150723)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function RelDRE_Exclui(ByVal sRelatorio As String, ByVal sModelo As String) As Long
'exclui os elementos do modelo em questão da tabela RelDRE
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sModelo1 As String

On Error GoTo Erro_RelDRE_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44557

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 44558

    sModelo1 = String(STRING_DRE_MODELO, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT Modelo FROM RelDRE WHERE Relatorio = ? AND Modelo = ?", 0, sModelo1, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 44559

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44560

    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_ExecutarPos(lComando1, "DELETE From RelDRE", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 44561

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44562

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    RelDRE_Exclui = SUCESSO

    Exit Function

Erro_RelDRE_Exclui:

    RelDRE_Exclui = Err

    Select Case Err

        Case 44557, 44558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44559, 44560, 44562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDRE_MODELO", Err, sModelo)

        Case 44561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELDRE", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150724)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function RelDREConta_Exclui(ByVal sRelatorio As String, ByVal sModelo As String) As Long
'exclui os elementos do modelo em questão da tabela RelDRE
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sModelo1 As String

On Error GoTo Erro_RelDREConta_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44563

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 44564

    sModelo1 = String(STRING_DRE_MODELO, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT Modelo FROM RelDREConta WHERE Relatorio = ? AND Modelo = ?", 0, sModelo1, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 44565

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44566

    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_ExecutarPos(lComando1, "DELETE From RelDREConta", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 44567

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44568

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    RelDREConta_Exclui = SUCESSO

    Exit Function

Erro_RelDREConta_Exclui:

    RelDREConta_Exclui = Err

    Select Case Err

        Case 44563, 44564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44565, 44566, 44568
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDRECONTA", Err, sModelo)

        Case 44567
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELDRECONTA", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150725)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function RelDREFormula_Exclui(ByVal sRelatorio As String, ByVal sModelo As String) As Long
'exclui os elementos do modelo em questão da tabela RelDRE
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sModelo1 As String

On Error GoTo Erro_RelDREFormula_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44569

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 44570

    sModelo1 = String(STRING_DRE_MODELO, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT Modelo FROM RelDREFormula WHERE Relatorio = ? AND Modelo = ?", 0, sModelo1, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 44571

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44572

    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_ExecutarPos(lComando1, "DELETE From RelDREFormula", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 44573

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44574

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    RelDREFormula_Exclui = SUCESSO

    Exit Function

Erro_RelDREFormula_Exclui:

    RelDREFormula_Exclui = Err

    Select Case Err

        Case 44569, 44570
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44571, 44572, 44574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDREFORMULA", Err, sModelo)

        Case 44573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELDREFORMULA", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150726)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function RelDRE_Insere(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRelDRE As Collection) As Long
'inclui os elementos de colRelDRE na tabela RelDRE
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim objRelDRE As ClassRelDRE

On Error GoTo Erro_RelDRE_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44575

    For Each objRelDRE In colRelDRE

        lErro = Comando_Executar(lComando, "INSERT INTO RelDRE ( Relatorio, Modelo, Codigo, Tipo, Nivel, Titulo, Posicao, Imprime, Exercicio ) VALUES ( ?,?,?,?,?,?,?,?,? )", sRelatorio, sModelo, objRelDRE.iCodigo, objRelDRE.iTipo, objRelDRE.iNivel, objRelDRE.sTitulo, objRelDRE.iPosicao, objRelDRE.iImprime, objRelDRE.iExercicio)
        If lErro <> AD_SQL_SUCESSO Then Error 44576

    Next

    Call Comando_Fechar(lComando)

    RelDRE_Insere = SUCESSO

    Exit Function

Erro_RelDRE_Insere:

    RelDRE_Insere = Err

    Select Case Err

        Case 44575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44576
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELDRE", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150727)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDREConta_Insere(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRelDREConta As Collection) As Long
'inclui os elementos de colRelDREConta na tabela RelDREConta
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim objRelDREConta As ClassRelDREConta

On Error GoTo Erro_RelDREConta_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44577

    For Each objRelDREConta In colRelDREConta

        lErro = Comando_Executar(lComando, "INSERT INTO RelDREConta ( Relatorio, Modelo, Codigo , Item, ContaInicial, ContaFinal, CclInicial, CclFinal ) VALUES ( ?,?,?,?,?,?,?,? )", sRelatorio, sModelo, objRelDREConta.iCodigo, objRelDREConta.iItem, objRelDREConta.sContaInicial, objRelDREConta.sContaFinal, objRelDREConta.sCclInicial, objRelDREConta.sCclFinal)
        If lErro <> AD_SQL_SUCESSO Then Error 44578

    Next

    Call Comando_Fechar(lComando)

    RelDREConta_Insere = SUCESSO

    Exit Function

Erro_RelDREConta_Insere:

    RelDREConta_Insere = Err

    Select Case Err

        Case 44577
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELDRECONTA", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150728)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDREFormula_Insere(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRelDREFormula As Collection) As Long
'inclui os elementos de colRelDREFormula na tabela RelDREFormula
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim objRelDREFormula As ClassRelDREFormula

On Error GoTo Erro_RelDREFormula_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44579

    For Each objRelDREFormula In colRelDREFormula

        lErro = Comando_Executar(lComando, "INSERT INTO RelDREFormula ( Relatorio, Modelo, Codigo , Item, CodigoFormula, Operacao, Formula ) VALUES ( ?,?,?,?,?,?,? )", sRelatorio, sModelo, objRelDREFormula.iCodigo, objRelDREFormula.iItem, objRelDREFormula.iCodigoFormula, objRelDREFormula.iOperacao, objRelDREFormula.sFormula)
        If lErro <> AD_SQL_SUCESSO Then Error 44580

    Next

    Call Comando_Fechar(lComando)

    RelDREFormula_Insere = SUCESSO

    Exit Function

Erro_RelDREFormula_Insere:

    RelDREFormula_Insere = Err

    Select Case Err

        Case 44579
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44580
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELDRE", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150729)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDRE_Exclui1(ByVal sRelatorio As String, ByVal sModelo As String) As Long
'Grava no BD o modelo em questão

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RelDRE_Exclui1

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 44583

    'exclui os registros do modelo em questao da tabela  RelDRE
    lErro = CF("RelDRE_Exclui", sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 44584

    'exclui os registros do modelo em questao da tabela  RelDREConta
    lErro = CF("RelDREConta_Exclui", sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 44585

    'exclui as formulas do modelo em questao da tabela  RelDREFormula
    lErro = CF("RelDREFormula_Exclui", sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 44586

    lErro = RelDREDetPer_Exclui(sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 44584

    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 44587

    RelDRE_Exclui1 = SUCESSO

    Exit Function

Erro_RelDRE_Exclui1:

    RelDRE_Exclui1 = Err

    Select Case Err

        Case 44583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 44584, 44585, 44586

        Case 44587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150730)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Exercicio_Instalacao_Filial(ByVal iFilialEmpresa As Integer) As Long
'Insere os registros em MvPerCta, MvPerCcl, ExercicioOrigem, ExerciciosFilial e PeriodosFilial para a filial em questão em todos os exercícios existentes.
'AVISO: ESSA ROTINA SO PODE SER CHAMADA COM A TRANSACAO ABERTA

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iExercicio As Integer
Dim iNumPeriodos As Integer
Dim iCcl As Integer
Dim lNumReg As Long
Dim iIncluiSaldo As Integer
Dim colFiliais As New Collection
Dim objAdmFiliais As New AdmFiliais

On Error GoTo Erro_Exercicio_Instalacao_Filial

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44701

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 44702

    'apenas prepara o lock na tabela Configuracao
    lErro = Comando_ExecutarPos(lComando1, "SELECT UsoCcl FROM Configuracao", 0, iCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 44703
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 44704
    
    'Lock da Configuracao, para não permitir a criacao simultanea de exercicios
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 44705

    objAdmFiliais.iCodFilial = iFilialEmpresa
    objAdmFiliais.colModulos.Add MODULO_CONTABILIDADE

    colFiliais.Add objAdmFiliais

    'Pesquisa os Exercicios. Está pesquisando inclusive os exercicios apurados ou fechados para não dar problema posterior com reabertura de exercicio
    lErro = Comando_Executar(lComando, "SELECT Exercicio, NumPeriodos FROM Exercicios E1 WHERE NOT EXISTS (SELECT * FROM ExerciciosFilial E2 WHERE FilialEmpresa = ? AND E1.Exercicio = E2.Exercicio) ORDER BY Exercicio", iExercicio, iNumPeriodos, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 44706

    'le o primeiro Exercicio aberto
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44707

    Do While lErro = AD_SQL_SUCESSO

        lErro = ExerciciosFilial_Insere_Exercicio(iExercicio, colFiliais)
        If lErro <> SUCESSO Then Error 44709

        'inicializa os saldos de conta
        lErro = MvPerCta_Insere_Exercicio(iExercicio, NAO_INCLUI_SALDO_INICIAL, colFiliais)
        If lErro <> SUCESSO Then Error 44710

        'inicializa os saldos de ccl
        lErro = MvPerCcl_Insere_Exercicio(iExercicio, NAO_INCLUI_SALDO_INICIAL, colFiliais)
        If lErro <> SUCESSO Then Error 44711

        'Insere os periodos das filiais
        lErro = CF("PeriodosFilial_Insere_Exercicio", iExercicio, iNumPeriodos, colFiliais)
        If lErro <> SUCESSO Then Error 44712
    
        'Insere exercicio origem
        lErro = ExercicioOrigem_Insere_Exercicio(iExercicio, iNumPeriodos, colFiliais)
        If lErro <> SUCESSO Then Error 44713

        'le o proximo Exercicio aberto
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44708

    Loop

    'Insere na tabela SaldoInicialConta os registros sobre saldos iniciais de todas as contas para a nova filial, a partir dos dados da EMPRESA_TODA
    lErro = SaldoInicialConta_Insere_NovaFilial(iFilialEmpresa)
    If lErro <> SUCESSO Then Error 55687
    
    'Insere na tabela SaldoInicialContaCcl os registros sobre saldos iniciais de todas os pares conta/ccl para a nova filial, a partir dos dados da EMPRESA_TODA
    lErro = SaldoInicialContaCcl_Insere_NovaFilial(iFilialEmpresa)
    If lErro <> SUCESSO Then Error 55688

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exercicio_Instalacao_Filial = SUCESSO

    Exit Function

Erro_Exercicio_Instalacao_Filial:

    Exercicio_Instalacao_Filial = Err

    Select Case Err

        Case 44701, 44702
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44703, 44704
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)

        Case 44705
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)

        Case 44706, 44707, 44708
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", Err)

        Case 44709, 44710, 44711, 44712, 44713, 55687, 55688

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150731)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Segmento_Grava_Ccl(ByVal colSegmentos As Collection) As Long
'grava os registros da colecao na tabela Segmento

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Segmento_Grava_Ccl

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 44715

    lErro = CF("Segmento_Grava_Ccl_Trans", colSegmentos)
    If lErro <> SUCESSO Then Error 44716
     
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 44717
     
    Segmento_Grava_Ccl = SUCESSO
    
    Exit Function

Erro_Segmento_Grava_Ccl:

    Segmento_Grava_Ccl = Err

    Select Case Err

        Case 44715
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 44716

        Case 44717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150732)

    End Select
    
    Call Transacao_Rollback

    Exit Function

End Function

Function Segmento_Grava_Ccl_Trans(ByVal colSegmentos As Collection) As Long
'grava os registros da colecao na tabela Segmento de Ccl
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim lErro As Long
Dim iPossui_Ccl As Integer
Dim iExercicio As Integer
Dim iIndice As Integer
Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim tSegmento As typeSegmento
Dim sCodigo As String
Dim sCodigo2 As String
Dim objSegmento As New ClassSegmento
Dim iCcl As Integer

On Error GoTo Erro_Segmento_Grava_Ccl_Trans

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44718
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 44719
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 44720
  
    sCodigo = String(STRING_SEGMENTO_CODIGO, 0)
    sCodigo2 = String(STRING_SEGMENTO_CODIGO, 0)
    tSegmento.sDelimitador = String(STRING_SEGMENTO_DELIMITADOR, 0)
    
    'pega um objeto da colSegmentos para saber qual o formato( = campo codigo na tabela Segmento)
    For Each objSegmento In colSegmentos
        'guarda o formato em sCodigo para depois fazer o Select na tabela Segmentos
        sCodigo = objSegmento.sCodigo
        Exit For
    Next
       
    If objSegmento.sCodigo <> SEGMENTO_CCL Then Error 44721
       
    'verifica existencia de pelo menos uma centro de custo cadastrado
    lErro = CF("Ccl_ExisteCcl", iPossui_Ccl)
    If lErro <> SUCESSO Then Error 44722

    If iPossui_Ccl = POSSUI_CCL Then
    'so faz modificacao no delimitador e no preenchimento da tabela Segmento
        
         lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, Nivel, Delimitador, Preenchimento FROM Segmento WHERE Codigo = ? ORDER BY Nivel", 0, sCodigo2, tSegmento.iNivel, tSegmento.sDelimitador, tSegmento.iPreenchimento, sCodigo)
         If lErro <> AD_SQL_SUCESSO Then Error 44723
     
         'percorre cada registro em colSegmentos para fazer a modificacao nos delimitadores e no preenchimento
         For Each objSegmento In colSegmentos
                     
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO Then Error 44724
            
            lErro = Comando_ExecutarPos(lComando2, "UPDATE Segmento SET Codigo =? , Nivel =? , Delimitador =? , Preenchimento =? ", lComando, sCodigo, objSegmento.iNivel, objSegmento.sDelimitador, objSegmento.iPreenchimento)
            If lErro <> AD_SQL_SUCESSO Then Error 44725
            
         Next
    
    Else
        'apaga toda tabela Segmento e insere os dados novos em Segmento
        
        'prepara o lock na tabela Configuracao
        lErro = Comando_ExecutarPos(lComando3, "SELECT UsoCcl FROM Configuracao", 0, iCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 44726
    
        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44727
    
        'Lock da Configuracao
        lErro = Comando_LockExclusive(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 44728
    
        'Seleciona registros na tabela de Segmento
        lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, Nivel FROM Segmento WHERE Codigo = ?", 0, sCodigo2, tSegmento.iNivel, sCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 44729
        
        'le o primeiro segmento
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44730
        
        'verifica se ainda existe registro a ser deletado
        Do While lErro <> AD_SQL_SEM_DADOS
           
            'exclui o segmento
            lErro = Comando_ExecutarPos(lComando2, "DELETE FROM Segmento", lComando)
            If lErro <> AD_SQL_SUCESSO Then Error 44732
        
            'le o proximo segmento
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44731
        
        Loop
        
        For Each objSegmento In colSegmentos
           
           lErro = Comando_Executar(lComando2, "INSERT INTO Segmento (Codigo,Nivel,Tipo,Tamanho,Delimitador,Preenchimento) VALUES (?,?,?,?,?,?)", sCodigo, objSegmento.iNivel, objSegmento.iTipo, objSegmento.iTamanho, objSegmento.sDelimitador, objSegmento.iPreenchimento)
           If lErro <> AD_SQL_SUCESSO Then Error 44733
           
        Next
        
    End If
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
     
    Segmento_Grava_Ccl_Trans = SUCESSO
    
    Exit Function

Erro_Segmento_Grava_Ccl_Trans:

    Segmento_Grava_Ccl_Trans = Err

    Select Case Err

        Case 44718, 44719, 44720
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 44721
            Call Rotina_Erro(vbOKOnly, "ERRO_SEGMENTO_CCL_INVALIDO", Err, sCodigo)

        Case 44722

        Case 44723, 44724, 44729, 44730, 44731
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SEGMENTO1", Err, sCodigo)
             
        Case 44725
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_SEGMENTO", Err, sCodigo, tSegmento.iNivel)
             
        Case 44726, 44727
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
        
        Case 44728
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)

        Case 44732
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_SEGMENTO", Err, sCodigo, tSegmento.iNivel)

        Case 44733
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SEGMENTO", Err, sCodigo, objSegmento.iNivel)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150733)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Voucher_Automatico_Obter_Batch(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal sOrigem As String, lDoc As Long, lComando As Long) As Long
'retorna em lDoc o proximo numero p/criacao de lancamentos pendentes
'Não verifica a tabela de Lancamentos ou Lancamentos Pendentes pois a origem tem que ser de lancamentos automaticos.
'ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long

On Error GoTo Erro_Voucher_Automatico_Obter_Batch

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 32138
    
    'se a numeração do voucher(documento) não é reinicializada a cada periodo ==> não leva em consideração o periodo passado como parametro
    If giSetupDocPorPeriodo = 0 Then
        iPeriodo = 0
    End If
    
    'Pesquisa o ultimo número de voucher gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT Doc FROM ExercicioOrigem WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ? AND Origem = ?", 0, lDoc, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 32140

    'le o registro com o número do voucher
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 32141

    'loca o registro que guarda o numero do voucher gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 32142
    
    Voucher_Automatico_Obter_Batch = SUCESSO
        
    Exit Function
        
Erro_Voucher_Automatico_Obter_Batch:

    Voucher_Automatico_Obter_Batch = Err

    Select Case Err
    
        Case 32138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 32140, 32141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOORIGEM", Err, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
        
        Case 32142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOORIGEM", Err, iFilialEmpresa, iExercicio, iPeriodo, sOrigem)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150734)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Function Voucher_Automatico_Gravar_Batch(ByVal objContabAutoAux As ClassContabAutoAux, ByVal lComando As Long, ByVal sOrigem As String) As Long
'grava o proximo numero p/criacao de lancamentos pendentes com o valor de lDoc
'Não verifica a tabela de Lancamentos ou Lancamentos Pendentes pois a origem tem que ser de lancamentos automaticos.
'ATENCAO: PARA EXECUTAR ESTA ROTINA O PROGRAMA TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando1 As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim iFilialEmpresa As Integer
Dim lDoc As Long

On Error GoTo Erro_Voucher_Automatico_Gravar_Batch

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 32139
        
    'atualiza o numero do voucher gerado automaticamente
    lErro = Comando_ExecutarPos(lComando1, "UPDATE ExercicioOrigem SET Doc = ?", lComando, objContabAutoAux.lProxDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 32143
        
    If giContabGerencial = 1 Then
        
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 193318
        Next
        
        If objContabAutoAux.iFilial > Abs(giFilialAuxiliar) Then
            iFilialEmpresa = objContabAutoAux.iFilial - Abs(giFilialAuxiliar)
        Else
            iFilialEmpresa = objContabAutoAux.iFilial + Abs(giFilialAuxiliar)
        End If
            
        'Pesquisa o ultimo número de voucher gerado automaticamente pelo sistema
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Doc FROM ExercicioOrigem WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ? AND Origem = ?", 0, lDoc, iFilialEmpresa, objContabAutoAux.iExercicio, objContabAutoAux.iPeriodo, sOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 193319
    
        'le o registro com o número do voucher
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 193320
    
        'loca o registro que guarda o numero do voucher gerado automaticamente
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 193321
        
        'atualiza o numero do voucher gerado automaticamente
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ExercicioOrigem SET Doc = ?", alComando(1), objContabAutoAux.lProxDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 193322
        
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Voucher_Automatico_Gravar_Batch = SUCESSO
        
    Exit Function
        
Erro_Voucher_Automatico_Gravar_Batch:

    Voucher_Automatico_Gravar_Batch = gErr

    Select Case gErr
    
        Case 32139, 193318
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 193319, 193320
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOORIGEM", gErr, iFilialEmpresa, objContabAutoAux.iExercicio, objContabAutoAux.iPeriodo, sOrigem)
        
        Case 193321
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOORIGEM", gErr, iFilialEmpresa, objContabAutoAux.iExercicio, objContabAutoAux.iPeriodo, sOrigem)
        
        Case 32143, 193322
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOORIGEM_BATCH", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150735)
        
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function
    
End Function

Function LotePendente_Grava_Trans(ByVal objLote As ClassLote) As Long
'grava o lote no banco de dados dentro de uma transacao
    
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lErro As Long
Dim iStatus As Integer
Dim iLoteAtualizado As Integer
Dim iIdAtualizacao As Integer
Dim objPeriodosFilial As New ClassPeriodosFilial
    
On Error GoTo Erro_LotePendente_Grava_Trans

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5410

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5415

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5745

    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 5770

    'verifica se o exercicio não está fechado
    lErro = Comando_ExecutarLockado(lComando3, "SELECT Status FROM Exercicios WHERE Exercicio=?", iStatus, objLote.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5746
    
    'le o exercicio, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO Then Error 5747
    
    lErro = Comando_LockShared(lComando3)
    If lErro <> AD_SQL_SUCESSO Then Error 5748
    
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If iStatus = EXERCICIO_FECHADO Then Error 5749
    End If
        
    objPeriodosFilial.iFilialEmpresa = objLote.iFilialEmpresa
    objPeriodosFilial.iExercicio = objLote.iExercicio
    objPeriodosFilial.iPeriodo = objLote.iPeriodo
    objPeriodosFilial.sOrigem = objLote.sOrigem

    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then Error 20855
    
    'se o periodo estiver fechado, não pode alterar os lançamentos do periodo
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then Error 5774
    End If
        
    'verifica se o lote  está atualizado
    lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
    If lErro <> SUCESSO Then Error 5992
    
    If iLoteAtualizado = LOTE_ATUALIZADO Then Error 5417
        
    'pesquisa no BD o lote em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT IdAtualizacao FROM LotePendente WHERE FilialEmpresa=? AND Origem=? AND Exercicio=? AND Periodo=? AND Lote=?", 0, iIdAtualizacao, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 5411

    'le o lote, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    
    'encontrou o lote ==> isto é uma alteração
    If lErro = AD_SQL_SUCESSO Then
            
        If iIdAtualizacao > 0 Then Error 10268
        
        'atualiza o lote
        lErro = Comando_ExecutarPos(lComando2, "UPDATE LotePendente SET TotInf=?, NumDocInf=?, IdOriginal=?, Status=? ", lComando1, objLote.dTotInf, objLote.iNumDocInf, objLote.sIdOriginal, objLote.iStatus)
        If lErro <> AD_SQL_SUCESSO Then Error 5418
        
    Else
        'não encontrou o lote ==>  é uma inserção de lote
        
        'verifica se já existe algum lote com essa numeração neste exercicio
        'caso esteja trabalhando com numeracao de lote incializando por exercicio
        lErro = CF("Lote_LotePendente_Le", objLote)
        If lErro <> SUCESSO Then Error 20769
    
        'insere o lote
        lErro = Comando_Executar(lComando2, "INSERT INTO LotePendente (FilialEmpresa,Origem,Exercicio,Periodo,Lote,TotInf,Status,IdOriginal,NumDocInf, IdAtualizacao) VALUES (?,?,?,?,?,?,?,?,?,?)", objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote, objLote.dTotInf, objLote.iStatus, objLote.sIdOriginal, objLote.iNumDocInf, 0)
        If lErro <> AD_SQL_SUCESSO Then Error 5419
    
    End If
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    LotePendente_Grava_Trans = SUCESSO

    Exit Function

Erro_LotePendente_Grava_Trans:
    
    LotePendente_Grava_Trans = Err
    
    Select Case Err
    
        Case 5410, 5415, 5745, 5770
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 5417
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_EDITAVEL", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 5418
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            
        Case 5419
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 5746, 5747
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO2", Err, objLote.iExercicio)
            
        Case 5748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", Err, objLote.iExercicio)
        
        Case 5749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_EXERCICIO_FECHADO", Err, objLote.iExercicio)
        
        Case 5771, 5772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo)
            
        Case 5773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODOSFILIAL", Err, objLote.iFilialEmpresa, objLote.iExercicio, objLote.iPeriodo)
        
        Case 5774
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_PERIODO_FECHADO", Err, objLote.iExercicio, objLote.iPeriodo)
        
        Case 5992, 20769
        
        Case 10268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_SENDO_ATUALIZADO", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150736)
        
    End Select
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    
    Exit Function

End Function

Function LotePendente_Grava_Totais_Trans(ByVal objLote As ClassLote) As Long
'atualiza qtdes e valoes de um lote dentro de uma transacao

Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim iStatus As Integer
Dim iLoteAtualizado As Integer
    
On Error GoTo Erro_LotePendente_Grava_Totais_Trans

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 8036

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 8037

    'verifica se o lote  está atualizado
    lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
    If lErro <> SUCESSO Then Error 5994
    
    If iLoteAtualizado = LOTE_ATUALIZADO Then Error 8040
    
    'pesquisa no BD o lote em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Status FROM LotePendente WHERE FilialEmpresa = ? AND Origem=? AND Exercicio=? AND Periodo=? AND Lote=?", 0, iStatus, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 8038

    'le o lote, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 8042
    
    'atualiza o lote
    lErro = Comando_ExecutarPos(lComando2, "UPDATE LotePendente SET TotCre=?, TotDeb=?, NumDocAtual=?", lComando1, objLote.dTotCre, objLote.dTotDeb, objLote.iNumDocAtual)
    If lErro <> AD_SQL_SUCESSO Then Error 8041
        
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando1)
    
    LotePendente_Grava_Totais_Trans = SUCESSO

    Exit Function

Erro_LotePendente_Grava_Totais_Trans:
    
    LotePendente_Grava_Totais_Trans = Err
    
    Select Case Err
    
        Case 5994
            
        Case 8036, 8037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 8038, 8042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 8040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_EDITAVEL", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 8041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150737)
        
    End Select
    
     Call Comando_Fechar(lComando1)
     Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function LotePendente_Grava_Totais_Auto(ByVal objLote As ClassLote) As Long
'atualiza qtdes e valoes de um lote automatico dentro de uma transacao
'ao contrario da gravacao normal vai atualizar tambem os valores "informados"

Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim iStatus As Integer
Dim iLoteAtualizado As Integer
    
On Error GoTo Erro_LotePendente_Grava_Totais_Auto

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 32170

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 32171

    'verifica se o lote  está atualizado
    lErro = CF("Lote_Critica_Atualizado", objLote, iLoteAtualizado)
    If lErro <> SUCESSO Then Error 32172
    
    If iLoteAtualizado = LOTE_ATUALIZADO Then Error 32173
    
    'pesquisa no BD o lote em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Status FROM LotePendente WHERE FilialEmpresa = ? AND Origem=? AND Exercicio=? AND Periodo=? AND Lote=?", 0, iStatus, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 32174

    'le o lote, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 32175
    
    'atualiza o lote
    lErro = Comando_ExecutarPos(lComando2, "UPDATE LotePendente SET TotCre=?, TotDeb=?, NumDocAtual=?, TotInf=?, NumDocInf=?, NumLancInf = ?", lComando1, objLote.dTotCre, objLote.dTotDeb, objLote.iNumDocAtual, objLote.dTotInf, objLote.iNumDocInf, objLote.iNumLancInf)
    If lErro <> AD_SQL_SUCESSO Then Error 32176
        
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando1)
    
    LotePendente_Grava_Totais_Auto = SUCESSO

    Exit Function

Erro_LotePendente_Grava_Totais_Auto:
    
    LotePendente_Grava_Totais_Auto = Err
    
    Select Case Err
    
        Case 32172
            
        Case 32170, 32171
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 32174, 32175
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 32173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_ATUALIZADO_NAO_EDITAVEL", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
        
        Case 32176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LOTEPENDENTE", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150738)
        
    End Select
    
     Call Comando_Fechar(lComando1)
     Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function ConfiguracaoSetup_Altera_Trans(ByVal objConfiguracao As ClassConfiguracao) As Long
'grava os registros do objeto na tabela Configuracao
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim iCcl As Integer
Dim iIndice As Integer
Dim lComando As Long
Dim lComando2 As Long

On Error GoTo Erro_ConfiguracaoSetup_Altera_Trans

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44977
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 44978
    
    'apenas prepara o lock na tabela Configuracao
    lErro = Comando_ExecutarPos(lComando, "SELECT UsoCcl FROM Configuracao", 0, iCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 44979
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44980
    
    'Lock da Configuracao
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 44981
           
    'faz as alteracoes em TipoConta e Natureza na tabela de Configurarcao
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Configuracao SET TipoContaDefault =? , NaturezaDefault =?, DocPorPeriodo=?, LotePorPeriodo=?, UsoCcl=?, ContabCentralizada=?", lComando, objConfiguracao.iTipoContaDefault, objConfiguracao.iNaturezaDefault, objConfiguracao.iDocPorPeriodo, objConfiguracao.iLotePorPeriodo, objConfiguracao.iUsoCcl, objConfiguracao.iContabCentralizada)
    If lErro <> AD_SQL_SUCESSO Then Error 44982
                     
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    ConfiguracaoSetup_Altera_Trans = SUCESSO
    
    Exit Function

Erro_ConfiguracaoSetup_Altera_Trans:

    ConfiguracaoSetup_Altera_Trans = Err

    Select Case Err

        Case 44977
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44978
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44979, 44980
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)

        Case 44981
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)

        Case 44982
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_CONFIGURACAO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150739)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function RelDRE_Grava_Valor(ByVal sRelatorio As String, ByVal colRelDRE As Collection, Optional ByVal iDetalhado As Integer) As Long
'Grava o Valor acumulado do Modelo na tabela RelDRE
    
Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim sModelo As String
Dim iCodigo As Integer
Dim objRelDRE As ClassRelDRE
Dim iTransacaoAbertaAnterior As Integer

On Error GoTo Erro_RelDRE_Grava_Valor

    'Inicializa comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 43596

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 43597

    'Inicia transação
    If GL_lTransacao = 0 Then
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 43598
    Else
        iTransacaoAbertaAnterior = 1
    End If
    
    If colRelDRE.Count > 0 Then
        If iDetalhado = MARCADO Then
            lErro = RelDREDetPer_Exclui(RELDRE, colRelDRE.Item(1).sModelo)
            If lErro <> SUCESSO Then Error 43589
        End If
    End If

    For Each objRelDRE In colRelDRE

        If iDetalhado = MARCADO Then
        
            If Not (objRelDRE.objRelDREDetPer Is Nothing) Then
                lErro = RelDREDetPer_Insere(objRelDRE.objRelDREDetPer)
                If lErro <> SUCESSO Then gError 200871
            End If
            
        End If
        
        sModelo = String(STRING_RELDRE_MODELO, 0)
    
        lErro = Comando_ExecutarPos(lComando1, "SELECT Modelo, Codigo FROM RelDRE WHERE Relatorio = ? AND Modelo = ? AND Codigo = ?", 0, sModelo, iCodigo, sRelatorio, objRelDRE.sModelo, objRelDRE.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 43599
        
        'Lê RelDRE se estiver no BD
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 43600
        
        'Se não encontrou o Modelo --> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 43601
        
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 43602
        
        'Atualiza o Valor do Modelo
        lErro = Comando_ExecutarPos(lComando2, "UPDATE RelDRE SET Valor = ?, ValorExercAnt = ?, ValorPerAcumAnt = ?, ValorPerAcumAtual = ?", lComando1, objRelDRE.dValor, objRelDRE.dValorExercAnt, objRelDRE.dValorPerAcumAnt, objRelDRE.dValorPerAcumAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 43603
        
    Next
        
    If iTransacaoAbertaAnterior = 0 Then
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 43604
    End If
    
    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    RelDRE_Grava_Valor = SUCESSO

    Exit Function

Erro_RelDRE_Grava_Valor:
    
    RelDRE_Grava_Valor = Err
    
    Select Case Err
    
        Case 43596, 43597
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 43598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 43599, 43600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDRE_MODELO", gErr, objRelDRE.sModelo)
        
        Case 43601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELDRE_MODELO_NAO_CADASTRADO", gErr, objRelDRE.sModelo, objRelDRE.iCodigo)
        
        Case 43602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RELDRE", gErr, objRelDRE.sModelo, objRelDRE.iCodigo)
        
        Case 43603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RELDRE", gErr, objRelDRE.sModelo, objRelDRE.iCodigo)
            
        Case 43604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 200871
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150740)
        
    End Select
    
    If iTransacaoAbertaAnterior = 0 Then
        Call Transacao_Rollback
    End If
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function MnemonicoCTBValor_Exclui() As Long
'Exclui toda a tabela MnemonicoCTBValor
'IMPORTANTE: TEM QUE ESTAR EM TRANSACAO

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_MnemonicoCTBValor_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 39780
           
    lErro = Comando_Executar(lComando, "DELETE FROM MnemonicoCTBValor")
    If lErro <> AD_SQL_SUCESSO Then Error 39781
    
    Call Comando_Fechar(lComando)
    
    MnemonicoCTBValor_Exclui = SUCESSO

        Exit Function

Erro_MnemonicoCTBValor_Exclui:

    MnemonicoCTBValor_Exclui = Err
    
    Select Case Err
        
        Case 39780
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 39781
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MNEMONICOCTBVALOR", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150741)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function MnemonicoCTBValor_Grava(ByVal colMnemonico As Collection) As Long
'Grava os dados  passada como parâmetro, no BD

Dim lErro As Long
Dim lComando As Long
Dim lTransacao As Long
Dim objMnemonico As ClassMnemonicoCTBValor

On Error GoTo Erro_MnemonicoCTBValor_Grava
        
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 39787
        
    lErro = CF("MnemonicoCTBValor_Exclui")
    If lErro <> SUCESSO Then Error 39783
        
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 39782
    
    For Each objMnemonico In colMnemonico
    
        'insere novos registros na tabela MnemonicoCTBValor
        lErro = Comando_Executar(lComando, "INSERT INTO MnemonicoCTBValor (Mnemonico,Descricao,Valor) VALUES (?,?,?)", objMnemonico.sMnemonico, objMnemonico.sDescricao, objMnemonico.sValor)
        If lErro <> AD_SQL_SUCESSO Then Error 39786
    
    Next
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 39788
    
    Call Comando_Fechar(lComando)
    
    MnemonicoCTBValor_Grava = SUCESSO

    Exit Function

Erro_MnemonicoCTBValor_Grava:

    MnemonicoCTBValor_Grava = Err

    Select Case Err
        
        Case 39782
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 39783
            
        Case 39786
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MNEMONICOCTBVALOR", Err)
        
        Case 39787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 39788
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150742)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Configuracao_Altera_DataInstalacao() As Long
'grava os registros do objeto na tabela Configuracao
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim iCcl As Integer
Dim iIndice As Integer
Dim lComando As Long
Dim lComando2 As Long

On Error GoTo Erro_Configuracao_Altera_DataInstalacao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55237
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 55238
    
    'apenas prepara o lock na tabela Configuracao
    lErro = Comando_ExecutarPos(lComando, "SELECT UsoCcl FROM Configuracao", 0, iCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 55239
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 55240
    
    'Lock da Configuracao
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 55241
           
    'faz as alteracoes em TipoConta e Natureza na tabela de Configurarcao
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Configuracao SET DataImplantacao =?", lComando, Date)
    If lErro <> AD_SQL_SUCESSO Then Error 55242
                     
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Configuracao_Altera_DataInstalacao = SUCESSO
    
    Exit Function

Erro_Configuracao_Altera_DataInstalacao:

    Configuracao_Altera_DataInstalacao = Err

    Select Case Err

        Case 55237, 55238
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55239, 55240
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)

        Case 55241
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)

        Case 55242
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_CONFIGURACAO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150743)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function RelDMPL_Exclui1(ByVal sRelatorio As String, ByVal sModelo As String) As Long
'Exclui do BD o modelo em questão

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RelDMPL_Exclui1

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 60538

    'exclui os registros do modelo em questao da tabela  RelDRE
    lErro = RelDMPL_Exclui(sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 60539

    'exclui os registros do modelo em questao da tabela  RelDREConta
    lErro = RelDMPLConta_Exclui(sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 60540

    'exclui as formulas do modelo em questao da tabela  RelDREFormula
    lErro = RelDMPLFormula_Exclui(sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 60541

    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 60542

    RelDMPL_Exclui1 = SUCESSO

    Exit Function

Erro_RelDMPL_Exclui1:

    RelDMPL_Exclui1 = Err

    Select Case Err

        Case 60538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 60539, 60540, 60541

        Case 60542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150744)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Private Function RelDMPL_Exclui(ByVal sRelatorio As String, ByVal sModelo As String) As Long
'exclui os elementos do modelo em questão da tabela RelDMPL
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sModelo1 As String

On Error GoTo Erro_RelDMPL_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60543

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 60544

    sModelo1 = String(STRING_DRE_MODELO, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT Modelo FROM RelDMPL WHERE Relatorio = ? AND Modelo = ?", 0, sModelo1, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 60545

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60546

    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_ExecutarPos(lComando1, "DELETE From RelDMPL", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 60547

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60548

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    RelDMPL_Exclui = SUCESSO

    Exit Function

Erro_RelDMPL_Exclui:

    RelDMPL_Exclui = Err

    Select Case Err

        Case 60543, 60544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 60545, 60546, 60548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDMPL_MODELO", Err, sModelo)

        Case 60547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELDMPL", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150745)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Function RelDMPLConta_Exclui(ByVal sRelatorio As String, ByVal sModelo As String) As Long
'exclui os elementos do modelo em questão da tabela RelDRE
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sModelo1 As String

On Error GoTo Erro_RelDMPLConta_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60549

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 60550

    sModelo1 = String(STRING_DRE_MODELO, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT Modelo FROM RelDMPLConta WHERE Relatorio=? AND Modelo = ?", 0, sModelo1, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 60551

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60552

    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_ExecutarPos(lComando1, "DELETE From RelDMPLConta", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 60553

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60554

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    RelDMPLConta_Exclui = SUCESSO

    Exit Function

Erro_RelDMPLConta_Exclui:

    RelDMPLConta_Exclui = Err

    Select Case Err

        Case 60549, 60550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 60551, 60552, 60554
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDMPLCONTA", Err, sModelo)

        Case 60553
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELDMPLCONTA", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150746)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Function RelDMPLFormula_Exclui(ByVal sRelatorio As String, ByVal sModelo As String) As Long
'exclui os elementos do modelo em questão da tabela RelDRE
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sModelo1 As String

On Error GoTo Erro_RelDMPLFormula_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60555

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 60556

    sModelo1 = String(STRING_DRE_MODELO, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT Modelo FROM RelDMPLFormula WHERE Relatorio=? AND Modelo = ?", 0, sModelo1, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then Error 60557

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60558

    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_ExecutarPos(lComando1, "DELETE From RelDMPLFormula", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 60559

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60560

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    RelDMPLFormula_Exclui = SUCESSO

    Exit Function

Erro_RelDMPLFormula_Exclui:

    RelDMPLFormula_Exclui = Err

    Select Case Err

        Case 60555, 60556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 60557, 60558, 60560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDMPLFORMULA", Err, sModelo)

        Case 60559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELDMPLFORMULA", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150747)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function RelDMPL_Grava(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRel As Collection, ByVal colRelConta As Collection, ByVal colRelFormula As Collection) As Long
'Grava no BD o modelo em questão

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RelDMPL_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 60632

    'exclui os registros do modelo em questao da tabela  RelDMPL
    lErro = RelDMPL_Exclui(sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 60633

    'exclui os registros do modelo em questao da tabela  RelDMPLConta
    lErro = RelDMPLConta_Exclui(sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 60634

    'exclui as formulas do modelo em questao da tabela  RelDMPLFormula
    lErro = RelDMPLFormula_Exclui(sRelatorio, sModelo)
    If lErro <> SUCESSO Then Error 60635

    'inclui os elementos do modelo em questao da tabela  RelDMPL
    lErro = RelDMPL_Insere(sRelatorio, sModelo, colRel)
    If lErro <> SUCESSO Then Error 60636

    'inclui os elementos do tipo conta do modelo em questao da tabela  RelDREConta
    lErro = RelDMPLConta_Insere(sRelatorio, sModelo, colRelConta)
    If lErro <> SUCESSO Then Error 60637

    'inclui os elementos do modelo em questao da tabela  RelDREFormula
    lErro = RelDMPLFormula_Insere(sRelatorio, sModelo, colRelFormula)
    If lErro <> SUCESSO Then Error 60638

    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 60639

    RelDMPL_Grava = SUCESSO

    Exit Function

Erro_RelDMPL_Grava:

    RelDMPL_Grava = Err

    Select Case Err

        Case 60632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 60633, 60634, 60635, 60636, 60637, 60638

        Case 60639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150748)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Private Function RelDMPL_Insere(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRel As Collection) As Long
'inclui os elementos de colRel na tabela RelDMPL
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim objRel As ClassRelDMPL

On Error GoTo Erro_RelDMPL_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60640

    For Each objRel In colRel

        lErro = Comando_Executar(lComando, "INSERT INTO RelDMPL ( Relatorio, Modelo, Linha, Coluna, Tipo, Titulo, Exercicio) VALUES (?,?,?,?,?,?,?)", sRelatorio, sModelo, objRel.iLinha, objRel.iColuna, objRel.iTipo, objRel.sTitulo, objRel.iExercicio)
        If lErro <> AD_SQL_SUCESSO Then Error 60641

    Next

    Call Comando_Fechar(lComando)

    RelDMPL_Insere = SUCESSO

    Exit Function

Erro_RelDMPL_Insere:

    RelDMPL_Insere = Err

    Select Case Err

        Case 60640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 60641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELDMPL", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150749)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function RelDMPLConta_Insere(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRelConta As Collection) As Long
'inclui os elementos de colRelDREConta na tabela RelDMPLConta
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim objRelConta As ClassRelDMPLConta

On Error GoTo Erro_RelDMPLConta_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60642

    For Each objRelConta In colRelConta

        lErro = Comando_Executar(lComando, "INSERT INTO RelDMPLConta ( Relatorio, Modelo, Linha, Coluna , Item, ContaInicial, ContaFinal ) VALUES ( ?,?,?,?,?,?,? )", sRelatorio, sModelo, objRelConta.iLinha, objRelConta.iColuna, objRelConta.iItem, objRelConta.sContaInicial, objRelConta.sContaFinal)
        If lErro <> AD_SQL_SUCESSO Then Error 60643

    Next

    Call Comando_Fechar(lComando)

    RelDMPLConta_Insere = SUCESSO

    Exit Function

Erro_RelDMPLConta_Insere:

    RelDMPLConta_Insere = Err

    Select Case Err

        Case 60642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 60643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELDMPLCONTA", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150750)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function RelDMPLFormula_Insere(ByVal sRelatorio As String, ByVal sModelo As String, ByVal colRelFormula As Collection) As Long
'inclui os elementos de colRelDREFormula na tabela RelDREFormula
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim objRelFormula As ClassRelDMPLFormula

On Error GoTo Erro_RelDMPLFormula_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60644

    For Each objRelFormula In colRelFormula

        lErro = Comando_Executar(lComando, "INSERT INTO RelDMPLFormula ( Relatorio, Modelo, Linha, Coluna , Item, LinhaFormula, ColunaFormula, Operacao ) VALUES ( ?,?,?,?,?,?,?,? )", sRelatorio, sModelo, objRelFormula.iLinha, objRelFormula.iColuna, objRelFormula.iItem, objRelFormula.iLinhaFormula, objRelFormula.iColunaFormula, objRelFormula.iOperacao)
        If lErro <> AD_SQL_SUCESSO Then Error 60645

    Next

    Call Comando_Fechar(lComando)

    RelDMPLFormula_Insere = SUCESSO

    Exit Function

Erro_RelDMPLFormula_Insere:

    RelDMPLFormula_Insere = Err

    Select Case Err

        Case 60644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 60645
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELDMPLFORMULA", Err, sModelo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150751)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RelDMPL_Grava_Valor(ByVal sRelatorio As String, ByVal colRel As Collection) As Long
'Grava o Valor acumulado do Modelo na tabela RelDMPL
    
Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim sModelo As String
Dim objRel As ClassRelDMPL

On Error GoTo Erro_RelDMPL_Grava_Valor

    'Inicializa comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 60681

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 60682

    'Inicia transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 60683

    For Each objRel In colRel

        sModelo = String(STRING_RELDRE_MODELO, 0)
    
        lErro = Comando_ExecutarPos(lComando1, "SELECT Modelo FROM RelDMPL WHERE Relatorio = ? AND Modelo = ? AND Linha = ? AND Coluna = ?", 0, sModelo, sRelatorio, objRel.sModelo, objRel.iLinha, objRel.iColuna)
        If lErro <> AD_SQL_SUCESSO Then Error 60684
        
        'Lê RelDRE se estiver no BD
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60685
        
        'Se não encontrou o Modelo --> Erro
        If lErro = AD_SQL_SEM_DADOS Then Error 60686
        
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 60687
        
        'Atualiza o Valor do Modelo
        lErro = Comando_ExecutarPos(lComando2, "UPDATE RelDMPL SET Valor = ?", lComando1, objRel.dValor)
        If lErro <> AD_SQL_SUCESSO Then Error 60688
        
    Next
        
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 60689
    
    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    RelDMPL_Grava_Valor = SUCESSO

    Exit Function

Erro_RelDMPL_Grava_Valor:
    
    RelDMPL_Grava_Valor = Err
    
    Select Case Err
    
        Case 60681, 60682
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 60683
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 60684, 60685
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDMPL_MODELO", Err, objRel.sModelo)
        
        Case 60686
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELDMPL_MODELO_NAO_CADASTRADO", Err, objRel.sModelo, objRel.iLinha, objRel.iColuna)
        
        Case 60687
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RELDMPL", Err, objRel.sModelo, objRel.iLinha, objRel.iColuna)
        
        Case 60688
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RELDMPL", Err, objRel.sModelo, objRel.iLinha, objRel.iColuna)
            
        Case 60689
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150752)
        
    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function Lancamento_Grava_Estorno_Doc_Basico(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal objLancamento_Cabecalho1 As ClassLancamento_Cabecalho) As Long
'Faz Atualiza os lancamentos
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim lErro As Long
Dim tLanc As typeLancamento
Dim colLancamento_Detalhe As New Collection
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim alComando(0 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Lancamento_Grava_Estorno_Doc_Basico

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 61160
    Next

    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    
    'reprocessamento
    'Pesquisa os Lançamentos do Voucher a ser extornado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Seq, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, ContaSimples, SeqContraPartida, Produto, ApropriaCRProd, Quantidade, DataEstoque, EscaninhoCusto, Status FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ? AND (Aglutinado = ? Or Aglutinado = ?) AND Status = 0", 0, _
    tLanc.iSeq, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.iApropriaCRProd, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iEscaninho_Custo, tLanc.iStatus, objLancamento_Cabecalho1.iFilialEmpresa, objLancamento_Cabecalho1.sOrigem, objLancamento_Cabecalho1.iExercicio, objLancamento_Cabecalho1.iPeriodoLan, objLancamento_Cabecalho1.lDoc, LANCAMENTO_AGLUTINA, LANCAMENTO_NAO_AGLUTINADO)
    If lErro <> AD_SQL_SUCESSO Then gError 61161
        
    'Lê o primeiro lançamento
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 61162
    
    If lErro = AD_SQL_SEM_DADOS Then gError 83883
    
    Do While lErro = AD_SQL_SUCESSO
        
        'marca o lançamento como estornado
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Lancamentos SET Status = ?", alComando(0), VOUCHER_ESTORNADO)
        If lErro <> SUCESSO Then gError 83880
        
        Set objLancamento_Detalhe = New ClassLancamento_Detalhe
        
        objLancamento_Detalhe.iSeq = tLanc.iSeq
        objLancamento_Detalhe.sConta = tLanc.sConta
        objLancamento_Detalhe.sCcl = tLanc.sCcl
        objLancamento_Detalhe.sHistorico = tLanc.sHistorico
        objLancamento_Detalhe.dValor = -tLanc.dValor
        objLancamento_Detalhe.lNumIntDoc = tLanc.lNumIntDoc
        objLancamento_Detalhe.iFilialCliForn = tLanc.iFilialCliForn
        objLancamento_Detalhe.lCliForn = tLanc.lCliForn
        objLancamento_Detalhe.iTransacao = tLanc.iTransacao
        objLancamento_Detalhe.iSeqContraPartida = tLanc.iSeqContraPartida
        'reprocessamento
        objLancamento_Detalhe.sProduto = tLanc.sProduto
        objLancamento_Detalhe.iApropriaCRProd = tLanc.iApropriaCRProd
        objLancamento_Detalhe.dQuantidade = tLanc.dQuantidade
        objLancamento_Detalhe.dtDataEstoque = tLanc.dtDataEstoque
        objLancamento_Detalhe.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLancamento_Detalhe.iStatus = VOUCHER_ESTORNADOR
        
        colLancamento_Detalhe.Add objLancamento_Detalhe
        
        'Lê o proximo lançamento
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 61163

    Loop
    
    lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe)
    If lErro <> SUCESSO Then gError 61164
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Lancamento_Grava_Estorno_Doc_Basico = SUCESSO
    
    Exit Function
    
Erro_Lancamento_Grava_Estorno_Doc_Basico:

    Lancamento_Grava_Estorno_Doc_Basico = gErr
    
    Select Case gErr
        
        Case 61160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 61161, 61162, 61163
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS8", gErr, objLancamento_Cabecalho1.iFilialEmpresa, objLancamento_Cabecalho1.sOrigem, objLancamento_Cabecalho1.iExercicio, objLancamento_Cabecalho1.iPeriodoLan, objLancamento_Cabecalho1.lDoc)
    
        Case 61164
            
        Case 83880
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANCAMENTO1", gErr, objLancamento_Cabecalho1.iFilialEmpresa, objLancamento_Cabecalho1.sOrigem, objLancamento_Cabecalho1.iExercicio, objLancamento_Cabecalho1.iPeriodoLan, objLancamento_Cabecalho1.lDoc, tLanc.iSeq)
            
        Case 83883
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_ESTORNADO", gErr, objLancamento_Cabecalho1.iFilialEmpresa, objLancamento_Cabecalho1.sOrigem, objLancamento_Cabecalho1.iExercicio, objLancamento_Cabecalho1.iPeriodoLan, objLancamento_Cabecalho1.lDoc)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150753)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function UsuarioLote_Grava_Lote(ByVal sCodUsuario As String, ByVal sOrigem As String, ByVal iLote As Integer) As Long
'se o lote padrão para o Usuario e Origem em questão for diferente do que está gravado, atualiza. Se não existir, insere.
'IMPORTANTE: TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim iLote1 As Integer
Dim iIndice As Integer
Dim lErro As Long
Dim alComando(0 To 1) As Long

On Error GoTo Erro_UsuarioLote_Grava_Lote

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 60896
    Next

    'verifica se o usuario está cadastrado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Lote FROM UsuarioLote WHERE CodUsuario = ? AND Origem=?", 0, iLote1, sCodUsuario, sOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 60898
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60899
    
    If lErro = SUCESSO Then
    
        If iLote <> iLote1 Then
        
            'existe, então alterar
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE UsuarioLote SET Lote = ?", alComando(0), iLote)
            If lErro <> AD_SQL_SUCESSO Then Error 60900
    
        End If
    
    Else
    
        'não existe, então incluir
        lErro = Comando_Executar(alComando(1), "INSERT INTO UsuarioLote (CodUsuario,Origem,Lote) VALUES (?,?,?)", sCodUsuario, sOrigem, iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 60901
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    UsuarioLote_Grava_Lote = SUCESSO
    
    Exit Function
    
Erro_UsuarioLote_Grava_Lote:

    UsuarioLote_Grava_Lote = Err
    
    Select Case Err
        
        Case 60897
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 60898, 60899
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOLOTE", Err, sCodUsuario, sOrigem)
            
        Case 60900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_USUARIOLOTE", Err, sCodUsuario, sOrigem)
                   
        Case 60901
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_USUARIOLOTE", Err, sCodUsuario, sOrigem)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150754)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Ccl_Critica_Vinculos_Compras(ByVal sCcl As String) As Long

Dim lErro As Long
Dim lComando As Long
Dim sCcls As String

On Error GoTo Erro_Ccl_Critica_Vinculos_Compras

    'abertura do comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 85221
    
    'inicializa a variàvel
    sCcls = String(STRING_CCL, 0)
    
    'busca na tabela ItensReqModelo dados com o ccl
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM ItensReqModelo WHERE Ccl =?", sCcls, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 85222
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 85223
    
    'se foi encontrado dados erro
    If lErro = AD_SQL_SUCESSO Then gError 85224
    
    'busca na view ItensReqCompraTodos dados com o ccl
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM ItensReqCompraTodos WHERE Ccl =?", sCcls, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 85225
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 85226
    
    'se foi encontrado dados erro
    If lErro = AD_SQL_SUCESSO Then gError 85227
    
    'busca na view RequisiçãoCompraTodas dados com o ccl
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM RequisicaoCompraTodas WHERE Ccl =?", sCcls, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 85228
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 85229
    
    'se foi encontrado dados erro
    If lErro = AD_SQL_SUCESSO Then gError 85230
    
    'busca na tabela RequisiçãoModelo dados com o ccl
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM RequisicaoModelo WHERE Ccl =?", sCcls, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 85231
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 85232
    
    'se foi encontrado dados erro
    If lErro = AD_SQL_SUCESSO Then gError 85233
    
    'busca na tabela Requisitante dados com o ccl
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM Requisitante WHERE Ccl =?", sCcls, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 85234
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 85235
    
    'se foi encontrado dados erro
    If lErro = AD_SQL_SUCESSO Then gError 85236
    
    'busca na tabela LocalizaçãoItensPC dados com o ccl
    lErro = Comando_Executar(lComando, "SELECT Ccl FROM LocalizacaoItensPC WHERE Ccl =?", sCcls, sCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 85237
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 85238
    
    'se foi encontrado dados erro
    If lErro = AD_SQL_SUCESSO Then gError 85239
    
    Call Comando_Fechar(lComando)
    
    Ccl_Critica_Vinculos_Compras = SUCESSO
    
    Exit Function
    
Erro_Ccl_Critica_Vinculos_Compras:
    
    Ccl_Critica_Vinculos_Compras = gErr
    
    Select Case gErr
    
        Case 85221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 85222, 85223
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQMODELO", gErr)
            
        Case 85224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCl_VINCULADO_ITENSREQMODELO", gErr, sCcl)
            
        Case 85225, 85226
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRATODOS", gErr)
            
        Case 85227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_VINCULADO_ITENSREQCOMPRATODOS", gErr, sCcl)
        
        Case 85228, 85229
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRATODAS", gErr)
            
        Case 85230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_VINCULADO_REQUISICAOCOMPRATODAS", gErr, sCcl)
            
        Case 85231, 85232
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOMODELO", gErr)
            
        Case 85233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_VINCULADO_REQUISICAOMODELO", gErr, sCcl)
            
        Case 85234, 85235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISITANTE", gErr)
            
        Case 85236
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_VINCULADO_REQUISITANTE", gErr, sCcl)
            
        Case 85237, 85238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOCALIZACAOITENSPC", gErr)
        
        Case 85239
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_VINCULADO_LOCALIZACAOITENSPC", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150755)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
        
            
End Function

Function Lancamento_Altera_Status(ByVal iFilialEmpresa As Integer, ByVal iTransacao As Integer, ByVal lNumIntDoc As Long, ByVal colLancamentos As Collection, ByVal iStatus As Integer) As Long
'le os lançamentos relativos ao documento passado como parametro ,coloca-os em colLancamentos e atualiza o status
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento

On Error GoTo Erro_Lancamento_Altera_Status

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 83872
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 83873
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    lErro = Comando_ExecutarPos(lComando, "SELECT  Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem FROM Lancamentos WHERE FilialEmpresa = ? AND Transacao = ? AND NumIntDoc = ? AND Status = 0 ORDER BY Origem, Exercicio, PeriodoLan, Doc", 0, _
    tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, iFilialEmpresa, iTransacao, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 83874
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83875
        
    If lErro = AD_SQL_SEM_DADOS Then gError 83876
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.iFilialEmpresa = iFilialEmpresa
        objLanc.sOrigem = tLanc.sOrigem
        objLanc.iExercicio = tLanc.iExercicio
        objLanc.iPeriodoLan = tLanc.iPeriodoLan
        objLanc.lDoc = tLanc.lDoc
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = tLanc.iLote
        objLanc.iPeriodoLote = tLanc.iPeriodoLote
        objLanc.dtData = tLanc.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        objLanc.iAglutina = tLanc.iAglutina
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.sProduto = tLanc.sProduto
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        objLanc.iApropriaCRProd = tLanc.iApropriaCRProd
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.sDocOrigem = tLanc.sDocOrigem
        
        colLancamentos.Add objLanc
        
        lErro = Comando_ExecutarPos(lComando1, "UPDATE Lancamentos SET Status = ?", lComando, iStatus)
        If lErro <> AD_SQL_SUCESSO Then gError 83877
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83878
        
    Loop
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Lancamento_Altera_Status = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Altera_Status:

    Lancamento_Altera_Status = gErr
    
    Select Case gErr
    
        Case 83872, 83873
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 83874, 83875, 83878
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)
           
        Case 83876
       
        Case 83877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANCAMENTO", gErr, iFilialEmpresa, iTransacao, lNumIntDoc)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150756)
    
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function Lancamento_Exclui_30(ByVal iFilialEmpresa As Integer, ByVal iTransacao As Integer, ByVal lNumIntDoc As Long, ByVal colLancamentos As Collection) As Long
'Exclui os lançamentos da transação em questão e acerta os totais dos lotes envolvidos
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
    
Dim lComando As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim dDebito As Double
Dim dCredito As Double
Dim sOrigem As String
Dim iExercicio As Integer
Dim iPeriodoLan As Integer
Dim iLote As Integer
Dim lDoc As Double
Dim iNumLancAtual As Integer
Dim iNumDocAtual As Integer

On Error GoTo Erro_Lancamento_Exclui_30

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 83891
    
    'le os lançamentos relativos ao documento passado como parametro e coloca-os em colLancamentos
    lErro = Lancamento_Exclui_31(iFilialEmpresa, iTransacao, lNumIntDoc, colLancamentos)
    If lErro <> SUCESSO Then gError 83892
    
    For Each objLanc In colLancamentos
    
        If objLanc.iLote <> 0 Then
    
            If objLanc.iFilialEmpresa <> iFilialEmpresa Or objLanc.sOrigem <> sOrigem Or objLanc.iExercicio <> iExercicio Or objLanc.iPeriodoLan <> iPeriodoLan Or objLanc.iLote <> iLote Then
                
                If iLote <> 0 Then
        
                    lErro = Comando_Executar(lComando, "UPDATE Lote SET TotCre = TotCre + ?, TotDeb = TotDeb + ?, NumLancAtual = NumLancAtual + ?, NumDocAtual = NumDocAtual + ? WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote = ?", dCredito, dDebito, iNumLancAtual, iNumDocAtual, iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, iLote)
                    If lErro <> AD_SQL_SUCESSO Then gError 83893
        
                End If
        
                iFilialEmpresa = objLanc.iFilialEmpresa
                sOrigem = objLanc.sOrigem
                iExercicio = objLanc.iExercicio
                iPeriodoLan = objLanc.iPeriodoLan
                iLote = objLanc.iLote
                lDoc = objLanc.lDoc
                dCredito = 0
                dDebito = 0
                iNumLancAtual = 0
                iNumDocAtual = -1
        
            End If
        
            If objLanc.lDoc <> lDoc Then
                iNumDocAtual = iNumDocAtual - 1
                lDoc = objLanc.lDoc
            End If
        
            
            If objLanc.dValor > 0 Then
                dCredito = dCredito - objLanc.dValor
            Else
                dDebito = dDebito + objLanc.dValor
            End If
            
            iNumLancAtual = iNumLancAtual - 1
    
        End If
    
    Next
  
    If iLote <> 0 Then
        lErro = Comando_Executar(lComando, "UPDATE Lote SET TotCre = TotCre + ?, TotDeb = TotDeb + ?, NumLancAtual = NumLancAtual + ?, NumDocAtual = NumDocAtual + ? WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote = ?", dCredito, dDebito, iNumLancAtual, iNumDocAtual, iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, iLote)
        If lErro <> AD_SQL_SUCESSO Then gError 83894
    End If
  
    Call Comando_Fechar(lComando)
    
    Lancamento_Exclui_30 = SUCESSO

    Exit Function

Erro_Lancamento_Exclui_30:
    
    Lancamento_Exclui_30 = gErr
    
    Select Case gErr
    
        Case 83891
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 83892
    
        Case 83893, 83894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LOTE", gErr, iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, iLote)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150757)
            
    End Select
   
    Call Comando_Fechar(lComando)
   
    Exit Function

End Function

Private Function Lancamento_Exclui_31(ByVal iFilialEmpresa As Integer, ByVal iTransacao As Integer, ByVal lNumIntDoc As Long, ByVal colLancamentos As Collection) As Long
'le os lançamentos relativos ao documento passado como parametro, exclui e coloca-os em colLancamentos
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO

Dim alComando(1 To 2) As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim objLanc1 As ClassLancamento_Detalhe
Dim tLanc As typeLancamento
Dim iIndice As Integer
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim sOrigem As String
Dim lDoc As Long
Dim colLancamentos1 As New Collection
Dim colLancamento_Detalhe As Collection

On Error GoTo Erro_Lancamento_Exclui_31

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 83886
        
    Next

    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT  Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, Status, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, EscaninhoCusto FROM Lancamentos WHERE FilialEmpresa = ? AND Transacao = ? AND NumIntDoc = ? ORDER BY FilialEmpresa, Origem, Exercicio, PeriodoLan, Lote, Doc", 0, _
    tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iStatus, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iEscaninho_Custo, iFilialEmpresa, iTransacao, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 83887
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83888
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objLanc = New ClassLancamento_Detalhe
        
        objLanc.iFilialEmpresa = iFilialEmpresa
        objLanc.sOrigem = tLanc.sOrigem
        objLanc.iExercicio = tLanc.iExercicio
        objLanc.iPeriodoLan = tLanc.iPeriodoLan
        objLanc.lDoc = tLanc.lDoc
        objLanc.iSeq = tLanc.iSeq
        objLanc.iLote = tLanc.iLote
        objLanc.iPeriodoLote = tLanc.iPeriodoLote
        objLanc.dtData = tLanc.dtData
        objLanc.sConta = tLanc.sConta
        objLanc.sCcl = tLanc.sCcl
        objLanc.sHistorico = tLanc.sHistorico
        objLanc.dValor = tLanc.dValor
        objLanc.lNumIntDoc = tLanc.lNumIntDoc
        objLanc.iFilialCliForn = tLanc.iFilialCliForn
        objLanc.lCliForn = tLanc.lCliForn
        objLanc.iTransacao = tLanc.iTransacao
        objLanc.iAglutina = tLanc.iAglutina
        objLanc.lContaSimples = tLanc.lContaSimples
        objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc.sProduto = tLanc.sProduto
        objLanc.dQuantidade = tLanc.dQuantidade
        objLanc.dtDataEstoque = tLanc.dtDataEstoque
        objLanc.iStatus = tLanc.iStatus
        objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc.sDocOrigem = tLanc.sDocOrigem
        objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo
        
        Set objLanc1 = New ClassLancamento_Detalhe
        
        objLanc1.iFilialEmpresa = iFilialEmpresa
        objLanc1.sOrigem = tLanc.sOrigem
        objLanc1.iExercicio = tLanc.iExercicio
        objLanc1.iPeriodoLan = tLanc.iPeriodoLan
        objLanc1.lDoc = tLanc.lDoc
        objLanc1.iSeq = tLanc.iSeq
        objLanc1.iLote = tLanc.iLote
        objLanc1.iPeriodoLote = tLanc.iPeriodoLote
        objLanc1.dtData = tLanc.dtData
        objLanc1.sConta = tLanc.sConta
        objLanc1.sCcl = tLanc.sCcl
        objLanc1.sHistorico = tLanc.sHistorico
        objLanc1.dValor = tLanc.dValor
        objLanc1.lNumIntDoc = tLanc.lNumIntDoc
        objLanc1.iFilialCliForn = tLanc.iFilialCliForn
        objLanc1.lCliForn = tLanc.lCliForn
        objLanc1.iTransacao = tLanc.iTransacao
        objLanc1.iAglutina = tLanc.iAglutina
        objLanc1.lContaSimples = tLanc.lContaSimples
        objLanc1.iSeqContraPartida = tLanc.iSeqContraPartida
        objLanc1.sProduto = tLanc.sProduto
        objLanc1.dQuantidade = tLanc.dQuantidade
        objLanc1.dtDataEstoque = tLanc.dtDataEstoque
        objLanc1.iStatus = tLanc.iStatus
        objLanc1.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
        objLanc1.iClienteFornecedor = tLanc.iClienteFornecedor
        objLanc1.sDocOrigem = tLanc.sDocOrigem
        objLanc1.iEscaninho_Custo = tLanc.iEscaninho_Custo
        
        'esta coleção será passada para fora da função e deve ser mantido tal qual os lançamentos foram gravados
        colLancamentos.Add objLanc
        
        'esta coleção servirá para criar lançamentos de estorno com sinal contrário ao original
        colLancamentos1.Add objLanc1
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83890
        
    Loop
    
    objLancamento_Cabecalho.iFilialEmpresa = iFilialEmpresa
    objLancamento_Cabecalho.sOrigem = "TMP"
    objLancamento_Cabecalho.lNumIntDoc = lNumIntDoc
    objLancamento_Cabecalho.iTransacao = iTransacao
    
    For Each objLanc1 In colLancamentos1
    
        If sOrigem <> objLanc1.sOrigem Or objLancamento_Cabecalho.dtData <> objLanc1.dtData Or lDoc <> objLanc1.lDoc Then
    
            If Not colLancamento_Detalhe Is Nothing Then
    
                'contabiliza os lançamentos imediatamente
                lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe)
                If lErro <> SUCESSO Then gError 83895
            
            End If

            Set colLancamento_Detalhe = New Collection
    
            sOrigem = objLanc1.sOrigem
            lDoc = objLanc1.lDoc
            objLancamento_Cabecalho.dtData = objLanc1.dtData
            objLancamento_Cabecalho.lDoc = objLancamento_Cabecalho.lDoc + 1
            
        End If
        
        objLanc1.iStatus = VOUCHER_EXCLUSAO
            
        colLancamento_Detalhe.Add objLanc1
            
    Next
            
    If Not colLancamento_Detalhe Is Nothing Then

        'contabiliza os lançamentos imediatamente
        lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe)
        If lErro <> SUCESSO Then gError 83896
    
    End If
            
    lErro = Comando_Executar(alComando(2), "DELETE FROM Lancamentos WHERE FilialEmpresa = ? AND Transacao = ? AND NumIntDoc = ?", iFilialEmpresa, iTransacao, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 83889
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Lancamento_Exclui_31 = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Exclui_31:

    Lancamento_Exclui_31 = gErr
    
    Select Case gErr
    
        Case 83886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 83887, 83888, 83890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS9", Err, iFilialEmpresa, iTransacao, lNumIntDoc)
           
        Case 83889
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LANCAMENTO1", Err, iFilialEmpresa, iTransacao, lNumIntDoc)

        Case 83895, 83896

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150758)
    
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Lancamento_Altera_Data(ByVal iFilialEmpresa As Integer, ByVal iTransacao As Integer, ByVal lNumIntDoc As Long, ByVal dtData As Date, ByVal dtDataEstoque As Date) As Long
'Exclui os lançamentos da transação em questão, acerta os totais dos lotes envolvidos e inclui um novo lançamento na nova data
'AVISO IMPORTANTE: PRECISA ESTAR EM TRANSAÇÃO PARA CHAMAR ESTA FUNÇÃO
    
Dim lErro As Long
Dim colLancamentos As New Collection
Dim objLanc As ClassLancamento_Detalhe
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim sOrigem As String
Dim dtData1 As Date
Dim colLancamento_Detalhe As Collection
Dim lDoc As Long
Dim lDoc1 As Long
Dim objPeriodo As New ClassPeriodo

On Error GoTo Erro_Lancamento_Altera_Data

    'Exclui os lançamentos da transação em questão, acerta os totais dos lotes envolvidos
    lErro = CF("Lancamento_Exclui_30", iFilialEmpresa, iTransacao, lNumIntDoc, colLancamentos)
    If lErro <> SUCESSO Then gError 89704

    'Preenche Objeto Lançamento_Cabeçalho
    objLancamento_Cabecalho.iFilialEmpresa = iFilialEmpresa
    objLancamento_Cabecalho.lNumIntDoc = lNumIntDoc
    objLancamento_Cabecalho.iTransacao = iTransacao
    objLancamento_Cabecalho.dtData = dtData
    
    lErro = CF("Periodo_Le1", objLancamento_Cabecalho.dtData, objPeriodo, objLancamento_Cabecalho.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 89705
    
    For Each objLanc In colLancamentos
    
        If objLancamento_Cabecalho.sOrigem <> objLanc.sOrigem Or dtData1 <> objLanc.dtData Or lDoc1 <> objLanc.lDoc Then
    
            If Not colLancamento_Detalhe Is Nothing Then
    
                'contabiliza os lançamentos imediatamente
                lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe)
                If lErro <> SUCESSO Then gError 89706
            
            End If

            Set colLancamento_Detalhe = New Collection
    
            objLancamento_Cabecalho.sOrigem = objLanc.sOrigem
            lDoc1 = objLanc.lDoc
            dtData1 = objLanc.dtData
            
            lErro = CF("Voucher_Automatico_Trans", objLancamento_Cabecalho.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, objLancamento_Cabecalho.sOrigem, lDoc)
            If lErro <> SUCESSO Then gError 89707
        
            objLancamento_Cabecalho.lDoc = lDoc
            
        End If
  
        objLanc.dtDataEstoque = dtDataEstoque
        objLanc.iLote = 0
            
        colLancamento_Detalhe.Add objLanc
            
    Next
  
    If Not colLancamento_Detalhe Is Nothing Then

        'contabiliza os lançamentos imediatamente
        lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe)
        If lErro <> SUCESSO Then gError 89708
    
    End If
  
    Lancamento_Altera_Data = SUCESSO

    Exit Function

Erro_Lancamento_Altera_Data:
    
    Lancamento_Altera_Data = gErr
    
    Select Case gErr
    
        Case 89704, 89705, 89706, 89707, 89708
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150759)
            
    End Select
   
    Exit Function

End Function

Function Lancamento_Altera_Historico(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection) As Long
'le os lançamentos relativos ao documento passado como parametro e altera o histórico

Dim lTransacao As Long
Dim lErro As Long
Dim colLancamento_Detalhe_Aux As New Collection
Dim objLancamento_Detalhe As ClassLancamento_Detalhe

On Error GoTo Erro_Lancamento_Altera_Historico

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92132

    Do While objLancamento_Cabecalho.iFilialEmpresa > 0 And objLancamento_Cabecalho.iFilialEmpresa < 100
       
        Set colLancamento_Detalhe_Aux = New Collection
       
        For Each objLancamento_Detalhe In colLancamento_Detalhe
        
            If giContabGerencial = 1 Then
                If (objLancamento_Detalhe.iGerencial = 0 Or objLancamento_Detalhe.iGerencial = 2) And ((objLancamento_Cabecalho.iFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (objLancamento_Cabecalho.iFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                    colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
                End If
                    
                If (objLancamento_Detalhe.iGerencial = 1 Or objLancamento_Detalhe.iGerencial = 2) And ((objLancamento_Cabecalho.iFilialEmpresa < giFilialAuxiliar And giFilialAuxiliar > 0) Or (objLancamento_Cabecalho.iFilialEmpresa > Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                    colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
                End If
                
            Else
                    
                'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
            End If
                
        Next
       
        lErro = CF("Lancamento_Altera_Historico_Trans", objLancamento_Cabecalho, colLancamento_Detalhe_Aux)
        If lErro <> SUCESSO Then gError 92133
    
        If giContabGerencial = 0 Then Exit Do

        objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa - giFilialAuxiliar
        
    Loop



    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 92134

    Lancamento_Altera_Historico = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Altera_Historico:

    Lancamento_Altera_Historico = gErr
    
    Select Case gErr

        Case 92132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92133

        Case 92134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150760)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Lancamento_Altera_Historico_Trans(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho, ByVal colLancamento_Detalhe As Collection) As Long
'le os lançamentos relativos ao documento passado como parametro e altera o histórico
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim iSeq As Integer

On Error GoTo Erro_Lancamento_Altera_Historico_Trans

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92096
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 92097
    
    For Each objLanc In colLancamento_Detalhe
    
        'Alteração Daniel em 30/08/2002
        If giHistoricoObrigatorio = MARCADO Then
            If Len(Trim(objLanc.sHistorico)) = 0 Then gError 106523
        End If
        'Fim da Alteração
        
        lErro = Comando_ExecutarPos(lComando, "SELECT  Seq FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ? AND Seq = ?", 0, _
        iSeq, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc, objLanc.iSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 92098
            
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92099
        
        If lErro = AD_SQL_SEM_DADOS Then gError 92100
        
        lErro = Comando_ExecutarPos(lComando1, "UPDATE Lancamentos SET Historico = ?", lComando, objLanc.sHistorico)
        If lErro <> AD_SQL_SUCESSO Then gError 92101
        
    Next
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Lancamento_Altera_Historico_Trans = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Altera_Historico_Trans:

    Lancamento_Altera_Historico_Trans = gErr
    
    Select Case gErr
    
    'Alteracao Daniel em 30/08/2002
        Case 106523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_SEM_HISTORICO", gErr)
    'Fim da Alteração Daniel em 30/08/02
        
        Case 92096, 92097
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 92098, 92099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS6", gErr, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLanc.lDoc, objLanc.iSeq)
           
        Case 92100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_INEXISTENTE1", gErr, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLanc.lDoc, objLanc.iSeq)
       
        Case 92101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANCAMENTO1", gErr, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLanc.lDoc, objLanc.iSeq)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150761)
    
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function RelCtaSaldo_Prepara(ByVal iFilialEmpresa As Integer, ByRef lNumIntRel As Long, ByVal sCtaInic As String, ByVal sCtaFim As String, ByVal dtData As Date, ByVal dtDataFim As Date) As Long
'preenche tabela com saldos de contas contabeis numa data para uso em relatorios

Dim lErro As Long, dSaldo As Double, iExerAnt As Integer
Dim objPer As New ClassPeriodo, alComando(1 To 4) As Long, iIndice As Integer
Dim lTransacao As Long, iExercicio As Integer, iStatus As Integer
Dim sConta As String, sContaAtual As String, dSldIni As Double, dDeb(1 To 12) As Double, dCre(1 To 12) As Double
Dim iMes As Integer

On Error GoTo Erro_RelCtaSaldo_Prepara

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81781
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 81782

    'obtem numintrel
    lErro = CF("Config_ObterNumInt", "CTBConfig", "NUM_PROX_REL_SLD_CTA", lNumIntRel)
    If lErro <> SUCESSO Then gError 81783
    
    lErro = CF("Periodo_Le", dtData, objPer)
    If lErro <> SUCESSO Then gError 81784
    
    'obter exercicio que contenha saldo inicial valido menor ou igual a iExercicioAtual
    lErro = RelCtaSaldo_Prepara1(objPer, alComando(1), iExerAnt)
    If lErro <> SUCESSO Then gError 81785
    
    sConta = String(STRING_CONTA, 0)
    
    If sCtaFim <> "" Then
        lErro = Comando_Executar(alComando(2), "SELECT Exercicio, Conta,SldIni,Deb01,Cre01,Deb02,Cre02,Deb03,Cre03,Deb04,Cre04,Deb05,Cre05,Deb06,Cre06,Deb07,Cre07,Deb08,Cre08,Deb09,Cre09,Deb10,Cre10,Deb11,Cre11,Deb12,Cre12 FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio >= ? AND Exercicio <= ? AND Conta >= ? AND Conta <= ? ORDER BY Conta, Exercicio", _
            iExercicio, sConta, dSldIni, dDeb(1), dCre(1), dDeb(2), dCre(2), dDeb(3), dCre(3), dDeb(4), dCre(4), dDeb(5), dCre(5), dDeb(6), dCre(6), dDeb(7), dCre(7), dDeb(8), dCre(8), dDeb(9), dCre(9), dDeb(10), dCre(10), dDeb(11), dCre(11), dDeb(12), dCre(12), iFilialEmpresa, iExerAnt, objPer.iExercicio, sCtaInic, sCtaFim)
        If lErro <> AD_SQL_SUCESSO Then gError 81786

    Else
        lErro = Comando_Executar(alComando(2), "SELECT Exercicio, Conta,SldIni,Deb01,Cre01,Deb02,Cre02,Deb03,Cre03,Deb04,Cre04,Deb05,Cre05,Deb06,Cre06,Deb07,Cre07,Deb08,Cre08,Deb09,Cre09,Deb10,Cre10,Deb11,Cre11,Deb12,Cre12 FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio >= ? AND Exercicio <= ? AND Conta >= ? ORDER BY Conta, Exercicio", _
            iExercicio, sConta, dSldIni, dDeb(1), dCre(1), dDeb(2), dCre(2), dDeb(3), dCre(3), dDeb(4), dCre(4), dDeb(5), dCre(5), dDeb(6), dCre(6), dDeb(7), dCre(7), dDeb(8), dCre(8), dDeb(9), dCre(9), dDeb(10), dCre(10), dDeb(11), dCre(11), dDeb(12), dCre(12), iFilialEmpresa, iExerAnt, objPer.iExercicio, sCtaInic)
        If lErro <> AD_SQL_SUCESSO Then gError 81787
            
    End If
        
    'percorrer registros gerando os saldos na data para as diversas contas
    sContaAtual = ""
    dSaldo = 0
    iMes = Month(dtData)
    
    Do While lErro = AD_SQL_SUCESSO
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81788
    
        If lErro = AD_SQL_SEM_DADOS Then
        
            lErro = RelCtaSaldo_Prepara2(iFilialEmpresa, objPer, dtData, lNumIntRel, dSaldo, sContaAtual, alComando())
            If lErro <> SUCESSO Then gError 81789
            
            Exit Do
            
        Else
        
            If sContaAtual = "" Then
                sContaAtual = sConta
                dSaldo = dSldIni
            End If
            
            'se terminou de ler os registros de uma conta
            If sContaAtual <> sConta Then
                    
                lErro = RelCtaSaldo_Prepara2(iFilialEmpresa, objPer, dtData, lNumIntRel, dSaldo, sContaAtual, alComando())
                If lErro <> SUCESSO Then gError 81790
                
                sContaAtual = sConta
                dSaldo = dSldIni
                
            End If
        
            If iExercicio <> objPer.iExercicio Then
            
                dSaldo = Round(dSaldo - dDeb(1) + dCre(1) - dDeb(2) + dCre(2) - dDeb(3) + dCre(3) - dDeb(4) + dCre(4) - dDeb(5) + dCre(5) - dDeb(6) + dCre(6) - dDeb(7) + dCre(7) - dDeb(8) + dCre(8) - dDeb(9) + dCre(9) - dDeb(10) + dCre(10) - dDeb(11) + dCre(11) - dDeb(12) + dCre(12), 2)
                
            Else

                iIndice = 1
                Do While iIndice < iMes
                    dSaldo = Round(dSaldo + dCre(iIndice) - dDeb(iIndice), 2)
                    iIndice = iIndice + 1
                Loop
                
            End If
        
        End If
    
    Loop
    
    lErro = CF("RelLcto_Prepara", iFilialEmpresa, lNumIntRel, sCtaInic, sCtaFim, dtData, dtDataFim)
    If lErro <> SUCESSO Then Error 106537
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 81791
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    RelCtaSaldo_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelCtaSaldo_Prepara:

    RelCtaSaldo_Prepara = gErr
     
    Select Case gErr
          
        Case 81783, 81784, 81785, 81789, 81790, 106537
        
        Case 81781
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 81782
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 81786, 81787, 81788
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVCTA_SALDO", gErr)
        
        Case 81791
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150762)
     
    End Select
     
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Function RelCtaSaldo_Prepara1(ByVal objPer As ClassPeriodo, ByVal lComando As Long, iExerAnt As Integer) As Long
'obtem exercicio que contenha saldo inicial valido menor ou igual ao exercicio do periodo passado

Dim lErro As Long, iStatus As Integer, iExercicio As Integer

On Error GoTo Erro_RelCtaSaldo_Prepara1
    
    lErro = Comando_Executar(lComando, "SELECT Exercicio, Status FROM Exercicios WHERE Exercicio <= ? ORDER BY Exercicio DESC", iExercicio, iStatus, objPer.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then gError 81775
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81776
    
    iExerAnt = iExercicio
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        If iStatus = 3 Then
            
            'se o exercicio atual está fechado o saldo inicial em mvepercta é valido
            If iExercicio = objPer.iExercicio Then Exit Do

            'se este exercicio está fechado entao o saldo inicial do exercicio seguinte é valido
            iExerAnt = iExercicio + 1
            Exit Do
            
        End If

        iExerAnt = iExercicio
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81777
    
    Loop
    
    RelCtaSaldo_Prepara1 = SUCESSO
     
    Exit Function
    
Erro_RelCtaSaldo_Prepara1:

    RelCtaSaldo_Prepara1 = gErr
     
    Select Case gErr
          
        Case 81775, 81776, 81777
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOS_CTB", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150763)
     
    End Select
     
    Exit Function

End Function

Function RelCtaSaldo_Prepara2(ByVal iFilialEmpresa As Integer, ByVal objPer As ClassPeriodo, ByVal dtData As Date, ByVal lNumIntRel As Long, dSaldo As Double, ByVal sContaAtual As String, alComando() As Long) As Long
'obtem valor movimentado no periodo (mes) até a data e insere registro em RelCtaSaldo

Dim lErro As Long, dValMvDia As Double

On Error GoTo Erro_RelCtaSaldo_Prepara2

    If objPer.dtDataInicio <> dtData Then
    
        'obtem valor movimentado no periodo (mes) até a data
        lErro = Comando_Executar(alComando(3), "SELECT SUM(Cre-Deb) FROM MvDiaCta WHERE FilialEmpresa = ? AND Data >= ? AND Data < ? AND Conta = ?", _
            dValMvDia, iFilialEmpresa, objPer.dtDataInicio, dtData, sContaAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 81778
                
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81779
        
        If lErro = AD_SQL_SUCESSO Then dSaldo = Round(dSaldo + dValMvDia, 2)
        
    End If
            
    lErro = Comando_Executar(alComando(4), "INSERT INTO RelCtaSaldo (NumIntRel, Conta, Saldo) VALUES (?,?,?)", lNumIntRel, sContaAtual, dSaldo)
    If lErro <> AD_SQL_SUCESSO Then gError 81780

    RelCtaSaldo_Prepara2 = SUCESSO
     
    Exit Function
    
Erro_RelCtaSaldo_Prepara2:

    RelCtaSaldo_Prepara2 = gErr
     
    Select Case gErr
          
        Case 81778, 81779
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIA_SALDO", gErr)
        
        Case 81780
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELCTASALDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150764)
     
    End Select
     
    Exit Function

End Function

Private Function CclHistorico_Grava(objCclHistorico As ClassCclHistorico) As Long
'Grava o Historico do Ccl a partir dos dados passados no objCclHistorico

Dim lErro As Long
Dim iSituacao As Integer
Dim lComando As Long

On Error GoTo Erro_CclHistorico_Grava

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97629
    
    'Inicializa a Situacao
    iSituacao = -1
        
    'Verifica se é um novo registro ou uma alteracao dos campos da CclHistorico
    lErro = CclHistorico_Verifica_Alteracao(objCclHistorico, iSituacao)
    If lErro <> SUCESSO Then gError 97630
    
    'Se foi alterado algum campo válido (vide CclHistorico) ou é um novo reg. => GRAVA
    If ((iSituacao = REGISTRO_INEXISTENTE) Or (iSituacao = REGISTRO_ALTERADO)) Then
    
        'Faz a Gravacao na Tabela de CclHistorico
        lErro = CclHistorico_Grava1(objCclHistorico)
        If lErro <> SUCESSO Then gError 97631
            
    End If

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    CclHistorico_Grava = SUCESSO
    
    Exit Function

Erro_CclHistorico_Grava:

    CclHistorico_Grava = gErr
    
    Select Case gErr
    
        Case 97630, 97631
        
        Case 97629
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150765)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Private Function CclHistorico_Verifica_Alteracao(objCclHistorico As ClassCclHistorico, iSituacao As Integer) As Long
'Verifica se está sendo feita alguma alteracao nos campos presentes da tabela CclHistorico

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tCclHistorico As typeCclHistorico

On Error GoTo Erro_CclHistorico_Verifica_Alteracao

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97632

    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT Ccl.Ccl, Ccl.TipoCcl, Ccl.DescCcl FROM Ccl WHERE Ccl.Ccl=?"
    
    With tCclHistorico
        
        'Inicializa as Strings do Type
        .sCodigoCcl = String(STRING_CCL, 0)
        .sDescCcl = String(STRING_CCL_DESCRICAO, 0)
    
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .sCodigoCcl, .iTipoCcl, .sDescCcl, objCclHistorico.sCodigoCcl)
        If lErro <> AD_SQL_SUCESSO Then gError 97633
    
        'Busca o Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97634
    
        'Se existe o registro => Verificar se houve alteracao considerável
        If lErro <> AD_SQL_SEM_DADOS Then
        
            'Verifica se houve alguma alteracao nos dados já existentes
            Call CclHistorico_Verifica_Alteracao1(objCclHistorico, tCclHistorico, iSituacao)
               
        ElseIf lErro = AD_SQL_SEM_DADOS Then
            
            iSituacao = REGISTRO_INEXISTENTE
            
        End If
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    CclHistorico_Verifica_Alteracao = SUCESSO
    
    Exit Function

Erro_CclHistorico_Verifica_Alteracao:

    CclHistorico_Verifica_Alteracao = gErr
    
    Select Case gErr
    
        Case 97633, 97634
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCLHISTORICO", gErr)
        
        Case 97632
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150766)
            
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Private Sub CclHistorico_Verifica_Alteracao1(objCclHistorico As ClassCclHistorico, tCclHistorico As typeCclHistorico, iSituacao As Integer)
'Verifica se houve alguma alteracao nos dados já existentes

On Error GoTo Erro_CclHistorico_Verifica_Alteracao1

    'Se houver alguma alteração => iSituacao = REGISTRO_ALTERADO
    With objCclHistorico
    
        If .sDescCcl <> tCclHistorico.sDescCcl Then gError 97635
        If .iTipoCcl <> tCclHistorico.iTipoCcl Then gError 97635
        
        iSituacao = REGISTRO_INALTERADO
            
    End With
    
    Exit Sub

Erro_CclHistorico_Verifica_Alteracao1:

    Select Case gErr
        
        Case 97635
            iSituacao = REGISTRO_ALTERADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150767)
            
    End Select
    
    Exit Sub

End Sub

Private Function CclHistorico_Grava1(objCclHistorico As ClassCclHistorico) As Long
'Insere o registro na tabela CclHistorico

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim sInsert As String

On Error GoTo Erro_CclHistorico_Grava1
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97636
    
    'Busca o Novo Codigo para insercao
    lErro = CF("Config_ObterAutomatico_EmTrans", "CTBConfig", "NUM_PROX_CCLHISTORICO", "CclHistorico", "NumIntDoc", lNumIntDoc)
    If lErro <> SUCESSO Then gError 97637
    
    With objCclHistorico
    
        .lNumIntDoc = lNumIntDoc
        
        'Guarda a expressão SQL
        sInsert = "INSERT INTO CclHistorico (NumIntDoc, DataAtualizacao, CodigoCcl, TipoCcl, DescCcl) VALUES (?,?,?,?,?)"
                
        'Grava historico na tabela de CclHistorico
        lErro = Comando_Executar(lComando, sInsert, .lNumIntDoc, .dtDataAtualizacao, .sCodigoCcl, .iTipoCcl, .sDescCcl)
        If lErro <> AD_SQL_SUCESSO Then gError 97638
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    CclHistorico_Grava1 = SUCESSO
    
    Exit Function

Erro_CclHistorico_Grava1:

    CclHistorico_Grava1 = gErr
    
    Select Case gErr
    
        Case 97637
        
        Case 97638
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CCLHISTORICO", gErr)
            
        Case 97636
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150768)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Public Function CclHistorico_Grava_Retroativos(ByVal dtDataAtualizacao As Date) As Long
'Grava o historico de Ccls já existentes no banco de dados
'ATENÇÃO: NÃO RETIRE O COMENTÁRIO ABAIXO
'Para chamar esta funcao pelo immediate ...
'call callbyname(createobject("RotinasContab.ClassGrava"),"CclHistorico_Grava_Retroativos",VbMethod,cdate("01/01/02"))

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CclHistorico_Grava_Retroativos

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 94886
    
    lErro = CF("CclHistorico_Grava_Retroativos_EmTrans", dtDataAtualizacao)
    If lErro <> SUCESSO Then gError 94887
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 94888
    
    CclHistorico_Grava_Retroativos = SUCESSO
    
    Exit Function
    
Erro_CclHistorico_Grava_Retroativos:

    CclHistorico_Grava_Retroativos = gErr

    Select Case gErr
        
        Case 94887
        
        Case 94886
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 94888
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150769)
    
    End Select

    Call Transacao_Rollback
    
End Function

Public Function CclHistorico_Grava_Retroativos_EmTrans(ByVal dtDataAtualizacao As Date) As Long
'Grava o historico de Ccls já existentes no banco de dados
'ATENÇÃO: ESTA FUNÇÃO DEVE SER CHAMADA EM TRANSAÇÃO !!

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tCclHistorico As typeCclHistorico
Dim objCclHistorico As New ClassCclHistorico

On Error GoTo Erro_CclHistorico_Grava_Retroativos_EmTrans

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97639
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT Ccl.Ccl, Ccl.TipoCcl, Ccl.DescCcl FROM Ccl"
    
    With tCclHistorico
        
        'Inicializa as Strings do Type
        .sCodigoCcl = String(STRING_CCL, 0)
        .sDescCcl = String(STRING_CCL_DESCRICAO, 0)
            
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .sCodigoCcl, .iTipoCcl, .sDescCcl)
        If lErro <> AD_SQL_SUCESSO Then gError 97641
    
    End With
        
        'Busca o Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97642
    
        'Enquanto tem Ccl => grava o registro
        Do While lErro <> AD_SQL_SEM_DADOS
                                        
            'Atualiza os dados para futura gravacao de histórico
            With objCclHistorico
        
                .dtDataAtualizacao = dtDataAtualizacao
                .iTipoCcl = tCclHistorico.iTipoCcl
                .sCodigoCcl = tCclHistorico.sCodigoCcl
                .sDescCcl = tCclHistorico.sDescCcl
                
            End With
        
            'FAZ A GRAVACAO DO HISTORICO
            lErro = CclHistorico_Grava1(objCclHistorico)
            If lErro <> SUCESSO Then gError 97643
        
            'Busca o Próximo Registro
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97644
            
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
        
    CclHistorico_Grava_Retroativos_EmTrans = SUCESSO
    
    Exit Function
    
Erro_CclHistorico_Grava_Retroativos_EmTrans:

    CclHistorico_Grava_Retroativos_EmTrans = gErr

    Select Case gErr
        
        Case 97639
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97641, 97642, 97644
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCLHISTORICO", gErr)
            
        Case 97643
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150770)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

Private Function PlanoContaHistorico_Grava(objPlanoContaHistorico As ClassPlanoContaHistorico) As Long
'Grava o Historico do PlanoConta a partir dos dados passados no objPlanoContaHistorico

Dim lErro As Long
Dim iSituacao As Integer
Dim lComando As Long

On Error GoTo Erro_PlanoContaHistorico_Grava

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97733
    
    'Inicializa a Situacao
    iSituacao = -1
        
    'Verifica se é um novo registro ou uma alteracao dos campos da PlanoContaHistorico
    lErro = PlanoContaHistorico_Verifica_Alteracao(objPlanoContaHistorico, iSituacao)
    If lErro <> SUCESSO Then gError 97734
    
    'Se foi alterado algum campo válido (vide PlanoContaHistorico) ou é um novo reg. => GRAVA
    If ((iSituacao = REGISTRO_INEXISTENTE) Or (iSituacao = REGISTRO_ALTERADO)) Then
    
        'Faz a Gravacao na Tabela de PlanoContaHistorico
        lErro = PlanoContaHistorico_Grava1(objPlanoContaHistorico)
        If lErro <> SUCESSO Then gError 97735
            
    End If

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    PlanoContaHistorico_Grava = SUCESSO
    
    Exit Function

Erro_PlanoContaHistorico_Grava:

    PlanoContaHistorico_Grava = gErr
    
    Select Case gErr
    
        Case 97734, 97735
        
        Case 97733
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150771)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Private Function PlanoContaHistorico_Verifica_Alteracao(objPlanoContaHistorico As ClassPlanoContaHistorico, iSituacao As Integer) As Long
'Verifica se está sendo feita alguma alteracao nos campos presentes da tabela PlanoContaHistorico

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tPlanoContaHistorico As typePlanoContaHistorico

On Error GoTo Erro_PlanoContaHistorico_Verifica_Alteracao

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97739

    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT PlanoConta.Conta, PlanoConta.DescConta, PlanoConta.TipoConta FROM PlanoConta WHERE PlanoConta.Conta=?"
    
    With tPlanoContaHistorico
        
        'Inicializa as Strings do Type
        .sCodigo = String(STRING_CONTA, 0)
        .sDescricao = String(STRING_CONTA_DESCRICAO, 0)
        
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .sCodigo, .sDescricao, .iTipo, objPlanoContaHistorico.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 97740
    
        'Busca o Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97741
    
        'Se existe o registro => Verificar se houve alteracao considerável
        If lErro <> AD_SQL_SEM_DADOS Then
        
            'Verifica se houve alguma alteracao nos dados já existentes
            Call PlanoContaHistorico_Verifica_Alteracao1(objPlanoContaHistorico, tPlanoContaHistorico, iSituacao)
               
        ElseIf lErro = AD_SQL_SEM_DADOS Then
            
            iSituacao = REGISTRO_INEXISTENTE
            
        End If
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    PlanoContaHistorico_Verifica_Alteracao = SUCESSO
    
    Exit Function

Erro_PlanoContaHistorico_Verifica_Alteracao:

    PlanoContaHistorico_Verifica_Alteracao = gErr
    
    Select Case gErr
    
        Case 97740, 97741
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAHISTORICO", gErr)
        
        Case 97739
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150772)
            
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Private Sub PlanoContaHistorico_Verifica_Alteracao1(objPlanoContaHistorico As ClassPlanoContaHistorico, tPlanoContaHistorico As typePlanoContaHistorico, iSituacao As Integer)
'Verifica se houve alguma alteracao nos dados já existentes

On Error GoTo Erro_PlanoContaHistorico_Verifica_Alteracao1

    'Se houver alguma alteração => iSituacao = REGISTRO_ALTERADO
    With objPlanoContaHistorico
    
        If .sDescricao <> tPlanoContaHistorico.sDescricao Then gError 97742
        If .iTipo <> tPlanoContaHistorico.iTipo Then gError 97742
        
        iSituacao = REGISTRO_INALTERADO
            
    End With
    
    Exit Sub

Erro_PlanoContaHistorico_Verifica_Alteracao1:

    Select Case gErr
        
        Case 97742
            iSituacao = REGISTRO_ALTERADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150773)
            
    End Select
    
    Exit Sub

End Sub

Private Function PlanoContaHistorico_Grava1(objPlanoContaHistorico As ClassPlanoContaHistorico) As Long
'Insere o registro na tabela PlanoContaHistorico

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim sInsert As String

On Error GoTo Erro_PlanoContaHistorico_Grava1
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97743
    
    'Busca o Novo Codigo para insercao
    lErro = CF("Config_ObterAutomatico_EmTrans", "CTBConfig", "NUM_PROX_PLANOCONTAHISTORICO", "PlanoContaHistorico", "NumIntDoc", lNumIntDoc)
    If lErro <> SUCESSO Then gError 97744
    
    With objPlanoContaHistorico
    
        .lNumIntDoc = lNumIntDoc
        
        'Guarda a expressão SQL
        sInsert = "INSERT INTO PlanoContaHistorico (NumIntDoc, DataAtualizacao, Codigo, Tipo, ContaPai, Descricao) VALUES (?,?,?,?,?,?)"
                
        'Grava historico na tabela de PlanoContaHistorico
        lErro = Comando_Executar(lComando, sInsert, .lNumIntDoc, .dtDataAtualizacao, .sCodigo, .iTipo, .sContaPai, .sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 97745
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    PlanoContaHistorico_Grava1 = SUCESSO
    
    Exit Function

Erro_PlanoContaHistorico_Grava1:

    PlanoContaHistorico_Grava1 = gErr
    
    Select Case gErr
    
        Case 97744
        
        Case 97745
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOCONTAHISTORICO", gErr)
            
        Case 97743
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150774)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Public Function PlanoContaHistorico_Grava_Retroativos(ByVal dtDataAtualizacao As Date) As Long
'Grava o historico de Plano de Contas já existentes no banco de dados
'ATENÇÃO: NÃO RETIRE O COMENTÁRIO ABAIXO
'Para chamar esta funcao pelo immediate ...
'call callbyname(createobject("RotinasContab.ClassGrava"),"PlanoContaHistorico_Grava_Retroativos",VbMethod,cdate("01/01/02"))

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PlanoContaHistorico_Grava_Retroativos

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 94889
    
    'Chama a função que faz a gravação
    lErro = CF("PlanoContaHistorico_Grava_Retroativos_EmTrans", dtDataAtualizacao)
    If lErro <> SUCESSO Then gError 94890
    
    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 94891
    
    PlanoContaHistorico_Grava_Retroativos = SUCESSO
    
    Exit Function
    
Erro_PlanoContaHistorico_Grava_Retroativos:

    PlanoContaHistorico_Grava_Retroativos = gErr

    Select Case gErr
        
        Case 94890
        
        Case 94889
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 94891
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150775)
    
    End Select

    Call Transacao_Rollback
    
End Function

Public Function PlanoContaHistorico_Grava_Retroativos_EmTrans(ByVal dtDataAtualizacao As Date) As Long
'Grava o historico de Plano de Contas já existentes no banco de dados
'ATENÇÃO: ESTA FUNÇÃO DEVE SER CHAMADA EM TRANSAÇÃO !!!

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tPlanoContaHistorico As typePlanoContaHistorico
Dim objPlanoContaHistorico As New ClassPlanoContaHistorico
Dim sContaPai As String

On Error GoTo Erro_PlanoContaHistorico_Grava_Retroativos_EmTrans

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97746
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT PlanoConta.Conta, PlanoConta.DescConta, PlanoConta.TipoConta FROM PlanoConta"
    
    With tPlanoContaHistorico
        
        'Inicializa as Strings do Type
        .sCodigo = String(STRING_CONTA, 0)
        .sDescricao = String(STRING_CONTA_DESCRICAO, 0)
        .sContaPai = String(STRING_CONTA, 0)
            
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .sCodigo, .sDescricao, .iTipo)
        If lErro <> AD_SQL_SUCESSO Then gError 97748
    
    End With
        
        'Busca o Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97749
    
        'Enquanto tem PlanoConta => grava o registro
        Do While lErro <> AD_SQL_SEM_DADOS
                                        
            'Atualiza os dados para futura gravacao de histórico
            With objPlanoContaHistorico
        
                'Guarda a ContaPai imediatamente acima da conta em questao
                lErro = Mascara_RetornaContaPai(tPlanoContaHistorico.sCodigo, sContaPai)
                If lErro <> SUCESSO Then gError 97753
                                
                .dtDataAtualizacao = dtDataAtualizacao
                .iTipo = tPlanoContaHistorico.iTipo
                .sCodigo = tPlanoContaHistorico.sCodigo
                .sDescricao = tPlanoContaHistorico.sDescricao
                .sContaPai = sContaPai
                
            End With
        
            'FAZ A GRAVACAO DO HISTORICO
            lErro = PlanoContaHistorico_Grava1(objPlanoContaHistorico)
            If lErro <> SUCESSO Then gError 97750
        
            'Busca o Próximo Registro
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97751
            
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    PlanoContaHistorico_Grava_Retroativos_EmTrans = SUCESSO
    
    Exit Function
    
Erro_PlanoContaHistorico_Grava_Retroativos_EmTrans:

    PlanoContaHistorico_Grava_Retroativos_EmTrans = gErr

    Select Case gErr
        
        Case 97746
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97748, 97749, 97751
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAHISTORICO", gErr)
            
        Case 97750
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150776)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

Public Function Gera_Historicos_Retroativos_IN86(ByVal dtDataGravacao As Date)
'Chama todas as funções de gravação que fazem o input dos dados
'das tabelas de históricos utilizadas pelo IN86
'dtDataGravacao (input): passa a data que será inserida nas tabelas de históricos como data de gravação

'ATENÇÃO: NÃO RETIRE O COMENTÁRIO ABAIXO
'Para chamar esta funcao pelo immediate ...
'call callbyname(createobject("RotinasContab.ClassGrava"),"Gera_Historicos_Retroativos_IN86",VbMethod,cdate("01/01/02"))

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Gera_Historicos_Retroativos_IN86
    
    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 94884
    
    'desativa os locks dos comandos a seguir
    lErro = Conexao_DesativarLocks(DESATIVAR_LOCKS)
    If lErro <> SUCESSO Then gError 94882
        
    lErro = CF("PlanoContaHistorico_Grava_Retroativos_EmTrans", dtDataGravacao)
    If lErro <> SUCESSO Then gError 97780
    
    lErro = CF("CclHistorico_Grava_Retroativos_EmTrans", dtDataGravacao)
    If lErro <> SUCESSO Then gError 97781
    
    lErro = CF("FornecedorHistorico_Grava_Retroativos_EmTrans", dtDataGravacao)
    If lErro <> SUCESSO Then gError 97782
    
    lErro = CF("TranspHistorico_Grava_Retroativos_EmTrans", dtDataGravacao)
    If lErro <> SUCESSO Then gError 97783
    
    lErro = CF("NaturezaOPHistorico_Grava_Retroativos_EmTrans", dtDataGravacao)
    If lErro <> SUCESSO Then gError 97784
    
    lErro = CF("ClienteHistorico_Grava_Retroativos_EmTrans", dtDataGravacao)
    If lErro <> SUCESSO Then gError 97785
    
    lErro = CF("ProdutoHistorico_Grava_Retroativos_EmTrans", dtDataGravacao)
    If lErro <> SUCESSO Then gError 97786
    
    'reativa os locks
    lErro = Conexao_DesativarLocks(REATIVAR_LOCKS)
    If lErro <> SUCESSO Then gError 94883
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 94885
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_OPERACAO_SUCESSO")
    
    Gera_Historicos_Retroativos_IN86 = SUCESSO
    
    Exit Function
    
Erro_Gera_Historicos_Retroativos_IN86:

    Gera_Historicos_Retroativos_IN86 = gErr
    
    Select Case gErr
    
        Case 94884
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 94882
            Call Rotina_Erro(vbOKOnly, "ERRO_DESATIVACAO_LOCKS", gErr)
             
        Case 94883
            Call Rotina_Erro(vbOKOnly, "ERRO_REATIVACAO_LOCKS", gErr)

        Case 94885
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 97780 To 97786
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150777)
        
    End Select

    'reativa os locks
    Call Conexao_DesativarLocks(REATIVAR_LOCKS)

    'Rollback
    Call Transacao_Rollback

End Function

Function RelCclCtaSt_Gera(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal sCclInicial As String, ByVal sCclFinal As String) As Long
'Le os registros da tabela MvPerCcl para a FilialEmpresa, Exercicio, faixa de centros de custos selecionados
'e que tenham o valor debito ou credito do periodo selecionado diferente de zero ordenados por conta
'Calcula Valor pela soma de SldIni com os valores de crédito e débito do Exercicio/Periodo em questao

Dim lErro As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim dSldIni As Double
Dim sSQL As String
Dim sSQL1 As String
Dim sSQL2 As String
Dim sSQLDebito As String
Dim sSQLCredito As String
Dim iPeriodo1 As Integer
Dim sConta As String
Dim dValorPeriodo As Double
Dim dValorAcumPeriodo As Double
Dim dDebito As Double
Dim dCredito As Double
Dim tProcessa_Lancamento As typeProcessa_Lancamento1

On Error GoTo Erro_RelCclCtaSt_Gera

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 93691
   
    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 93692
    
    tProcessa_Lancamento.sConta = String(STRING_CONTA, 0)
    
    sSQL = "SUM(SldIni + "
    
    For iPeriodo1 = 1 To iPeriodo - 1
        sSQL = sSQL + " Cre" + Format(iPeriodo1, "00") + " - Deb" + Format(iPeriodo1, "00") + " + "
    Next
            
    sSQL = left(sSQL, Len(sSQL) - 3) + ")"
            
    'debito
    sSQLDebito = " Deb" + Format(iPeriodo, "00")
        
    'credito
    sSQLCredito = " Cre" + Format(iPeriodo, "00")
        
    If Len(Trim(sCclInicial)) <> 0 And Len(Trim(sCclFinal)) <> 0 Then
    
        lErro = Comando_Executar(lComando1, "SELECT Conta, " + sSQL + ", SUM(" + sSQLDebito + "), SUM(" + sSQLCredito + ") FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Ccl >= ? AND Ccl <= ? AND Conta <> '' GROUP BY  Conta ORDER BY Conta", _
        tProcessa_Lancamento.sConta, tProcessa_Lancamento.dSaldoAnterior, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito, iFilialEmpresa, iExercicio, sCclInicial, sCclFinal)
        'HAVING (SUM(" + sSQLcredito + ") <> 0 Or SUM(" + sSQLdebito + ") <> 0)
    
    ElseIf Len(Trim(sCclInicial)) <> 0 And Len(Trim(sCclFinal)) = 0 Then
        
        lErro = Comando_Executar(lComando1, "SELECT Conta, " + sSQL + ", SUM(" + sSQLDebito + ") , SUM(" + sSQLCredito + ") FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Ccl >= ? AND Conta <> '' GROUP BY Conta ORDER BY Conta", _
        tProcessa_Lancamento.sConta, tProcessa_Lancamento.dSaldoAnterior, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito, iFilialEmpresa, iExercicio, sCclInicial)
        'HAVING (SUM(" + sSQLcredito + ") <> 0 Or SUM(" + sSQLdebito + ") <> 0)
        
    ElseIf Len(Trim(sCclInicial)) = 0 And Len(Trim(sCclFinal)) <> 0 Then
        
        lErro = Comando_Executar(lComando1, "SELECT Conta, " + sSQL + ", SUM(" + sSQLDebito + ") , SUM(" + sSQLCredito + ") FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Ccl <= ? AND Conta <> '' GROUP BY Conta ORDER BY Conta", _
        tProcessa_Lancamento.sConta, tProcessa_Lancamento.dSaldoAnterior, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito, iFilialEmpresa, iExercicio, sCclFinal)
        'Having (Sum(" + sSQLcredito + ") <> 0 Or Sum(" + sSQLdebito + ") <> 0)
        
    Else
        
        lErro = Comando_Executar(lComando1, "SELECT Conta, " + sSQL + ", SUM(" + sSQLDebito + ") , SUM(" + sSQLCredito + ")  FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta <> '' GROUP BY Conta ORDER BY Conta", _
        tProcessa_Lancamento.sConta, tProcessa_Lancamento.dSaldoAnterior, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito, iFilialEmpresa, iExercicio)
        'HAVING (SUM(" + sSQLcredito + ") <> 0 Or SUM(" + sSQLdebito + ") <> 0)
        
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 93693
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 93694

    tProcessa_Lancamento.iFilialEmpresa = iFilialEmpresa
    tProcessa_Lancamento.iExercicio = iExercicio
    tProcessa_Lancamento.iPeriodo = iPeriodo
    tProcessa_Lancamento.lComando = lComando1
       
    tProcessa_Lancamento.iFim_de_Arquivo = 1
    
    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE
       
        'processa os lancamentos
        lErro = Processa_Lancamento_ContaMes1_1(1, tProcessa_Lancamento)
        If lErro <> SUCESSO Then gError 93695

    Loop
    
    'Confirma a transção
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 93696

    'Libera comando
    Call Comando_Fechar(lComando1)
    
    RelCclCtaSt_Gera = SUCESSO
    
    Exit Function
      
Erro_RelCclCtaSt_Gera:

    RelCclCtaSt_Gera = gErr
    
    Select Case gErr
    
        Case 93691, 93692
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 93693, 93694
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", gErr)
        
        Case 93695
        
        Case 93696
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150778)
        
    End Select
    
    'Rollback
    Call Transacao_Rollback
    
    'Libera comando
    Call Comando_Fechar(lComando1)

    Exit Function
    
End Function

Function RelCclCtaSt_Exclui(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer) As Long
'Limpa os registros de RelCclCtaSt para a FilialEmpresa, Exercicio e Períodos especificados

Dim lErro As Long
Dim iCodigo As Integer
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lNumIntCom As Long
Dim sConta As String

On Error GoTo Erro_RelCclCtaSt_Exclui

    'Abre o comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 93682
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 93683
        
    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 93684

    sConta = String(STRING_CONTA, 0)
    
    'Pesquisa  RelCclCta no BD
    lErro = Comando_ExecutarPos(lComando1, "SELECT Conta FROM RelCclCtaSt WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ?", 0, sConta, iFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93685
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93686
    
    Do While lErro <> AD_SQL_SEM_DADOS
                                 
        'Lock em RelCclCtaSt
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 93687
                                 
        'Exclui os registros da tabela RelCclCtaSt
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM RelCclCtaSt", lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 93688
            
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93689
              
    Loop
            
    'Confirma a transção
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 93690
    
    'libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    RelCclCtaSt_Exclui = SUCESSO
    
    Exit Function

Erro_RelCclCtaSt_Exclui:

    RelCclCtaSt_Exclui = gErr

    Select Case gErr

        Case 93682, 93683
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 93684
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 93685, 93686, 93689
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELCCTLCTAST", gErr)

        Case 93687
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RELCCTLCTAST", gErr)
                         
        Case 93688
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELCCTLCTAST", gErr)
               
        Case 93690
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150779)

    End Select

    Call Transacao_Rollback
    
    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function
    
End Function

Private Function Processa_Lancamento_ContaMes1_1(ByVal iNivel As Integer, tProcessa_Lancamento As typeProcessa_Lancamento1) As Long

Dim dDebito1 As Double
Dim dCredito1 As Double
Dim dSaldoAnterior1 As Double
Dim sConta As String
Dim sConta1 As String
Dim sConta2 As String
Dim iExiste_Proximo_Nivel As Integer
Dim lErro As Long
Dim lComando1 As Long

On Error GoTo Erro_Processa_Lancamento_ContaMes1_1

    'Abre o comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 93697

    'inicializa os acumuladores de debito e credito
    dDebito1 = 0
    dCredito1 = 0
    dSaldoAnterior1 = 0

    sConta = String(STRING_CONTA, 0)
    
    'guarda o número da conta que está sendo processada
    lErro = Mascara_RetornaContaNoNivel(iNivel, tProcessa_Lancamento.sConta, sConta)
    If lErro <> SUCESSO Then gError 93698
    
    sConta1 = sConta

    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sConta = sConta1

        'verifica se a conta possui um nivel mais "profundo"
        iExiste_Proximo_Nivel = Mascara_ExisteProxNivelCta(iNivel, tProcessa_Lancamento.sConta)

        If iExiste_Proximo_Nivel = SUCESSO Then

            'se existe um nivel mais profundo, processa-o
            lErro = Processa_Lancamento_ContaMes1_1(iNivel + 1, tProcessa_Lancamento)
            If lErro <> SUCESSO Then gError 93699

            'acumula debitos e creditos
            dDebito1 = dDebito1 + tProcessa_Lancamento.dAcumDebito
            dCredito1 = dCredito1 + tProcessa_Lancamento.dAcumCredito
            dSaldoAnterior1 = dSaldoAnterior1 + tProcessa_Lancamento.dAcumSaldoAnterior
    
            If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then
                    
                'verifica se a conta possui o nivel em questão. Se não possuir sair do loop
                iExiste_Proximo_Nivel = Mascara_ExisteProxNivelCta(iNivel - 1, tProcessa_Lancamento.sConta)
                If iExiste_Proximo_Nivel <> SUCESSO Then Exit Do
    
                sConta = String(STRING_CONTA, 0)
    
                'armazena a conta do nivel em questão
                lErro = Mascara_RetornaContaNoNivel(iNivel, tProcessa_Lancamento.sConta, sConta)
                If lErro <> SUCESSO Then gError 93700
    
            End If

        Else
            
            
            dDebito1 = tProcessa_Lancamento.dDebito
            dCredito1 = tProcessa_Lancamento.dCredito
            dSaldoAnterior1 = tProcessa_Lancamento.dSaldoAnterior
            
            lErro = Comando_BuscarProximo(tProcessa_Lancamento.lComando)
            If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 93701
                
            If lErro = AD_SQL_SEM_DADOS Then tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_FALSE
                       
            Exit Do

        End If

    Loop

    tProcessa_Lancamento.dAcumDebito = dDebito1
    tProcessa_Lancamento.dAcumCredito = dCredito1
    tProcessa_Lancamento.dAcumSaldoAnterior = dSaldoAnterior1

    'insere os registros na tabela
    If (tProcessa_Lancamento.dAcumDebito <> 0 Or tProcessa_Lancamento.dAcumCredito <> 0 Or tProcessa_Lancamento.dAcumSaldoAnterior <> 0) Then
    
        lErro = Comando_Executar(lComando1, "INSERT INTO RelCclCtaSt (FilialEmpresa, Exercicio, Periodo, Conta, SldAnt, Debito, Credito) VALUES (?,?,?,?,?,?,?)", tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.iPeriodo, sConta1, tProcessa_Lancamento.dAcumSaldoAnterior, tProcessa_Lancamento.dAcumDebito, tProcessa_Lancamento.dAcumCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 93702

    End If

    Call Comando_Fechar(lComando1)

    Processa_Lancamento_ContaMes1_1 = SUCESSO

    Exit Function

Erro_Processa_Lancamento_ContaMes1_1:

    Processa_Lancamento_ContaMes1_1 = gErr

    Select Case gErr
        
        Case 93699
        
        Case 93697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case 93698, 93700
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", gErr, tProcessa_Lancamento.sConta, iNivel)
        
        Case 93701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", gErr)
        
        Case 93702
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELCCLCTASL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 150780)

    End Select

    Exit Function

End Function

Function RelLcto_Prepara(ByVal iFilialEmpresa As Integer, lNumIntRel As Long, ByVal sCtaInic As String, ByVal sCtaFim As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date) As Long
'preenche tabela com lancamentos contabeis numa faixa de contas / periodo para uso em relatorios

Dim lErro As Long
Dim alComando(1 To 2) As Long, iIndice As Integer, sConta As String
Dim tLanc As typeLancamento, sSQL As String

On Error GoTo Erro_RelLcto_Prepara



    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 106532
    Next

    sConta = String(STRING_CONTA, 0)
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
        sSQL = "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, SeqAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, ApropriaCrProd, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Status, Modelo FROM Lancamentos WHERE Data >= ? AND Data <= ?"
    Else
        sSQL = "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, SeqAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, ApropriaCrProd, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Status, Modelo FROM Lancamentos WHERE FilialEmpresa = ? And Data >= ? AND Data <= ?"
    End If
    
    If sCtaInic <> "" Then

        sSQL = sSQL & " AND Conta >= ?"
        
    End If
    
    If sCtaFim <> "" Then
    
        sSQL = sSQL & " AND Conta <= ?"
    
    End If
    
    If iFilialEmpresa = EMPRESA_TODA Then
    
        If sCtaInic <> "" And sCtaFim <> "" Then
            lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaInic, sCtaFim)
        Else
        
            If sCtaInic <> "" Then
                    lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaInic)
            ElseIf sCtaFim <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaFim)
            Else
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim)
            End If
            
        End If
    
    Else
    
        If sCtaInic <> "" And sCtaFim <> "" Then
            lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaInic, sCtaFim)
        Else
        
            If sCtaInic <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaInic)
            ElseIf sCtaFim <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaFim)
            Else
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim)
            End If
            
        End If
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 106533
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106534
        
    Do While lErro = AD_SQL_SUCESSO
    
        'transfere o lancamento do cadastro de lancamentos pendentes para o cadastro de lancamentos contabilizados
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelLcto (NumIntRel, FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Seq,Lote,PeriodoLote,Data,Conta,Ccl,Historico,Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, DocAglutinado, SeqAglutinado, Aglutinado, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Produto, ApropriaCRProd, Quantidade, DataEstoque, Status, Modelo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntRel, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.iAglutina, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.iEscaninho_Custo, _
            tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.sProduto, tLanc.iApropriaCRProd, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iStatus, tLanc.sModelo)
        If lErro <> AD_SQL_SUCESSO Then gError 106535
                
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106536
        
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    RelLcto_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelLcto_Prepara:

    RelLcto_Prepara = gErr
     
    Select Case gErr
          
        Case 106532
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 106533, 106534, 106536
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTO", gErr)
        
        Case 106535
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELLCTO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150781)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Function Ccl_Importa() As Long

Dim lErro As Long
Dim colCcl As New Collection
Dim objCcl As ClassCcl

On Error GoTo Erro_Ccl_Importa

    lErro = ImportCcl_Le(colCcl)
    If lErro <> SUCESSO Then gError 182656
    
    For Each objCcl In colCcl
    
        Call Ccl_Insere(objCcl)

    Next

    Ccl_Importa = SUCESSO
     
    Exit Function
    
Erro_Ccl_Importa:

    Ccl_Importa = gErr
     
    Select Case gErr
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182657)
     
    End Select
   
    Exit Function
    
End Function

Function ImportCcl_Le(ByVal colCcl As Collection) As Long
'le o ccl a partir da chave fornecida em objCcl. Devolve os dados em objCcl

Dim lComando As Long
Dim lErro As Long
Dim tCcl As typeCcl
Dim objCcl As ClassCcl
Dim sCclBD As String, iCclPreenchida As Integer

On Error GoTo Erro_ImportCcl_Le

    tCcl.sDescCcl = String(STRING_CCL_DESCRICAO, 0)
    tCcl.sCcl = String(STRING_CCL, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 182659

    'Verifica se o Ccl existe
    lErro = Comando_Executar(lComando, "SELECT Ccl, DescCcl, TipoCcl, AtivoCcl FROM ImportCcl ", tCcl.sCcl, tCcl.sDescCcl, tCcl.iTipoCcl, tCcl.iAtivo)
    If lErro <> AD_SQL_SUCESSO Then gError 182660

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182661
        
    'Ccl não está cadastrado
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objCcl = New ClassCcl

        lErro = CF("Ccl_Formata", tCcl.sCcl, sCclBD, iCclPreenchida)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objCcl.sCcl = sCclBD
        objCcl.sDescCcl = tCcl.sDescCcl
        objCcl.iTipoCcl = tCcl.iTipoCcl
        objCcl.iAtivo = tCcl.iAtivo
        
        colCcl.Add objCcl
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182662

    Loop
    
    Call Comando_Fechar(lComando)

    ImportCcl_Le = SUCESSO

Exit Function

Erro_ImportCcl_Le:

    ImportCcl_Le = gErr
    
    Select Case gErr

        Case 182659
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 182660 To 182662
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", gErr)
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182663)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PlanoConta_Importa() As Long

Dim lErro As Long
Dim colPlanoConta As New Collection
Dim objPlanoConta As ClassPlanoConta
Dim iMaxLen As Integer
Dim sConta As String
Dim iNivel As Integer

On Error GoTo Erro_PlanoConta_Importa

    lErro = ImportPlanoConta_Le(colPlanoConta, iMaxLen)
    If lErro <> SUCESSO Then gError 182656
    
    For Each objPlanoConta In colPlanoConta
        
        If Len(objPlanoConta.sConta) <> iMaxLen Then
            objPlanoConta.sConta = objPlanoConta.sConta & String(iMaxLen - Len(objPlanoConta.sConta), "0")
        End If
        
        sConta = objPlanoConta.sConta
    
        Call Mascara_Conta_ObterNivel(sConta, iNivel)

        objPlanoConta.iNivelConta = iNivel
        
        objPlanoConta.sDescConta = left(objPlanoConta.sDescConta, STRING_CONTA_DESCRICAO)
    
        'verifica se a conta possui uma conta pai
        lErro = CF("Conta_Critica_ContaPai", objPlanoConta.sConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO Then
            gError 5487
        End If
        
        'verifica se a conta é de nivel 1 e tem a categoria preenchida.
        lErro = CF("Conta_Critica_Categoria", objPlanoConta.sConta, objPlanoConta.iCategoria)
        If lErro <> SUCESSO Then
            gError 9717
        End If
    
        'verifica se a conta já está cadastrada
        lErro = CF("PlanoConta_Le_Conta", sConta)
        If lErro <> SUCESSO And lErro <> 10051 Then
            gError 5490
        End If
        
        'se estiver não cadastrada
        If lErro <> SUCESSO Then
    
            lErro = PlanoConta_Insere(objPlanoConta)
            If lErro <> SUCESSO Then
                gError 9999
            End If

        End If

    Next

    PlanoConta_Importa = SUCESSO
     
    Exit Function
    
Erro_PlanoConta_Importa:

    Call MsgBox(objPlanoConta.sConta)

    PlanoConta_Importa = gErr
     
    Select Case gErr
    
        Case 5487, 182656, 9717
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182657)
     
    End Select
   
    Exit Function
    
End Function

Function ImportPlanoConta_Le(ByVal colPlanoConta As Collection, iMaxLen As Integer) As Long
'le o PlanoConta a partir da chave fornecida em objPlanoConta. Devolve os dados em objPlanoConta

Dim lComando As Long
Dim lErro As Long
Dim tPlanoConta As typePlanoConta
Dim objPlanoConta As ClassPlanoConta
Dim sContaBD As String, iContaPreenchida As Integer

On Error GoTo Erro_ImportPlanoConta_Le

    tPlanoConta.sDescConta = String(255, 0)
    tPlanoConta.sConta = String(255, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 182659

    With tPlanoConta
        'Verifica se o PlanoConta existe
        lErro = Comando_Executar(lComando, "SELECT Conta,DescConta,TipoConta,ContaSimples,Natureza,HistPadrao,Ativo,Categoria,UsaContaSimples,CP,CR,TES,FAT,EST,COM,LJ,FIS,FluxoCaixa, SldIni FROM ImportPlanoConta ORDER BY Conta ", _
                .sConta, .sDescConta, .iTipoConta, .lContaSimples, .iNatureza, .iHistPadrao, .iAtivo, .iCategoria, .iUsaContaSimples, .iCP, .iCR, .iTES, .iFAT, .iEST, .iCOM, .iLJ, .iFIS, .iFluxoCaixa, .dSldIni)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 182660

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182661
        
    'PlanoConta não está cadastrado
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objPlanoConta = New ClassPlanoConta
        
        With objPlanoConta
            .iFilialEmpresa = giFilialEmpresa

            lErro = CF("Conta_Formata", tPlanoConta.sConta, sContaBD, iContaPreenchida)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            .sConta = sContaBD
            .iNivelConta = tPlanoConta.iNivelConta
            .sDescConta = tPlanoConta.sDescConta
            .iTipoConta = tPlanoConta.iTipoConta
            .lContaSimples = tPlanoConta.lContaSimples
            .iNatureza = tPlanoConta.iNatureza
            .iHistPadrao = tPlanoConta.iHistPadrao
            .iAtivo = tPlanoConta.iAtivo
            .iCategoria = tPlanoConta.iCategoria
            .iUsaContaSimples = tPlanoConta.iUsaContaSimples
            .iFluxoCaixa = tPlanoConta.iFluxoCaixa
            .dSldIni = tPlanoConta.dSldIni
            
            If tPlanoConta.iCP = MARCADO Then .colModulo.Add "CP"
            If tPlanoConta.iCR = MARCADO Then .colModulo.Add "CR"
            If tPlanoConta.iTES = MARCADO Then .colModulo.Add "TES"
            If tPlanoConta.iFAT = MARCADO Then .colModulo.Add "FAT"
            If tPlanoConta.iEST = MARCADO Then .colModulo.Add "EST"
            If tPlanoConta.iCOM = MARCADO Then .colModulo.Add "COM"
            If tPlanoConta.iLJ = MARCADO Then .colModulo.Add "LJ"
            If tPlanoConta.iFIS = MARCADO Then .colModulo.Add "FIS"
            
            If iMaxLen < Len(.sConta) Then iMaxLen = Len(.sConta)

        End With
        
        colPlanoConta.Add objPlanoConta
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182662

    Loop
    
    Call Comando_Fechar(lComando)

    ImportPlanoConta_Le = SUCESSO

Exit Function

Erro_ImportPlanoConta_Le:

    ImportPlanoConta_Le = Err
    
    Select Case Err

        Case 182659
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 182660 To 182662
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", gErr)
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182663)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Lancamento_Exclui_4(ByVal objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long
'le os lançamentos relativos ao documento passado como parametro e  exclui antes de excluir estorna-os

Dim alComando(1 To 2) As Long
Dim lErro As Long
Dim objLanc As ClassLancamento_Detalhe
Dim tLanc As typeLancamento
Dim iIndice As Integer
Dim sOrigem As String
Dim colLancamento_Detalhe As New Collection
Dim lTransacao As Long
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_Lancamento_Exclui_4

    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 183099

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 183092
        
    Next
    
    objPeriodosFilial.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
    objPeriodosFilial.iExercicio = objLancamento_Cabecalho.iExercicio
    objPeriodosFilial.iPeriodo = objLancamento_Cabecalho.iPeriodoLan
    objPeriodosFilial.sOrigem = objLancamento_Cabecalho.sOrigem

    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then gError 207116
    
    'se o periodo estiver fechado, não pode alterar os lançamentos do periodo
    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then gError 207117
    End If

'    tLanc.sOrigem = String(STRING_ORIGEM, 0)
'    tLanc.sConta = String(STRING_CONTA, 0)
'    tLanc.sCcl = String(STRING_CCL, 0)
'    tLanc.sHistorico = String(STRING_HISTORICO, 0)
'    tLanc.sProduto = String(STRING_PRODUTO, 0)
'    tLanc.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    sOrigem = objLancamento_Cabecalho.sOrigem

    Do While objLancamento_Cabecalho.iFilialEmpresa > 0 And objLancamento_Cabecalho.iFilialEmpresa < 100

        tLanc.sOrigem = String(STRING_ORIGEM, 0)
        tLanc.sConta = String(STRING_CONTA, 0)
        tLanc.sCcl = String(STRING_CCL, 0)
        tLanc.sHistorico = String(STRING_HISTORICO, 0)
        tLanc.sProduto = String(STRING_PRODUTO, 0)
        tLanc.sDocOrigem = String(STRING_DOCORIGEM, 0)

        lErro = Comando_ExecutarPos(alComando(1), "SELECT  Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, Status, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, EscaninhoCusto FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ?  And Doc = ?", 0, _
        tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iStatus, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iEscaninho_Custo, _
        objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 183093

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183094

        Do While lErro <> AD_SQL_SEM_DADOS

            Set objLanc = New ClassLancamento_Detalhe

            objLanc.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
            objLanc.sOrigem = tLanc.sOrigem
            objLanc.iExercicio = tLanc.iExercicio
            objLanc.iPeriodoLan = tLanc.iPeriodoLan
            objLanc.lDoc = tLanc.lDoc
            objLanc.iSeq = tLanc.iSeq
            objLanc.iLote = tLanc.iLote
            objLanc.iPeriodoLote = tLanc.iPeriodoLote
            objLanc.dtData = tLanc.dtData
            objLanc.sConta = tLanc.sConta
            objLanc.sCcl = tLanc.sCcl
            objLanc.sHistorico = tLanc.sHistorico
            objLanc.dValor = tLanc.dValor '-tLanc.dValor
            objLanc.lNumIntDoc = tLanc.lNumIntDoc
            objLanc.iFilialCliForn = tLanc.iFilialCliForn
            objLanc.lCliForn = tLanc.lCliForn
            objLanc.iTransacao = tLanc.iTransacao
            objLanc.iAglutina = tLanc.iAglutina
            objLanc.lContaSimples = tLanc.lContaSimples
            objLanc.iSeqContraPartida = tLanc.iSeqContraPartida
            objLanc.sProduto = tLanc.sProduto
            objLanc.dQuantidade = tLanc.dQuantidade
            objLanc.dtDataEstoque = tLanc.dtDataEstoque
            objLanc.iStatus = tLanc.iStatus
            objLanc.dValorLivroAuxiliar = tLanc.dValorLivroAuxiliar
            objLanc.iClienteFornecedor = tLanc.iClienteFornecedor
            objLanc.sDocOrigem = tLanc.sDocOrigem
            objLanc.iEscaninho_Custo = tLanc.iEscaninho_Custo

            'esta coleção será passada para fora da função e deve ser mantido tal qual os lançamentos foram gravados
            colLancamento_Detalhe.Add objLanc

            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183095

        Loop

        objLancamento_Cabecalho.sOrigem = "TMP"
        objLancamento_Cabecalho.dtData = objLanc.dtData
                
        'contabiliza os lançamentos imediatamente
        'Sem o true final acertará mvdiacta e mvpercta pelo estorno e não pela exclusão
        lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe, True)
        If lErro <> SUCESSO Then gError 183096
        
        lErro = Comando_Executar(alComando(2), "DELETE FROM Lancamentos WHERE FilialEmpresa = ? AND (Origem = ? Or Origem = ?) AND Exercicio = ? And PeriodoLan = ?  And Doc = ?", objLancamento_Cabecalho.iFilialEmpresa, sOrigem, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 183097
    
        If giContabGerencial = 0 Then Exit Do
        
        objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa - giFilialAuxiliar
    
        Set colLancamento_Detalhe = New Collection
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transção
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 183100
    
    Lancamento_Exclui_4 = SUCESSO
    
    Exit Function
 
Erro_Lancamento_Exclui_4:

    Lancamento_Exclui_4 = gErr
    
    Select Case gErr
    
        Case 183092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 183093, 183094, 183095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS8", gErr, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)
           
        Case 183096
           
        Case 183097
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LANCAMENTO", gErr, objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc)

        Case 183099
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 183100
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 207116
                        
        Case 207117
            Call Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", gErr, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.iExercicio)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183098)
    
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Rollback
    Call Transacao_Rollback
    
    Exit Function

End Function

Function ImportacaoCTB_Grava(objImportacaoCTB As ClassImportacaoCTB) As Long
'Insere o registro na tabela ImportacaoCTB

Dim lErro As Long
Dim lComando As Long
Dim sInsert As String

On Error GoTo Erro_ImportacaoCTB_Grava
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 186543
    
    'Guarda a expressão SQL
    sInsert = "INSERT INTO ImportacaoCTB (FilialEmpresa, Arquivo, NumIntDoc, Data, Hora) VALUES (?,?,?,?,?)"
            
    'Grava historico na tabela de CclHistorico
    lErro = Comando_Executar(lComando, sInsert, objImportacaoCTB.iFilialEmpresa, objImportacaoCTB.sArquivo, objImportacaoCTB.lNumIntDoc, objImportacaoCTB.dtData, CDbl(objImportacaoCTB.dtHora))
    If lErro <> AD_SQL_SUCESSO Then gError 186544
        
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ImportacaoCTB_Grava = SUCESSO
    
    Exit Function

Erro_ImportacaoCTB_Grava:

    ImportacaoCTB_Grava = gErr
    
    Select Case gErr
    
        Case 186543
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 186544
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_IMPORTACAOCTB", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186545)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Function ImportacaoCTB_Exclui(objImportacaoCTB As ClassImportacaoCTB) As Long
'Exclui o registro na tabela ImportacaoCTB se existir e o documento contabil associado

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2)  As Long
Dim lNumIntDoc As Long
Dim dtData As Date
Dim dHora As Double
Dim iTransacaoCTB As Integer
Dim colLancamentos As New Collection
Dim sSelect As String

On Error GoTo Erro_ImportacaoCTB_Exclui
    
    'Abre o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186546
    Next
    
    'Guarda a expressão SQL
    sSelect = "SELECT NumIntDoc, Data, Hora FROM ImportacaoCTB WHERE FilialEmpresa = ? AND Arquivo = ?"
            
    'Grava historico na tabela de CclHistorico
    lErro = Comando_ExecutarPos(alComando(1), sSelect, 0, lNumIntDoc, dtData, dHora, objImportacaoCTB.iFilialEmpresa, objImportacaoCTB.sArquivo)
    If lErro <> AD_SQL_SUCESSO Then gError 186547
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186548
        
    If lErro = AD_SQL_SUCESSO Then
        
        iTransacaoCTB = 75
    
        'exclui os lancamentos contabeis vinculados a ImportacaoCTB que esta sendo excluida
        lErro = CF("Lancamento_Exclui_30", objImportacaoCTB.iFilialEmpresa, iTransacaoCTB, lNumIntDoc, colLancamentos)
        If lErro <> SUCESSO Then gError 186549
        
        If giContabGerencial = 1 And ((objImportacaoCTB.iFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (objImportacaoCTB.iFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
        
            lErro = CF("Lancamento_Exclui_30", objImportacaoCTB.iFilialEmpresa - giFilialAuxiliar, iTransacaoCTB, lNumIntDoc, colLancamentos)
            If lErro <> SUCESSO Then gError 186549
        
        End If
        
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ImportacaoCTB", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 186550
        
    End If
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ImportacaoCTB_Exclui = SUCESSO
    
    Exit Function

Erro_ImportacaoCTB_Exclui:

    ImportacaoCTB_Exclui = gErr
    
    Select Case gErr
    
        Case 186546
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 186547, 186548
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTACAOCTB", gErr)
        
        Case 186549
        
        Case 186550
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_IMPORTACAOCTB", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186551)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Public Function Excel_Le_Planilha(ByVal iFilialEmpresa As Integer, ByVal sArquivo As String, ByVal objMsg As Object, ByVal objProgresso As Object, ByVal objTela As Object) As Long

Dim lErro As Long

'Declarações para manipulação do Excel
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
'Dim objFaixaCelulasPreencher As Object 'Excel.Range
Dim sConta As String
Dim sCCusto As String
Dim sCredDeb As String
Dim dValor As Double
Dim sHistorico As String
Dim iSeq As Integer
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim colLancamento_Detalhe As New Collection
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim dtData As Date
Dim lDoc As Long
Dim objPeriodo As New ClassPeriodo
Dim iIndice As Integer
Dim dValorTotal As Double
Dim lTransacao As Long
Dim sArq As String
Dim iPos As Integer
Dim iPosNew As Integer
Dim sDir As String
Dim objImportacaoCTB As New ClassImportacaoCTB
Dim iLinhas As Integer
Dim iGerencial As Integer
Dim colLancamento_Detalhe_Aux As Collection
Dim lNumIntDoc As Long
Dim dValorTotal0 As Double
Dim dValorTotal1 As Double
Dim objExcelApp As New ClassExcelApp

On Error GoTo Erro_Excel_Le_Planilha

    iPosNew = 1
    iPos = 1
    
   
    Do While iPosNew <> 0

        iPosNew = InStr(iPos, sArquivo, Chr(0))
   
        If iPosNew = 0 Then

            sArq = sDir & Mid(sArquivo, iPos)

        Else
        
            If iPos = 1 Then
                sDir = Mid(sArquivo, iPos, iPosNew - iPos) & "\"
            End If
        
            iPos = iPosNew + 1
            iPosNew = InStr(iPos, sArquivo, Chr(0))
        
            sArq = sDir & Mid(sArquivo, iPos, iPosNew - iPos)
            iPos = iPosNew + 1

        End If
        
        Set colLancamento_Detalhe = New Collection
        Set objLancamento_Cabecalho = New ClassLancamento_Cabecalho


        'Abrir transacao
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 186554
    
'        'Abre o excel
'        lErro = CF("Excel_Abrir")
'        If lErro <> 0 Then gError 186555
'
'        'Desabilita as mensagens do Excel
'        GL_objExcelSist.DisplayAlerts = False
        
        'Abre o excel
        lErro = objExcelApp.Abrir()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        objMsg.AddItem "Abrindo a planilha " & sArq
        
'        Set objPastaTrabalho = GL_objExcelSist.Workbooks.Open(sArq)
        
        'Abre o excel
        lErro = objExcelApp.Abrir_Planilha(sArq)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        DoEvents
        
        If objTela.giStop = 1 Then gError 186556
        
        objMsg.AddItem "Planilha aberta, iniciando importação"
        
        'Seleciona a planilha ativa na pasta de trabalho criada
        'Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
        
        'dtData = StrParaDate(objPlanilhaExcel.Cells(4, 2).Value)
        dtData = StrParaDate(objExcelApp.Obtem_Valor_Celula(4, 2))
        
        lErro = CF("Periodo_Le", dtData, objPeriodo)
        If lErro <> SUCESSO Then gError 186557
    
        lErro = CF("Voucher_Automatico_Trans", iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, "CTB", lDoc)
        If lErro <> SUCESSO Then gError 186558
        
        'Preenche Objeto Lançamento_Cabeçalho
        objLancamento_Cabecalho.iFilialEmpresa = iFilialEmpresa
        objLancamento_Cabecalho.sOrigem = "CTB"
        objLancamento_Cabecalho.iLote = 0
        objLancamento_Cabecalho.lDoc = lDoc
        objLancamento_Cabecalho.dtData = dtData
        
        iIndice = 6
        iSeq = 0
        iLinhas = 0
        
        'sConta = objPlanilhaExcel.Cells(7, 1).Value
        sConta = objExcelApp.Obtem_Valor_Celula(7, 1)
        
        Do While sConta <> ""
            iLinhas = iLinhas + 1
            'sConta = objPlanilhaExcel.Cells(iLinhas + 7, 1).Value
            sConta = objExcelApp.Obtem_Valor_Celula(iLinhas + 7, 1)
        Loop
        
        objProgresso.Max = iLinhas
        objProgresso.Min = 0
        objProgresso.Value = 0
        
        sConta = objExcelApp.Obtem_Valor_Celula(7, 1)
        
        lErro = CF("Config_ObterNumInt", "CTBConfig", "NUM_PROX_IMPORTACAO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 81783
        
        Do While sConta <> ""
        
            iIndice = iIndice + 1
        
'            sCCusto = objPlanilhaExcel.Cells(iIndice, 2).Value
'            sCredDeb = objPlanilhaExcel.Cells(iIndice, 3).Value
'            dValor = StrParaDbl(objPlanilhaExcel.Cells(iIndice, 4).Value)
'            sHistorico = objPlanilhaExcel.Cells(iIndice, 5).Value
'            If giContabGerencial = 1 Then iGerencial = StrParaInt(objPlanilhaExcel.Cells(iIndice, 6).Value)
            sCCusto = objExcelApp.Obtem_Valor_Celula(iIndice, 2)
            sCredDeb = objExcelApp.Obtem_Valor_Celula(iIndice, 3)
            dValor = StrParaDbl(objExcelApp.Obtem_Valor_Celula(iIndice, 4))
            sHistorico = objExcelApp.Obtem_Valor_Celula(iIndice, 5)
            If giContabGerencial = 1 Then iGerencial = StrParaInt(objExcelApp.Obtem_Valor_Celula(iIndice, 6))
        
            iSeq = iSeq + 1
        
            lErro = CF("Conta_Formata", sConta, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then gError 186559
            
            If iContaPreenchida = CONTA_VAZIA Then gError 186560
        
            Set objLancamento_Detalhe = New ClassLancamento_Detalhe
        
            objLancamento_Detalhe.sConta = sContaFormatada
        
            objLancamento_Detalhe.sCcl = sCCusto
        
            If UCase(sCredDeb) = "C" Then
                objLancamento_Detalhe.dValor = Round(dValor, 2)
            ElseIf UCase(sCredDeb) = "D" Then
                objLancamento_Detalhe.dValor = Round(-dValor, 2)
            Else
                gError 186561
            End If
            
            dValorTotal = dValorTotal + Round(objLancamento_Detalhe.dValor, 2)
        
            If giContabGerencial = 1 Then
                If iGerencial = 0 Or iGerencial = 2 Then
                    dValorTotal0 = dValorTotal0 + Round(objLancamento_Detalhe.dValor, 2)
                End If
                
                If iGerencial = 1 Or iGerencial = 2 Then
                    dValorTotal0 = dValorTotal0 + Round(objLancamento_Detalhe.dValor, 2)
                End If
            End If
        
            If dValor = 0 Then gError 186562
        
            objLancamento_Detalhe.sHistorico = sHistorico
            
            If giContabGerencial = 1 Then objLancamento_Detalhe.iGerencial = iGerencial
            
            objLancamento_Detalhe.iSeq = iSeq
            
            'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
            colLancamento_Detalhe.Add objLancamento_Detalhe
        
            'sConta = objPlanilhaExcel.Cells(iIndice + 1, 1).Value
            sConta = objExcelApp.Obtem_Valor_Celula(iIndice + 1, 1)
        
            objLancamento_Detalhe.lNumIntDoc = lNumIntDoc
            
            objLancamento_Detalhe.iTransacao = 75
            
            objProgresso.Value = objProgresso.Value + 1
            
            DoEvents
            
            If objTela.giStop = 1 Then gError 186563
        
        Loop
        
        If Abs(dValorTotal) > DELTA_VALORMONETARIO Then gError 188298
        
        If Abs(dValorTotal0) > DELTA_VALORMONETARIO Then gError 207327
        
        If Abs(dValorTotal1) > DELTA_VALORMONETARIO Then gError 207328
        
        Do While objLancamento_Cabecalho.iFilialEmpresa > 0 And objLancamento_Cabecalho.iFilialEmpresa < 100
           
            Set colLancamento_Detalhe_Aux = New Collection
           
            For Each objLancamento_Detalhe In colLancamento_Detalhe
            
                If giContabGerencial = 1 Then
                    If (objLancamento_Detalhe.iGerencial = 0 Or objLancamento_Detalhe.iGerencial = 2) And ((objLancamento_Cabecalho.iFilialEmpresa > giFilialAuxiliar And giFilialAuxiliar > 0) Or (objLancamento_Cabecalho.iFilialEmpresa < Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                        'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                        colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
                    End If
                        
                    If (objLancamento_Detalhe.iGerencial = 1 Or objLancamento_Detalhe.iGerencial = 2) And ((objLancamento_Cabecalho.iFilialEmpresa < giFilialAuxiliar And giFilialAuxiliar > 0) Or (objLancamento_Cabecalho.iFilialEmpresa > Abs(giFilialAuxiliar) And giFilialAuxiliar < 0)) Then
                        colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
                    End If
                    
                Else
                        
                    'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
                    colLancamento_Detalhe_Aux.Add objLancamento_Detalhe
                End If
                    
            Next
           
            lErro = CF("Lancamento_Atualiza0", objLancamento_Cabecalho, colLancamento_Detalhe_Aux)
            If lErro <> SUCESSO Then gError 186565
        
            If giContabGerencial = 0 Then Exit Do
    
            objLancamento_Cabecalho.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa - giFilialAuxiliar
            
        Loop
        
        DoEvents
        
        If objTela.giStop = 1 Then gError 186565
        
        'Call CF("Excel_Fechar")
        
        lErro = objExcelApp.Fechar()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objImportacaoCTB.iFilialEmpresa = iFilialEmpresa
        
        objImportacaoCTB.sArquivo = Dir(sArq)
        
        objImportacaoCTB.lNumIntDoc = lNumIntDoc
        
        objImportacaoCTB.dtData = gdtDataAtual
        
        objImportacaoCTB.dtHora = Time
        
        lErro = CF("ImportacaoCTB_Exclui", objImportacaoCTB)
        If lErro <> SUCESSO Then gError 186566
        
        lErro = CF("ImportacaoCTB_Grava", objImportacaoCTB)
        If lErro <> SUCESSO Then gError 186567
        
        'Fechar transacao
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 186568
    
        objMsg.AddItem "Arquivo " & sArq & " importado com sucesso."
    
    Loop
    
    Excel_Le_Planilha = SUCESSO
    
    Exit Function
    
Erro_Excel_Le_Planilha:

    Excel_Le_Planilha = gErr
    
    Select Case gErr

        Case 186554
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 186555, 186557 To 186559
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_LINHA", gErr, iIndice)

        Case 186556, 186563, 186565
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_INTERROMPIDA_USUARIO", gErr)
        
        Case 186560
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_PREENCHIDA_LINHA", gErr, iIndice)

        Case 186561
            Call Rotina_Erro(vbOKOnly, "ERRO_COLUNA_DC", gErr, iIndice)

        Case 186562
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ZERADO_LINHA", gErr, iIndice)

        Case 186564, 186566, 186567
        
        Case 186568
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 188298, 207327, 207328
            Call Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_BALANCEADO", gErr, objLancamento_Cabecalho.lDoc)
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186569)

    End Select

    Call Transacao_Rollback

    'Fecha o Excel
    'Call CF("Excel_Fechar")
    Call objExcelApp.Fechar

    Exit Function

End Function

Function RelDRE_Calcula(ByVal sModelo As String, ByVal iExercicio As Integer, ByVal iFilialEmpresa As Integer, Optional ByVal iDetalhado As Integer = DESMARCADO, Optional ByVal iGrupoEmpresarial As Integer = 0) As Long
'Calcula Valor correspondente ao Modelo

Dim lErro As Long, bAchou As Boolean
Dim iOperacao As Integer
Dim dValorExercAtual As Double, dValorFormula As Double, dValorFormulaExercAnt As Double
Dim dValorExercAnt As Double
Dim colRelDRE As New Collection
Dim colRelDREConta As New Collection
Dim colRelDREFormula As New Collection
Dim objRelDRE As New ClassRelDRE
Dim objRelDRE1 As New ClassRelDRE
Dim objRelDREConta As New ClassRelDREConta
Dim objRelDREFormula As New ClassRelDREFormula
Dim sAuxValor1 As String, sAuxValor2 As String
Dim colGrupoEmpAux As Collection, objGrupoEmpAux As ClassGrupoEmpAux
    
On Error GoTo Erro_RelDRE_Calcula
    
    'Lê registros da tabela RelDRE para o Modelo passado como parâmetro
    lErro = CF("RelDRE_Le_Modelo", RELDRE, sModelo, colRelDRE)
    If lErro <> SUCESSO Then Error 43587
    
    'Lê registros da tabela RelDREConta para o Modelo passado como parâmetro
    lErro = CF("RelDREConta_Le_Modelo", RELDRE, sModelo, colRelDREConta)
    If lErro <> SUCESSO Then Error 43588
    
    'Lê registros da tabela RelDREFormula para o Modelo passado como parâmetro
    lErro = CF("RelDREFormula_Le_Modelo", RELDRE, sModelo, colRelDREFormula)
    If lErro <> SUCESSO Then Error 43589
       
    'procurar filiaisempresa que sao extensoes da empresa passada dentro do grupo empresarial, ordenadas por empresa
    'para cada uma colocar numa colecao qual a empresa, filial e lcomando
    lErro = CF("GrupoEmp_AbreComandos", iGrupoEmpresarial, glEmpresa, iFilialEmpresa, 1, colGrupoEmpAux)
    If lErro <> SUCESSO Then Error 43589
    
    For Each objRelDRE In colRelDRE
        
        If iDetalhado = MARCADO Then
        
            Set objRelDRE.objRelDREDetPer = New ClassRelDREDetPer
            
            objRelDRE.objRelDREDetPer.sModelo = objRelDRE.sModelo
            objRelDRE.objRelDREDetPer.sRelatorio = RELDRE
            objRelDRE.objRelDREDetPer.iCodigo = objRelDRE.iCodigo

        End If

        If objRelDRE.iTipo = DRE_TIPO_CONTA Then
        
            For Each objRelDREConta In colRelDREConta
                
                'Se o elemento tiver o mesmo código do elemento da coleção RelDRE
                If objRelDREConta.iCodigo = objRelDRE.iCodigo Then
                    
                    For Each objGrupoEmpAux In colGrupoEmpAux
                    
                        If objRelDRE.iExercicio = CONTAS_EXERCICIO_ANTERIOR Then
                        
                            'Calcula o Valor
                            lErro = CF("MvPerCcl_Calcula_Valor", objGrupoEmpAux.iFilialEmpresa, iExercicio - 1, objRelDREConta.sContaInicial, objRelDREConta.sContaFinal, dValorExercAtual, dValorExercAnt, objRelDREConta.sCclInicial, objRelDREConta.sCclFinal, objGrupoEmpAux.alComando(1))
                            If lErro <> SUCESSO Then Error 60730
                        
                        Else
                        
                            'Calcula o Valor
                            lErro = CF("MvPerCcl_Calcula_Valor", objGrupoEmpAux.iFilialEmpresa, iExercicio, objRelDREConta.sContaInicial, objRelDREConta.sContaFinal, dValorExercAtual, dValorExercAnt, objRelDREConta.sCclInicial, objRelDREConta.sCclFinal, objGrupoEmpAux.alComando(1))
                            If lErro <> SUCESSO Then Error 43590
                        
                        End If
                        
                        'Aculmula o Valor
                        objRelDRE.dValor = objRelDRE.dValor + dValorExercAtual
                        objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt + dValorExercAnt
                        
                        If iDetalhado = MARCADO Then
                           
                            'Calcula o Valor
                            If objRelDRE.iExercicio = CONTAS_EXERCICIO_ANTERIOR Then
                                lErro = RelDREDetPer_Calcula_Valor(objRelDRE.objRelDREDetPer, iFilialEmpresa, iExercicio - 1, objRelDREConta.sContaInicial, objRelDREConta.sContaFinal, objRelDREConta.sCclInicial, objRelDREConta.sCclFinal)
                            Else
                                lErro = RelDREDetPer_Calcula_Valor(objRelDRE.objRelDREDetPer, iFilialEmpresa, iExercicio, objRelDREConta.sContaInicial, objRelDREConta.sContaFinal, objRelDREConta.sCclInicial, objRelDREConta.sCclFinal)
                            End If
                            If lErro <> SUCESSO Then Error 43590
                        
                        End If
                    
                    Next
                    
                End If
                
            Next
            
        ElseIf objRelDRE.iTipo = DRE_TIPO_FORMULA Then
        
            iOperacao = DRE_OPERACAO_SOMA
            
            For Each objRelDREFormula In colRelDREFormula
                
                'Se o elemento tiver o mesmo código do elemento da coleção RelDRE
                If objRelDREFormula.iCodigo = objRelDRE.iCodigo Then
                                        
                    dValorFormula = 0
                    dValorFormulaExercAnt = 0
                    
                    bAchou = False
                    For Each objRelDRE1 In colRelDRE
                        
                        'Pesquisar na coleção RelDRE o elemento que tenha o Código igual ao
                        'código da fórmula do elemento da coleção RelDREFormula
                        If objRelDRE1.iCodigo = objRelDREFormula.iCodigoFormula Then
                        
                            dValorFormula = objRelDRE1.dValor
                            dValorFormulaExercAnt = objRelDRE1.dValorExercAnt
                            bAchou = True
                            Exit For
                            
                        End If
                        
                    Next
                        
                    If bAchou = False Then
                        
                        If left(objRelDREFormula.sFormula, 1) = "*" Then
                        
                            sAuxValor1 = left(objRelDREFormula.sFormula, 5)
                            sAuxValor2 = Mid(objRelDREFormula.sFormula, 6)
                            If Not IsNumeric(sAuxValor2) Then sAuxValor2 = "0"
                            
                            Select Case sAuxValor1
                                Case "*PANT"
                                    dValorFormulaExercAnt = CDbl(sAuxValor2)
                                Case "*PATU"
                                    dValorFormula = CDbl(sAuxValor2)
                            End Select
                        
                        End If
                        
                        Set objRelDRE1 = New ClassRelDRE
                        
                    End If
                    
                    If iOperacao = DRE_OPERACAO_SOMA Then
                        objRelDRE.dValor = objRelDRE.dValor + dValorFormula
                        objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt + dValorFormulaExercAnt
                    Else
                        objRelDRE.dValor = objRelDRE.dValor - dValorFormula
                        objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt - dValorFormulaExercAnt
                    End If
                    
                    If iDetalhado = MARCADO Then
                        
                        If iOperacao = DRE_OPERACAO_SOMA Then
                            objRelDRE.objRelDREDetPer.dValorOrc1 = objRelDRE.objRelDREDetPer.dValorOrc1 + objRelDRE1.objRelDREDetPer.dValorOrc1
                            objRelDRE.objRelDREDetPer.dValorOrc2 = objRelDRE.objRelDREDetPer.dValorOrc2 + objRelDRE1.objRelDREDetPer.dValorOrc2
                            objRelDRE.objRelDREDetPer.dValorOrc3 = objRelDRE.objRelDREDetPer.dValorOrc3 + objRelDRE1.objRelDREDetPer.dValorOrc3
                            objRelDRE.objRelDREDetPer.dValorOrc4 = objRelDRE.objRelDREDetPer.dValorOrc4 + objRelDRE1.objRelDREDetPer.dValorOrc4
                            objRelDRE.objRelDREDetPer.dValorOrc5 = objRelDRE.objRelDREDetPer.dValorOrc5 + objRelDRE1.objRelDREDetPer.dValorOrc5
                            objRelDRE.objRelDREDetPer.dValorOrc6 = objRelDRE.objRelDREDetPer.dValorOrc6 + objRelDRE1.objRelDREDetPer.dValorOrc6
                            objRelDRE.objRelDREDetPer.dValorOrc7 = objRelDRE.objRelDREDetPer.dValorOrc7 + objRelDRE1.objRelDREDetPer.dValorOrc7
                            objRelDRE.objRelDREDetPer.dValorOrc8 = objRelDRE.objRelDREDetPer.dValorOrc8 + objRelDRE1.objRelDREDetPer.dValorOrc8
                            objRelDRE.objRelDREDetPer.dValorOrc9 = objRelDRE.objRelDREDetPer.dValorOrc9 + objRelDRE1.objRelDREDetPer.dValorOrc9
                            objRelDRE.objRelDREDetPer.dValorOrc10 = objRelDRE.objRelDREDetPer.dValorOrc10 + objRelDRE1.objRelDREDetPer.dValorOrc10
                            objRelDRE.objRelDREDetPer.dValorOrc11 = objRelDRE.objRelDREDetPer.dValorOrc11 + objRelDRE1.objRelDREDetPer.dValorOrc11
                            objRelDRE.objRelDREDetPer.dValorOrc12 = objRelDRE.objRelDREDetPer.dValorOrc12 + objRelDRE1.objRelDREDetPer.dValorOrc12
                                    
                            objRelDRE.objRelDREDetPer.dValorReal1 = objRelDRE.objRelDREDetPer.dValorReal1 + objRelDRE1.objRelDREDetPer.dValorReal1
                            objRelDRE.objRelDREDetPer.dValorReal2 = objRelDRE.objRelDREDetPer.dValorReal2 + objRelDRE1.objRelDREDetPer.dValorReal2
                            objRelDRE.objRelDREDetPer.dValorReal3 = objRelDRE.objRelDREDetPer.dValorReal3 + objRelDRE1.objRelDREDetPer.dValorReal3
                            objRelDRE.objRelDREDetPer.dValorReal4 = objRelDRE.objRelDREDetPer.dValorReal4 + objRelDRE1.objRelDREDetPer.dValorReal4
                            objRelDRE.objRelDREDetPer.dValorReal5 = objRelDRE.objRelDREDetPer.dValorReal5 + objRelDRE1.objRelDREDetPer.dValorReal5
                            objRelDRE.objRelDREDetPer.dValorReal6 = objRelDRE.objRelDREDetPer.dValorReal6 + objRelDRE1.objRelDREDetPer.dValorReal6
                            objRelDRE.objRelDREDetPer.dValorReal7 = objRelDRE.objRelDREDetPer.dValorReal7 + objRelDRE1.objRelDREDetPer.dValorReal7
                            objRelDRE.objRelDREDetPer.dValorReal8 = objRelDRE.objRelDREDetPer.dValorReal8 + objRelDRE1.objRelDREDetPer.dValorReal8
                            objRelDRE.objRelDREDetPer.dValorReal9 = objRelDRE.objRelDREDetPer.dValorReal9 + objRelDRE1.objRelDREDetPer.dValorReal9
                            objRelDRE.objRelDREDetPer.dValorReal10 = objRelDRE.objRelDREDetPer.dValorReal10 + objRelDRE1.objRelDREDetPer.dValorReal10
                            objRelDRE.objRelDREDetPer.dValorReal11 = objRelDRE.objRelDREDetPer.dValorReal11 + objRelDRE1.objRelDREDetPer.dValorReal11
                            objRelDRE.objRelDREDetPer.dValorReal12 = objRelDRE.objRelDREDetPer.dValorReal12 + objRelDRE1.objRelDREDetPer.dValorReal12
                                    
                        Else
                            objRelDRE.objRelDREDetPer.dValorOrc1 = objRelDRE.objRelDREDetPer.dValorOrc1 - objRelDRE1.objRelDREDetPer.dValorOrc1
                            objRelDRE.objRelDREDetPer.dValorOrc2 = objRelDRE.objRelDREDetPer.dValorOrc2 - objRelDRE1.objRelDREDetPer.dValorOrc2
                            objRelDRE.objRelDREDetPer.dValorOrc3 = objRelDRE.objRelDREDetPer.dValorOrc3 - objRelDRE1.objRelDREDetPer.dValorOrc3
                            objRelDRE.objRelDREDetPer.dValorOrc4 = objRelDRE.objRelDREDetPer.dValorOrc4 - objRelDRE1.objRelDREDetPer.dValorOrc4
                            objRelDRE.objRelDREDetPer.dValorOrc5 = objRelDRE.objRelDREDetPer.dValorOrc5 - objRelDRE1.objRelDREDetPer.dValorOrc5
                            objRelDRE.objRelDREDetPer.dValorOrc6 = objRelDRE.objRelDREDetPer.dValorOrc6 - objRelDRE1.objRelDREDetPer.dValorOrc6
                            objRelDRE.objRelDREDetPer.dValorOrc7 = objRelDRE.objRelDREDetPer.dValorOrc7 - objRelDRE1.objRelDREDetPer.dValorOrc7
                            objRelDRE.objRelDREDetPer.dValorOrc8 = objRelDRE.objRelDREDetPer.dValorOrc8 - objRelDRE1.objRelDREDetPer.dValorOrc8
                            objRelDRE.objRelDREDetPer.dValorOrc9 = objRelDRE.objRelDREDetPer.dValorOrc9 - objRelDRE1.objRelDREDetPer.dValorOrc9
                            objRelDRE.objRelDREDetPer.dValorOrc10 = objRelDRE.objRelDREDetPer.dValorOrc10 - objRelDRE1.objRelDREDetPer.dValorOrc10
                            objRelDRE.objRelDREDetPer.dValorOrc11 = objRelDRE.objRelDREDetPer.dValorOrc11 - objRelDRE1.objRelDREDetPer.dValorOrc11
                            objRelDRE.objRelDREDetPer.dValorOrc12 = objRelDRE.objRelDREDetPer.dValorOrc12 - objRelDRE1.objRelDREDetPer.dValorOrc12
                                    
                            objRelDRE.objRelDREDetPer.dValorReal1 = objRelDRE.objRelDREDetPer.dValorReal1 - objRelDRE1.objRelDREDetPer.dValorReal1
                            objRelDRE.objRelDREDetPer.dValorReal2 = objRelDRE.objRelDREDetPer.dValorReal2 - objRelDRE1.objRelDREDetPer.dValorReal2
                            objRelDRE.objRelDREDetPer.dValorReal3 = objRelDRE.objRelDREDetPer.dValorReal3 - objRelDRE1.objRelDREDetPer.dValorReal3
                            objRelDRE.objRelDREDetPer.dValorReal4 = objRelDRE.objRelDREDetPer.dValorReal4 - objRelDRE1.objRelDREDetPer.dValorReal4
                            objRelDRE.objRelDREDetPer.dValorReal5 = objRelDRE.objRelDREDetPer.dValorReal5 - objRelDRE1.objRelDREDetPer.dValorReal5
                            objRelDRE.objRelDREDetPer.dValorReal6 = objRelDRE.objRelDREDetPer.dValorReal6 - objRelDRE1.objRelDREDetPer.dValorReal6
                            objRelDRE.objRelDREDetPer.dValorReal7 = objRelDRE.objRelDREDetPer.dValorReal7 - objRelDRE1.objRelDREDetPer.dValorReal7
                            objRelDRE.objRelDREDetPer.dValorReal8 = objRelDRE.objRelDREDetPer.dValorReal8 - objRelDRE1.objRelDREDetPer.dValorReal8
                            objRelDRE.objRelDREDetPer.dValorReal9 = objRelDRE.objRelDREDetPer.dValorReal9 - objRelDRE1.objRelDREDetPer.dValorReal9
                            objRelDRE.objRelDREDetPer.dValorReal10 = objRelDRE.objRelDREDetPer.dValorReal10 - objRelDRE1.objRelDREDetPer.dValorReal10
                            objRelDRE.objRelDREDetPer.dValorReal11 = objRelDRE.objRelDREDetPer.dValorReal11 - objRelDRE1.objRelDREDetPer.dValorReal11
                            objRelDRE.objRelDREDetPer.dValorReal12 = objRelDRE.objRelDREDetPer.dValorReal12 - objRelDRE1.objRelDREDetPer.dValorReal12
                        
                        End If
                    
                    End If
                            
                    iOperacao = objRelDREFormula.iOperacao
                
                End If
                
            Next
        
        End If
        
    Next
    
    lErro = CF("GrupoEmp_FechaComandos", colGrupoEmpAux)
    If lErro <> SUCESSO Then Error 43595
    Set colGrupoEmpAux = Nothing
    
    'Grava o Valor acumulado na tabela RelDRE
    lErro = CF("RelDRE_Grava_Valor", RELDRE, colRelDRE, iDetalhado)
    If lErro <> SUCESSO Then Error 43595
    
    RelDRE_Calcula = SUCESSO
    
    Exit Function
      
Erro_RelDRE_Calcula:

    RelDRE_Calcula = Err
    
    Select Case Err
    
        Case 43587, 43588, 43589, 43590, 43595, 60730
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168173)
        
    End Select
    
    Call CF("GrupoEmp_FechaComandos", colGrupoEmpAux)
    
    Exit Function
    
End Function

Function RelDRP_Calcula(ByVal sModelo As String, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, ByVal iFilialEmpresa As Integer, Optional ByVal iGrupoEmpresarial As Integer = 0) As Long
'Calcula Valor correspondente ao Modelo

Dim lErro As Long, bAchou As Boolean
Dim iOperacao As Integer
Dim dValorFormula As Double, dValorFormulaExercAnt As Double, dValorFormulaPerAcumAnt As Double, dValorFormulaPerAcumAtual As Double
Dim dValorPerAnt As Double
Dim dValorPerAtual As Double
Dim dValorPerAcumAnt As Double
Dim dValorPerAcumAtual As Double
Dim colRelDRE As New Collection
Dim colRelDREConta As New Collection
Dim colRelDREFormula As New Collection
Dim objRelDRE As New ClassRelDRE
Dim objRelDRE1 As New ClassRelDRE
Dim objRelDREConta As New ClassRelDREConta
Dim objRelDREFormula As New ClassRelDREFormula
Dim sAuxValor1 As String, sAuxValor2 As String
Dim colGrupoEmpAux As Collection, objGrupoEmpAux As ClassGrupoEmpAux

On Error GoTo Erro_RelDRP_Calcula
    
    'Lê registros da tabela RelDRE para o Modelo passado como parâmetro
    lErro = CF("RelDRE_Le_Modelo", RELDRE, sModelo, colRelDRE)
    If lErro <> SUCESSO Then Error 55893
    
    'Lê registros da tabela RelDREConta para o Modelo passado como parâmetro
    lErro = CF("RelDREConta_Le_Modelo", RELDRE, sModelo, colRelDREConta)
    If lErro <> SUCESSO Then Error 55894
    
    'Lê registros da tabela RelDREFormula para o Modelo passado como parâmetro
    lErro = CF("RelDREFormula_Le_Modelo", RELDRE, sModelo, colRelDREFormula)
    If lErro <> SUCESSO Then Error 55895
    
    'procurar filiaisempresa que sao extensoes da empresa passada dentro do grupo empresarial, ordenadas por empresa
    'para cada uma colocar numa colecao qual a empresa, filial e lcomando
    lErro = CF("GrupoEmp_AbreComandos", iGrupoEmpresarial, glEmpresa, iFilialEmpresa, 1, colGrupoEmpAux)
    If lErro <> SUCESSO Then Error 55895
    
    For Each objRelDRE In colRelDRE
        
        If objRelDRE.iTipo = DRE_TIPO_CONTA Then
        
            For Each objRelDREConta In colRelDREConta
                
                'Se o elemento tiver o mesmo código do elemento da coleção RelDRE
                If objRelDREConta.iCodigo = objRelDRE.iCodigo Then
                    
                    For Each objGrupoEmpAux In colGrupoEmpAux
                    
                        If objRelDRE.iExercicio = CONTAS_EXERCICIO_ANTERIOR Then
                        
                            'Calcula o Valor
                            lErro = CF("MvPerCcl_Calcula_Valor_Periodo", objGrupoEmpAux.iFilialEmpresa, iExercicio - 1, iPeriodo, objRelDREConta.sContaInicial, objRelDREConta.sContaFinal, dValorPerAnt, dValorPerAtual, dValorPerAcumAnt, dValorPerAcumAtual, objRelDREConta.sCclInicial, objRelDREConta.sCclFinal, objGrupoEmpAux.alComando(1))
                            If lErro <> SUCESSO Then Error 60731
                            
                        Else
                        
                            'Calcula o Valor
                            lErro = CF("MvPerCcl_Calcula_Valor_Periodo", objGrupoEmpAux.iFilialEmpresa, iExercicio, iPeriodo, objRelDREConta.sContaInicial, objRelDREConta.sContaFinal, dValorPerAnt, dValorPerAtual, dValorPerAcumAnt, dValorPerAcumAtual, objRelDREConta.sCclInicial, objRelDREConta.sCclFinal, objGrupoEmpAux.alComando(1))
                            If lErro <> SUCESSO Then Error 55896
                        
                        End If
                        
                        'Aculmula o Valor
                        objRelDRE.dValor = objRelDRE.dValor + dValorPerAtual
                        objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt + dValorPerAnt
                        objRelDRE.dValorPerAcumAnt = objRelDRE.dValorPerAcumAnt + dValorPerAcumAnt
                        objRelDRE.dValorPerAcumAtual = objRelDRE.dValorPerAcumAtual + dValorPerAcumAtual
                    
                    Next
                                    
                End If
                
            Next
            
        ElseIf objRelDRE.iTipo = DRE_TIPO_FORMULA Then
        
            iOperacao = DRE_OPERACAO_SOMA
            
            For Each objRelDREFormula In colRelDREFormula
                
                'Se o elemento tiver o mesmo código do elemento da coleção RelDRE
                If objRelDREFormula.iCodigo = objRelDRE.iCodigo Then
                                        
                    bAchou = False
                    For Each objRelDRE1 In colRelDRE
                        
                        'Pesquisar na coleção RelDRE o elemento que tenha o Código igual ao
                        'código da fórmula do elemento da coleção RelDREFormula
                        If objRelDRE1.iCodigo = objRelDREFormula.iCodigoFormula Then
                        
                            dValorFormula = objRelDRE1.dValor
                            dValorFormulaExercAnt = objRelDRE1.dValorExercAnt
                            dValorFormulaPerAcumAnt = objRelDRE1.dValorPerAcumAnt
                            dValorFormulaPerAcumAtual = objRelDRE1.dValorPerAcumAtual
                            bAchou = True
                            Exit For
                            

                        End If
                        
                    Next
                        
                    If bAchou = False Then
                        
                        dValorFormula = 0
                        dValorFormulaExercAnt = 0
                        dValorFormulaPerAcumAnt = 0
                        dValorFormulaPerAcumAtual = 0
                        
                        If left(objRelDREFormula.sFormula, 1) = "*" Then
                        
                            sAuxValor1 = left(objRelDREFormula.sFormula, 5)
                            sAuxValor2 = Mid(objRelDREFormula.sFormula, 6)
                            If Not IsNumeric(sAuxValor2) Then sAuxValor2 = "0"
                            
                            Select Case sAuxValor1
                            
                                Case "*AATU"
                                    dValorFormulaPerAcumAtual = CDbl(sAuxValor2)
                                Case "*AANT"
                                    dValorFormulaPerAcumAnt = CDbl(sAuxValor2)
                                Case "*PANT"
                                    dValorFormulaExercAnt = CDbl(sAuxValor2)
                                Case "*PATU"
                                    dValorFormula = CDbl(sAuxValor2)
                                    
                            End Select
                        
                        End If
                        
                    End If
                    
                    If iOperacao = DRE_OPERACAO_SOMA Then
                        objRelDRE.dValor = objRelDRE.dValor + dValorFormula
                        objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt + dValorFormulaExercAnt
                        objRelDRE.dValorPerAcumAnt = objRelDRE.dValorPerAcumAnt + dValorFormulaPerAcumAnt
                        objRelDRE.dValorPerAcumAtual = objRelDRE.dValorPerAcumAtual + dValorFormulaPerAcumAtual
                    Else
                        objRelDRE.dValor = objRelDRE.dValor - dValorFormula
                        objRelDRE.dValorExercAnt = objRelDRE.dValorExercAnt - dValorFormulaExercAnt
                        objRelDRE.dValorPerAcumAnt = objRelDRE.dValorPerAcumAnt - dValorFormulaPerAcumAnt
                        objRelDRE.dValorPerAcumAtual = objRelDRE.dValorPerAcumAtual - dValorFormulaPerAcumAtual
                    End If
                            
                    iOperacao = objRelDREFormula.iOperacao
                
                End If
                
            Next
                        
        End If
        
    Next
    
    lErro = CF("GrupoEmp_FechaComandos", colGrupoEmpAux)
    If lErro <> SUCESSO Then Error 55897
    Set colGrupoEmpAux = Nothing
    
    'Grava o Valor acumulado na tabela RelDRE
    lErro = CF("RelDRE_Grava_Valor", RELDRE, colRelDRE)
    If lErro <> SUCESSO Then Error 55897
    
    RelDRP_Calcula = SUCESSO
    
    Exit Function
      
Erro_RelDRP_Calcula:

    RelDRP_Calcula = Err
    
    Select Case Err
    
        Case 55893, 55894, 55895, 55896, 55897, 60731
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 168182)
        
    End Select
    
    Call CF("GrupoEmp_FechaComandos", colGrupoEmpAux)
    
    Exit Function
    
End Function

Public Function Excel_Le_Rateio(ByVal iFilialEmpresa As Integer, ByVal sArquivo As String, ByVal objMsg As Object, ByVal objProgresso As Object, ByVal objTela As Object) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long

'Declarações para manipulação do Excel
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
'Dim objFaixaCelulasPreencher As Object 'Excel.Range
Dim iIndice As Integer
Dim lTransacao As Long
Dim sArq As String
Dim iLinha As Integer
Dim iColuna As Integer
Dim iLinhas As Integer
Dim lCodigo As Long
Dim dRateio As Double
Dim dTotalRateio As Double
Dim objExcelApp As New ClassExcelApp

On Error GoTo Erro_Excel_Le_Rateio

    sArq = sArquivo
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 188349

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 188358
        
    Next

    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'    'Abre o excel
'    lErro = CF("Excel_Abrir")
'    If lErro <> 0 Then gError 188350
'
'    'Desabilita as mensagens do Excel
'    GL_objExcelSist.DisplayAlerts = False
    
    objMsg.AddItem "Abrindo a planilha " & sArq
    
    'Set objPastaTrabalho = GL_objExcelSist.Workbooks.Open(sArq)
    
    'Abre o excel
    lErro = objExcelApp.Abrir_Planilha(sArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    DoEvents
    
    If objTela.giStop = 1 Then gError 188351
    
    objMsg.AddItem "Planilha aberta, iniciando importação"
    
    'Seleciona a planilha ativa na pasta de trabalho criada
    'Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
    
    iColuna = 2
    'lCodigo = StrParaLong(objPlanilhaExcel.Cells(1, iColuna).Value)
    lCodigo = StrParaLong(objExcelApp.Obtem_Valor_Celula(1, iColuna))
    iLinhas = 1
    
    Do While lCodigo <> 0
        iLinhas = iLinhas + 1
        iColuna = iColuna + 1
        'lCodigo = StrParaLong(objPlanilhaExcel.Cells(1, iColuna).Value)
        lCodigo = StrParaLong(objExcelApp.Obtem_Valor_Celula(1, iColuna))
    Loop
    
    objProgresso.Max = iLinhas
    objProgresso.Min = 0
    objProgresso.Value = 0
    
    dTotalRateio = 0
    
    For iLinha = 1 To 8
    
        'dRateio = StrParaDbl(objPlanilhaExcel.Cells(iLinha + 1, 1).Value) / 100
        dRateio = StrParaDbl(objExcelApp.Obtem_Valor_Celula(iLinha + 1, 1)) / 100
    
        'Faz a atualizacao do rateio
        lErro = Comando_Executar(alComando(1), "UPDATE RateioOff SET Percentual = ? WHERE Seq = ?", dRateio, iLinha)
        If lErro <> AD_SQL_SUCESSO Then gError 188356
        
        dTotalRateio = dTotalRateio + dRateio
                    
    Next
    
    If Abs(dTotalRateio - 1) > DELTA_VALORMONETARIO Then gError 188359
    
    objProgresso.Value = objProgresso.Value + 1
    
    DoEvents
    
    If objTela.giStop = 1 Then gError 188352
                    
    
    iColuna = 2
    'lCodigo = StrParaLong(objPlanilhaExcel.Cells(1, iColuna).Value)
    lCodigo = StrParaLong(objExcelApp.Obtem_Valor_Celula(1, iColuna))
    
    Do While lCodigo <> 0
    
        dTotalRateio = 0
                    
        For iLinha = 1 To 8
        
            'dRateio = StrParaDbl(objPlanilhaExcel.Cells(iLinha + 1, iColuna).Value) / 100
            dRateio = StrParaDbl(objExcelApp.Obtem_Valor_Celula(iLinha + 1, iColuna)) / 100
        
            'Faz a atualizacao do rateio
            lErro = Comando_Executar(alComando(1), "UPDATE RateioOff SET Percentual = ? WHERE Codigo = ?  AND Seq = ?", dRateio, lCodigo, iLinha)
            If lErro <> AD_SQL_SUCESSO Then gError 188357
                        
            dTotalRateio = dTotalRateio + dRateio
                        
        Next
    
        If Abs(dTotalRateio - 1) > DELTA_VALORMONETARIO Then gError 188360
    
        iColuna = iColuna + 1
    
        'lCodigo = StrParaLong(objPlanilhaExcel.Cells(1, iColuna).Value)
        lCodigo = StrParaLong(objExcelApp.Obtem_Valor_Celula(1, iColuna))
    
        objProgresso.Value = objProgresso.Value + 1
        
        DoEvents
        
        If objTela.giStop = 1 Then gError 188353
                    
    Loop
    
    lErro = objExcelApp.Fechar()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Call CF("Excel_Fechar")
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 188354

    objMsg.AddItem "Arquivo " & sArq & " importado com sucesso."
    
    Excel_Le_Rateio = SUCESSO
    
    Exit Function
    
Erro_Excel_Le_Rateio:

    Excel_Le_Rateio = gErr
    
    Select Case gErr

        Case 188349
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 188350
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_LINHA", gErr, iIndice)

        Case 188351, 188352, 188353
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_INTERROMPIDA_USUARIO", gErr)
        
        Case 188354
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 188356, 188357
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_RATEIOOFF", gErr)

        Case 188358
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 188359, 188360
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PERCENTUAL_NAO_VALIDA", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188355)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    'Fecha o Excel
    Call objExcelApp.Fechar

    'Call CF("Excel_Fechar")

    Exit Function

End Function

Public Function Excel_Le_Planilha_Ccl(ByVal iFilialEmpresa As Integer, ByVal sArquivo As String, ByVal objMsg As Object, ByVal objProgresso As Object, ByVal objTela As Object) As Long

Dim lErro As Long

'Declarações para manipulação do Excel
'Dim objPastaTrabalho As Object 'Excel.Workbook
'Dim objPlanilhaExcel As Object 'Excel.Worksheet
'Dim objFaixaCelulasPreencher As Object 'Excel.Range
Dim sConta As String
Dim sCCusto As String
Dim sCredDeb As String
Dim dValor As Double
Dim sHistorico As String
Dim iSeq As Integer
Dim objLancamento_Detalhe As ClassLancamento_Detalhe
Dim colLancamento_Detalhe As New Collection
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim dtData As Date
Dim lDoc As Long
Dim objPeriodo As New ClassPeriodo
Dim iIndice As Integer
Dim dValorTotal As Double
Dim lTransacao As Long
Dim sArq As String
Dim iPos As Integer
Dim iPosNew As Integer
Dim sDir As String
Dim objImportacaoCTB As New ClassImportacaoCTB
Dim iLinhas As Integer
Dim iGerencial As Integer
Dim colLancamento_Detalhe_Aux As Collection
Dim lNumIntDoc As Long
Dim objCcl As New ClassCcl
Dim sCcl As String
Dim alComando(1 To 7) As Long
Dim sCclAux As String
Dim sEmpresa As String
Dim sSegmento As String
Dim sCliente As String
Dim sDescEmpresa As String
Dim sDescSegmento As String
Dim sDescCliente As String
Dim iPosEmpresa As Integer
Dim iPosSegmento As Integer
Dim iPosCliente As Integer
Dim sDescricao As String
Dim objExcelApp As New ClassExcelApp

On Error GoTo Erro_Excel_Le_Planilha_Ccl

    iPosNew = 1
    iPos = 1
    
    Do While iPosNew <> 0

        iPosNew = InStr(iPos, sArquivo, Chr(0))
   
        If iPosNew = 0 Then

            sArq = sDir & Mid(sArquivo, iPos)

        Else
        
            If iPos = 1 Then
                sDir = Mid(sArquivo, iPos, iPosNew - iPos) & "\"
            End If
        
            iPos = iPosNew + 1
            iPosNew = InStr(iPos, sArquivo, Chr(0))
        
            sArq = sDir & Mid(sArquivo, iPos, iPosNew - iPos)
            iPos = iPosNew + 1

        End If
        

        'Abrir transacao
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 186554
    
        For iIndice = LBound(alComando) To UBound(alComando)
        
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 188358
            
        Next
    
'        'Abre o excel
'        lErro = CF("Excel_Abrir")
'        If lErro <> 0 Then gError 186555
'
'        'Desabilita as mensagens do Excel
'        GL_objExcelSist.DisplayAlerts = False
        
        'Abre o excel
        lErro = objExcelApp.Abrir()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objMsg.AddItem "Abrindo a planilha " & sArq
        
        'Set objPastaTrabalho = GL_objExcelSist.Workbooks.Open(sArq)
        
        'Abre o excel
        lErro = objExcelApp.Abrir_Planilha(sArq)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        DoEvents
        
        If objTela.giStop = 1 Then gError 186556
        
        objMsg.AddItem "Planilha aberta, iniciando importação"
        
        'Seleciona a planilha ativa na pasta de trabalho criada
        'Set objPlanilhaExcel = objPastaTrabalho.ActiveSheet
        
        
        iIndice = 0
        iSeq = 0
        iLinhas = 0
        
        'sCcl = objPlanilhaExcel.Cells(1, 1).Value
        sCcl = objExcelApp.Obtem_Valor_Celula(1, 1)
        
        Do While sCcl <> ""
            iLinhas = iLinhas + 1
            'sCcl = objPlanilhaExcel.Cells(iLinhas, 1).Value
            sCcl = objExcelApp.Obtem_Valor_Celula(iLinhas, 1)
        Loop
        
        objProgresso.Max = iLinhas
        objProgresso.Min = 0
        objProgresso.Value = 0
        
        'sCcl = objPlanilhaExcel.Cells(1, 1).Value
        sCcl = objExcelApp.Obtem_Valor_Celula(1, 1)
        
        Do While sCcl <> ""
        
            iIndice = iIndice + 1
            
            'sDescricao = objPlanilhaExcel.Cells(iIndice, 2).Value
            sDescricao = objExcelApp.Obtem_Valor_Celula(iIndice, 2)
            
            If Len(sCcl) = 7 Then
                sCcl = Mid(sCcl, 1, 1) & " " & Mid(sCcl, 2, 1) & " " & Mid(sCcl, 3)
            Else
                sCcl = Mid(sCcl, 1, 1) & " " & Mid(sCcl, 2)
            End If
                
            sEmpresa = Mid(sCcl, 1, 2) + Space(7)
            sSegmento = Mid(sCcl, 1, 4) + Space(5)
            sCliente = Mid(sCcl, 1, 7) + Space(2)
        
            sCclAux = String(20, 0)
        
            lErro = Comando_Executar(alComando(1), "SELECT Ccl From Ccl WHERE Ccl = ?", sCclAux, sEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 1
            
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 10
            
            If lErro = AD_SQL_SEM_DADOS Then
            
                iPosEmpresa = InStr(1, sDescricao, "/")
            
                sDescEmpresa = Mid(sDescricao, 1, iPosEmpresa - 1)
            
                lErro = Comando_Executar(alComando(2), "INSERT INTO Ccl (Ccl, DescCcl, TipoCcl, AtivoCcl) VALUES (?,?,1,1)", sEmpresa, sDescEmpresa)
                If lErro <> AD_SQL_SUCESSO Then gError 2
                         
            End If
        
        
            sCclAux = String(20, 0)
        
            lErro = Comando_Executar(alComando(3), "SELECT Ccl From Ccl WHERE Ccl = ?", sCclAux, sSegmento)
            If lErro <> AD_SQL_SUCESSO Then gError 3
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 11
            
            If lErro = AD_SQL_SEM_DADOS Then
            
                iPosEmpresa = InStr(1, sDescricao, "/")
                iPosSegmento = InStr(iPosEmpresa + 1, sDescricao, "/")
            
                sDescSegmento = Mid(sDescricao, 1, iPosSegmento - 1)
            
                lErro = Comando_Executar(alComando(4), "INSERT INTO Ccl (Ccl, DescCcl, TipoCcl, AtivoCcl) VALUES (?,?,1,1)", sSegmento, sDescSegmento)
                If lErro <> AD_SQL_SUCESSO Then gError 2
                         
            End If
        
            sCclAux = String(20, 0)
        
            lErro = Comando_Executar(alComando(5), "SELECT Ccl From Ccl WHERE Ccl = ?", sCclAux, sCliente)
            If lErro <> AD_SQL_SUCESSO Then gError 4
            
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 12
            
            If lErro = AD_SQL_SEM_DADOS Then
            
                iPosEmpresa = InStr(1, sDescricao, "/")
                iPosSegmento = InStr(iPosEmpresa + 1, sDescricao, "/")
                iPosCliente = InStr(iPosSegmento + 1, sDescricao, "/")
            
                sDescCliente = Mid(sDescricao, 1, iPosCliente - 1)
            
                lErro = Comando_Executar(alComando(6), "INSERT INTO Ccl (Ccl, DescCcl, TipoCcl, AtivoCcl) VALUES (?,?,1,1)", sCliente, sDescCliente)
                If lErro <> AD_SQL_SUCESSO Then gError 5
                         
            End If
        
        
            lErro = Comando_Executar(alComando(7), "INSERT INTO Ccl (Ccl, DescCcl, TipoCcl, AtivoCcl) VALUES (?,?,3,1)", sCcl, sDescricao)
            If lErro <> AD_SQL_SUCESSO Then gError 6
        
            'sCcl = objPlanilhaExcel.Cells(iIndice + 1, 1).Value
            sCcl = objExcelApp.Obtem_Valor_Celula(iIndice + 1, 1)
        
            objProgresso.Value = objProgresso.Value + 1
            
            DoEvents
            
            If objTela.giStop = 1 Then gError 186563
        
        Loop
        
        
        'Call CF("Excel_Fechar")
        lErro = objExcelApp.Fechar()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
        'Fechar transacao
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 186568
    
        objMsg.AddItem "Arquivo " & sArq & " importado com sucesso."
    
    Loop
    
    Excel_Le_Planilha_Ccl = SUCESSO
    
    Exit Function
    
Erro_Excel_Le_Planilha_Ccl:

    Excel_Le_Planilha_Ccl = gErr
    
    Select Case gErr

        Case 186554
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 186555, 186557 To 186559
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_LINHA", gErr, iIndice)

        Case 186556, 186563, 186565
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_INTERROMPIDA_USUARIO", gErr)
        
        Case 186560
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_PREENCHIDA_LINHA", gErr, iIndice)

        Case 186561
            Call Rotina_Erro(vbOKOnly, "ERRO_COLUNA_DC", gErr, iIndice)

        Case 186562
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_ZERADO_LINHA", gErr, iIndice)

        Case 186564, 186566, 186567
        
        Case 186568
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 188298
            Call Rotina_Erro(vbOKOnly, "ERRO_DOCUMENTO_NAO_BALANCEADO", gErr, objLancamento_Cabecalho.lDoc)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186569)

    End Select

    Call Transacao_Rollback

    'Fecha o Excel
    'Call CF("Excel_Fechar")
    Call objExcelApp.Fechar

    Exit Function

End Function

Function RateioOff_Automatico_Trans(lCodigo As Long) As Long
'retorna o número do proximo código de Rateio Offline disponivel

Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim lErro As Long
Dim lCodigo1 As Long
Dim lCodigo2 As Long
Dim iIndice As Integer

On Error GoTo Erro_RateioOff_Automatico_Trans

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 36781
    Next

    'Pesquisa o ultimo codigo gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(alComando(1), "SELECT CodRateioOff FROM Configuracao", 0, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 36783

    'le o registro com o número do codigo
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 36784

    'loca o registro que guarda o codigo gerado automaticamente
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 36785

    'verifica se este número ainda não foi utilizado
    'Pesquisa os códigos maiores ou iguais ao codigo gerado automaticamente
    lErro = Comando_Executar(alComando(2), "SELECT Codigo FROM RateioOff WHERE  Codigo >= ? ORDER BY Codigo", lCodigo1, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 36786
        
    'le o primeiro codigo >= ao codigo automatico
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36787
    
    'pesquisa o proximo código disponível
    Do While lErro <> AD_SQL_SEM_DADOS
        
        If lCodigo < lCodigo1 Then
              
            'Se ainda não foi usado ==> este é o número a ser usado
             Exit Do
        
        Else
            'se o número for igual ==> passa para o proximo
            lCodigo = lCodigo + 1
        
            lCodigo2 = lCodigo1
        
            'pesquisa o proximo codigo
            Do While lCodigo2 = lCodigo1 And lErro <> AD_SQL_SEM_DADOS
        
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36788
            
            Loop
            
        End If
        
    Loop
    
    'atualiza o codigo gerado automaticamente
    lErro = Comando_ExecutarPos(alComando(3), "UPDATE Configuracao SET CodRateioOff = ?", alComando(1), lCodigo + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 36789
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
        
    RateioOff_Automatico_Trans = SUCESSO
        
    Exit Function
    
    
Erro_RateioOff_Automatico_Trans:

    RateioOff_Automatico_Trans = Err

    Select Case Err
    
        Case 36781
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 36783, 36784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
        
        Case 36785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONFIGURACAO", Err)
        
        Case 36786, 36787, 36788
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RATEIOOFF1", Err)
        
        Case 36789
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONFIGURACAO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 150717)
        
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
        
End Function

Function RelLcto_Prepara_1(ByVal iFilialEmpresa As Integer, lNumIntRel As Long, ByVal sCtaInic As String, ByVal sCtaFim As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date) As Long
'preenche tabela com lancamentos contabeis numa faixa de contas / periodo para uso em relatorios

Dim lErro As Long
Dim alComando(1 To 2) As Long, iIndice As Integer, sConta As String
Dim tLanc As typeLancamento, sSQL As String
Dim lTransacao As Long

On Error GoTo Erro_RelLcto_Prepara_1

    If lNumIntRel = 0 Then
    
        'Inicia a Transacao
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 199018
    
        'obtem numintrel
        lErro = CF("Config_ObterNumInt", "CTBConfig", "NUM_PROX_REL_SLD_CTA", lNumIntRel)
        If lErro <> SUCESSO Then gError 199019
    
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197980
    Next

    sConta = String(STRING_CONTA, 0)
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
        sSQL = "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, SeqAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, ApropriaCrProd, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Status, Modelo FROM Lancamentos AS L WHERE FilialEmpresa < 50 AND Data >= ? AND Data <= ? AND NOT EXISTS " & _
               " (SELECT Doc FROM Lancamentos AS L1 WHERE L.FilialEmpresa = L1.FilialEmpresa AND L.Origem = L1.Origem AND L.Exercicio = L1.Exercicio AND L.PeriodoLan = L1.PeriodoLan AND L.Doc = L1.Doc AND L1.Conta IN (SELECT Conta FROM ContasTransitorias))"
    Else
        sSQL = "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, SeqAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, ApropriaCrProd, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Status, Modelo FROM Lancamentos AS L WHERE FilialEmpresa = ? And Data >= ? AND Data <= ? AND NOT EXISTS " & _
               " (SELECT Doc FROM Lancamentos AS L1 WHERE L.FilialEmpresa = L1.FilialEmpresa AND L.Origem = L1.Origem AND L.Exercicio = L1.Exercicio AND L.PeriodoLan = L1.PeriodoLan AND L.Doc = L1.Doc AND L1.Conta IN (SELECT Conta FROM ContasTransitorias))"
    
    End If
    
    If sCtaInic <> "" Then

        sSQL = sSQL & " AND Conta >= ?"
        
    End If
    
    If sCtaFim <> "" Then
    
        sSQL = sSQL & " AND Conta <= ?"
    
    End If
    
    If iFilialEmpresa = EMPRESA_TODA Then
    
        If sCtaInic <> "" And sCtaFim <> "" Then
            lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaInic, sCtaFim)
        Else
        
            If sCtaInic <> "" Then
                    lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaInic)
            ElseIf sCtaFim <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaFim)
            Else
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim)
            End If
            
        End If
    
    Else
    
        If sCtaInic <> "" And sCtaFim <> "" Then
            lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaInic, sCtaFim)
        Else
        
            If sCtaInic <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaInic)
            ElseIf sCtaFim <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaFim)
            Else
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim)
            End If
            
        End If
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 197981
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197982
        
    Do While lErro = AD_SQL_SUCESSO
    
        'transfere o lancamento do cadastro de lancamentos pendentes para o cadastro de lancamentos contabilizados
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelLcto (NumIntRel, FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Seq,Lote,PeriodoLote,Data,Conta,Ccl,Historico,Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, DocAglutinado, SeqAglutinado, Aglutinado, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Produto, ApropriaCRProd, Quantidade, DataEstoque, Status, Modelo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntRel, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.iAglutina, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.iEscaninho_Custo, _
            tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.sProduto, tLanc.iApropriaCRProd, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iStatus, tLanc.sModelo)
        If lErro <> AD_SQL_SUCESSO Then gError 197983
                
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197984
        
    Loop
    
    
    
    
    sConta = String(STRING_CONTA, 0)
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
        sSQL = "SELECT 0, '', MAX(Exercicio), MAX(PeriodoLan), 0,0, 0, 0, Data, Conta, '', '', SUM(Valor), 0, 0 , 0, 0, 0, 0, 0, 0, 0, '', 0, MAX(DataEstoque), MAX(ApropriaCrProd), MAX(EscaninhoCusto), MAX(ValorLivroAuxiliar), MAX(ClienteFornecedor), MAX(DocOrigem), MAX(Status), MAX(Modelo) FROM Lancamentos AS L WHERE FilialEmpresa < 50 AND Data >= ? AND Data <= ? AND Valor > 0 AND EXISTS " & _
               " (SELECT Doc FROM Lancamentos AS L1 WHERE L.FilialEmpresa = L1.FilialEmpresa AND L.Origem = L1.Origem AND L.Exercicio = L1.Exercicio AND L.PeriodoLan = L1.PeriodoLan AND L.Doc = L1.Doc AND L1.Conta IN (SELECT Conta FROM ContasTransitorias)) GROUP BY Data, Conta"
    Else
        sSQL = "SELECT MAX(FilialEmpresa), '', MAX(Exercicio), MAX(PeriodoLan), 0, 0, 0, 0, Data, Conta, '', '', SUM(Valor), 0, 0 , 0, 0, 0, 0, 0, 0, 0, '', 0, MAX(DataEstoque), MAX(ApropriaCrProd), MAX(EscaninhoCusto), MAX(ValorLivroAuxiliar), MAX(ClienteFornecedor), MAX(DocOrigem), MAX(Status), MAX(Modelo) FROM Lancamentos AS L WHERE FilialEmpresa = ? And Data >= ? AND Data <= ? AND Valor > 0 AND EXISTS " & _
               " (SELECT Doc FROM Lancamentos AS L1 WHERE L.FilialEmpresa = L1.FilialEmpresa AND L.Origem = L1.Origem AND L.Exercicio = L1.Exercicio AND L.PeriodoLan = L1.PeriodoLan AND L.Doc = L1.Doc AND L1.Conta IN (SELECT Conta FROM ContasTransitorias)) GROUP BY Data, Conta"
    
    End If
    
    If sCtaInic <> "" Then

        sSQL = sSQL & " AND Conta >= ?"
        
    End If
    
    If sCtaFim <> "" Then
    
        sSQL = sSQL & " AND Conta <= ?"
    
    End If
    
    If iFilialEmpresa = EMPRESA_TODA Then
    
        If sCtaInic <> "" And sCtaFim <> "" Then
            lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaInic, sCtaFim)
        Else
        
            If sCtaInic <> "" Then
                    lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaInic)
            ElseIf sCtaFim <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaFim)
            Else
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim)
            End If
            
        End If
    
    Else
    
        If sCtaInic <> "" And sCtaFim <> "" Then
            lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaInic, sCtaFim)
        Else
        
            If sCtaInic <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaInic)
            ElseIf sCtaFim <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaFim)
            Else
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim)
            End If
            
        End If
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 197985
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197986
        
    Do While lErro = AD_SQL_SUCESSO
    
        'transfere o lancamento do cadastro de lancamentos pendentes para o cadastro de lancamentos contabilizados
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelLcto (NumIntRel, FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Seq,Lote,PeriodoLote,Data,Conta,Ccl,Historico,Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, DocAglutinado, SeqAglutinado, Aglutinado, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Produto, ApropriaCRProd, Quantidade, DataEstoque, Status, Modelo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntRel, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.iAglutina, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.iEscaninho_Custo, _
            tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.sProduto, tLanc.iApropriaCRProd, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iStatus, tLanc.sModelo)
        If lErro <> AD_SQL_SUCESSO Then gError 197987
                
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197988
        
    Loop
    
    sConta = String(STRING_CONTA, 0)
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
        sSQL = "SELECT 0, '', MAX(Exercicio), MAX(PeriodoLan), 0,1, 0, 0, Data, Conta, '', '', SUM(Valor), 0, 0 , 0, 0, 0, 0, 0, 0, 0, '', 0, MAX(DataEstoque), MAX(ApropriaCrProd), MAX(EscaninhoCusto), MAX(ValorLivroAuxiliar), MAX(ClienteFornecedor), MAX(DocOrigem), MAX(Status), MAX(Modelo) FROM Lancamentos AS L WHERE FilialEmpresa < 50 AND Data >= ? AND Data <= ? AND Valor < 0 AND EXISTS " & _
               " (SELECT Doc FROM Lancamentos AS L1 WHERE L.FilialEmpresa = L1.FilialEmpresa AND L.Origem = L1.Origem AND L.Exercicio = L1.Exercicio AND L.PeriodoLan = L1.PeriodoLan AND L.Doc = L1.Doc AND L1.Conta IN (SELECT Conta FROM ContasTransitorias)) GROUP BY Data, Conta"
    Else
        sSQL = "SELECT MAX(FilialEmpresa), '', MAX(Exercicio), MAX(PeriodoLan), 0, 1, 0, 0, Data, Conta, '', '', SUM(Valor), 0, 0 , 0, 0, 0, 0, 0, 0, 0, '', 0, MAX(DataEstoque), MAX(ApropriaCrProd), MAX(EscaninhoCusto), MAX(ValorLivroAuxiliar), MAX(ClienteFornecedor), MAX(DocOrigem), MAX(Status), MAX(Modelo) FROM Lancamentos AS L WHERE FilialEmpresa = ? And Data >= ? AND Data <= ? AND Valor < 0 AND EXISTS " & _
               " (SELECT Doc FROM Lancamentos AS L1 WHERE L.FilialEmpresa = L1.FilialEmpresa AND L.Origem = L1.Origem AND L.Exercicio = L1.Exercicio AND L.PeriodoLan = L1.PeriodoLan AND L.Doc = L1.Doc AND L1.Conta IN (SELECT Conta FROM ContasTransitorias)) GROUP BY Data, Conta"
    
    End If
    
    If sCtaInic <> "" Then

        sSQL = sSQL & " AND Conta >= ?"
        
    End If
    
    If sCtaFim <> "" Then
    
        sSQL = sSQL & " AND Conta <= ?"
    
    End If
    
    If iFilialEmpresa = EMPRESA_TODA Then
    
        If sCtaInic <> "" And sCtaFim <> "" Then
            lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaInic, sCtaFim)
        Else
        
            If sCtaInic <> "" Then
                    lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaInic)
            ElseIf sCtaFim <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaFim)
            Else
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim)
            End If
            
        End If
    
    Else
    
        If sCtaInic <> "" And sCtaFim <> "" Then
            lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaInic, sCtaFim)
        Else
        
            If sCtaInic <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaInic)
            ElseIf sCtaFim <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaFim)
            Else
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim)
            End If
            
        End If
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 197985
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197986
        
    Do While lErro = AD_SQL_SUCESSO
    
        'transfere o lancamento do cadastro de lancamentos pendentes para o cadastro de lancamentos contabilizados
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelLcto (NumIntRel, FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Seq,Lote,PeriodoLote,Data,Conta,Ccl,Historico,Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, DocAglutinado, SeqAglutinado, Aglutinado, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Produto, ApropriaCRProd, Quantidade, DataEstoque, Status, Modelo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntRel, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.iAglutina, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.iEscaninho_Custo, _
            tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.sProduto, tLanc.iApropriaCRProd, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iStatus, tLanc.sModelo)
        If lErro <> AD_SQL_SUCESSO Then gError 197987
                
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197988
        
    Loop
    
    
    If lTransacao <> 0 Then
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 199028
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    RelLcto_Prepara_1 = SUCESSO
     
    Exit Function
    
Erro_RelLcto_Prepara_1:

    RelLcto_Prepara_1 = gErr
     
    Select Case gErr
          
        Case 197980
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 197981, 197982, 197984, 197985, 197986, 197988
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTO", gErr)
        
        Case 197983, 197987
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELLCTO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197989)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Function RelCtaSaldo_Prepara_1(ByVal iFilialEmpresa As Integer, ByRef lNumIntRel As Long, ByVal sCtaInic As String, ByVal sCtaFim As String, ByVal dtData As Date, ByVal dtDataFim As Date) As Long
'preenche tabela com saldos de contas contabeis numa data para uso em relatorios

Dim lErro As Long, dSaldo As Double, iExerAnt As Integer
Dim objPer As New ClassPeriodo, alComando(1 To 4) As Long, iIndice As Integer
Dim lTransacao As Long, iExercicio As Integer, iStatus As Integer
Dim sConta As String, sContaAtual As String, dSldIni As Double, dDeb(1 To 12) As Double, dCre(1 To 12) As Double
Dim iMes As Integer

On Error GoTo Erro_RelCtaSaldo_Prepara_1

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 199017
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 199018

    'obtem numintrel
    lErro = CF("Config_ObterNumInt", "CTBConfig", "NUM_PROX_REL_SLD_CTA", lNumIntRel)
    If lErro <> SUCESSO Then gError 199019
    
    lErro = CF("Periodo_Le", dtData, objPer)
    If lErro <> SUCESSO Then gError 199020
    
    'obter exercicio que contenha saldo inicial valido menor ou igual a iExercicioAtual
    lErro = RelCtaSaldo_Prepara1(objPer, alComando(1), iExerAnt)
    If lErro <> SUCESSO Then gError 199021
    
    sConta = String(STRING_CONTA, 0)
    
    If sCtaFim <> "" Then
        lErro = Comando_Executar(alComando(2), "SELECT Exercicio, Conta,SldIni,Deb01,Cre01,Deb02,Cre02,Deb03,Cre03,Deb04,Cre04,Deb05,Cre05,Deb06,Cre06,Deb07,Cre07,Deb08,Cre08,Deb09,Cre09,Deb10,Cre10,Deb11,Cre11,Deb12,Cre12 FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio >= ? AND Exercicio <= ? AND Conta >= ? AND Conta <= ? ORDER BY Conta, Exercicio", _
            iExercicio, sConta, dSldIni, dDeb(1), dCre(1), dDeb(2), dCre(2), dDeb(3), dCre(3), dDeb(4), dCre(4), dDeb(5), dCre(5), dDeb(6), dCre(6), dDeb(7), dCre(7), dDeb(8), dCre(8), dDeb(9), dCre(9), dDeb(10), dCre(10), dDeb(11), dCre(11), dDeb(12), dCre(12), iFilialEmpresa, iExerAnt, objPer.iExercicio, sCtaInic, sCtaFim)
        If lErro <> AD_SQL_SUCESSO Then gError 199022

    Else
        lErro = Comando_Executar(alComando(2), "SELECT Exercicio, Conta,SldIni,Deb01,Cre01,Deb02,Cre02,Deb03,Cre03,Deb04,Cre04,Deb05,Cre05,Deb06,Cre06,Deb07,Cre07,Deb08,Cre08,Deb09,Cre09,Deb10,Cre10,Deb11,Cre11,Deb12,Cre12 FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio >= ? AND Exercicio <= ? AND Conta >= ? ORDER BY Conta, Exercicio", _
            iExercicio, sConta, dSldIni, dDeb(1), dCre(1), dDeb(2), dCre(2), dDeb(3), dCre(3), dDeb(4), dCre(4), dDeb(5), dCre(5), dDeb(6), dCre(6), dDeb(7), dCre(7), dDeb(8), dCre(8), dDeb(9), dCre(9), dDeb(10), dCre(10), dDeb(11), dCre(11), dDeb(12), dCre(12), iFilialEmpresa, iExerAnt, objPer.iExercicio, sCtaInic)
        If lErro <> AD_SQL_SUCESSO Then gError 199023
            
    End If
        
    'percorrer registros gerando os saldos na data para as diversas contas
    sContaAtual = ""
    dSaldo = 0
    iMes = Month(dtData)
    
    Do While lErro = AD_SQL_SUCESSO
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199024
    
        If lErro = AD_SQL_SEM_DADOS Then
        
            lErro = RelCtaSaldo_Prepara2(iFilialEmpresa, objPer, dtData, lNumIntRel, dSaldo, sContaAtual, alComando())
            If lErro <> SUCESSO Then gError 199025
            
            Exit Do
            
        Else
        
            If sContaAtual = "" Then
                sContaAtual = sConta
                dSaldo = dSldIni
            End If
            
            'se terminou de ler os registros de uma conta
            If sContaAtual <> sConta Then
                    
                lErro = RelCtaSaldo_Prepara2(iFilialEmpresa, objPer, dtData, lNumIntRel, dSaldo, sContaAtual, alComando())
                If lErro <> SUCESSO Then gError 199026
                
                sContaAtual = sConta
                dSaldo = dSldIni
                
            End If
        
            If iExercicio <> objPer.iExercicio Then
            
                dSaldo = Round(dSaldo - dDeb(1) + dCre(1) - dDeb(2) + dCre(2) - dDeb(3) + dCre(3) - dDeb(4) + dCre(4) - dDeb(5) + dCre(5) - dDeb(6) + dCre(6) - dDeb(7) + dCre(7) - dDeb(8) + dCre(8) - dDeb(9) + dCre(9) - dDeb(10) + dCre(10) - dDeb(11) + dCre(11) - dDeb(12) + dCre(12), 2)
                
            Else
            
                iIndice = 1
                Do While iIndice < iMes
                    dSaldo = Round(dSaldo + dCre(iIndice) - dDeb(iIndice), 2)
                    iIndice = iIndice + 1
                Loop
                
            End If
        
        End If
    
    Loop
    
    lErro = CF("RelLcto_Prepara_1", iFilialEmpresa, lNumIntRel, sCtaInic, sCtaFim, dtData, dtDataFim)
    If lErro <> SUCESSO Then gError 199027
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 199028
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    RelCtaSaldo_Prepara_1 = SUCESSO
     
    Exit Function
    
Erro_RelCtaSaldo_Prepara_1:

    RelCtaSaldo_Prepara_1 = gErr
     
    Select Case gErr
          
        Case 199017
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 199018
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 199019 To 199021, 199025 To 199027

        Case 199022 To 199024
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVCTA_SALDO", gErr)
        
        Case 199028
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199029)
     
    End Select
     
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function



'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function PlanoContaRefModelo_Exclui(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PlanoContaRefModelo_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 200899

    lErro = CF("PlanoContaRefModelo_Exclui_EmTrans", objPlanoContaRefModelo)
    If lErro <> SUCESSO Then gError 200900

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 200901

    PlanoContaRefModelo_Exclui = SUCESSO

    Exit Function

Erro_PlanoContaRefModelo_Exclui:

    PlanoContaRefModelo_Exclui = gErr

    Select Case gErr

        Case 200899
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 200900

        Case 200901
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200902)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PlanoContaRefModelo_Exclui_EmTrans(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iOficial As Integer

On Error GoTo Erro_PlanoContaRefModelo_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200891
    Next

    'Le a tabela PlanoContaRefModelo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Oficial FROM PlanoContaRefModelo WHERE Codigo= ? ", _
                0, iOficial, objPlanoContaRefModelo.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 200892

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200893

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 200894

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 200895

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoContaRefModelo", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 200896
    
    lErro = CF("PlanoContaRef_Exclui", objPlanoContaRefModelo)
    If lErro <> SUCESSO Then gError 200897

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PlanoContaRefModelo_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_PlanoContaRefModelo_Exclui_EmTrans:

    PlanoContaRefModelo_Exclui_EmTrans = gErr

    Select Case gErr

        Case 200891
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200892, 200893
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFMODELO", gErr)

        Case 200894
            Call Rotina_Erro(vbOKOnly, "ERRO_PLANOCONTAREFMODELO_NAO_CADASTRADO", gErr, objPlanoContaRefModelo.lCodigo)

        Case 200895
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PLANOCONTAREFMODELO", gErr)

        Case 200896
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOCONTAREFMODELO", gErr)
            
        Case 200897

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200898)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PlanoContaRef_Exclui(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim sConta As String
Dim objPlanoContaRef As ClassPlanoContaRef

On Error GoTo Erro_PlanoContaRef_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200903
    Next
    
    sConta = String(STRING_CONTA, 0)

    'Le a tabela PlanoContaRef
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Conta FROM PlanoContaRef WHERE CodigoModelo= ? ", _
                0, sConta, objPlanoContaRefModelo.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 200904

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200905

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 200906
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoContaRef", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 200907
        
        Set objPlanoContaRef = New ClassPlanoContaRef
        
        objPlanoContaRef.lCodigoModelo = objPlanoContaRefModelo.lCodigo
        objPlanoContaRef.sConta = sConta
        
        lErro = CF("PlanoContaRefConfig_Exclui", objPlanoContaRef)
        If lErro <> SUCESSO Then gError 200909
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200908
                
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PlanoContaRef_Exclui = SUCESSO

    Exit Function

Erro_PlanoContaRef_Exclui:

    PlanoContaRef_Exclui = gErr

    Select Case gErr

        Case 200903
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200904, 200905, 200908
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREF", gErr)

        Case 200906
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PLANOCONTAREF", gErr)

        Case 200907
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOCONTAREF", gErr)
            
        Case 200909

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200910)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PlanoContaRefConfig_Exclui(ByVal objPlanoContaRef As ClassPlanoContaRef) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lCodigoModelo As Long

On Error GoTo Erro_PlanoContaRefConfig_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200911
    Next

    'Le a tabela PlanoContaRefConfig
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodigoModelo FROM PlanoContaRefConfig WHERE CodigoModelo= ? AND Conta = ? ", _
                0, lCodigoModelo, objPlanoContaRef.lCodigoModelo, objPlanoContaRef.sConta)
    If lErro <> AD_SQL_SUCESSO Then gError 200912

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200913

    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 200914
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PlanoContaRefConfig", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 200915

        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200916

    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PlanoContaRefConfig_Exclui = SUCESSO

    Exit Function

Erro_PlanoContaRefConfig_Exclui:

    PlanoContaRefConfig_Exclui = gErr

    Select Case gErr

        Case 200911
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200912, 200913, 200916
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFCONFIG", gErr)

        Case 200914
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PLANOCONTAREFCONFIG", gErr)

        Case 200915
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOCONTAREFCONFIG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200917)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function PlanoContaRefModelo_Grava(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PlanoContaRefModelo_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 200918

    lErro = CF("PlanoContaRefModelo_Grava_EmTrans", objPlanoContaRefModelo)
    If lErro <> SUCESSO Then gError 200919

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 200920

    PlanoContaRefModelo_Grava = SUCESSO

    Exit Function

Erro_PlanoContaRefModelo_Grava:

    PlanoContaRefModelo_Grava = gErr

    Select Case gErr

        Case 200918
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 200919

        Case 200920
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200921)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function PlanoContaRefModelo_Grava_EmTrans(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim iOficial As Integer
Dim lCodigo As Long

On Error GoTo Erro_PlanoContaRefModelo_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200922
    Next
    
    If objPlanoContaRefModelo.iOficial = MARCADO Then

        'Le a tabelaPlanoContaRefModelo
        lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo FROM PlanoContaRefModelo WHERE Codigo <> ? AND Oficial = ? AND AnoVigencia = ? ", _
                    0, lCodigo, objPlanoContaRefModelo.lCodigo, MARCADO, objPlanoContaRefModelo.iAnoVigencia)
        If lErro <> AD_SQL_SUCESSO Then gError 200923
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200924
        
        If lErro = AD_SQL_SUCESSO Then gError 200932 'Não deixa gravar um outro modelo oficial para o mesmo ano

        lErro = PlanoContaRef_Grava_ContaCclRef(objPlanoContaRefModelo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    'Le a tabela ContaRefModelo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Oficial FROM PlanoContaRefModelo WHERE Codigo= ? ", _
                0, iOficial, objPlanoContaRefModelo.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 200925

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200926

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 200927

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE PlanoContaRefModelo SET Descricao= ? , Oficial= ? , AnoVigencia= ? , " & _
                "UsuarioAlt= ? , DataAlt= ? , HoraAlt= ? ", alComando(0), _
                objPlanoContaRefModelo.sDescricao, objPlanoContaRefModelo.iOficial, objPlanoContaRefModelo.iAnoVigencia, gsUsuario, Date, CDbl(Time))
        If lErro <> AD_SQL_SUCESSO Then gError 200928

        'lErro = CF("PlanoContaRef_Exclui", objPlanoContaRefModelo)
        'If lErro <> SUCESSO Then gError 200929
        
        'Não exclui mais PlanoContaRef mas tem que excluir PlanoContaRefConfig para poder reinserir depois
        lErro = Comando_Executar(alComando(3), "DELETE FROM PlanoContaRefConfig WHERE CodigoModelo = ?", objPlanoContaRefModelo.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 200928
    Else


        lErro = Comando_Executar(alComando(1), "INSERT INTO PlanoContaRefModelo( Codigo, Descricao, Oficial, AnoVigencia, " & _
                "UsuarioCad, DataCad, HoraCad, UsuarioAlt, DataAlt, " & _
                "HoraAlt) VALUES (?,?,?,?,?,?,?,?,?,?)", _
                objPlanoContaRefModelo.lCodigo, objPlanoContaRefModelo.sDescricao, objPlanoContaRefModelo.iOficial, objPlanoContaRefModelo.iAnoVigencia, _
                gsUsuario, Date, CDbl(Time), gsUsuario, Date, CDbl(Time))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200930

    End If

    'Apenas acerta as configurações pois mão é mais possível incluir ou excluir conta do plano de contas referencial
    lErro = CF("PlanoContaRef_Insere", objPlanoContaRefModelo)
    If lErro <> SUCESSO Then gError 200931

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PlanoContaRefModelo_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PlanoContaRefModelo_Grava_EmTrans:

    PlanoContaRefModelo_Grava_EmTrans = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 200922
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200923 To 200926
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFMODELO", gErr)

        Case 200927
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PLANOCONTAREFMODELO", gErr)

        Case 200928
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PLANOCONTAREFMODELO", gErr)
            
        Case 200929, 200931

        Case 200930
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOCONTAREFMODELO", gErr)

        Case 200932
            Call Rotina_Erro(vbOKOnly, "ERRO_PLANOCONTAREFMODELO_OFICIAL_REPETIDO", gErr, lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200933)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function PlanoContaRef_Grava_ContaCclRef(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long
Dim alComando(1 To 8) As Long
Dim iIndice As Integer
Dim objPlanoContaRef As ClassPlanoContaRef
Dim sConta As String, sCcl As String, sContaRef As String, sContaRefNovo As String
Dim objPlanoContaRefConfig As ClassPlanoContaRefConfig
Dim sMsgErro As String, sContaF As String, sCclF As String
Dim objCTBConfig As New ClassCTBConfig
Dim objExercicio As New ClassExercicio, dtData As Date

On Error GoTo Erro_PlanoContaRef_Grava_ContaCclRef

    objCTBConfig.sCodigo = PLANOCTBREF_VALIDA_CONTA_CCL
    
    lErro = CF("CTBConfig_Le", objCTBConfig)
    If lErro <> SUCESSO And lErro <> 9755 Then gError ERRO_SEM_MENSAGEM

    If StrParaInt(objCTBConfig.sConteudo) <> DESMARCADO Then
    
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 206799
        Next
        
        dtData = StrParaDate("01/01/" & objPlanoContaRefModelo.iAnoVigencia)
        
        lErro = CF("Exercicio_Le_Data", dtData, objExercicio)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = Comando_Executar(alComando(1), "DELETE FROM ContaCclRef")
        If lErro <> AD_SQL_SUCESSO Then gError 206800
        
'        sConta = String(STRING_CONTA, 0)
'
'        'Le a tabela PlanoConta
'        lErro = Comando_Executar(alComando(2), "SELECT Conta FROM PlanoConta WHERE TipoConta = ? ", _
'                    sConta, CONTA_ANALITICA)
'        If lErro <> AD_SQL_SUCESSO Then gError 206801
'
'        'Busca Primeiro
'        lErro = Comando_BuscarPrimeiro(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206802
'
'        Do While lErro <> AD_SQL_SEM_DADOS
'
'            lErro = Comando_Executar(alComando(3), "INSERT INTO ContaCclRef (Conta, Ccl, ContaRef) VALUES (?,?,?) ", sConta, "", "")
'            If lErro <> AD_SQL_SUCESSO Then gError 206803
'
'            'Busca Próximo
'            lErro = Comando_BuscarProximo(alComando(2))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206804
'
'        Loop

         lErro = Comando_Executar(alComando(3), "INSERT INTO ContaCclRef (Conta, Ccl, ContaRef) (SELECT Conta, '', '' FROM PlanoConta WHERE TipoConta = ?) ", CONTA_ANALITICA)
         If lErro <> AD_SQL_SUCESSO Then gError 206803

        
'        sConta = String(STRING_CONTA, 0)
'        sCcl = String(STRING_CCL, 0)
'
'        'Le a tabela ContaCcl
'        lErro = Comando_Executar(alComando(4), "SELECT Conta, Ccl FROM ContaCcl ", _
'                    sConta, sCcl)
'        If lErro <> AD_SQL_SUCESSO Then gError 206805
'
'        'Busca Primeiro
'        lErro = Comando_BuscarPrimeiro(alComando(4))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206806
'
'        Do While lErro <> AD_SQL_SEM_DADOS
'
'            lErro = Comando_Executar(alComando(5), "INSERT INTO ContaCclRef (Conta, Ccl, ContaRef) VALUES (?,?,?) ", sConta, sCcl, "")
'            If lErro <> AD_SQL_SUCESSO Then gError 206807
'
'            'Busca Próximo
'            lErro = Comando_BuscarProximo(alComando(4))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206808
'
'        Loop
        
        lErro = Comando_Executar(alComando(5), "INSERT INTO ContaCclRef (Conta, Ccl, ContaRef) (SELECT DISTINCT Conta, Ccl, '' FROM MvPerCcl WHERE Exercicio >= ? AND Conta <> '' ) ", objExercicio.iExercicio)
        If lErro <> AD_SQL_SUCESSO Then gError 206803
                
        For Each objPlanoContaRef In objPlanoContaRefModelo.colContas
                   
            For Each objPlanoContaRefConfig In objPlanoContaRef.colConfig
        
'                sContaRef = String(255, 0)
'
'                With objPlanoContaRefConfig
'                    lErro = Comando_ExecutarPos(alComando(6), "SELECT ContaRef FROM ContaCclRef WHERE (Conta >= ? OR '' = ?) AND (Conta <= ? OR '' = ?) AND (Ccl >= ? OR '' = ?) AND (Ccl <= ? OR '' = ?)", 0, _
'                                sContaRef, .sContaInicial, .sContaInicial, .sContaFinal, .sContaFinal, .sCclInicial, .sCclInicial, .sCclFinal, .sCclFinal)
'                End With
'                If lErro <> AD_SQL_SUCESSO Then gError 206809
'
'                'Busca Primeiro
'                lErro = Comando_BuscarPrimeiro(alComando(6))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206810
'
'                Do While lErro <> AD_SQL_SEM_DADOS
'
'                    sContaRefNovo = ""
'                    If objPlanoContaRefConfig.iSubtrai = MARCADO Then
'                        sContaRefNovo = Replace(Replace(sContaRef, SEPARADOR & objPlanoContaRef.sContaImp, ""), objPlanoContaRef.sContaImp, "")
'                    Else
'                        sContaRefNovo = IIf(Len(Trim(sContaRef)) = 0, "", sContaRef & SEPARADOR) & objPlanoContaRef.sContaImp
'                    End If
'
'                    lErro = Comando_ExecutarPos(alComando(7), "UPDATE ContaCclRef SET ContaRef= ?  ", alComando(6), sContaRefNovo)
'                    If lErro <> AD_SQL_SUCESSO Then gError 206811
'
'                    'Busca Próximo
'                    lErro = Comando_BuscarProximo(alComando(6))
'                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206812
'
'                Loop
    
                If objPlanoContaRefConfig.iSubtrai = MARCADO Then
                    With objPlanoContaRefConfig
                        lErro = Comando_Executar(alComando(7), "UPDATE ContaCclRef SET ContaRef = REPLACE(REPLACE(ContaRef,?,''),?,'')  WHERE (Conta >= ? OR '' = ?) AND (Conta <= ? OR '' = ?) AND (Ccl >= ? OR '' = ?) AND (Ccl <= ? OR '' = ?)", SEPARADOR & objPlanoContaRef.sContaImp, objPlanoContaRef.sContaImp, .sContaInicial, .sContaInicial, .sContaFinal, .sContaFinal, .sCclInicial, .sCclInicial, .sCclFinal, .sCclFinal)
                    End With
                    If lErro <> AD_SQL_SUCESSO Then gError 206811
                Else
                    With objPlanoContaRefConfig
                        lErro = Comando_Executar(alComando(4), "UPDATE ContaCclRef SET ContaRef = ContaRef + ?  WHERE (Conta >= ? OR '' = ?) AND (Conta <= ? OR '' = ?) AND (Ccl >= ? OR '' = ?) AND (Ccl <= ? OR '' = ?) AND ContaRef <> ''", SEPARADOR & objPlanoContaRef.sContaImp, .sContaInicial, .sContaInicial, .sContaFinal, .sContaFinal, .sCclInicial, .sCclInicial, .sCclFinal, .sCclFinal)
                    End With
                    If lErro <> AD_SQL_SUCESSO Then gError 206811
                
                    With objPlanoContaRefConfig
                        lErro = Comando_Executar(alComando(6), "UPDATE ContaCclRef SET ContaRef = ContaRef + ?  WHERE (Conta >= ? OR '' = ?) AND (Conta <= ? OR '' = ?) AND (Ccl >= ? OR '' = ?) AND (Ccl <= ? OR '' = ?) AND ContaRef = ''", objPlanoContaRef.sContaImp, .sContaInicial, .sContaInicial, .sContaFinal, .sContaFinal, .sCclInicial, .sCclInicial, .sCclFinal, .sCclFinal)
                    End With
                    If lErro <> AD_SQL_SUCESSO Then gError 206811
                
                End If
    
    
            Next
        
        Next
        
        If StrParaInt(objCTBConfig.sConteudo) = MARCADO Then
            
            sConta = String(STRING_CONTA, 0)
            sCcl = String(STRING_CCL, 0)
            sContaRef = String(255, 0)
            
            lErro = Comando_Executar(alComando(8), "SELECT Conta, Ccl, ContaRef FROM ContaCclRef WHERE ContaRef LIKE ? ORDER BY Conta, Ccl", _
                            sConta, sCcl, sContaRef, "%-%")
            If lErro <> AD_SQL_SUCESSO Then gError 206813
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206814
            
        '    Do While lErro <> AD_SQL_SEM_DADOS
        
            If lErro <> AD_SQL_SEM_DADOS Then
            
                sContaF = String(STRING_CONTA, 0)
            
                lErro = Mascara_RetornaContaTela(sConta, sContaF)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                If Len(Trim(sCcl)) > 0 Then
                    
                    'mascara o ccl
                    sCclF = String(STRING_CCL, 0)
            
                    lErro = Mascara_RetornaCclEnxuta(sCcl, sCclF)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Else
                    sCclF = ""
                End If
            
                sMsgErro = IIf(Len(Trim(sMsgErro)) = 0, "", sMsgErro & "|") & "Conta:" & sContaF & " Ccl:" & sCclF & "  Contas Ref:" & sContaRef
        
        '        'Busca Próximo
        '        lErro = Comando_BuscarProximo(alComando(8))
        '        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206815
        '
        '    Loop
        
            End If
            
            If Len(Trim(sMsgErro)) > 0 Then gError 206816
            
        End If
    
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If

    PlanoContaRef_Grava_ContaCclRef = SUCESSO

    Exit Function

Erro_PlanoContaRef_Grava_ContaCclRef:

    PlanoContaRef_Grava_ContaCclRef = gErr

    Select Case gErr
    
        Case 206799
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206800
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANOCONTAREFMODELO", gErr)
            
        Case 206801, 206802, 206804
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", gErr)
        
        Case 206802, 206807
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOCONTAREFMODELO", gErr)

        Case 206805, 206806, 206808
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", gErr)
            
        Case 206809, 206810, 206812 To 206815
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFMODELO", gErr)
            
        Case 206811
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PLANOCONTAREFMODELO", gErr)

        Case 206816
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTAREF_MULTIREFERENCIA_CONTA", gErr, sMsgErro)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206817)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PlanoContaRef_Insere(ByVal objPlanoContaRefModelo As ClassPlanoContaRefModelo) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objPlanoContaRef As ClassPlanoContaRef

On Error GoTo Erro_PlanoContaRef_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200934
    Next

    For Each objPlanoContaRef In objPlanoContaRefModelo.colContas
    
        objPlanoContaRef.lCodigoModelo = objPlanoContaRefModelo.lCodigo

        'lErro = Comando_Executar(alComando(1), "INSERT INTO PlanoContaRef( CodigoModelo, Conta, Descricao, Orientacao, " & _
        '        "Tipo, ContaImp, TipoImp, ValidadeDe, ValidadeAte, Conta0) VALUES (?,?,?,?,?,?,?,?,?,?)", objPlanoContaRef.lCodigoModelo, objPlanoContaRef.sConta, objPlanoContaRef.sDescricao, objPlanoContaRef.sOrientacao, objPlanoContaRef.iTipo, objPlanoContaRef.sContaImp, objPlanoContaRef.iTipoImp, objPlanoContaRef.dtValidadeDe, objPlanoContaRef.dtValidadeAte, objPlanoContaRef.sConta0)
        'If lErro <> SUCESSO Then gError 200935
        
        lErro = CF("PlanoContaRefConfig_Insere", objPlanoContaRef)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200936

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PlanoContaRef_Insere = SUCESSO

    Exit Function

Erro_PlanoContaRef_Insere:

    PlanoContaRef_Insere = gErr

    Select Case gErr

        Case 200934
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200935
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOCONTAREF", gErr)
            
        Case 200936

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200937)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function PlanoContaRefConfig_Insere(ByVal objPlanoContaRef As ClassPlanoContaRef) As Long

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objPlanoContaRefConfig As ClassPlanoContaRefConfig

On Error GoTo Erro_PlanoContaRefConfig_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200938
    Next

    For Each objPlanoContaRefConfig In objPlanoContaRef.colConfig

        objPlanoContaRefConfig.lCodigoModelo = objPlanoContaRef.lCodigoModelo
        objPlanoContaRefConfig.sConta = objPlanoContaRef.sConta

        lErro = Comando_Executar(alComando(1), "INSERT INTO PlanoContaRefConfig( CodigoModelo, Conta, Seq, ContaInicial, " & _
                "ContaFinal, CclInicial, CclFinal, Subtrai) VALUES (?,?,?,?,?,?,?,?)", _
                objPlanoContaRefConfig.lCodigoModelo, objPlanoContaRefConfig.sConta, objPlanoContaRefConfig.iSeq, objPlanoContaRefConfig.sContaInicial, _
                objPlanoContaRefConfig.sContaFinal, objPlanoContaRefConfig.sCclInicial, objPlanoContaRefConfig.sCclFinal, objPlanoContaRefConfig.iSubtrai)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200939

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PlanoContaRefConfig_Insere = SUCESSO

    Exit Function

Erro_PlanoContaRefConfig_Insere:

    PlanoContaRefConfig_Insere = gErr

    Select Case gErr

        Case 200938
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200939
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANOCONTAREFCONFIG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200940)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RelDREDetPer_Exclui(ByVal sRelatorio As String, ByVal sModelo As String) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_RelDREDetPer_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200871
    Next

    'Le a tabelaRelDREDetPer
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM RelDREDetPer WHERE Relatorio= ?  AND Modelo= ? ", _
                0, iAux, sRelatorio, sModelo)
    If lErro <> AD_SQL_SUCESSO Then gError 200872

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200873

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS

'        'Trava registro contra alterações/Leituras
'        lErro = Comando_LockExclusive(alComando(0))
'        If lErro <> SUCESSO Then gError 200874
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RelDREDetPer", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 200875
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200876
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelDREDetPer_Exclui = SUCESSO

    Exit Function

Erro_RelDREDetPer_Exclui:

    RelDREDetPer_Exclui = gErr

    Select Case gErr

        Case 200871
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200872, 200873, 200876
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELDREDETPER", gErr)

        Case 200874
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_RELDREDETPER", gErr)

        Case 200875
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELDREDETPER", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200877)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RelDREDetPer_Insere(ByVal objRelDREDetPer As ClassRelDREDetPer) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_RelDREDetPer_Insere

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200878
    Next
    
    objRelDREDetPer.dValorOrcAte1 = objRelDREDetPer.dValorOrc1
    objRelDREDetPer.dValorOrcAte2 = objRelDREDetPer.dValorOrcAte1 + objRelDREDetPer.dValorOrc2
    objRelDREDetPer.dValorOrcAte3 = objRelDREDetPer.dValorOrcAte2 + objRelDREDetPer.dValorOrc3
    objRelDREDetPer.dValorOrcAte4 = objRelDREDetPer.dValorOrcAte3 + objRelDREDetPer.dValorOrc4
    objRelDREDetPer.dValorOrcAte5 = objRelDREDetPer.dValorOrcAte4 + objRelDREDetPer.dValorOrc5
    objRelDREDetPer.dValorOrcAte6 = objRelDREDetPer.dValorOrcAte5 + objRelDREDetPer.dValorOrc6
    objRelDREDetPer.dValorOrcAte7 = objRelDREDetPer.dValorOrcAte6 + objRelDREDetPer.dValorOrc7
    objRelDREDetPer.dValorOrcAte8 = objRelDREDetPer.dValorOrcAte7 + objRelDREDetPer.dValorOrc8
    objRelDREDetPer.dValorOrcAte9 = objRelDREDetPer.dValorOrcAte8 + objRelDREDetPer.dValorOrc9
    objRelDREDetPer.dValorOrcAte10 = objRelDREDetPer.dValorOrcAte9 + objRelDREDetPer.dValorOrc10
    objRelDREDetPer.dValorOrcAte11 = objRelDREDetPer.dValorOrcAte10 + objRelDREDetPer.dValorOrc11
    objRelDREDetPer.dValorOrcAte12 = objRelDREDetPer.dValorOrcAte11 + objRelDREDetPer.dValorOrc12
    
    objRelDREDetPer.dValorRealAte1 = objRelDREDetPer.dValorReal1
    objRelDREDetPer.dValorRealAte2 = objRelDREDetPer.dValorRealAte1 + objRelDREDetPer.dValorReal2
    objRelDREDetPer.dValorRealAte3 = objRelDREDetPer.dValorRealAte2 + objRelDREDetPer.dValorReal3
    objRelDREDetPer.dValorRealAte4 = objRelDREDetPer.dValorRealAte3 + objRelDREDetPer.dValorReal4
    objRelDREDetPer.dValorRealAte5 = objRelDREDetPer.dValorRealAte4 + objRelDREDetPer.dValorReal5
    objRelDREDetPer.dValorRealAte6 = objRelDREDetPer.dValorRealAte5 + objRelDREDetPer.dValorReal6
    objRelDREDetPer.dValorRealAte7 = objRelDREDetPer.dValorRealAte6 + objRelDREDetPer.dValorReal7
    objRelDREDetPer.dValorRealAte8 = objRelDREDetPer.dValorRealAte7 + objRelDREDetPer.dValorReal8
    objRelDREDetPer.dValorRealAte9 = objRelDREDetPer.dValorRealAte8 + objRelDREDetPer.dValorReal9
    objRelDREDetPer.dValorRealAte10 = objRelDREDetPer.dValorRealAte9 + objRelDREDetPer.dValorReal10
    objRelDREDetPer.dValorRealAte11 = objRelDREDetPer.dValorRealAte10 + objRelDREDetPer.dValorReal11
    objRelDREDetPer.dValorRealAte12 = objRelDREDetPer.dValorRealAte11 + objRelDREDetPer.dValorReal12
    
    lErro = Comando_Executar(alComando(1), "INSERT INTO RelDREDetPer( Relatorio, Modelo, Codigo, ValorOrc1, " & _
            "ValorOrc2, ValorOrc3, ValorOrc4, ValorOrc5, ValorOrc6, " & _
            "ValorOrc7, ValorOrc8, ValorOrc9, ValorOrc10, ValorOrc11, " & _
            "ValorOrc12, ValorReal1, ValorReal2, ValorReal3, ValorReal4, " & _
            "ValorReal5, ValorReal6, ValorReal7, ValorReal8, ValorReal9, " & _
            "ValorReal10, ValorReal11, ValorReal12, ValorOrcEx, ValorRealEx, " & _
            "ValorOrcAte1, ValorOrcAte2, ValorOrcAte3, ValorOrcAte4, ValorOrcAte5, " & _
            "ValorOrcAte6, ValorOrcAte7, ValorOrcAte8, ValorOrcAte9, ValorOrcAte10, " & _
            "ValorOrcAte11, ValorOrcAte12, ValorRealAte1, ValorRealAte2, ValorRealAte3, " & _
            "ValorRealAte4, ValorRealAte5, ValorRealAte6, ValorRealAte7, ValorRealAte8, " & _
            "ValorRealAte9, ValorRealAte10, ValorRealAte11, ValorRealAte12, Conta) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objRelDREDetPer.sRelatorio, objRelDREDetPer.sModelo, objRelDREDetPer.iCodigo, objRelDREDetPer.dValorOrc1, _
            objRelDREDetPer.dValorOrc2, objRelDREDetPer.dValorOrc3, objRelDREDetPer.dValorOrc4, objRelDREDetPer.dValorOrc5, objRelDREDetPer.dValorOrc6, _
            objRelDREDetPer.dValorOrc7, objRelDREDetPer.dValorOrc8, objRelDREDetPer.dValorOrc9, objRelDREDetPer.dValorOrc10, objRelDREDetPer.dValorOrc11, _
            objRelDREDetPer.dValorOrc12, objRelDREDetPer.dValorReal1, objRelDREDetPer.dValorReal2, objRelDREDetPer.dValorReal3, objRelDREDetPer.dValorReal4, _
            objRelDREDetPer.dValorReal5, objRelDREDetPer.dValorReal6, objRelDREDetPer.dValorReal7, objRelDREDetPer.dValorReal8, objRelDREDetPer.dValorReal9, _
            objRelDREDetPer.dValorReal10, objRelDREDetPer.dValorReal11, objRelDREDetPer.dValorReal12, objRelDREDetPer.dValorOrcEx, objRelDREDetPer.dValorRealEx, _
            objRelDREDetPer.dValorOrcAte1, objRelDREDetPer.dValorOrcAte2, objRelDREDetPer.dValorOrcAte3, objRelDREDetPer.dValorOrcAte4, objRelDREDetPer.dValorOrcAte5, _
            objRelDREDetPer.dValorOrcAte6, objRelDREDetPer.dValorOrcAte7, objRelDREDetPer.dValorOrcAte8, objRelDREDetPer.dValorOrcAte9, objRelDREDetPer.dValorOrcAte10, _
            objRelDREDetPer.dValorOrcAte11, objRelDREDetPer.dValorOrcAte12, objRelDREDetPer.dValorRealAte1, objRelDREDetPer.dValorRealAte2, objRelDREDetPer.dValorRealAte3, _
            objRelDREDetPer.dValorRealAte4, objRelDREDetPer.dValorRealAte5, objRelDREDetPer.dValorRealAte6, objRelDREDetPer.dValorRealAte7, objRelDREDetPer.dValorRealAte8, _
            objRelDREDetPer.dValorRealAte9, objRelDREDetPer.dValorRealAte10, objRelDREDetPer.dValorRealAte11, objRelDREDetPer.dValorRealAte12, objRelDREDetPer.sConta)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200879

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelDREDetPer_Insere = SUCESSO

    Exit Function

Erro_RelDREDetPer_Insere:

    RelDREDetPer_Insere = gErr

    Select Case gErr

        Case 200878
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200879
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELDREDETPER", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200880)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RelDREDetPer_Calcula_Valor(ByVal objRelDREDetPer As ClassRelDREDetPer, ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal sContaInicial As String, ByVal sContaFinal As String, ByVal sCclInicial As String, ByVal sCclFinal As String) As Long
'Calcula Valor pela soma de SldIni com os valores de crédito e subtrai os débitos

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim adValor(1 To 12) As Double
Dim dValorOrc As Double
Dim iIndice As Integer
Dim dValor As Double
Dim iExercicio1 As Integer
Dim iPeriodo As Integer
Dim sContaMasc As String
    
On Error GoTo Erro_RelDREDetPer_Calcula_Valor

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200882
    Next
    
    If Len(Trim(sContaInicial)) > 0 And Len(Trim(sCclInicial)) = 0 Then
    
        lErro = Comando_Executar(alComando(1), "SELECT Exercicio, SUM((Cre01-Deb01) * (((Natureza -1)* -2) +1)), SUM((Cre02-Deb02) * (((Natureza -1)* -2) +1)), SUM((Cre03-Deb03) * (((Natureza -1)* -2) +1)), SUM((Cre04-Deb04) * (((Natureza -1)* -2) +1)), SUM((Cre05-Deb05) * (((Natureza -1)* -2) +1)), SUM((Cre06-Deb06) * (((Natureza -1)* -2) +1)) , SUM((Cre07-Deb07) * (((Natureza -1)* -2) +1)), SUM((Cre08-Deb08) * (((Natureza -1)* -2) +1)), SUM((Cre09-Deb09) * (((Natureza -1)* -2) +1)), SUM((Cre10-Deb10) * (((Natureza -1)* -2) +1)), SUM((Cre11-Deb11) * (((Natureza -1)* -2) +1)), SUM((Cre12-Deb12) * (((Natureza -1)* -2) +1)) FROM MvPerCta, PlanoConta WHERE MvPerCta.Conta = PlanoConta.Conta AND TipoConta = 3 AND FilialEmpresa = ? AND Exercicio = ? AND MvPerCta.Conta >= ? AND MvPerCta.Conta <= ? GROUP BY Exercicio", _
            iExercicio1, adValor(1), adValor(2), adValor(3), adValor(4), adValor(5), adValor(6), adValor(7), adValor(8), adValor(9), adValor(10), adValor(11), adValor(12), iFilialEmpresa, iExercicio, sContaInicial, sContaFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 200883
    
    ElseIf Len(Trim(sContaInicial)) = 0 And Len(Trim(sCclInicial)) > 0 Then
    
        lErro = Comando_Executar(alComando(1), "SELECT Exercicio, SUM(Cre01-Deb01), SUM(Cre02-Deb02), SUM(Cre03-Deb03), SUM(Cre04-Deb04), SUM(Cre05-Deb05), SUM(Cre06-Deb06), SUM(Cre07-Deb07), SUM(Cre08-Deb08), SUM(Cre09-Deb09), SUM(Cre10-Deb10), SUM(Cre11-Deb11), SUM(Cre12-Deb12) FROM MvPerCcl, Ccl WHERE MvPerCcl.Ccl = Ccl.Ccl AND TipoCcl = 3 AND FilialEmpresa = ? AND Exercicio = ? AND MvPerCcl.Ccl >= ? AND MvPerCcl.Ccl <= ? AND MvPerCcl.Conta <> '' GROUP BY Exercicio", _
            iExercicio1, adValor(1), adValor(2), adValor(3), adValor(4), adValor(5), adValor(6), adValor(7), adValor(8), adValor(9), adValor(10), adValor(11), adValor(12), iFilialEmpresa, iExercicio, sCclInicial, sCclFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 200884
    
    ElseIf Len(Trim(sContaInicial)) > 0 And Len(Trim(sCclInicial)) > 0 Then
    
        lErro = Comando_Executar(alComando(1), "SELECT Exercicio, SUM((Cre01-Deb01) * (((Natureza -1)* -2) +1)), SUM((Cre02-Deb02) * (((Natureza -1)* -2) +1)), SUM((Cre03-Deb03) * (((Natureza -1)* -2) +1)), SUM((Cre04-Deb04) * (((Natureza -1)* -2) +1)), SUM((Cre05-Deb05) * (((Natureza -1)* -2) +1)), SUM((Cre06-Deb06) * (((Natureza -1)* -2) +1)) , SUM((Cre07-Deb07) * (((Natureza -1)* -2) +1)), SUM((Cre08-Deb08) * (((Natureza -1)* -2) +1)), SUM((Cre09-Deb09) * (((Natureza -1)* -2) +1)), SUM((Cre10-Deb10) * (((Natureza -1)* -2) +1)), SUM((Cre11-Deb11) * (((Natureza -1)* -2) +1)), SUM((Cre12-Deb12) * (((Natureza -1)* -2) +1)) FROM MvPerCcl, Ccl, PlanoConta WHERE MvPerCcl.Ccl = Ccl.Ccl AND MvPerCcl.Conta = PlanoConta.Conta AND TipoCcl = 3 AND TipoConta = 3 AND FilialEmpresa = ? AND Exercicio = ? AND MvPerCcl.Ccl >= ? AND MvPerCcl.Ccl <= ? AND MvPerCcl.Conta >= ? AND MvPerCcl.Conta <= ? GROUP BY Exercicio", _
            iExercicio1, adValor(1), adValor(2), adValor(3), adValor(4), adValor(5), adValor(6), adValor(7), adValor(8), adValor(9), adValor(10), adValor(11), adValor(12), iFilialEmpresa, iExercicio, sCclInicial, sCclFinal, sContaInicial, sContaFinal)
        If lErro <> AD_SQL_SUCESSO Then gError 200885
    
    End If
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 200886
   
    objRelDREDetPer.dValorReal1 = objRelDREDetPer.dValorReal1 + adValor(1)
    objRelDREDetPer.dValorReal2 = objRelDREDetPer.dValorReal2 + adValor(2)
    objRelDREDetPer.dValorReal3 = objRelDREDetPer.dValorReal3 + adValor(3)
    objRelDREDetPer.dValorReal4 = objRelDREDetPer.dValorReal4 + adValor(4)
    objRelDREDetPer.dValorReal5 = objRelDREDetPer.dValorReal5 + adValor(5)
    objRelDREDetPer.dValorReal6 = objRelDREDetPer.dValorReal6 + adValor(6)
    objRelDREDetPer.dValorReal7 = objRelDREDetPer.dValorReal7 + adValor(7)
    objRelDREDetPer.dValorReal8 = objRelDREDetPer.dValorReal8 + adValor(8)
    objRelDREDetPer.dValorReal9 = objRelDREDetPer.dValorReal9 + adValor(9)
    objRelDREDetPer.dValorReal10 = objRelDREDetPer.dValorReal10 + adValor(10)
    objRelDREDetPer.dValorReal11 = objRelDREDetPer.dValorReal11 + adValor(11)
    objRelDREDetPer.dValorReal12 = objRelDREDetPer.dValorReal12 + adValor(12)
    
    For iPeriodo = 1 To 12
    
        dValorOrc = 0
    
        If Len(Trim(sContaInicial)) > 0 And Len(Trim(sCclInicial)) = 0 Then
        
            lErro = Comando_Executar(alComando(2), "SELECT Exercicio, SUM(Valor) FROM Orcamento, PlanoConta WHERE Orcamento.Conta = PlanoConta.Conta AND FilialEmpresa = ? AND Exercicio = ? AND Orcamento.Conta >= ? AND Orcamento.Conta <= ? AND Periodo = ? GROUP BY Exercicio", _
                iExercicio1, dValorOrc, iFilialEmpresa, iExercicio, sContaInicial, sContaFinal, iPeriodo)
            If lErro <> AD_SQL_SUCESSO Then gError 200887
        
        ElseIf Len(Trim(sContaInicial)) = 0 And Len(Trim(sCclInicial)) > 0 Then
        
            lErro = Comando_Executar(alComando(2), "SELECT Exercicio, SUM(Valor) FROM Orcamento, Ccl WHERE Orcamento.Ccl = Ccl.Ccl AND TipoCcl = 3 AND FilialEmpresa = ? AND Exercicio = ? AND Orcamento.Ccl >= ? AND Orcamento.Ccl <= ? AND Orcamento.Conta <> '' AND Periodo = ? GROUP BY Exercicio", _
                iExercicio1, dValorOrc, iFilialEmpresa, iExercicio, sCclInicial, sCclFinal, iPeriodo)
            If lErro <> AD_SQL_SUCESSO Then gError 200888
        
        ElseIf Len(Trim(sContaInicial)) > 0 And Len(Trim(sCclInicial)) > 0 Then
        
            lErro = Comando_Executar(alComando(2), "SELECT Exercicio, SUM(Valor) FROM Orcamento, Ccl, PlanoConta WHERE Orcamento.Ccl = Ccl.Ccl AND Orcamento.Conta = PlanoConta.Conta AND FilialEmpresa = ? AND Exercicio = ? AND Orcamento.Ccl >= ? AND Orcamento.Ccl <= ? AND Orcamento.Conta >= ? AND Orcamento.Conta <= ? AND Periodo = ? GROUP BY Exercicio", _
                iExercicio1, dValorOrc, iFilialEmpresa, iExercicio, sCclInicial, sCclFinal, sContaInicial, sContaFinal, iPeriodo)
            If lErro <> AD_SQL_SUCESSO Then gError 200889
        
        End If
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 200890
        
        Select Case iPeriodo
            Case 1
                objRelDREDetPer.dValorOrc1 = objRelDREDetPer.dValorOrc1 + dValorOrc
            Case 2
                objRelDREDetPer.dValorOrc2 = objRelDREDetPer.dValorOrc2 + dValorOrc
            Case 3
                objRelDREDetPer.dValorOrc3 = objRelDREDetPer.dValorOrc3 + dValorOrc
            Case 4
                objRelDREDetPer.dValorOrc4 = objRelDREDetPer.dValorOrc4 + dValorOrc
            Case 5
                objRelDREDetPer.dValorOrc5 = objRelDREDetPer.dValorOrc5 + dValorOrc
            Case 6
                objRelDREDetPer.dValorOrc6 = objRelDREDetPer.dValorOrc6 + dValorOrc
            Case 7
                objRelDREDetPer.dValorOrc7 = objRelDREDetPer.dValorOrc7 + dValorOrc
            Case 8
                objRelDREDetPer.dValorOrc8 = objRelDREDetPer.dValorOrc8 + dValorOrc
            Case 9
                objRelDREDetPer.dValorOrc9 = objRelDREDetPer.dValorOrc9 + dValorOrc
            Case 10
                objRelDREDetPer.dValorOrc10 = objRelDREDetPer.dValorOrc10 + dValorOrc
            Case 11
                objRelDREDetPer.dValorOrc11 = objRelDREDetPer.dValorOrc11 + dValorOrc
            Case 12
                objRelDREDetPer.dValorOrc12 = objRelDREDetPer.dValorOrc12 + dValorOrc
        End Select
        
    Next
        
    If Len(Trim(sContaInicial)) > 0 And sContaFinal = sContaInicial Then
        If Len(Trim(objRelDREDetPer.sConta)) > 0 Then
            'Se tiver várias contas marca cmo várias
            objRelDREDetPer.sConta = "VÁRIAS"
        Else
            Call Mascara_RetornaContaTela(sContaInicial, sContaMasc)
            objRelDREDetPer.sConta = sContaMasc
        End If
    Else
        objRelDREDetPer.sConta = "VÁRIAS"
    End If
        
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelDREDetPer_Calcula_Valor = SUCESSO
    
    Exit Function
      
Erro_RelDREDetPer_Calcula_Valor:

    RelDREDetPer_Calcula_Valor = gErr
    
    Select Case gErr
           
        Case 200882
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 200883 To 200886
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", gErr)
        
        Case 200887 To 200890
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200891)
        
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function Processa_Lancamento_Analitico_ContaDia_Cust(ByVal sConta As String) As Long
    Processa_Lancamento_Analitico_ContaDia_Cust = SUCESSO
End Function

Public Function RelMvPerCtaRef_Prepara(ByVal iExercicio As Integer, ByVal iPeriodoDe As Integer, ByVal iPeriodoAte As Integer, ByVal iFilialEmpresa As Integer, lNumIntRel As Long) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RelMvPerCtaRef_Prepara

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 200918

    'lErro = CF("RelMvPerCtaRef_Prepara_EmTrans", iExercicio, iPeriodoDe, iPeriodoAte, iFilialEmpresa, lNumIntRel)
    lErro = RelMvPerCtaRef_Prepara_EmTrans2(iExercicio, iPeriodoDe, iPeriodoAte, iFilialEmpresa, lNumIntRel)
    If lErro <> SUCESSO Then gError 200919

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 200920

    RelMvPerCtaRef_Prepara = SUCESSO

    Exit Function

Erro_RelMvPerCtaRef_Prepara:

    RelMvPerCtaRef_Prepara = gErr

    Select Case gErr

        Case 200918
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 200919

        Case 200920
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200921)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function RelMvPerCtaRef_Prepara_EmTrans(ByVal iExercicio As Integer, ByVal iPeriodoDe As Integer, ByVal iPeriodoAte As Integer, ByVal iFilialEmpresa As Integer, lNumIntRel As Long) As Long

Dim lErro As Long
Dim alComando(0 To 9) As Long
Dim iIndice As Integer
Dim objMvPer As ClassRelMvPerCtaRef
Dim objMvPerAux As ClassRelMvPerCtaRef
Dim objMvDia As ClassRelMvDiaCtaRef
Dim objMvDiaAux As ClassRelMvDiaCtaRef
Dim objMvLcto As ClassRelMvLctoCtaRef
Dim colPer As New Collection
Dim lCodigoModelo As Long, sConta As String, sDescricao As String, iTipo As Integer, sContaImp As String, iTipoImp As Integer
Dim sContaInicial As String, sContaFinal As String, sCclInicial As String, sCclFinal  As String, iSubtrai As Integer
Dim tLanc As typeLancamento, sContaRefMask As String, sContaAnt As String
Dim iPeriodoAnt, dtDataAnt As Date, bAchou As Boolean
Dim bComecou1 As Boolean, colPerOrdPer As New Collection, colCampos As New Collection
Dim iPeriodoAux As Integer

On Error GoTo Erro_RelMvPerCtaRef_Prepara_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205572
    Next
    
    bComecou1 = False
    
    'Le o Modelo a ser usado
    lErro = Comando_Executar(alComando(0), "SELECT M.Codigo FROM PlanoContaRefModelo As M, Exercicios AS E WHERE M.Oficial = ? AND M.AnoVigencia = YEAR(E.DataInicio) AND E.Exercicio <= ? ORDER BY M.AnoVigencia DESC ", _
                lCodigoModelo, MARCADO, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then gError 205573

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205574
    
    If lErro <> AD_SQL_SUCESSO Then gError 205575 'Não tem um modelo oficial válido

    lErro = CF("Config_ObterNumInt", "FISConfig", "NUM_PROX_REL_MVPERCTAREF", lNumIntRel)
    If lErro <> SUCESSO Then gError 205632

    sConta = String(STRING_CONTA, 0)
    sContaImp = String(STRING_CONTA, 0)
    sDescricao = String(STRING_PLANOCONTAREF_DESCRICAO, 0)

    'Le as contas referenciais do modelo oficial (primeiro as analíticas "TipoImp")
    lErro = Comando_Executar(alComando(1), "SELECT Conta,Descricao,Tipo,ContaImp,TipoImp FROM PlanoContaRef WHERE CodigoModelo = ? ORDER BY TipoImp DESC, Conta ASC ", _
                sConta, sDescricao, iTipo, sContaImp, iTipoImp, lCodigoModelo)
    If lErro <> AD_SQL_SUCESSO Then gError 205576

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205577

    'Enquanto houver dados
    Do While lErro = AD_SQL_SUCESSO

        iPeriodoAnt = 0
        If iTipoImp = CONTA_ANALITICA Then

            For iPeriodoAux = 1 To iPeriodoAte

                sContaInicial = String(STRING_CONTA, 0)
                sContaFinal = String(STRING_CONTA, 0)
                sCclInicial = String(STRING_CCL, 0)
                sCclFinal = String(STRING_CCL, 0)
        
                'Le as contas referenciais
                lErro = Comando_Executar(alComando(2), "SELECT ContaInicial, ContaFinal, CclInicial, CclFinal, Subtrai FROM PlanoContaRefConfig WHERE CodigoModelo = ? AND Conta = ? ORDER BY Seq ", _
                            sContaInicial, sContaFinal, sCclInicial, sCclFinal, iSubtrai, lCodigoModelo, sConta)
                If lErro <> AD_SQL_SUCESSO Then gError 205578
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205579
            
                Do While lErro = AD_SQL_SUCESSO
            
            
                    tLanc.sOrigem = String(STRING_ORIGEM, 0)
                    tLanc.sHistorico = String(STRING_HISTORICO, 0)
            
                    'Le os lançamentos
                    lErro = Comando_Executar(alComando(3), "SELECT FilialEmpresa, Origem, PeriodoLan, Doc, Seq, Data, Valor, Historico FROM Lancamentos WHERE Exercicio = ? AND PeriodoLan = ? AND (FilialEmpresa = ? OR 0 = ?) AND (Conta >= ? OR '' = ?) AND (Conta <= ? OR '' = ?) AND (Ccl >= ? OR '' = ?) AND (Ccl <= ? OR '' = ?) ", _
                                tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.dtData, tLanc.dValor, tLanc.sHistorico, iExercicio, iPeriodoAux, iFilialEmpresa, iFilialEmpresa, sContaInicial, sContaInicial, sContaFinal, sContaFinal, sCclInicial, sCclInicial, sCclFinal, sCclFinal)
                    If lErro <> AD_SQL_SUCESSO Then gError 205580
                
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(3))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205581
                    
                    Do While lErro = AD_SQL_SUCESSO
                    
                        'Se mudou de período seta o obj novo de período e coloca na coleção
                        If iPeriodoAnt <> tLanc.iPeriodoLan Then
                            Set objMvPer = New ClassRelMvPerCtaRef
                            colPer.Add objMvPer
                            objMvPer.iExercicio = iExercicio
                            objMvPer.iPeriodo = tLanc.iPeriodoLan
                            objMvPer.iFilialEmpresa = iFilialEmpresa
                            objMvPer.iSeqCtaPer = colPer.Count
                            objMvPer.sContaRef = sConta
                            objMvPer.sContaRefMask = sContaImp
                            objMvPer.sDescricao = sDescricao
                            'Se é a mesma conta o Saldo Anterior do período atual
                            'é o saldo anterior do periodo anterios + créditos - débitos
                            If sContaAnt = sConta And objMvPer.iPeriodo > 1 Then
                                objMvPer.dSaldoAnterior = colPer.Item(objMvPer.iSeqCtaPer - 1).dSaldoAnterior + colPer.Item(objMvPer.iSeqCtaPer - 1).dCredito - colPer.Item(objMvPer.iSeqCtaPer - 1).dDebito
                            Else
                                objMvPer.dSaldoAnterior = 0
                            End If
                        End If
                        'Se mudou de período seta o obj novo de dia e coloca na coleção
                        If dtDataAnt <> tLanc.dtData Then
                            Set objMvDia = New ClassRelMvDiaCtaRef
                            objMvPer.colDia.Add objMvDia
                            objMvDia.dtData = tLanc.dtData
                            objMvDia.iSeqCtaDia = objMvPer.colDia.Count
                            objMvDia.iSeqCtaPer = objMvPer.iSeqCtaPer
                            'Se é a mesma conta
                            If sContaAnt = sConta Then
                                'Se não é o primeiro dia do período
                                '   Pega o saldo anterior do dia anterior e soma créditos e débitos
                                'Se é o primeiro dia
                                '   Pega o saldo anterior do último dia do período anterior e soma créditos e débitos
'                                If objMvDia.iSeqCtaDia <> 1 Then
'                                    objMvDia.dSaldoAnterior = objMvPer.colDia.Item(objMvDia.iSeqCtaDia - 1).dSaldoAnterior + objMvPer.colDia.Item(objMvDia.iSeqCtaDia - 1).dCredito - objMvPer.colDia.Item(objMvDia.iSeqCtaDia - 1).dDebito
'                                Else
'                                    objMvDia.dSaldoAnterior = colPer.Item(objMvPer.iSeqCtaPer - 1).colDia.Item(colPer.Item(objMvPer.iSeqCtaPer - 1).colDia.Count).dSaldoAnterior + colPer.Item(objMvPer.iSeqCtaPer - 1).colDia.Item(colPer.Item(objMvPer.iSeqCtaPer - 1).colDia.Count).dCredito - colPer.Item(objMvPer.iSeqCtaPer - 1).colDia.Item(colPer.Item(objMvPer.iSeqCtaPer - 1).colDia.Count).dDebito
'                                End If
                            Else
                                objMvDia.dSaldoAnterior = 0
                            End If
                        End If
                        sContaAnt = sConta
                        dtDataAnt = tLanc.dtData
                        iPeriodoAnt = tLanc.iPeriodoLan
                        'Preenche as informações do Lançamento
                        Set objMvLcto = New ClassRelMvLctoCtaRef
                        objMvDia.colLcto.Add objMvLcto
                        objMvLcto.iFilialEmpresaLcto = tLanc.iFilialEmpresa
                        objMvLcto.iSeqCtaDia = objMvDia.iSeqCtaDia
                        objMvLcto.iSeqCtaPer = objMvDia.iSeqCtaPer
                        objMvLcto.iSeqCtaLcto = objMvDia.colLcto.Count
                        objMvLcto.iSeqLcto = tLanc.iSeq
                        objMvLcto.lDocLcto = tLanc.lDoc
                        objMvLcto.sOrigemLcto = tLanc.sOrigem
                        objMvLcto.dValorLcto = tLanc.dValor
                        objMvLcto.dtDataLcto = tLanc.dtData
                        objMvLcto.sHistoricoLcto = tLanc.sHistorico
                        If objMvLcto.dValorLcto > 0 Then
                            objMvPer.dCredito = objMvPer.dCredito + objMvLcto.dValorLcto
                            objMvDia.dCredito = objMvDia.dCredito + objMvLcto.dValorLcto
                        Else
                            objMvPer.dDebito = objMvPer.dDebito + Abs(objMvLcto.dValorLcto)
                            objMvDia.dDebito = objMvDia.dDebito + Abs(objMvLcto.dValorLcto)
                        End If
            
                        'Busca Próximo
                        lErro = Comando_BuscarProximo(alComando(3))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205582
                
                    Loop
                    
           
                    'Busca Próximo
                    lErro = Comando_BuscarProximo(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205583
        
                Loop
        
        
            Next
        
        Else
        
            'Se passou a tratar as contas sintéticas
            If Not bComecou1 Then
                sContaAnt = ""
                colCampos.Add "iPeriodo"
                'Ordena por período não importa a conta
                Call Ordena_Colecao(colPer, colPerOrdPer, colCampos)
            End If
        
            bComecou1 = True
            
            'Para cada informação já obtida
            For Each objMvPerAux In colPerOrdPer

                'Se é conta filha
                If Trim(sConta) = left(Trim(objMvPerAux.sContaRef), Len(Trim(sConta))) And objMvPerAux.sContaRef <> sConta Then
                    
                    'Verifica se já existe dados para a conta sintética no período
                    'Senão Cria
                    bAchou = False
                    For Each objMvPer In colPer
                        If objMvPer.iPeriodo = objMvPerAux.iPeriodo And objMvPer.sContaRef = sConta Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    If Not bAchou Then
                        'Não tem o período
                        Set objMvPer = New ClassRelMvPerCtaRef
                        colPer.Add objMvPer
                        objMvPer.iSeqCtaPer = colPer.Count
                        objMvPer.iExercicio = iExercicio
                        objMvPer.iPeriodo = objMvPerAux.iPeriodo
                        objMvPer.iFilialEmpresa = iFilialEmpresa
                        objMvPer.sContaRef = sConta
                        objMvPer.sContaRefMask = sContaImp
                        objMvPer.sDescricao = sDescricao
                        'Se é a mesma conta o Saldo Anterior do período atual
                        'é o saldo anterior do periodo anterios + créditos - débitos
                        If sContaAnt = sConta And objMvPer.iPeriodo > 1 Then
                            objMvPer.dSaldoAnterior = colPer.Item(objMvPer.iSeqCtaPer - 1).dSaldoAnterior + colPer.Item(objMvPer.iSeqCtaPer - 1).dCredito - colPer.Item(objMvPer.iSeqCtaPer - 1).dDebito
                        Else
                            objMvPer.dSaldoAnterior = 0
                        End If
                    End If
                    'Soma os créditos e débitos dos filhos
                    objMvPer.dCredito = objMvPer.dCredito + objMvPerAux.dCredito
                    objMvPer.dDebito = objMvPer.dDebito + objMvPerAux.dDebito
                    
                    'Para cada data no período
                    For Each objMvDiaAux In objMvPerAux.colDia
                    bAchou = False
                        For Each objMvDia In objMvPer.colDia
                            If objMvDia.dtData = objMvDiaAux.dtData Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                        If Not bAchou Then
                            Set objMvDia = New ClassRelMvDiaCtaRef
                            objMvPer.colDia.Add objMvDia
                            objMvDia.dtData = objMvDiaAux.dtData
                            objMvDia.iSeqCtaDia = objMvPer.colDia.Count
                            objMvDia.iSeqCtaPer = objMvPer.iSeqCtaPer
                            
                            'OBTER SALDO
                            '***********
                            'NÃO ESTÁ CALCULANDO O SALDO ANTERIOR NAS CONTAS SINTÉTICAS A NÍVEL DE DIA
                            'O CÁLCULO SERIA UM POUCO MAS COMPLEXO E NÃO HOUVE A NECESSIDADE POR ENQUANTO
                        End If
                        objMvDia.dCredito = objMvDia.dCredito + objMvDiaAux.dCredito
                        objMvDia.dDebito = objMvDia.dDebito + objMvDiaAux.dDebito
                                                
                    Next

                End If
                sContaAnt = sConta
                
            Next

        End If
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205584

    Loop
    
    'Para cada conta\período
    For Each objMvPer In colPer
    
        'Se existe crédito ou débito
        If objMvPer.iPeriodo <= iPeriodoAte And (objMvPer.dCredito > DELTA_VALORMONETARIO Or objMvPer.dDebito > DELTA_VALORMONETARIO) Then
        
            objMvPer.lNumIntRel = lNumIntRel
            objMvPer.dSaldoPeriodo = objMvPer.dCredito - objMvPer.dDebito
            objMvPer.dSaldoAtual = objMvPer.dSaldoAnterior + objMvPer.dSaldoPeriodo
            'Se não tem a máscara oficial mascara a conta do cadastro
            If objMvPer.sContaRefMask = "" Then
                If objMvPer.sContaRef <> "" Then
                    Call Retorna_ContaRef_Tela(objMvPer.sContaRef, sContaRefMask)
                    objMvPer.sContaRefMask = sContaRefMask
                End If
            End If
            
            'INSERE
            lErro = Comando_Executar(alComando(4), "INSERT INTO RelMvPerCtaRef( NumIntRel, SeqCtaPer, Exercicio, Periodo, " & _
                    "ContaRef, SaldoAnterior, Credito, Debito, SaldoPeriodo, " & _
                    "SaldoAtual, FilialEmpresa, ContaRefMask, Descricao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    objMvPer.lNumIntRel, objMvPer.iSeqCtaPer, objMvPer.iExercicio, objMvPer.iPeriodo, _
                    objMvPer.sContaRef, objMvPer.dSaldoAnterior, objMvPer.dCredito, objMvPer.dDebito, objMvPer.dSaldoPeriodo, _
                    objMvPer.dSaldoAtual, objMvPer.iFilialEmpresa, objMvPer.sContaRefMask, objMvPer.sDescricao)
            If lErro <> AD_SQL_SUCESSO Then gError 205585
            
            For Each objMvDia In objMvPer.colDia
                objMvDia.dSaldoAtual = objMvDia.dSaldoAnterior + objMvDia.dCredito - objMvDia.dDebito
                objMvDia.lNumIntRel = lNumIntRel
                
                'INSERE
                lErro = Comando_Executar(alComando(5), "INSERT INTO RelMvDiaCtaRef( NumIntRel, SeqCtaPer, SeqCtaDia, Data, " & _
                        "Credito, Debito, SaldoAnterior, SaldoAtual) VALUES (?,?,?,?,?,?,?,?)", _
                        objMvDia.lNumIntRel, objMvDia.iSeqCtaPer, objMvDia.iSeqCtaDia, objMvDia.dtData, _
                        objMvDia.dCredito, objMvDia.dDebito, objMvDia.dSaldoAnterior, objMvDia.dSaldoAtual)
                If lErro <> AD_SQL_SUCESSO Then gError 205586
            
                For Each objMvLcto In objMvDia.colLcto
                    objMvLcto.lNumIntRel = lNumIntRel
                
                    'INSERE
                    lErro = Comando_Executar(alComando(6), "INSERT INTO RelMvLctoCtaRef( NumIntRel, SeqCtaPer, SeqCtaDia, SeqCtaLcto, " & _
                            "FilialEmpresaLcto, OrigemLcto, DocLcto, SeqLcto,ValorLcto, DataLcto,HistoricoLcto) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                            objMvLcto.lNumIntRel, objMvLcto.iSeqCtaPer, objMvLcto.iSeqCtaDia, objMvLcto.iSeqCtaLcto, _
                            objMvLcto.iFilialEmpresaLcto, objMvLcto.sOrigemLcto, objMvLcto.lDocLcto, objMvLcto.iSeqLcto, objMvLcto.dValorLcto, objMvLcto.dtDataLcto, objMvLcto.sHistoricoLcto)
                    If lErro <> AD_SQL_SUCESSO Then gError 205587
                    
                Next
            Next
        End If
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelMvPerCtaRef_Prepara_EmTrans = SUCESSO

    Exit Function

Erro_RelMvPerCtaRef_Prepara_EmTrans:

    RelMvPerCtaRef_Prepara_EmTrans = gErr

    Select Case gErr

        Case 205572
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205573, 205574
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFMODELO", gErr)

        Case 205575 'Não existe modelo válido para o exercício
        
        Case 205576, 205577, 205584
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREF", gErr)

        Case 205578, 205579, 205583
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFCONFIG", gErr)

        Case 205580, 205581, 205582
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS", gErr)

        Case 205585
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMVPERCTAREF", gErr)

        Case 205586
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMVDIACTAREF", gErr)

        Case 205587
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMVLCTOCTAREF", gErr)

        Case 205632

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205588)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RelMvPerCtaRef_Prepara_EmTrans2(ByVal iExercicio As Integer, ByVal iPeriodoDe As Integer, ByVal iPeriodoAte As Integer, ByVal iFilialEmpresa As Integer, lNumIntRel As Long) As Long

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim lCodigoModelo As Long
Dim tContaRef As typePlanoContaRef
Dim sSQLObterSaldo As String, sSQLCredito As String, sSQLDebito As String
Dim dSaldoAnterior As Double, dCredito As Double, dDebito As Double
Dim dSaldoAnterior1 As Double, dCredito1 As Double, dDebito1 As Double
Dim dSaldoAnterior2 As Double, dCredito2 As Double, dDebito2 As Double
Dim iSeqCtaPer As Integer

On Error GoTo Erro_RelMvPerCtaRef_Prepara_EmTrans2

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 205572
    Next
        
    'Le o Modelo a ser usado
    lErro = Comando_Executar(alComando(0), "SELECT M.Codigo FROM PlanoContaRefModelo As M, Exercicios AS E WHERE M.Oficial = ? AND M.AnoVigencia = YEAR(E.DataInicio) AND E.Exercicio <= ? ORDER BY M.AnoVigencia DESC ", _
                lCodigoModelo, MARCADO, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then gError 205573

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205574
    
    If lErro <> AD_SQL_SUCESSO Then gError 205575 'Não tem um modelo oficial válido

    lErro = CF("Config_ObterNumInt", "FISConfig", "NUM_PROX_REL_MVPERCTAREF", lNumIntRel)
    If lErro <> SUCESSO Then gError 205632
    
    sSQLObterSaldo = "SldIni"
    For iIndice = 1 To iPeriodoDe - 1
        sSQLObterSaldo = sSQLObterSaldo & "+(Cre" & Format(iIndice, "00") & " - Deb" & Format(iIndice, "00") & ")"
    Next
    sSQLCredito = "Cre" & Format(iPeriodoAte, "00")
    sSQLDebito = "Deb" & Format(iPeriodoAte, "00")

    tContaRef.sConta = String(STRING_CONTA, 0)
    tContaRef.sContaImp = String(STRING_CONTA, 0)
    tContaRef.sDescricao = String(STRING_PLANOCONTAREF_DESCRICAO, 0)
    tContaRef.sContaPai = String(STRING_CONTA, 0)

    'Le as contas referenciais do modelo oficial (primeiro as analíticas "TipoImp")
    lErro = Comando_Executar(alComando(1), "SELECT Conta,Descricao,Tipo,ContaImp,TipoImp,ContaPai FROM PlanoContaRef WHERE CodigoModelo = ? ORDER BY Conta DESC ", _
                tContaRef.sConta, tContaRef.sDescricao, tContaRef.iTipo, tContaRef.sContaImp, tContaRef.iTipoImp, tContaRef.sContaPai, lCodigoModelo)
    If lErro <> AD_SQL_SUCESSO Then gError 205576

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205577

    'Enquanto houver dados
    Do While lErro = AD_SQL_SUCESSO

        If tContaRef.iTipoImp = CONTA_ANALITICA Then
        
            'Saldo das contas com ccl
            dSaldoAnterior1 = 0
            dCredito1 = 0
            dDebito1 = 0
        
            If iFilialEmpresa = EMPRESA_TODA Then
                lErro = Comando_Executar(alComando(3), "SELECT SUM(" & sSQLObterSaldo & "), SUM(" & sSQLCredito & "), SUM(" & sSQLDebito & ") FROM MvPerCcl As L, ContaCclRef AS X WHERE L.FilialEmpresa <> 0 AND L.Exercicio = ? AND L.Ccl = X.Ccl AND L.Conta = X.Conta AND X.ContaRef = ? AND X.Ccl <> '' ", _
                            dSaldoAnterior1, dCredito1, dDebito1, iExercicio, tContaRef.sContaImp)
            Else
                lErro = Comando_Executar(alComando(3), "SELECT SUM(" & sSQLObterSaldo & "), SUM(" & sSQLCredito & "), SUM(" & sSQLDebito & ") FROM MvPerCcl As L, ContaCclRef AS X WHERE L.FilialEmpresa = ? AND L.Exercicio = ? AND L.Ccl = X.Ccl AND L.Conta = X.Conta AND X.ContaRef = ? AND X.Ccl <> '' ", _
                            dSaldoAnterior1, dCredito1, dDebito1, iFilialEmpresa, iExercicio, tContaRef.sContaImp)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 205580
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205581
            
            'Saldo das contas sem ccl
            dSaldoAnterior2 = 0
            dCredito2 = 0
            dDebito2 = 0
            
            If iFilialEmpresa = EMPRESA_TODA Then
                lErro = Comando_Executar(alComando(5), "SELECT SUM(" & sSQLObterSaldo & "), SUM(" & sSQLCredito & "), SUM(" & sSQLDebito & ") FROM MvPerCta As L, ContaCclRef AS X WHERE L.FilialEmpresa <> 0 AND L.Exercicio = ? AND L.Conta = X.Conta AND X.ContaRef = ? AND X.Ccl = '' ", _
                            dSaldoAnterior2, dCredito2, dDebito2, iExercicio, tContaRef.sContaImp)
            Else
                lErro = Comando_Executar(alComando(5), "SELECT SUM(" & sSQLObterSaldo & "), SUM(" & sSQLCredito & "), SUM(" & sSQLDebito & ") FROM MvPerCta As L, ContaCclRef AS X WHERE L.FilialEmpresa = ? AND L.Exercicio = ? AND L.Conta = X.Conta AND X.ContaRef = ? AND X.Ccl = '' ", _
                            dSaldoAnterior2, dCredito2, dDebito2, iFilialEmpresa, iExercicio, tContaRef.sContaImp)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 205580
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205581
            
            dSaldoAnterior = dSaldoAnterior1 + dSaldoAnterior2
            dCredito = dCredito1 + dCredito2
            dDebito = dDebito1 + dDebito2
            
            iSeqCtaPer = iSeqCtaPer - 1
                    
            lErro = Comando_Executar(alComando(4), "INSERT INTO RelMvPerCtaRef( NumIntRel, SeqCtaPer, Exercicio, Periodo, " & _
                    "ContaRef, SaldoAnterior, Credito, Debito, SaldoPeriodo, " & _
                    "SaldoAtual, FilialEmpresa, ContaRefMask, Descricao, ContaPai) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    lNumIntRel, iSeqCtaPer, iExercicio, iPeriodoAte, _
                    tContaRef.sConta, dSaldoAnterior, dCredito, dDebito, dCredito - dDebito, _
                    dSaldoAnterior + dCredito - dDebito, iFilialEmpresa, tContaRef.sContaImp, tContaRef.sDescricao, tContaRef.sContaPai)
            If lErro <> AD_SQL_SUCESSO Then gError 205585
            
        Else
        'Contas sintéticas
        
            dSaldoAnterior = 0
            dCredito = 0
            dDebito = 0
        
            lErro = Comando_Executar(alComando(2), "SELECT SUM(SaldoAnterior), SUM(Credito), SUM(Debito) FROM RelMvPerCtaRef WHERE NumIntRel = ? AND ContaPai = ? ", _
                        dSaldoAnterior, dCredito, dDebito, lNumIntRel, tContaRef.sContaImp)
            If lErro <> AD_SQL_SUCESSO Then gError 205580
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205581
            
            iSeqCtaPer = iSeqCtaPer - 1
                    
            lErro = Comando_Executar(alComando(4), "INSERT INTO RelMvPerCtaRef( NumIntRel, SeqCtaPer, Exercicio, Periodo, " & _
                    "ContaRef, SaldoAnterior, Credito, Debito, SaldoPeriodo, " & _
                    "SaldoAtual, FilialEmpresa, ContaRefMask, Descricao, ContaPai) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    lNumIntRel, iSeqCtaPer, iExercicio, iPeriodoAte, _
                    tContaRef.sConta, dSaldoAnterior, dCredito, dDebito, dCredito - dDebito, _
                    dSaldoAnterior + dCredito - dDebito, iFilialEmpresa, tContaRef.sContaImp, tContaRef.sDescricao, tContaRef.sContaPai)
            If lErro <> AD_SQL_SUCESSO Then gError 205585
        
        End If
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 205584

    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelMvPerCtaRef_Prepara_EmTrans2 = SUCESSO

    Exit Function

Erro_RelMvPerCtaRef_Prepara_EmTrans2:

    RelMvPerCtaRef_Prepara_EmTrans2 = gErr

    Select Case gErr

        Case 205572
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 205573, 205574
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFMODELO", gErr)

        Case 205575 'Não existe modelo válido para o exercício
        
        Case 205576, 205577, 205584
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREF", gErr)

        Case 205578, 205579, 205583
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTAREFCONFIG", gErr)

        Case 205580, 205581, 205582
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS", gErr)

        Case 205585
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMVPERCTAREF", gErr)

        Case 205586
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMVDIACTAREF", gErr)

        Case 205587
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMVLCTOCTAREF", gErr)

        Case 205632

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205588)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Function Lancamento_Exclui_Cust() As Long
    Lancamento_Exclui_Cust = SUCESSO
End Function

Public Function Executa_ComissaoRegra1_Cust(bForcaInclusao As Boolean) As Long
    bForcaInclusao = False
    Executa_ComissaoRegra1_Cust = SUCESSO
End Function

Private Function RelCtaSaldoTot_Prepara2(ByVal iFilialEmpresa As Integer, ByVal objPer As ClassPeriodo, ByVal dtData As Date, ByVal lNumIntRel As Long, dSaldo As Double, ByVal sContaAtual As String, alComando() As Long) As Long
'obtem valor movimentado no periodo (mes) até a data e insere registro em RelCtaSaldo

Dim lErro As Long, dValMvDia As Double, dValLanPend As Double

On Error GoTo Erro_RelCtaSaldoTot_Prepara2

    If objPer.dtDataInicio <> dtData Then
    
        'obtem valor movimentado no periodo (mes) até a data
        lErro = Comando_Executar(alComando(3), "SELECT SUM(Cre-Deb) FROM MvDiaCta WHERE FilialEmpresa = ? AND Data >= ? AND Data < ? AND Conta = ?", _
            dValMvDia, iFilialEmpresa, objPer.dtDataInicio, dtData, sContaAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 81778
                
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81779
        
        If lErro = AD_SQL_SUCESSO Then dSaldo = Round(dSaldo + dValMvDia, 2)
        
    End If
            
    If iFilialEmpresa = EMPRESA_TODA Then
    
        lErro = Comando_Executar(alComando(5), "SELECT SUM(Valor) FROM LanPendente WHERE Data < ? AND Conta = ?", _
            dValLanPend, dtData, sContaAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 81778
                
        lErro = Comando_BuscarProximo(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81779
        
        If lErro = AD_SQL_SUCESSO Then dSaldo = Round(dSaldo + dValLanPend, 2)
    
    Else
    
        lErro = Comando_Executar(alComando(6), "SELECT SUM(Valor) FROM LanPendente WHERE FilialEmpresa = ? AND Data < ? AND Conta = ?", _
            dValLanPend, iFilialEmpresa, dtData, sContaAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 81778
                
        lErro = Comando_BuscarProximo(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81779
        
        If lErro = AD_SQL_SUCESSO Then dSaldo = Round(dSaldo + dValLanPend, 2)
    
    End If
    
    lErro = Comando_Executar(alComando(4), "INSERT INTO RelCtaSaldo (NumIntRel, Conta, Saldo) VALUES (?,?,?)", lNumIntRel, sContaAtual, dSaldo)
    If lErro <> AD_SQL_SUCESSO Then gError 81780

    RelCtaSaldoTot_Prepara2 = SUCESSO
     
    Exit Function
    
Erro_RelCtaSaldoTot_Prepara2:

    RelCtaSaldoTot_Prepara2 = gErr
     
    Select Case gErr
          
        Case 81778, 81779
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIA_SALDO", gErr)
        
        Case 81780
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELCTASALDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Exit Function

End Function

Function RelCtaSaldoTot_Prepara(ByVal iFilialEmpresa As Integer, ByRef lNumIntRel As Long, ByVal sCtaInic As String, ByVal sCtaFim As String, ByVal dtData As Date, ByVal dtDataFim As Date) As Long
'preenche tabela com saldos de contas contabeis numa data para uso em relatorios

Dim lErro As Long, dSaldo As Double, iExerAnt As Integer
Dim objPer As New ClassPeriodo, alComando(1 To 6) As Long, iIndice As Integer
Dim lTransacao As Long, iExercicio As Integer, iStatus As Integer
Dim sConta As String, sContaAtual As String, dSldIni As Double, dDeb(1 To 12) As Double, dCre(1 To 12) As Double
Dim iMes As Integer

On Error GoTo Erro_RelCtaSaldoTot_Prepara

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81781
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 81782

    'obtem numintrel
    lErro = CF("Config_ObterNumInt", "CTBConfig", "NUM_PROX_REL_SLD_CTA", lNumIntRel)
    If lErro <> SUCESSO Then gError 81783
    
    lErro = CF("Periodo_Le", dtData, objPer)
    If lErro <> SUCESSO Then gError 81784
    
    'obter exercicio que contenha saldo inicial valido menor ou igual a iExercicioAtual
    lErro = RelCtaSaldo_Prepara1(objPer, alComando(1), iExerAnt)
    If lErro <> SUCESSO Then gError 81785
    
    sConta = String(STRING_CONTA, 0)
    
    If sCtaFim <> "" Then
        lErro = Comando_Executar(alComando(2), "SELECT Exercicio, Conta,SldIni,Deb01,Cre01,Deb02,Cre02,Deb03,Cre03,Deb04,Cre04,Deb05,Cre05,Deb06,Cre06,Deb07,Cre07,Deb08,Cre08,Deb09,Cre09,Deb10,Cre10,Deb11,Cre11,Deb12,Cre12 FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio >= ? AND Exercicio <= ? AND Conta >= ? AND Conta <= ? ORDER BY Conta, Exercicio", _
            iExercicio, sConta, dSldIni, dDeb(1), dCre(1), dDeb(2), dCre(2), dDeb(3), dCre(3), dDeb(4), dCre(4), dDeb(5), dCre(5), dDeb(6), dCre(6), dDeb(7), dCre(7), dDeb(8), dCre(8), dDeb(9), dCre(9), dDeb(10), dCre(10), dDeb(11), dCre(11), dDeb(12), dCre(12), iFilialEmpresa, iExerAnt, objPer.iExercicio, sCtaInic, sCtaFim)
        If lErro <> AD_SQL_SUCESSO Then gError 81786

    Else
        lErro = Comando_Executar(alComando(2), "SELECT Exercicio, Conta,SldIni,Deb01,Cre01,Deb02,Cre02,Deb03,Cre03,Deb04,Cre04,Deb05,Cre05,Deb06,Cre06,Deb07,Cre07,Deb08,Cre08,Deb09,Cre09,Deb10,Cre10,Deb11,Cre11,Deb12,Cre12 FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio >= ? AND Exercicio <= ? AND Conta >= ? ORDER BY Conta, Exercicio", _
            iExercicio, sConta, dSldIni, dDeb(1), dCre(1), dDeb(2), dCre(2), dDeb(3), dCre(3), dDeb(4), dCre(4), dDeb(5), dCre(5), dDeb(6), dCre(6), dDeb(7), dCre(7), dDeb(8), dCre(8), dDeb(9), dCre(9), dDeb(10), dCre(10), dDeb(11), dCre(11), dDeb(12), dCre(12), iFilialEmpresa, iExerAnt, objPer.iExercicio, sCtaInic)
        If lErro <> AD_SQL_SUCESSO Then gError 81787
            
    End If
        
    'percorrer registros gerando os saldos na data para as diversas contas
    sContaAtual = ""
    dSaldo = 0
    iMes = Month(dtData)
    
    Do While lErro = AD_SQL_SUCESSO
        
        If lErro = AD_SQL_SEM_DADOS Then
        
            lErro = RelCtaSaldoTot_Prepara2(iFilialEmpresa, objPer, dtData, lNumIntRel, dSaldo, sContaAtual, alComando())
            If lErro <> SUCESSO Then gError 81789
            
            Exit Do
            
        Else
        
            If sContaAtual = "" Then
                sContaAtual = sConta
                dSaldo = dSldIni
            End If
            
            'se terminou de ler os registros de uma conta
            If sContaAtual <> sConta Then
                    
                lErro = RelCtaSaldoTot_Prepara2(iFilialEmpresa, objPer, dtData, lNumIntRel, dSaldo, sContaAtual, alComando())
                If lErro <> SUCESSO Then gError 81790
                
                sContaAtual = sConta
                dSaldo = dSldIni
                
            End If
        
            If iExercicio <> objPer.iExercicio Then
            
                dSaldo = Round(dSaldo - dDeb(1) + dCre(1) - dDeb(2) + dCre(2) - dDeb(3) + dCre(3) - dDeb(4) + dCre(4) - dDeb(5) + dCre(5) - dDeb(6) + dCre(6) - dDeb(7) + dCre(7) - dDeb(8) + dCre(8) - dDeb(9) + dCre(9) - dDeb(10) + dCre(10) - dDeb(11) + dCre(11) - dDeb(12) + dCre(12), 2)
                
            Else
            
                iIndice = 1
                Do While iIndice < iMes
                    dSaldo = Round(dSaldo + dCre(iIndice) - dDeb(iIndice), 2)
                    iIndice = iIndice + 1
                Loop
                
            End If
        
        End If
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81788
    
    Loop
    
    lErro = RelLctoTot_Prepara(iFilialEmpresa, lNumIntRel, sCtaInic, sCtaFim, dtData, dtDataFim)
    If lErro <> SUCESSO Then Error 106537
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 81791
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    RelCtaSaldoTot_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelCtaSaldoTot_Prepara:

    RelCtaSaldoTot_Prepara = gErr
     
    Select Case gErr
          
        Case 81783, 81784, 81785, 81789, 81790, 106537
        
        Case 81781
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 81782
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 81786, 81787, 81788
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVCTA_SALDO", gErr)
        
        Case 81791
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Private Function RelLctoTot_Prepara(ByVal iFilialEmpresa As Integer, lNumIntRel As Long, ByVal sCtaInic As String, ByVal sCtaFim As String, ByVal dtDataIni As Date, ByVal dtDataFim As Date) As Long
'preenche tabela com lancamentos contabeis numa faixa de contas / periodo para uso em relatorios

Dim lErro As Long
Dim alComando(1 To 2) As Long, iIndice As Integer, sConta As String
Dim tLanc As typeLancamento, sSQL As String

On Error GoTo Erro_RelLctoTot_Prepara

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 106532
    Next

    sConta = String(STRING_CONTA, 0)
    
    tLanc.sOrigem = String(STRING_ORIGEM, 0)
    tLanc.sConta = String(STRING_CONTA, 0)
    tLanc.sCcl = String(STRING_CCL, 0)
    tLanc.sHistorico = String(STRING_HISTORICO, 0)
    tLanc.sProduto = String(STRING_PRODUTO, 0)
    tLanc.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tLanc.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    If iFilialEmpresa = EMPRESA_TODA Then
        sSQL = "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, SeqAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, ApropriaCrProd, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Status, Modelo FROM LancamentosTodos2 WHERE Data >= ? AND Data <= ?"
    Else
        sSQL = "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn , CliForn, Transacao, Aglutinado, DocAglutinado, SeqAglutinado, ContaSimples, SeqContraPartida, Produto, Quantidade, DataEstoque, ApropriaCrProd, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Status, Modelo FROM LancamentosTodos2 WHERE FilialEmpresa = ? And Data >= ? AND Data <= ?"
    End If
    
    If sCtaInic <> "" Then

        sSQL = sSQL & " AND Conta >= ?"
        
    End If
    
    If sCtaFim <> "" Then
    
        sSQL = sSQL & " AND Conta <= ?"
    
    End If
    
    If iFilialEmpresa = EMPRESA_TODA Then
    
        If sCtaInic <> "" And sCtaFim <> "" Then
            lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaInic, sCtaFim)
        Else
        
            If sCtaInic <> "" Then
                    lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaInic)
            ElseIf sCtaFim <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim, sCtaFim)
            Else
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, dtDataIni, dtDataFim)
            End If
            
        End If
    
    Else
    
        If sCtaInic <> "" And sCtaFim <> "" Then
            lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaInic, sCtaFim)
        Else
        
            If sCtaInic <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaInic)
            ElseIf sCtaFim <> "" Then
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim, sCtaFim)
            Else
                lErro = Comando_Executar(alComando(1), sSQL, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.iAglutina, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.sProduto, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iApropriaCRProd, tLanc.iEscaninho_Custo, tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.iStatus, tLanc.sModelo, iFilialEmpresa, dtDataIni, dtDataFim)
            End If
            
        End If
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 106533
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106534
        
    Do While lErro = AD_SQL_SUCESSO
    
        'transfere o lancamento do cadastro de lancamentos pendentes para o cadastro de lancamentos contabilizados
        lErro = Comando_Executar(alComando(2), "INSERT INTO RelLcto (NumIntRel, FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Seq,Lote,PeriodoLote,Data,Conta,Ccl,Historico,Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, DocAglutinado, SeqAglutinado, Aglutinado, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Produto, ApropriaCRProd, Quantidade, DataEstoque, Status, Modelo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntRel, tLanc.iFilialEmpresa, tLanc.sOrigem, tLanc.iExercicio, tLanc.iPeriodoLan, tLanc.lDoc, tLanc.iSeq, tLanc.iLote, tLanc.iPeriodoLote, tLanc.dtData, tLanc.sConta, tLanc.sCcl, tLanc.sHistorico, tLanc.dValor, tLanc.lNumIntDoc, tLanc.iFilialCliForn, tLanc.lCliForn, tLanc.iTransacao, tLanc.lDocAglutinado, tLanc.iSeqAglutinado, tLanc.iAglutina, tLanc.lContaSimples, tLanc.iSeqContraPartida, tLanc.iEscaninho_Custo, _
            tLanc.dValorLivroAuxiliar, tLanc.iClienteFornecedor, tLanc.sDocOrigem, tLanc.sProduto, tLanc.iApropriaCRProd, tLanc.dQuantidade, tLanc.dtDataEstoque, tLanc.iStatus, tLanc.sModelo)
        If lErro <> AD_SQL_SUCESSO Then gError 106535
                
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106536
        
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    RelLctoTot_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelLctoTot_Prepara:

    RelLctoTot_Prepara = gErr
     
    Select Case gErr
          
        Case 106532
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 106533, 106534, 106536
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTO", gErr)
        
        Case 106535
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELLCTO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Public Function Lancamento_Valida_ContaCcl(ByVal sConta As String, ByVal sCcl As String) As Long
'Validações extras do preenchimento entre conta e ccl

Dim lErro As Long
Dim alComando(0 To 0) As Long, iIndice As Integer
Dim iNaturezaSped As Integer

On Error GoTo Erro_Lancamento_Valida_ContaCcl

    'Faz o teste da conta de resultado só se o ccl estiver em branco e usar ccl
    If Len(Trim(sCcl)) = 0 And giSetupUsoCcl <> CCL_NAO_USA Then
        
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 209259
        Next
    
        'Pesquisa Conta associada a um Lançamento
        lErro = Comando_Executar(alComando(0), "SELECT NaturezaSped FROM PlanoConta WHERE Conta=?", iNaturezaSped, sConta)
        If lErro <> AD_SQL_SUCESSO Then gError 209260
         
        'Lê a Conta
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209261
         
        'Se é uma conta de resultado
        If iNaturezaSped = 4 Then gError 209262
        
        For iIndice = LBound(alComando) To UBound(alComando)
             Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If
   
    Lancamento_Valida_ContaCcl = SUCESSO
     
    Exit Function
    
Erro_Lancamento_Valida_ContaCcl:

    Lancamento_Valida_ContaCcl = gErr
     
    Select Case gErr
          
        Case 209259
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209260, 209261
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA3", gErr, sConta)
            
        Case 209262
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTARESULTADO_SEM_CCL", gErr, sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209263)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Function Movimento_Verifica_Bloqueio_Data(ByVal dtData As Date, Optional ByVal iTipo As Integer = 1, Optional ByVal iTestaPeriodoContabil As Integer = 0, Optional ByVal dtDataLimite As Date = DATA_NULA) As Long

Dim lErro As Long, iExercicio As Integer, iPeriodo As Integer, iStatus As Integer
Dim objPeriodosFilial As New ClassPeriodosFilial, sTipo As String
Dim alComando(1 To 3) As Long, iIndice As Integer

On Error GoTo Erro_Movimento_Verifica_Bloqueio_Data

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213146
    Next
    
    Select Case iTipo
    
        Case 1
            sTipo = "Movimentos de Estoque"
    
        Case 2
            sTipo = "Movimentos Financeiros"
            
    End Select
    
    If iTestaPeriodoContabil = MARCADO And gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        
        'Obtém o exercício
        lErro = Comando_Executar(alComando(1), "SELECT Exercicio, Status FROM Exercicios WHERE DataInicio <= ? AND DataFim >= ?", iExercicio, iStatus, dtData, dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 213147
        
        'le o exercicio, se estiver no BD
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213148
        
        'se o exercicio estiver fechado, erro
        If giDesconsideraFechamentoPeriodo <> MARCADO Then
            If iStatus = EXERCICIO_FECHADO Then gError 213149
        End If
        
        'verifica qual o periodo em questão
        lErro = Comando_Executar(alComando(2), "SELECT Periodo FROM Periodo WHERE Exercicio=? AND DataInicio <= ? AND DataFim >= ?", iPeriodo, iExercicio, dtData, dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 213150
        
        'le o periodo, se estiver no BD
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213151
        
        objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
        objPeriodosFilial.iExercicio = iExercicio
        objPeriodosFilial.iPeriodo = iPeriodo
    
        lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'se o periodo estiver fechado, erro
        If giDesconsideraFechamentoPeriodo <> MARCADO Then
            If objPeriodosFilial.iFechado = PERIODO_FECHADO Then gError 213152
        End If
        
    End If
    
    If dtDataLimite <> DATA_NULA Then
        If dtDataLimite > dtData Then gError 213154
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Movimento_Verifica_Bloqueio_Data = SUCESSO

    Exit Function

Erro_Movimento_Verifica_Bloqueio_Data:
    
    Movimento_Verifica_Bloqueio_Data = gErr
    
    Select Case gErr
    
        Case 213146
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 213147, 213148
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO1", gErr)
    
        Case 213149
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_FECHADO", gErr, iExercicio)
    
        Case 213150, 213151
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO", gErr)
    
        Case 213152
            Call Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", gErr, iPeriodo, iExercicio)
    
        Case 213154
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVTO_BLOQUEIA_DATA", gErr, sTipo, dtDataLimite)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213153)
        
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function ImportCtb_Valida_Emp(ByVal sNomeArq As String, ByRef bValidaEmp As Boolean) As Long
    bValidaEmp = True
    ImportCtb_Valida_Emp = SUCESSO
End Function

Public Function ImportCtb_Insere_Lcto(ByVal sNomeArq As String, ByVal iFilialEmpresa As Integer, ByVal sCcl As String, ByRef bImportaLcto As Boolean) As Long
    bImportaLcto = True
    ImportCtb_Insere_Lcto = SUCESSO
End Function

Public Function ImportCtb_Trata_Conta_Ccl(ByRef sConta As String, ByRef sCcl As String) As Long
    ImportCtb_Trata_Conta_Ccl = SUCESSO
End Function


