Attribute VB_Name = "Module1"
'Atualização e Desatualização de Lote

Option Explicit

'????Transferir para ErrosContab
Public Const ERRO_LEITURA_LANCAMENTOS9 = 0 'Parametros: lNumIntDocOrigem, iOrigemLcto
'Ocorreu um erro na leitura da tabela de Lançamentos e Visão TransacaoCTBCodigo. NumIntDocOrigem = %l, Origem do Lançamento = %i (TransacaoCTBOrigem).
Public Const ERRO_LEITURA_LANCAMENTOS10 = 0 'Parametros: lNumIntDocOrigem, iOrigemLcto
'Ocorreu um erro na leitura da tabela de Lançamentos Pendentes e Visão TransacaoCTBCodigo. NumIntDocOrigem = %l, Origem do Lançamento = %i (TransacaoCTBOrigem).
Public Const ERRO_LANCAMENTO_NAO_CADASTRADO = 0 'Parâmetros: iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, lDoc, iSeq
'Não foi encontrado o lançamento contábil no banco de dados. Filial = %i, Origem = %s, Exercício = %i, Período = %i, Documento = %l, Sequencial = %i.
Public Const ERRO_LEITURA_LANPENDENTE6 = 0 'Parametros iFilial, sOrigem, iExercicio, iPeriodo, lDoc, iSeq
'Ocorreu um erro na leitura da tabela de Lançamentos Pendentes. Filial = %i, Origem = %s, Exericicio = %i, Periodo = %i, Documento = %l, Sequencial = %i.
Public Const ERRO_LANPENDENTE_NAO_CADASTRADO = 0 'Parâmetros: iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, lDoc, iSeq
'Não foi encontrado o lançamento contábil pendente no banco de dados. Filial = %i, Origem = %s, Exercício = %i, Período = %i, Documento = %l, Sequencial = %i.
Public Const ERRO_EXCLUSAO_LANPENDENTE = 0 'Parametros: iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, lDoc, iSeq
'Ocorreu um erro na exclusão de um lançamento da tabela de lançamentos pendentes. FilialEmpresa = %i, Origem =%s, Exercício = %i, Periodo = %i, Documento = %l, Sequencial = %i.
Public Const ERRO_ALTERACAO_LANCAMENTO = 0 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc, iSeq
'Ocorreu um erro na alteração de um lançamento contábil. Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l, Seq = %i.
Public Const ERRO_LANCAMENTO_NAO_CREDDEB = 0 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc, iSeq
'O lançamento não é a credito nem a debito. Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l, Seq = %i.


Const CLIENTE_ORIGEM_DOCUMENTO = 1
Const FORNECEDOR_ORIGEM_DOCUMENTO = 2

Function Rotina_Atualizacao_Int(iID_Atualizacao As Integer) As Long

Dim lComando As Long
Dim lComando2 As Long
Dim tLote As typeLote_batch
Dim lErro As Long
Dim iUsoCcl As Integer
Dim objLote As ClassLote
Dim colLote As New Collection
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo

On Error GoTo Erro_Rotina_Atualizacao_Int

    lComando = 0
    lComando2 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5036

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 9394

    lErro = Comando_Executar(lComando2, "SELECT UsoCcl FROM Configuracao", iUsoCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9395
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 9396
    
    tLote.sOrigem = String(STRING_ORIGEM, 0)

    'Pesquisa os lotes no banco de dados
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Origem, Exercicio, Periodo, Lote FROM LotePendente WHERE IdAtualizacao = ? ORDER BY FilialEmpresa, Origem, Exercicio, Periodo, Lote", tLote.iFilialEmpresa, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote, iID_Atualizacao)
    If lErro <> AD_SQL_SUCESSO Then Error 5037
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9299
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objLote = New ClassLote
        
        objLote.iFilialEmpresa = tLote.iFilialEmpresa
        objLote.sOrigem = tLote.sOrigem
        objLote.iExercicio = tLote.iExercicio
        objLote.iPeriodo = tLote.iPeriodo
        objLote.iLote = tLote.iLote
        
        colLote.Add objLote
        
        'le o proximo lote
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9300
        
    Loop
    
    TelaAcompanhaBatch.dValorTotal = colLote.Count
    
    For Each objLote In colLote
    
        tLote.iFilialEmpresa = objLote.iFilialEmpresa
        tLote.sOrigem = objLote.sOrigem
        tLote.iExercicio = objLote.iExercicio
        tLote.iPeriodo = objLote.iPeriodo
        tLote.iLote = objLote.iLote
        
        'Processa a Atualização do Lote
        lErro = Atualiza_Lote(tLote, iUsoCcl, iID_Atualizacao)
        If lErro <> SUCESSO Then Error 20339
        
        TelaAcompanhaBatch.dValorAtual = TelaAcompanhaBatch.dValorAtual + 1
        
        TelaAcompanhaBatch.ProgressBar1.Value = CInt((TelaAcompanhaBatch.dValorAtual / TelaAcompanhaBatch.dValorTotal) * 100)

    Next
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Rotina_Atualizacao_Int = SUCESSO
    
    'Alteracao Daniel em 07/05/02
    If colLote.Count = 1 Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_LOTE_ATUALIZADO")
    ElseIf colLote.Count > 1 Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_LOTES_ATUALIZADOS")
    End If
    
    Exit Function

Erro_Rotina_Atualizacao_Int:

    Rotina_Atualizacao_Int = Err

    Select Case Err

        Case 5036, 9394
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5037, 9299, 9300
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE1", Err)

        Case 9395, 9396
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
            
        Case 20339
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_BATCH", Err, objLote.iFilialEmpresa, objLote.sOrigem, objLote.iExercicio, objLote.iPeriodo, objLote.iLote)
            Resume Next
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154862)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Atualiza_Lote(tLote As typeLote_batch, ByVal iUsoCcl As Integer, ByVal iID_Atualizacao As Integer) As Long

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim iPeriodo As Integer
Dim iExercicio As Integer
Dim lTransacao As Long
Dim iStatus As Integer
Dim iApurado As Integer
Dim iID_Atualizacao1 As Integer
Dim tLote1 As typeLote

On Error GoTo Erro_Atualiza_Lote

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9429

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5010
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5247
    
    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 5245

    tLote1.sIdOriginal = String(STRING_IDORIGINAL, 0)

    'Pesquisa o lote no banco de dados
    lErro = Comando_ExecutarPos(lComando, "SELECT TotCre, TotDeb, TotInf, Status, IdOriginal, NumLancInf, NumLancAtual, IdAtualizacao, NumDocInf, NumDocAtual FROM LotePendente WHERE FilialEmpresa=? AND Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote = ?", 0, tLote1.dTotCre, tLote1.dTotDeb, tLote1.dTotInf, tLote1.iStatus, tLote1.sIdOriginal, tLote1.iNumLancInf, tLote1.iNumLancAtual, tLote1.iIdAtualizacao, tLote1.iNumDocInf, tLote1.iNumDocAtual, tLote.iFilialEmpresa, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 9431
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9299
    
    tLote1.iFilialEmpresa = tLote.iFilialEmpresa
    tLote1.sOrigem = tLote.sOrigem
    tLote1.iExercicio = tLote.iExercicio
    tLote1.iPeriodo = tLote.iPeriodo
    tLote1.iLote = tLote.iLote
    
    'Verifica se o lote está desatualizado
    If tLote1.iStatus <> LOTE_DESATUALIZADO Then Error 9432
    
    'Verifica se o lote deve ser atualizado por este processo
    If iID_Atualizacao <> tLote1.iIdAtualizacao Then Error 9433
    
    'Lock do Lote
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9430

    lErro = Atualiza_Lote1(tLote, iUsoCcl, iID_Atualizacao)
    If lErro <> SUCESSO Then Error 9525

    'marca o lote como atualizado
    lErro = Comando_ExecutarPos(lComando1, "DELETE From LotePendente", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5248

    lErro = Comando_Executar(lComando2, "INSERT INTO Lote (FilialEmpresa, Origem, Exercicio, Periodo, Lote, TotCre, TotDeb, TotInf, Status, IdOriginal, NumLancInf, NumLancAtual, IdAtualizacao, NumDocInf, NumDocAtual, Usuario, DataRegistro, HoraRegistro) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", tLote1.iFilialEmpresa, tLote1.sOrigem, tLote1.iExercicio, tLote1.iPeriodo, tLote1.iLote, tLote1.dTotCre, tLote1.dTotDeb, tLote1.dTotInf, LOTE_ATUALIZADO, tLote1.sIdOriginal, tLote1.iNumLancInf, tLote1.iNumLancAtual, tLote1.iIdAtualizacao, tLote1.iNumDocInf, tLote1.iNumDocAtual, gsUsuario, Date, CDbl(Time))
    If lErro <> AD_SQL_SUCESSO Then Error 9438

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 5249
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Atualiza_Lote = SUCESSO

    Exit Function

Erro_Atualiza_Lote:

    Atualiza_Lote = Err
    lErro = Err

    Select Case lErro

        Case 5010, 5247, 9429, 20338
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", lErro)

        Case 5245
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", lErro)

        Case 5248
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LOTEPENDENTE", lErro, tLote.iFilialEmpresa, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote)

        Case 5249
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", lErro)
            
        Case 9299, 9431
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE", lErro, tLote.iFilialEmpresa, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote)
            
        Case 9430
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LOTE", lErro, tLote1.iFilialEmpresa, tLote1.sOrigem, tLote1.iExercicio, tLote1.iPeriodo, tLote1.iLote)
            
        Case 9432
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_DESATUALIZADO", lErro, tLote1.iFilialEmpresa, tLote1.sOrigem, tLote1.iExercicio, tLote1.iPeriodo, tLote1.iLote)
            
        Case 9433
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_SENDO_ATUALIZADO", lErro, tLote1.iFilialEmpresa, tLote1.sOrigem, tLote1.iExercicio, tLote1.iPeriodo, tLote1.iLote)
            
        Case 9438
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOTE", lErro, tLote1.iFilialEmpresa, tLote1.sOrigem, tLote1.iExercicio, tLote1.iPeriodo, tLote1.iLote)
            
        Case 9525

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 154863)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Atualiza_Lote1(tLote As typeLote_batch, ByVal iUsoCcl As Integer, ByVal iID_Atualizacao As Integer) As Long

Dim alComando(1 To 6) As Long
Dim alComando1(1 To 3) As Long
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Atualiza_Lote1

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 9526
    Next
    
    For iIndice = 1 To 3
        alComando1(iIndice) = alComando(iIndice)
    Next
    
    lErro = Atualiza_ExercicioFilial_Nao_Apurado(alComando1(), tLote.iExercicio, tLote.iFilialEmpresa)
    If lErro <> SUCESSO Then Error 10663

    For iIndice = 1 To 3
        alComando1(iIndice) = alComando(iIndice + 3)
    Next
    
    lErro = Atualiza_Periodo_Não_Apurado(alComando1(), tLote.iExercicio, tLote.iPeriodo, tLote.iFilialEmpresa)
    If lErro <> SUCESSO Then Error 10652

    'Processa os lançamentos do lote
    lErro = Processa_Lancamentos(tLote, ATUALIZACAO, iUsoCcl, 0)
    If lErro <> SUCESSO Then Error 5246

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Atualiza_Lote1 = SUCESSO

    Exit Function

Erro_Atualiza_Lote1:

    Atualiza_Lote1 = Err

    Select Case Err

        Case 5246, 10652, 10663

        Case 9526
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", lErro)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 154864)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Atualiza_ExercicioFilial_Nao_Apurado(alComando() As Long, iExercicio As Integer, iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim iStatus As Integer

On Error GoTo Erro_Atualiza_ExercicioFilial_Nao_Apurado

    'Pesquisa o exercicio referente ao lote em questão no banco de dados
    lErro = Comando_ExecutarLockado(alComando(1), "SELECT Status FROM Exercicios WHERE Exercicio = ?", iStatus, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5067
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 5068

    'não permite a mudança no status do exercicio para fechado
    lErro = Comando_LockShared(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 9435

    'Se o exercicio estiver fechado
    If iStatus = EXERCICIO_FECHADO Then Error 9434

    'Pesquisa o ExercicioFilial em questão
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Status FROM ExerciciosFilial WHERE (FilialEmpresa=? Or FilialEmpresa=?) AND Exercicio = ?", 0, iStatus, iFilialEmpresa, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 10656

    'Le o ExercicioFilial
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 10657

    'Lock do ExercicioFilial
    lErro = Comando_LockExclusive(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 10658

    If iStatus = EXERCICIO_APURADO Then
    
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE ExerciciosFilial SET Status = ?", alComando(2), EXERCICIO_ABERTO)
        If lErro <> AD_SQL_SUCESSO Then Error 10659
        
    End If

    'Le o  proximo ExercicioFilial
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 10660

    'Lock do ExercicioFilial
    lErro = Comando_LockExclusive(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 10661

    If iStatus = EXERCICIO_APURADO Then
    
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE ExerciciosFilial SET Status = ?", alComando(2), EXERCICIO_ABERTO)
        If lErro <> AD_SQL_SUCESSO Then Error 10662
        
    End If

    Atualiza_ExercicioFilial_Nao_Apurado = SUCESSO
    
    Exit Function
    
Erro_Atualiza_ExercicioFilial_Nao_Apurado:

    Atualiza_ExercicioFilial_Nao_Apurado = Err
    
    Select Case Err
    
        Case 5067, 5068
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO", lErro, iExercicio)

        Case 9434
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_FECHADO", lErro, iExercicio)

        Case 9435
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", lErro, iExercicio)
            
        Case 10656, 10657, 10660
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL", lErro, iFilialEmpresa, iExercicio)

        Case 10658, 10661
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOSFILIAL", lErro, iFilialEmpresa, iExercicio)
            
        Case 10659, 10662
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOSFILIAL", lErro, iFilialEmpresa, iExercicio)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 154865)
    
    End Select
    
    Exit Function
    
End Function

Function Atualiza_Periodo_Não_Apurado(alComando() As Long, iExercicio As Integer, iPeriodo As Integer, iFilialEmpresa As Integer) As Long

Dim lErro As Long
Dim iApurado As Integer
Dim iPeriodoAux As Integer

On Error GoTo Erro_Atualiza_Periodo_Não_Apurado

    'Pesquisa o periodo em questão
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Periodo FROM Periodo WHERE Exercicio = ? AND Periodo = ?", 0, iPeriodoAux, iExercicio, iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 10613

    'Le o periodo
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 10614

    'Lock do Periodo
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 10615
    
    'Pesquisa o periodoFilial em questão
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Apurado FROM PeriodosFilial WHERE (FilialEmpresa=? Or FilialEmpresa=?) AND Exercicio = ? AND Periodo = ?", 0, iApurado, iFilialEmpresa, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio, iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 5005

    'Le o periodoFilial
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 5006

    'Lock do PeriodoFilial
    lErro = Comando_LockExclusive(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 5007

    If iApurado <> PERIODO_NAO_APURADO Then
    
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE PeriodosFilial SET Apurado = ?", alComando(2), PERIODO_NAO_APURADO)
        If lErro <> AD_SQL_SUCESSO Then Error 9524
        
    End If

    'Le o periodoFilial
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 10653

    'Lock do PeriodoFilial
    lErro = Comando_LockExclusive(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 10654

    If iApurado <> PERIODO_NAO_APURADO Then
    
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE PeriodosFilial SET Apurado = ?", alComando(2), PERIODO_NAO_APURADO)
        If lErro <> AD_SQL_SUCESSO Then Error 10655
        
    End If

    Atualiza_Periodo_Não_Apurado = SUCESSO
    
    Exit Function
    
Erro_Atualiza_Periodo_Não_Apurado:

    Atualiza_Periodo_Não_Apurado = Err

    Select Case Err

        Case 5005, 5006, 10653
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL", lErro, iFilialEmpresa, iExercicio, iPeriodo)

        Case 5007, 10654
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODOSFILIAL", lErro, iFilialEmpresa, iExercicio, iPeriodo)

        Case 9524, 10655
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PERIODOSFILIAL", lErro, iPeriodo, iExercicio, iFilialEmpresa)

        Case 10613, 10614
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOS", lErro, iExercicio, iPeriodo)

        Case 10615
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODO", lErro, iExercicio, iPeriodo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 154866)

    End Select

    Exit Function
    
End Function

Function Processa_Lancamento_Analitico_ContaDia(tProcessa_Lancamento As typeProcessa_Lancamento, dtData As Date, sConta As String) As Long

Dim tLancamento As typeLancamento
Dim lPosicao As Long
Dim dDebito1 As Double
Dim dCredito1 As Double
Dim lErro As Long
Dim tLancamento_Sort As typeLancamento_Sort
Dim vbMesRes As VbMsgBoxResult
Dim iSeq As Integer

On Error GoTo Erro_Processa_Analitico_ContaDia
    
    'inicializa os acumuladores de debito e credito
    dDebito1 = 0
    dCredito1 = 0

    'Se trabalhar com aglutinacao, descobre o doc utilizado para aglutinacao. Se ainda não houver nenhum aloca-o
    lErro = Inicializa_Aglutinacao(tProcessa_Lancamento)
    If lErro <> SUCESSO Then gError 20497
        
    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And dtData = tProcessa_Lancamento.tLancamento.dtData And sConta = tProcessa_Lancamento.tLancamento.sConta

        If tProcessa_Lancamento.tLancamento.dValor > 0 Then
            tProcessa_Lancamento.tLancamento.iCredDeb = CONTA_CREDITO
        Else
            tProcessa_Lancamento.tLancamento.iCredDeb = CONTA_DEBITO
        End If

        If tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar > 0 Then
            tProcessa_Lancamento.tLancamento.iCredDebLivroAuxiliar = CONTA_CREDITO
        Else
            tProcessa_Lancamento.tLancamento.iCredDebLivroAuxiliar = CONTA_DEBITO
        End If


        'se for um lançamento de custo (produto associado) e não se tratar de uma exclusão de lançamento
        If Len(tProcessa_Lancamento.tLancamento.sProduto) <> 0 And tProcessa_Lancamento.iOperacao1 <> ROTINA_EXCLUSAO_LANCAMENTOS Then
        
            'calcular o valor do custo
            lErro = Calcula_Custo(tProcessa_Lancamento)
            If lErro <> SUCESSO Then gError 20514
            
        End If
    
        'se a operação for exclusão de lançamento já contabilizado ==> inverte o valor do lançamento para retira-lo dos registros de saldos consolidados (MvPerCta, MvPerCcl, Aglutinação, ...)
        If tProcessa_Lancamento.iOperacao1 = ROTINA_EXCLUSAO_LANCAMENTOS Or tProcessa_Lancamento.tLancamento.iStatus = VOUCHER_EXCLUSAO Then
            tProcessa_Lancamento.tLancamento.dValor = -tProcessa_Lancamento.tLancamento.dValor
            tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar = -tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar
        End If
    
        tLancamento = tProcessa_Lancamento.tLancamento
    
        'se não for reprocessamento ou
        'se for reprocessamento com o produto preenchido e a apropriação sendo custo medio ou standard
        'ou se for reprocesssamento com o produto preenchido e a apropriação sendo custo de produção
        '==> processa o lançamento
        If tProcessa_Lancamento.iOperacao1 <> ROTINA_REPROC_CUSTO_MEDIO And tProcessa_Lancamento.iOperacao1 <> ROTINA_REPROC_CUSTO_PRODUCAO Or _
        (tProcessa_Lancamento.iOperacao1 = ROTINA_REPROC_CUSTO_MEDIO And Len(Trim(tProcessa_Lancamento.tLancamento.sProduto)) > 0 And (tProcessa_Lancamento.tLancamento.iApropriacao = APROPR_CUSTO_MEDIO Or tProcessa_Lancamento.tLancamento.iApropriacao = APROPR_CUSTO_STANDARD)) Or _
        (tProcessa_Lancamento.iOperacao1 = ROTINA_REPROC_CUSTO_PRODUCAO And Len(Trim(tProcessa_Lancamento.tLancamento.sProduto)) > 0 And (tProcessa_Lancamento.tLancamento.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO Or tProcessa_Lancamento.tLancamento.iApropriacao = APROPR_CUSTO_REAL)) Then
    
            tLancamento_Sort.iFilialEmpresa = tLancamento.iFilialEmpresa
            tLancamento_Sort.dtData = tLancamento.dtData
            tLancamento_Sort.dValor = tLancamento.dValor
            tLancamento_Sort.iExercicio = tLancamento.iExercicio
            tLancamento_Sort.iLote = tLancamento.iLote
            tLancamento_Sort.iPeriodoLan = tLancamento.iPeriodoLan
            tLancamento_Sort.iPeriodoLote = tLancamento.iPeriodoLote
            tLancamento_Sort.iSeq = tLancamento.iSeq
            tLancamento_Sort.lDoc = tLancamento.lDoc
            tLancamento_Sort.sCcl = tLancamento.sCcl + Chr(0)
            tLancamento_Sort.sConta = tLancamento.sConta + Chr(0)
            tLancamento_Sort.sHistorico = tLancamento.sHistorico + Chr(0)
            tLancamento_Sort.sOrigem = tLancamento.sOrigem + Chr(0)
            tLancamento_Sort.iCredDeb = tLancamento.iCredDeb
            tLancamento_Sort.iGerencial = tLancamento.iGerencial
    
            'insere o lançamento no arquivo temporario
            lErro = Arq_Temp_Inserir(tProcessa_Lancamento.lID_Arq_Temp, tLancamento_Sort, lPosicao)
            If lErro <> AD_BOOL_TRUE Then gError 5015
    
            'insere a chave conta no arquivo de sort
            lErro = Sort_Inserir(tProcessa_Lancamento.lID_Arq_Sort, lPosicao, tLancamento.sConta)
            If lErro <> AD_BOOL_TRUE Then gError 5016
    
            'se o lançamento tiver ccl
            If (tProcessa_Lancamento.iUsoCcl = CCL_USA_CONTABIL Or tProcessa_Lancamento.iUsoCcl = CCL_USA_EXTRACONTABIL) And Len(tLancamento.sCcl) > 0 Then
    
                'insere a chave ccl+data+conta no arquivo de sort
                lErro = Sort_Inserir(tProcessa_Lancamento.lID_Arq_Sort1, lPosicao, tLancamento.sCcl, tLancamento.dtData, tLancamento.sConta)
                If lErro <> AD_BOOL_TRUE Then gError 5017
    
                'insere a chave ccl+conta no arquivo de sort
                lErro = Sort_Inserir(tProcessa_Lancamento.lID_Arq_Sort2, lPosicao, tLancamento.sCcl, tLancamento.sConta)
                If lErro <> AD_BOOL_TRUE Then gError 10517
    
            End If
            
            'acumula o valor do lançamento
            'o campo tProcessa_Lancamento.iOperacao indica se a operacao que está sendo executada
            ' é uma atualizacao ou desatualizacao de um lote
            If tLancamento.iCredDeb = CONTA_CREDITO Then
                dCredito1 = dCredito1 + (tLancamento.dValor * tProcessa_Lancamento.iOperacao)
            ElseIf tLancamento.iCredDeb = CONTA_DEBITO Then
                dDebito1 = dDebito1 - (tLancamento.dValor * tProcessa_Lancamento.iOperacao)
            Else
                gError 89214
            End If
    
            If tProcessa_Lancamento.iOperacao1 <> ROTINA_REPROCESSAMENTO_BATCH Then
    
                'verifica se é um lançamento aglutinado e se for acumula o valor do lançamento.
                lErro = Processa_Lancamento_Aglutinado(tProcessa_Lancamento)
                If lErro <> AD_SQL_SUCESSO Then gError 20478
                
                If tProcessa_Lancamento.iOperacao1 <> ROTINA_REPROC_CUSTO_MEDIO And tProcessa_Lancamento.iOperacao1 <> ROTINA_REPROC_CUSTO_PRODUCAO And tProcessa_Lancamento.iOperacao1 <> ROTINA_EXCLUSAO_LANCAMENTOS Then
                
                    lErro = CF("Processa_Lancamento_Analitico_ContaDia_Cust", tLancamento.sConta)
                    If lErro <> SUCESSO Then gError 20478
                
                    'exclui o lancamento do cadastro de lancamentos pendentes
                    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando5, "DELETE FROM LanPendente", tProcessa_Lancamento.lComando2)
                    If lErro <> AD_SQL_SUCESSO Then gError 5344
                    
                    lErro = DataContabil_Valida(tLancamento.dtData, tLancamento.sOrigem)
                    If lErro <> SUCESSO Then gError 185073
                    
                    If gobjCTB.giValidaCtaCcl = MARCADO And tLancamento.sOrigem <> "APE" Then
                        lErro = CF("Lancamento_Valida_ContaCcl", tLancamento.sConta, tLancamento.sCcl)
                        If lErro <> SUCESSO Then Error 185073
                    End If
                    
                    If tProcessa_Lancamento.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA And tLancamento.iAglutina = LANCAMENTO_AGLUTINA Then
        
                        'transfere o lancamento do cadastro de lancamentos pendentes para o cadastro de lancamentos contabilizados
                        lErro = Comando_Executar(tProcessa_Lancamento.lComando6, "INSERT INTO Lancamentos (FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Seq,Lote,PeriodoLote,Data,Conta,Ccl,Historico,Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, DocAglutinado, SeqAglutinado, Aglutinado, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Produto, ApropriaCRProd, Quantidade, DataEstoque, Status, Modelo, Gerencial, SubTipo, Usuario, DataRegistro, HoraRegistro) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                        tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq, tLancamento.iLote, tLancamento.iPeriodoLote, tLancamento.dtData, tLancamento.sConta, tLancamento.sCcl, tLancamento.sHistorico, tLancamento.dValor, tLancamento.lNumIntDoc, tLancamento.iFilialCliForn, tLancamento.lCliForn, tLancamento.iTransacao, tProcessa_Lancamento.lDocAglutinado, tProcessa_Lancamento.iSeqAglutinadoContaCcl, LANCAMENTO_AGLUTINA, tLancamento.lContaSimples, tLancamento.iSeqContraPartida, tLancamento.iEscaninho_Custo, _
                        tLancamento.dValorLivroAuxiliar, tLancamento.iClienteFornecedor, tLancamento.sDocOrigem, tLancamento.sProduto, tLancamento.iApropriaCRProd, tLancamento.dQuantidade, tLancamento.dtDataEstoque, tLancamento.iStatus, tLancamento.sModelo, tLancamento.iGerencial, tLancamento.iSubTipo, IIf(Trim(tLancamento.sUsuario) = "", gsUsuario, Trim(tLancamento.sUsuario)), Date, CDbl(Time))
                        If lErro <> AD_SQL_SUCESSO Then gError 5345
                        
                    Else
                    
                        'transfere o lancamento do cadastro de lancamentos pendentes para o cadastro de lancamentos contabilizados
                        lErro = Comando_Executar(tProcessa_Lancamento.lComando6, "INSERT INTO Lancamentos (FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Seq,Lote,PeriodoLote,Data,Conta,Ccl,Historico,Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, DocAglutinado, SeqAglutinado, Aglutinado, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Produto, ApropriaCRProd, Quantidade, DataEstoque, Status, Modelo, Gerencial, SubTipo, Usuario, DataRegistro, HoraRegistro) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                        tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq, tLancamento.iLote, tLancamento.iPeriodoLote, tLancamento.dtData, tLancamento.sConta, tLancamento.sCcl, tLancamento.sHistorico, tLancamento.dValor, tLancamento.lNumIntDoc, tLancamento.iFilialCliForn, tLancamento.lCliForn, tLancamento.iTransacao, tProcessa_Lancamento.lDocAglutinado, tProcessa_Lancamento.iSeqAglutinado, LANCAMENTO_NAO_AGLUTINADO, tLancamento.lContaSimples, tLancamento.iSeqContraPartida, tLancamento.iEscaninho_Custo, _
                        tLancamento.dValorLivroAuxiliar, tLancamento.iClienteFornecedor, tLancamento.sDocOrigem, tLancamento.sProduto, tLancamento.iApropriaCRProd, tLancamento.dQuantidade, tLancamento.dtDataEstoque, tLancamento.iStatus, tLancamento.sModelo, tLancamento.iGerencial, tLancamento.iSubTipo, IIf(Trim(tLancamento.sUsuario) = "", gsUsuario, Trim(tLancamento.sUsuario)), Date, CDbl(Time))
                        If lErro <> AD_SQL_SUCESSO Then gError 20481
                        
                    End If
                    
                End If
        
                If tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar <> 0 Then
                
                    lErro = Atualiza_CliForn(tProcessa_Lancamento)
                    If lErro <> SUCESSO Then gError 20800
                
                End If
    
                If tProcessa_Lancamento.iOperacao1 = ROTINA_REPROC_CUSTO_MEDIO Then
            
                    'verifica se o exercicio está habilitado a receber lançamentos, ou seja, está presente em colExercicio senão verifica se o exercicio está fechado se estiver ==> erro,
                    'se não estiver coloca-o como aberto e na colecao colExercicio
                    lErro = Verifica_Exercicio_Fechado(tProcessa_Lancamento.colExercicio, tLancamento.iExercicio, tLancamento.iFilialEmpresa)
                    If lErro <> SUCESSO Then gError 83800
                
                    'atualiza o valor no reprocessamento pela diferença entre o valor antigo e o novo
                    lErro = Comando_ExecutarPos(tProcessa_Lancamento.alComando(12), "UPDATE Lancamentos SET Valor = Valor + ?", tProcessa_Lancamento.lComando2, tLancamento.dValor)
                    If lErro <> AD_SQL_SUCESSO Then gError 83801
    
                ElseIf tProcessa_Lancamento.iOperacao1 = ROTINA_REPROC_CUSTO_PRODUCAO Then
            
                    'verifica se o exercicio está habilitado a receber lançamentos, ou seja, está presente em colExercicio senão verifica se o exercicio está fechado se estiver ==> erro,
                    'se não estiver coloca-o como aberto e na colecao colExercicio
                    lErro = Verifica_Exercicio_Fechado(tProcessa_Lancamento.colExercicio, tLancamento.iExercicio, tLancamento.iFilialEmpresa)
                    If lErro <> SUCESSO Then gError 20815
                
                    'atualiza o valor no reprocessamento pela diferença entre o valor antigo e o novo
                    lErro = Comando_ExecutarPos(tProcessa_Lancamento.alComando(12), "UPDATE Lancamentos SET Valor = Valor + ?", tProcessa_Lancamento.lComando2, tLancamento.dValor)
                    If lErro <> AD_SQL_SUCESSO Then gError 83802
            
                ElseIf tProcessa_Lancamento.iOperacao1 = ROTINA_EXCLUSAO_LANCAMENTOS Then
            
                    'verifica se o exercicio está habilitado a receber lançamentos, ou seja, está presente em colExercicio senão verifica se o exercicio está fechado se estiver ==> erro,
                    'se não estiver coloca-o como aberto e na colecao colExercicio
                    lErro = Verifica_Exercicio_Fechado(tProcessa_Lancamento.colExercicio, tLancamento.iExercicio, tLancamento.iFilialEmpresa)
                    If lErro <> SUCESSO Then gError 20818
                
                    'exclui o lançamento do bd
                    lErro = Lancamento_Excluir(tProcessa_Lancamento.lComando2, tProcessa_Lancamento.alComando(14), tLancamento)
                    If lErro <> SUCESSO Then gError 20817
                
                End If
            
            End If
            
        End If
        
        'le o proximo lançamento
        lErro = Comando_BuscarProximo(tProcessa_Lancamento.lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9393

        If lErro = AD_SQL_SUCESSO Then
            tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE
        Else
            tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_FALSE
        End If
                
        'insere/altera o valor do lançamento aglutinado
        lErro = Grava_Lancamento_Aglutinado(tProcessa_Lancamento, tLancamento)
        If lErro <> SUCESSO Then gError 20479

        If tProcessa_Lancamento.iOperacao1 <> ROTINA_ATUALIZACAO_ONLINE Then
                
            TelaAcompanhaBatch.TotReg.Caption = CStr(StrParaLong(TelaAcompanhaBatch.TotReg.Caption) + 1)
                
            DoEvents
        
            If TelaAcompanhaBatch.iCancelaBatch = CANCELA_BATCH Then
            
                vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_ATUALIZACAO_LOTES")
                
                If vbMesRes = vbYes Then gError 28816
                
                TelaAcompanhaBatch.iCancelaBatch = 0
                
            End If
            
        End If
        
    Loop

    tProcessa_Lancamento.dDebito = dDebito1
    tProcessa_Lancamento.dCredito = dCredito1
    
    Processa_Lancamento_Analitico_ContaDia = SUCESSO

    Exit Function

Erro_Processa_Analitico_ContaDia:
    
    Processa_Lancamento_Analitico_ContaDia = gErr

    Select Case gErr
    
        Case 5015
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ARQUIVO_TEMPORARIO", gErr)

        Case 5016, 5017, 5018, 10517
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SORT", gErr)

        Case 5339, 5372
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACTA1", gErr, tProcessa_Lancamento.iFilialEmpresa, sConta, CStr(dtData))

        Case 5340
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACTA", gErr, tProcessa_Lancamento.iFilialEmpresa, sConta, CStr(dtData))

        Case 5344
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LANCAMENTO", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)

        Case 5345, 20481
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTO", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
        Case 5374
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVDIACTA", gErr, tProcessa_Lancamento.iFilialEmpresa, sConta, CStr(dtData))

        Case 9393
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS", gErr, tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sOrigem, tProcessa_Lancamento.tLancamento.iExercicio, tProcessa_Lancamento.tLancamento.iPeriodoLote, tProcessa_Lancamento.tLancamento.iLote)
            
        Case 20478, 20479, 20497, 20514, 20800, 20815, 20817, 20818, 28816, 83800, 185073
                                                                                    
        Case 83798, 83799
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS6", gErr, tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sOrigem, tProcessa_Lancamento.tLancamento.iExercicio, tProcessa_Lancamento.tLancamento.iPeriodoLan, tProcessa_Lancamento.tLancamento.lDoc, tProcessa_Lancamento.tLancamento.iSeq)
                                                                                    
        Case 83801, 83802
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANCAMENTO", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
                                                                                    
        Case 89214
            Call Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_NAO_CREDDEB", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154867)

    End Select

    Exit Function

End Function

Private Function Calcula_Custo(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'se o produto estiver preenchido ==> indica que tem que processar o custo médio de produção ou custo real de produção

Dim lErro As Long
Dim dCustoRProducao As Double
Dim dCustoMRProducao As Double
Dim iMes As Integer
Dim iAno As Integer
Dim objLancamento_Detalhe As New ClassLancamento_Detalhe
Dim dValor As Double
Dim dCustoUnitario As Double

On Error GoTo Erro_Calcula_Custo

    'Lê o atributo Apropriação do produto, cujo codigo foi passado como parâmetro
    lErro = CF("Produto_Le_Apropriacao", tProcessa_Lancamento.alComando(11), tProcessa_Lancamento.tLancamento.sProduto, tProcessa_Lancamento.tLancamento.iApropriacao)
    If lErro <> SUCESSO Then gError 83526

    If tProcessa_Lancamento.iOperacao1 <> ROTINA_REPROC_CUSTO_MEDIO Then
    
        iMes = Month(tProcessa_Lancamento.tLancamento.dtDataEstoque)
        iAno = Year(tProcessa_Lancamento.tLancamento.dtDataEstoque)
        
        'se for um produto apropriado pelo custo de produção
        If tProcessa_Lancamento.tLancamento.iApropriacao = APROPR_CUSTO_REAL Or tProcessa_Lancamento.tLancamento.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO Then
        
            lErro = CF("Retorna_CustoUnitario", tProcessa_Lancamento.tLancamento.iTransacao, tProcessa_Lancamento.tLancamento.lNumIntDoc, tProcessa_Lancamento.tLancamento.sProduto, dCustoUnitario)
            If lErro <> SUCESSO Then gError 92992
        
            If tProcessa_Lancamento.iOperacao1 <> ROTINA_REPROC_CUSTO_PRODUCAO Then

                tProcessa_Lancamento.tLancamento.dValor = tProcessa_Lancamento.tLancamento.dValor * dCustoUnitario

            Else

                'se estiver reprocessando os lançamentos ligados ao custo de produção
                'computa o valor como sendo a diferença entre o valor apurado e o atual
                tProcessa_Lancamento.tLancamento.dValor = (tProcessa_Lancamento.tLancamento.dQuantidade * dCustoUnitario) - tProcessa_Lancamento.tLancamento.dValor

            End If
        
        
'            'se o lançamento se refere a um movimento que usa o custo real de produção
'            If tProcessa_Lancamento.tLancamento.iApropriaCRProd = LANPENDENTE_APROPR_CRPROD Then
'
'                'seleciona o custo real de produção de SldMesEst relativo ao Ano, FilialEmpresa, Produto, Mes passados como parametro
'                lErro = Comando_Executar(tProcessa_Lancamento.lComando9, "SELECT CustoProducao" + CStr(iMes) + " FROM SldMesEst WHERE Ano=? AND FilialEmpresa=? AND Produto=?", dCustoRProducao, iAno, tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sProduto)
'                If lErro <> AD_SQL_SUCESSO Then gError 20515
'
'                'le o SldMesEst
'                lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando9)
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 20516
'
'                If lErro = AD_SQL_SEM_DADOS Then gError 20517
'
'                'se o custo real de producao estiver zerado ==> que o custo  não foi digitado
'                If dCustoRProducao = 0 Then gError 20518
'
'                If tProcessa_Lancamento.iOperacao1 <> ROTINA_REPROC_CUSTO_PRODUCAO Then
'
'                    tProcessa_Lancamento.tLancamento.dValor = tProcessa_Lancamento.tLancamento.dValor * dCustoRProducao
'
'                Else
'
'                    'se estiver reprocessando os lançamentos ligados ao custo de produção
'                    'computa o valor como sendo a diferença entre o valor apurado e o atual
'                    tProcessa_Lancamento.tLancamento.dValor = (tProcessa_Lancamento.tLancamento.dQuantidade * dCustoRProducao) - tProcessa_Lancamento.tLancamento.dValor
'
'                End If
'
'            Else
'
'                'calcula o custo medio de produção ou o custo médio do escaninho do produto em questão
'                lErro = CMProdApurado_Escaninho_Le_Mes(tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sProduto, iMes, iAno, dCustoMRProducao, tProcessa_Lancamento.tLancamento.iEscaninho_Custo, tProcessa_Lancamento.alComando())
'                If lErro <> SUCESSO Then gError 20778
'
'                'se o o custo medio real de producao ou o custo do escaninho estiver zerado ==> o programa que calcula o custo não foi executado
'                If dCustoMRProducao = 0 Then gError 20531
'
'                If tProcessa_Lancamento.iOperacao1 <> ROTINA_REPROC_CUSTO_PRODUCAO Then
'
'                    'se o lançamento se refere a um movimento que usa o custo médio real de produção ou o custo dos escaninhos
'                    tProcessa_Lancamento.tLancamento.dValor = tProcessa_Lancamento.tLancamento.dValor * dCustoMRProducao
'
'                Else
'
'                    dValor = (tProcessa_Lancamento.tLancamento.dQuantidade * dCustoMRProducao)
'
''                    If tProcessa_Lancamento.tLancamento.dValor < 0 Then dValor = -dValor
'
'                    'se estiver reprocessando os lançamentos ligados ao custo de produção
'                    'computa o valor como sendo a diferença entre o valor apurado e o atual
'                    tProcessa_Lancamento.tLancamento.dValor = dValor - tProcessa_Lancamento.tLancamento.dValor
'
'                End If
'
'            End If
        
        End If
    
    ElseIf tProcessa_Lancamento.iOperacao1 = ROTINA_REPROC_CUSTO_MEDIO And (tProcessa_Lancamento.tLancamento.iApropriacao = APROPR_CUSTO_MEDIO Or tProcessa_Lancamento.tLancamento.iApropriacao = APROPR_CUSTO_STANDARD) Then


        lErro = CF("Retorna_CustoUnitario", tProcessa_Lancamento.tLancamento.iTransacao, tProcessa_Lancamento.tLancamento.lNumIntDoc, tProcessa_Lancamento.tLancamento.sProduto, dCustoUnitario)
        If lErro <> SUCESSO Then gError 92993

        'se estiver reprocessando os lançamentos ligados ao custo de produção
        'computa o valor como sendo a diferença entre o valor apurado e o atual
        tProcessa_Lancamento.tLancamento.dValor = (tProcessa_Lancamento.tLancamento.dQuantidade * dCustoUnitario) - tProcessa_Lancamento.tLancamento.dValor

'        'se estiver reprocessando os lançamentos com produto associados ao custo medio ou standard
'        objLancamento_Detalhe.iFilialEmpresa = tProcessa_Lancamento.tLancamento.iFilialEmpresa
'        objLancamento_Detalhe.dtDataEstoque = tProcessa_Lancamento.tLancamento.dtDataEstoque
'        objLancamento_Detalhe.sProduto = tProcessa_Lancamento.tLancamento.sProduto
'        objLancamento_Detalhe.iEscaninho_Custo = tProcessa_Lancamento.tLancamento.iEscaninho_Custo
'        objLancamento_Detalhe.dValor = tProcessa_Lancamento.tLancamento.dQuantidade
'
'        'le o custo medio, standard ou custo dos escaninhos
'        lErro = CF("CustoMedio_Le", tProcessa_Lancamento.alComando1(), objLancamento_Detalhe, tProcessa_Lancamento.tLancamento.iApropriacao)
'        If lErro <> SUCESSO Then gError 83525
'
'        dValor = objLancamento_Detalhe.dValor
'
''        If tProcessa_Lancamento.tLancamento.dValor < 0 Then dValor = -dValor
'
'        'computa o valor como sendo a diferença entre o valor apurado e o valor atual
'        tProcessa_Lancamento.tLancamento.dValor = dValor - tProcessa_Lancamento.tLancamento.dValor
    
    End If
    
    Calcula_Custo = SUCESSO
    
    Exit Function
    
Erro_Calcula_Custo:

    Calcula_Custo = gErr
    
    Select Case gErr
    
        Case 20515, 20516
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, Year(tProcessa_Lancamento.tLancamento.dtData), tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sProduto)
    
        Case 20517
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_INEXISTENTE", gErr, Year(tProcessa_Lancamento.tLancamento.dtData), tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sProduto)
    
        Case 20518
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_CUSTORPRODUCAO_ZERADO", gErr, Year(tProcessa_Lancamento.tLancamento.dtData), tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sProduto, Month(tProcessa_Lancamento.tLancamento.dtData))
    
        Case 20531
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_CUSTOMRPRODUCAO_ZERADO", gErr, Year(tProcessa_Lancamento.tLancamento.dtData), tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sProduto, Month(tProcessa_Lancamento.tLancamento.dtData))

        Case 20778, 83525, 83526

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154868)

    End Select
    
    Exit Function

End Function

Private Function Grava_Lancamento_Aglutinado(tProcessa_Lancamento As typeProcessa_Lancamento, tLancamento As typeLancamento) As Long
'insere/altera um lançamento aglutinador, se o modulo trabalha com lançamentos aglutinados e tem algo para aglutinar

Dim lErro As Long
Dim dValor As Double
Dim objLancamento As New ClassLancamentos

On Error GoTo Erro_Grava_Lancamento_Aglutinado

    'se o modulo trabalha com lançamentos aglutinados e tem algo para aglutinar
    If tProcessa_Lancamento.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA And tProcessa_Lancamento.iTemLancamAglutinado = TEM_LANCAMENTO_AGLUTINADO Then

        'se o arquivo de lançamentos acabou ou mudou a data ou mudou a conta ou mudou o ccl (se usa ccl extra-contabil) ==> grava o lancamento aglutinado
        If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_FALSE Or tProcessa_Lancamento.tLancamento.sConta <> tLancamento.sConta Or tProcessa_Lancamento.tLancamento.dtData <> tLancamento.dtData Or (giSetupUsoCcl = CCL_USA_EXTRACONTABIL And tProcessa_Lancamento.tLancamento.sCcl <> tLancamento.sCcl) Then

            'se há indicação que o lançamento de aglutinação já existe ==> atualiza-o
            If tProcessa_Lancamento.iInsereLancamAglutinado <> INSERE_LANCAMENTO_AGLUTINADO Then
            
                'Se o valor a ser atualizado é diferente de zero
                If tProcessa_Lancamento.dValorAglutinado <> 0 Then
    
                    lErro = Comando_LockExclusive(tProcessa_Lancamento.lComando8)
                    If lErro <> AD_SQL_SUCESSO Then Error 20503

                    'atualiza o valor do lancamento aglutinado
                    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando6, "UPDATE Lancamentos SET Valor = Valor + ?", tProcessa_Lancamento.lComando8, tProcessa_Lancamento.dValorAglutinado)
                    If lErro <> AD_SQL_SUCESSO Then Error 20504
    
                    lErro = Comando_Unlock(tProcessa_Lancamento.lComando8)
                    If lErro <> AD_SQL_SUCESSO Then Error 20505
                    
                End If
                
            Else
            
                'Alteração Daniel em 03/09/2002
                'Preenche o Histórico do Lançamento Aglutinado
                lErro = CF("Preenche_Historico_Lanc_Aglutinado", objLancamento)
                If lErro <> SUCESSO Then Error 32281
                'Fim da Alteração Daniel em 03/09/2002
                
                lErro = DataContabil_Valida(tLancamento.dtData, tLancamento.sOrigem)
                If lErro <> SUCESSO Then Error 32281
                
                If gobjCTB.giValidaCtaCcl = MARCADO And tLancamento.sOrigem <> "APE" Then
                    lErro = CF("Lancamento_Valida_ContaCcl", tLancamento.sConta, tLancamento.sCcl)
                    If lErro <> SUCESSO Then Error 32281
                End If
                
                'Se o lançamento aglutinado não existe ==> insere o lançamento aglutinado
                lErro = Comando_Executar(tProcessa_Lancamento.lComando6, "INSERT INTO Lancamentos (FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Seq,Lote,PeriodoLote,Data,Conta,Ccl,Valor,DocAglutinado, SeqAglutinado, Aglutinado, Historico, Gerencial) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", tLancamento.iFilialEmpresa, gcolModulo_sOrigemAglutina(tLancamento.sOrigem), tLancamento.iExercicio, tLancamento.iPeriodoLan, tProcessa_Lancamento.lDocAglutinado, tProcessa_Lancamento.iSeqAglutinadoContaCcl, 0, tLancamento.iPeriodoLote, tLancamento.dtData, tLancamento.sConta, tLancamento.sCcl, tProcessa_Lancamento.dValorAglutinado, tProcessa_Lancamento.lDocAglutinado, tProcessa_Lancamento.iSeqAglutinadoContaCcl, LANCAMENTO_AGLUTINADO, objLancamento.sHistorico, tLancamento.iGerencial)
                If lErro <> AD_SQL_SUCESSO Then Error 20487
            
            End If
            
            tProcessa_Lancamento.sCclAglutinado = "***"
            tProcessa_Lancamento.dValorAglutinado = 0
            tProcessa_Lancamento.iTemLancamAglutinado = 0
            tProcessa_Lancamento.iInsereLancamAglutinado = 0
            tProcessa_Lancamento.iSeqAglutinadoContaCcl = 0
            
        End If
            
    End If
    
    Grava_Lancamento_Aglutinado = SUCESSO
    
    Exit Function
    
Erro_Grava_Lancamento_Aglutinado:

    Grava_Lancamento_Aglutinado = Err
    
    Select Case Err
    
        Case 20487
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTO", Err, tLancamento.iFilialEmpresa, gcolModulo_sOrigemAglutina(tLancamento.sOrigem), tLancamento.iExercicio, tLancamento.iPeriodoLan, tProcessa_Lancamento.lDocAglutinado, tProcessa_Lancamento.iSeqAglutinado)
        
        Case 20503
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LANCAMENTOS", Err, tLancamento.iFilialEmpresa, gcolModulo_sOrigemAglutina(tLancamento.sOrigem), tLancamento.iExercicio, tLancamento.iPeriodoLan, tProcessa_Lancamento.lDocAglutinado, tProcessa_Lancamento.iSeqAglutinado)
    
        Case 20504
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANCAMENTO_AGLUTINADO", Err, tLancamento.iFilialEmpresa, gcolModulo_sOrigemAglutina(tLancamento.sOrigem), tLancamento.iExercicio, tLancamento.iPeriodoLan, tProcessa_Lancamento.lDocAglutinado, tProcessa_Lancamento.iSeqAglutinado)
    
        Case 20505
            Call Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_LANCAMENTOS", Err, tLancamento.iFilialEmpresa, gcolModulo_sOrigemAglutina(tLancamento.sOrigem), tLancamento.iExercicio, tLancamento.iPeriodoLan, tProcessa_Lancamento.lDocAglutinado, tProcessa_Lancamento.iSeqAglutinado)
    
        Case 32281
        '???
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154869)

    End Select

    Exit Function

End Function

Private Function Processa_Lancamento_Aglutinado(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'verifica se é um lançamento aglutinado e se for acumula o valor do lançamento.
'Se o lançamento aglutinado ainda não for conhecido, descobre seu doc e sequencial.
'Se o lançamento aglutinado não estiver cadastrado, insere-o.

Dim sConta As String
Dim sCcl As String
Dim lErro As Long

On Error GoTo Erro_Processa_Lancamento_Aglutinado

    If tProcessa_Lancamento.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA Then
    
        If tProcessa_Lancamento.tLancamento.iAglutina = LANCAMENTO_AGLUTINA Then
        
            'se usa centro de custo extra contabil e a data, conta ou ccl são diferentes do ultimo encontrado (sCclAglutinado é limpo nestas ocasiões)
'            If giSetupUsoCcl = CCL_USA_EXTRACONTABIL And tProcessa_Lancamento.tLancamento.sCcl <> tProcessa_Lancamento.sCclAglutinado Then
            If tProcessa_Lancamento.tLancamento.sCcl <> tProcessa_Lancamento.sCclAglutinado Then
            
                'Pesquisa o lançamento aglutinado da conta/ccl/data em questao
                lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando8, "SELECT Doc, Seq FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Data = ? AND Conta = ? AND Ccl = ?", 0, tProcessa_Lancamento.lDocAglutinado, tProcessa_Lancamento.iSeqAglutinadoContaCcl, tProcessa_Lancamento.tLancamento.iFilialEmpresa, gcolModulo_sOrigemAglutina(tProcessa_Lancamento.tLancamento.sOrigem), tProcessa_Lancamento.tLancamento.dtData, tProcessa_Lancamento.tLancamento.sConta, tProcessa_Lancamento.tLancamento.sCcl)
                If lErro <> AD_SQL_SUCESSO Then Error 20484
                
                'Le o lançamento aglutinado
                lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando8)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20485

                'se ainda não houver um lançamento aglutinado para a conta/data em questao ==> alocar um novo sequencial para o novo lançamento aglutinado
                If lErro <> AD_SQL_SUCESSO Then
            
                    tProcessa_Lancamento.iUltSeqAglutinado = tProcessa_Lancamento.iUltSeqAglutinado + 1
                    tProcessa_Lancamento.iSeqAglutinadoContaCcl = tProcessa_Lancamento.iUltSeqAglutinado
                    tProcessa_Lancamento.iInsereLancamAglutinado = INSERE_LANCAMENTO_AGLUTINADO
                
                End If
                
                tProcessa_Lancamento.sCclAglutinado = tProcessa_Lancamento.tLancamento.sCcl
                
            End If
             
            tProcessa_Lancamento.dValorAglutinado = tProcessa_Lancamento.dValorAglutinado + (tProcessa_Lancamento.tLancamento.dValor * tProcessa_Lancamento.iOperacao)
            tProcessa_Lancamento.iTemLancamAglutinado = TEM_LANCAMENTO_AGLUTINADO
            
        Else
        
            'se o modulo aglutina por dia mas o lançamento não é aglutinado ==> devolve o proximo numero sequencial do documento aglutinador
            If tProcessa_Lancamento.iUltSeqAglutinado < 32767 Then
                tProcessa_Lancamento.iUltSeqAglutinado = tProcessa_Lancamento.iUltSeqAglutinado + 1
                tProcessa_Lancamento.iSeqAglutinado = tProcessa_Lancamento.iUltSeqAglutinado
            Else
                tProcessa_Lancamento.iSeqAglutinado = 0
            End If
        
'            'se o modulo aglutina por dia mas o lançamento não é aglutinado ==> devolve o proximo numero sequencial do documento aglutinador
'            tProcessa_Lancamento.iUltSeqAglutinado = tProcessa_Lancamento.iUltSeqAglutinado + 1
'            tProcessa_Lancamento.iSeqAglutinado = tProcessa_Lancamento.iUltSeqAglutinado
            
        End If
    
    End If
    
    Processa_Lancamento_Aglutinado = SUCESSO

    Exit Function

Erro_Processa_Lancamento_Aglutinado:
    
    Processa_Lancamento_Aglutinado = Err

    Select Case Err
    
        Case 20484, 20485
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS5", Err, tProcessa_Lancamento.tLancamento.iFilialEmpresa, gcolModulo_sOrigemAglutina(tProcessa_Lancamento.tLancamento.sOrigem), CStr(tProcessa_Lancamento.tLancamento.dtData))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154870)

    End Select

    Exit Function

End Function

Function Inicializa_Aglutinacao(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'descobre o doc de aglutinacao para a data em questao.
'Se ainda não houver um doc de aglutinacao, escolher um doc.
'retorna o proximo sequencial disponivel para este doc.

Dim lErro As Long

On Error GoTo Erro_Inicializa_Aglutinacao

    If tProcessa_Lancamento.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA Then

        'se a data do lançamento é diferente da data de aglutinação (já que vai existir um doc por data de aglutinacao para cada origem)
        'procura o novo documento de aglutinação
        If tProcessa_Lancamento.tLancamento.dtData <> tProcessa_Lancamento.dtDataAglutinado Then
        
            tProcessa_Lancamento.dtDataAglutinado = tProcessa_Lancamento.tLancamento.dtData
        
            'Pesquisa o lançamento aglutinado da data em questao com o maior sequencial
            lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando8, "SELECT DocAglutinado, SeqAglutinado FROM Lancamentos WHERE FilialEmpresa = ? AND Data = ? AND (Origem = ? Or Origem = ?) AND DocAglutinado <> 0 ORDER BY SeqAglutinado DESC", 0, tProcessa_Lancamento.lDocAglutinado, tProcessa_Lancamento.iUltSeqAglutinado, tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.dtData, gcolModulo_sOrigemAglutina(tProcessa_Lancamento.tLancamento.sOrigem), tProcessa_Lancamento.tLancamento.sOrigem)
            If lErro <> AD_SQL_SUCESSO Then Error 20494
                
            'Le o lançamento aglutinado com sequencial maior
            lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando8)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20495
    
            'se ainda não houver um lançamento aglutinado para a data em questao ==> alocar um novo doc
            If lErro <> AD_SQL_SUCESSO Then
            
                'mostra número do proximo voucher(documento) disponível
                lErro = CF("Voucher_Automatico1", tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.iExercicio, tProcessa_Lancamento.tLancamento.iPeriodoLan, gcolModulo_sOrigemAglutina(tProcessa_Lancamento.tLancamento.sOrigem), tProcessa_Lancamento.lDocAglutinado)
                If lErro <> SUCESSO Then Error 20496
        
                tProcessa_Lancamento.iUltSeqAglutinado = 0
                    
            End If
                
        End If
        
'        'se não estiver trabalhando com ccl extra-contabil,
'        'pesquisa se já existe um lançamento aglutinador para a conta/data em questão e se tiver devolve o doc e o sequencial
'        If giSetupUsoCcl <> CCL_USA_EXTRACONTABIL Then
'
'            'Pesquisa o lançamento aglutinador da conta/data em questao (utiliza a origem para identifica-lo)
'            lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando8, "SELECT Doc, Seq FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Data = ? AND Conta = ?", 0, tProcessa_Lancamento.lDocAglutinado, tProcessa_Lancamento.iSeqAglutinadoContaCcl, tProcessa_Lancamento.tLancamento.iFilialEmpresa, gcolModulo_sOrigemAglutina(tProcessa_Lancamento.tLancamento.sOrigem), tProcessa_Lancamento.tLancamento.dtData, tProcessa_Lancamento.tLancamento.sConta)
'            If lErro <> AD_SQL_SUCESSO Then Error 20499
'
'            'Le o lançamento aglutinador
'            lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando8)
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20500
'
'            'se ainda não tiver o lançamento aglutinador para a conta/data em questão ==> aloca um novo sequencial
'            If lErro <> AD_SQL_SUCESSO Then
'
'                tProcessa_Lancamento.iUltSeqAglutinado = tProcessa_Lancamento.iUltSeqAglutinado + 1
'                tProcessa_Lancamento.iSeqAglutinadoContaCcl = tProcessa_Lancamento.iUltSeqAglutinado
'                tProcessa_Lancamento.iInsereLancamAglutinado = INSERE_LANCAMENTO_AGLUTINADO
'
'            End If
'
'        End If
        
    End If
    
    Inicializa_Aglutinacao = SUCESSO

    Exit Function

Erro_Inicializa_Aglutinacao:
    
    Inicializa_Aglutinacao = Err

    Select Case Err
    
        Case 20494, 20495
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS5", Err, tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sOrigem, CStr(tProcessa_Lancamento.tLancamento.dtData))

        Case 20496
                                                                                    
        Case 20499, 20500
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS5", Err, tProcessa_Lancamento.tLancamento.iFilialEmpresa, gcolModulo_sOrigemAglutina(tProcessa_Lancamento.tLancamento.sOrigem), CStr(tProcessa_Lancamento.tLancamento.dtData))
                                                                                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154871)

    End Select

    Exit Function

End Function

Function Verifica_Aglutina_Lancam_Por_Dia(sSiglaModulo As String, iAglutinaLancamPorDia As Integer) As Long

Dim objCTBGlobal As New ClassCTBGlobal

On Error GoTo Erro_Verifica_Aglutina_Lancam_Por_Dia

    If sSiglaModulo = MODULO_CONTASAPAGAR Or sSiglaModulo = MODULO_BATCHCONTASAPAGAR Or sSiglaModulo = MODULO_CUSTOCP Then
    
        iAglutinaLancamPorDia = objCTBGlobal.gobjCTB.iCPAglutinaLancamPorDia
        
    ElseIf sSiglaModulo = MODULO_CONTASARECEBER Or sSiglaModulo = MODULO_BATCHCONTASARECEBER Or sSiglaModulo = MODULO_CUSTOCR Then
    
        iAglutinaLancamPorDia = objCTBGlobal.gobjCTB.iCRAglutinaLancamPorDia
        
    ElseIf sSiglaModulo = MODULO_TESOURARIA Or sSiglaModulo = MODULO_CUSTOTES Then

        iAglutinaLancamPorDia = objCTBGlobal.gobjCTB.iTESAglutinaLancamPorDia

    ElseIf sSiglaModulo = MODULO_FATURAMENTO Or sSiglaModulo = MODULO_CUSTOFAT Then

        iAglutinaLancamPorDia = objCTBGlobal.gobjCTB.iFATAglutinaLancamPorDia

    ElseIf sSiglaModulo = MODULO_ESTOQUE Or sSiglaModulo = MODULO_CUSTOEST Then

        iAglutinaLancamPorDia = objCTBGlobal.gobjCTB.iESTAglutinaLancamPorDia

    End If

    Verifica_Aglutina_Lancam_Por_Dia = SUCESSO

    Exit Function

Erro_Verifica_Aglutina_Lancam_Por_Dia:

    Verifica_Aglutina_Lancam_Por_Dia = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154872)

    End Select

    Exit Function

End Function

Function Processa_Lancamento_Analitico_CclDia(tProcessa_Lancamento As typeProcessa_Lancamento, sCcl As String, dtData As Date) As Long

Dim tLancamento As typeLancamento
Dim tLancamento_Sort As typeLancamento_Sort
Dim lPosicao As Long
Dim dDebito1 As Double
Dim dCredito1 As Double
Dim dDebitoAcum As Double
Dim dCreditoAcum As Double
Dim lErro As Long
Dim sConta As String
Dim vbMesRes As VbMsgBoxResult

On Error GoTo Erro_Processa_Lancamento_Analitico_CclDia

    tLancamento = tProcessa_Lancamento.tLancamento

    dDebitoAcum = 0
    dCreditoAcum = 0

    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sCcl = tLancamento.sCcl And dtData = tLancamento.dtData

        sConta = tLancamento.sConta

        'inicializa acumuladores de debito e credito
        dDebito1 = 0
        dCredito1 = 0

        Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sCcl = tLancamento.sCcl And dtData = tLancamento.dtData And sConta = tLancamento.sConta

            'acumula debito e credito
            If tLancamento.iCredDeb = CONTA_CREDITO Then
                dCredito1 = dCredito1 + (tLancamento.dValor * tProcessa_Lancamento.iOperacao)
            ElseIf tLancamento.iCredDeb = CONTA_DEBITO Then
                dDebito1 = dDebito1 - (tLancamento.dValor * tProcessa_Lancamento.iOperacao)
            Else
                gError 89215
            End If
    
            'le a chave do proximo lancamento ordenado por ccl+conta
            tProcessa_Lancamento.iFim_de_Arquivo = Sort_Ler(tProcessa_Lancamento.lID_Arq_Sort1, lPosicao)
    
            If tProcessa_Lancamento.iOperacao1 <> ROTINA_ATUALIZACAO_ONLINE Then
    
                TelaAcompanhaBatch.TotReg.Caption = CStr(CLng(TelaAcompanhaBatch.TotReg.Caption) + 1)
        
                DoEvents
                
                If TelaAcompanhaBatch.iCancelaBatch = CANCELA_BATCH Then
            
                    vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_ATUALIZACAO_LOTES")
                
                    If vbMesRes = vbYes Then gError 28817
                    
                    TelaAcompanhaBatch.iCancelaBatch = 0
                        
                End If
    
            End If
            
            If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then
    
                'le o proximo lancamento
                lErro = Arq_Temp_Ler(tProcessa_Lancamento.lID_Arq_Temp, tLancamento_Sort, lPosicao)
                If lErro <> SUCESSO Then gError 10641
    
                tProcessa_Lancamento.tLancamento.iFilialEmpresa = tLancamento_Sort.iFilialEmpresa
                tProcessa_Lancamento.tLancamento.dtData = tLancamento_Sort.dtData
                tProcessa_Lancamento.tLancamento.dValor = tLancamento_Sort.dValor
                tProcessa_Lancamento.tLancamento.iExercicio = tLancamento_Sort.iExercicio
                tProcessa_Lancamento.tLancamento.iLote = tLancamento_Sort.iLote
                tProcessa_Lancamento.tLancamento.iPeriodoLan = tLancamento_Sort.iPeriodoLan
                tProcessa_Lancamento.tLancamento.iPeriodoLote = tLancamento_Sort.iPeriodoLote
                tProcessa_Lancamento.tLancamento.iSeq = tLancamento_Sort.iSeq
                tProcessa_Lancamento.tLancamento.lDoc = tLancamento_Sort.lDoc
                tProcessa_Lancamento.tLancamento.sCcl = StringZ(tLancamento_Sort.sCcl)
                tProcessa_Lancamento.tLancamento.sConta = StringZ(tLancamento_Sort.sConta)
                tProcessa_Lancamento.tLancamento.sHistorico = StringZ(tLancamento_Sort.sHistorico)
                tProcessa_Lancamento.tLancamento.sOrigem = StringZ(tLancamento_Sort.sOrigem)
                tProcessa_Lancamento.tLancamento.iCredDeb = tLancamento_Sort.iCredDeb
                tProcessa_Lancamento.tLancamento.iGerencial = tLancamento_Sort.iGerencial
                
                tLancamento = tProcessa_Lancamento.tLancamento
    
            End If

        Loop
        
        
        tProcessa_Lancamento.dDebito = dDebito1
        tProcessa_Lancamento.dCredito = dCredito1
            
        dDebitoAcum = dDebitoAcum + dDebito1
        dCreditoAcum = dCreditoAcum + dCredito1

        lErro = Atualiza_CclDia(tProcessa_Lancamento, sCcl, sConta, dtData)
        If lErro <> SUCESSO Then gError 10642

        sConta = tLancamento.sConta

    Loop

    tProcessa_Lancamento.dDebito = dDebitoAcum
    tProcessa_Lancamento.dCredito = dCreditoAcum
    
    Processa_Lancamento_Analitico_CclDia = SUCESSO
    
    Exit Function

Erro_Processa_Lancamento_Analitico_CclDia:
    
    Processa_Lancamento_Analitico_CclDia = gErr

    Select Case gErr

        Case 10641
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQUIVO_TEMP", gErr)

        Case 10642
        
        Case 28817

        Case 89215
            Call Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_NAO_CREDDEB", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154873)

    End Select

    Exit Function

End Function

Function Processa_Lancamento_Analitico_ContaMes(tProcessa_Lancamento As typeProcessa_Lancamento, sConta As String) As Long

Dim tLancamento As typeLancamento
Dim tLancamento_Sort As typeLancamento_Sort
Dim lPosicao As Long
Dim dDebito1 As Double
Dim dCredito1 As Double
Dim lErro As Long
Dim vbMesRes As VbMsgBoxResult

On Error GoTo Erro_Processa_Analitico_ContaMes

    tLancamento = tProcessa_Lancamento.tLancamento

    'inicializa acumuladores de debito e credito
    dDebito1 = 0
    dCredito1 = 0

    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sConta = tLancamento.sConta

        'acumula debito e credito
        If tLancamento.iCredDeb = CONTA_CREDITO Then
            dCredito1 = dCredito1 + (tLancamento.dValor * tProcessa_Lancamento.iOperacao)
        ElseIf tLancamento.iCredDeb = CONTA_DEBITO Then
            dDebito1 = dDebito1 - (tLancamento.dValor * tProcessa_Lancamento.iOperacao)
        Else
            gError 89216
        End If

        'le a chave do proximo lancamento ordenado por conta
        tProcessa_Lancamento.iFim_de_Arquivo = Sort_Ler(tProcessa_Lancamento.lID_Arq_Sort, lPosicao)

        If tProcessa_Lancamento.iOperacao1 <> ROTINA_ATUALIZACAO_ONLINE Then

            TelaAcompanhaBatch.TotReg.Caption = CStr(CLng(TelaAcompanhaBatch.TotReg.Caption) + 1)
    
            DoEvents
            
            If TelaAcompanhaBatch.iCancelaBatch = CANCELA_BATCH Then
            
                vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_ATUALIZACAO_LOTES")
                
                If vbMesRes = vbYes Then gError 28818
                
                TelaAcompanhaBatch.iCancelaBatch = 0
               
            End If

        End If
        
        If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then

            'le o proximo lancamento
            lErro = Arq_Temp_Ler(tProcessa_Lancamento.lID_Arq_Temp, tLancamento_Sort, lPosicao)
            If lErro <> SUCESSO Then gError 5030

            tProcessa_Lancamento.tLancamento.iFilialEmpresa = tLancamento_Sort.iFilialEmpresa
            tProcessa_Lancamento.tLancamento.dtData = tLancamento_Sort.dtData
            tProcessa_Lancamento.tLancamento.dValor = tLancamento_Sort.dValor
            tProcessa_Lancamento.tLancamento.iExercicio = tLancamento_Sort.iExercicio
            tProcessa_Lancamento.tLancamento.iLote = tLancamento_Sort.iLote
            tProcessa_Lancamento.tLancamento.iPeriodoLan = tLancamento_Sort.iPeriodoLan
            tProcessa_Lancamento.tLancamento.iPeriodoLote = tLancamento_Sort.iPeriodoLote
            tProcessa_Lancamento.tLancamento.iSeq = tLancamento_Sort.iSeq
            tProcessa_Lancamento.tLancamento.lDoc = tLancamento_Sort.lDoc
            tProcessa_Lancamento.tLancamento.sCcl = StringZ(tLancamento_Sort.sCcl)
            tProcessa_Lancamento.tLancamento.sConta = StringZ(tLancamento_Sort.sConta)
            tProcessa_Lancamento.tLancamento.sHistorico = StringZ(tLancamento_Sort.sHistorico)
            tProcessa_Lancamento.tLancamento.sOrigem = StringZ(tLancamento_Sort.sOrigem)
            tProcessa_Lancamento.tLancamento.iCredDeb = tLancamento_Sort.iCredDeb
            tProcessa_Lancamento.tLancamento.iGerencial = tLancamento_Sort.iGerencial
            
            tLancamento = tProcessa_Lancamento.tLancamento

        End If

    Loop

    tProcessa_Lancamento.dDebito = dDebito1
    tProcessa_Lancamento.dCredito = dCredito1
    
    Processa_Lancamento_Analitico_ContaMes = SUCESSO
    
    Exit Function

Erro_Processa_Analitico_ContaMes:
    
    Processa_Lancamento_Analitico_ContaMes = gErr

    Select Case gErr

        Case 5030
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQUIVO_TEMP", gErr)
            
        Case 28818

        Case 89216
            Call Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_NAO_CREDDEB", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154874)

    End Select

    Exit Function

End Function

Function Processa_Lancamento_Analitico_CclMes(tProcessa_Lancamento As typeProcessa_Lancamento, sCcl As String) As Long

Dim tLancamento As typeLancamento
Dim tLancamento_Sort As typeLancamento_Sort
Dim lPosicao As Long
Dim dDebito1 As Double
Dim dCredito1 As Double
Dim dDebitoAcum As Double
Dim dCreditoAcum As Double
Dim lErro As Long
Dim sConta As String
Dim vbMesRes As VbMsgBoxResult

On Error GoTo Erro_Processa_Analitico_CclMes

    tLancamento = tProcessa_Lancamento.tLancamento

    dDebitoAcum = 0
    dCreditoAcum = 0

    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sCcl = tLancamento.sCcl

        sConta = tLancamento.sConta

        'inicializa acumuladores de debito e credito
        dDebito1 = 0
        dCredito1 = 0

        Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sCcl = tLancamento.sCcl And sConta = tLancamento.sConta

            'acumula debito e credito
            If tLancamento.iCredDeb = CONTA_CREDITO Then
                dCredito1 = dCredito1 + (tLancamento.dValor * tProcessa_Lancamento.iOperacao)
            ElseIf tLancamento.iCredDeb = CONTA_DEBITO Then
                dDebito1 = dDebito1 - (tLancamento.dValor * tProcessa_Lancamento.iOperacao)
            Else
                gError 89217
            End If
    
            'le a chave do proximo lancamento ordenado por ccl+conta
            tProcessa_Lancamento.iFim_de_Arquivo = Sort_Ler(tProcessa_Lancamento.lID_Arq_Sort2, lPosicao)
    
            If tProcessa_Lancamento.iOperacao1 <> ROTINA_ATUALIZACAO_ONLINE Then
    
                TelaAcompanhaBatch.TotReg.Caption = CStr(CLng(TelaAcompanhaBatch.TotReg.Caption) + 1)
    
                DoEvents
                
                If TelaAcompanhaBatch.iCancelaBatch = CANCELA_BATCH Then
            
                    vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_ATUALIZACAO_LOTES")
                
                    If vbMesRes = vbYes Then gError 28819
                    
                    TelaAcompanhaBatch.iCancelaBatch = 0
                    
                End If

            End If
            
            If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then
    
                'le o proximo lancamento
                lErro = Arq_Temp_Ler(tProcessa_Lancamento.lID_Arq_Temp, tLancamento_Sort, lPosicao)
                If lErro <> SUCESSO Then gError 10518
    
                tProcessa_Lancamento.tLancamento.iFilialEmpresa = tLancamento_Sort.iFilialEmpresa
                tProcessa_Lancamento.tLancamento.dtData = tLancamento_Sort.dtData
                tProcessa_Lancamento.tLancamento.dValor = tLancamento_Sort.dValor
                tProcessa_Lancamento.tLancamento.iExercicio = tLancamento_Sort.iExercicio
                tProcessa_Lancamento.tLancamento.iLote = tLancamento_Sort.iLote
                tProcessa_Lancamento.tLancamento.iPeriodoLan = tLancamento_Sort.iPeriodoLan
                tProcessa_Lancamento.tLancamento.iPeriodoLote = tLancamento_Sort.iPeriodoLote
                tProcessa_Lancamento.tLancamento.iSeq = tLancamento_Sort.iSeq
                tProcessa_Lancamento.tLancamento.lDoc = tLancamento_Sort.lDoc
                tProcessa_Lancamento.tLancamento.sCcl = StringZ(tLancamento_Sort.sCcl)
                tProcessa_Lancamento.tLancamento.sConta = StringZ(tLancamento_Sort.sConta)
                tProcessa_Lancamento.tLancamento.sHistorico = StringZ(tLancamento_Sort.sHistorico)
                tProcessa_Lancamento.tLancamento.sOrigem = StringZ(tLancamento_Sort.sOrigem)
                tProcessa_Lancamento.tLancamento.iCredDeb = tLancamento_Sort.iCredDeb
                tProcessa_Lancamento.tLancamento.iGerencial = tLancamento_Sort.iGerencial
                
                tLancamento = tProcessa_Lancamento.tLancamento
    
            End If

        Loop
        
        
        tProcessa_Lancamento.dDebito = dDebito1
        tProcessa_Lancamento.dCredito = dCredito1
            
        dDebitoAcum = dDebitoAcum + dDebito1
        dCreditoAcum = dCreditoAcum + dCredito1

        lErro = Atualiza_CclMes(tProcessa_Lancamento, sCcl, sConta)
        If lErro <> SUCESSO Then gError 10629

        sConta = tLancamento.sConta

    Loop

    tProcessa_Lancamento.dDebito = dDebitoAcum
    tProcessa_Lancamento.dCredito = dCreditoAcum
    
    Processa_Lancamento_Analitico_CclMes = SUCESSO
    
    Exit Function

Erro_Processa_Analitico_CclMes:
    
    Processa_Lancamento_Analitico_CclMes = gErr

    Select Case gErr

        Case 10518
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQUIVO_TEMP", gErr)

        Case 10629
        
        Case 28819

        Case 89217
            Call Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_NAO_CREDDEB", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154875)

    End Select

    Exit Function

End Function


'Function Processa_Lancamento_CclMes(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'
'Dim lErro As Long
'Dim lPosicao As Long
'Dim lComando3 As Long, lComando4 As Long
'Dim lComando5 As Long, lComando6 As Long, lComando7 As Long
'Dim sConta As String
'Dim sCcl As String
'Dim sPeriodo As String
'Dim tLancamento_Sort As typeLancamento_Sort
'Dim iExercicio As Integer
'
'On Error GoTo Erro_Processa_Lancamento_CclMes
'
'    lComando3 = 0
'    lComando4 = 0
'    lComando5 = 0
'    lComando6 = 0
'    lComando7 = 0
'
'    'classifica os lançamentos por ccl, conta, data
'    lErro = Sort_Classificar(tProcessa_Lancamento.lID_Arq_Sort1)
'    If lErro <> AD_BOOL_TRUE Then Error 5051
'
'    'le a chave do primeiro lancamento
'    tProcessa_Lancamento.iFim_de_Arquivo = Sort_Ler(tProcessa_Lancamento.lID_Arq_Sort1, lPosicao)
'
'    If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then
'
'        lComando3 = Comando_Abrir()
'        If lComando3 = 0 Then Error 5049
'
'        tProcessa_Lancamento.lComando3 = lComando3
'
'        lComando4 = Comando_Abrir()
'        If lComando4 = 0 Then Error 5351
'
'        tProcessa_Lancamento.lComando4 = lComando4
'
'        lComando5 = Comando_Abrir()
'        If lComando5 = 0 Then Error 5349
'
'        lComando6 = Comando_Abrir()
'        If lComando6 = 0 Then Error 5050
'
'        lComando7 = Comando_Abrir()
'        If lComando7 = 0 Then Error 5377
'
'        tProcessa_Lancamento.lComando7 = lComando7
'
'        'le o primeiro lancamento
'        lErro = Arq_Temp_Ler(tProcessa_Lancamento.lID_Arq_Temp, tLancamento_Sort, lPosicao)
'        If lErro <> SUCESSO Then Error 5052
'
'        tProcessa_Lancamento.tLancamento.iFilialEmpresa = tLancamento_Sort.iFilialEmpresa
'        tProcessa_Lancamento.tLancamento.dtData = tLancamento_Sort.dtData
'        tProcessa_Lancamento.tLancamento.dValor = tLancamento_Sort.dValor
'        tProcessa_Lancamento.tLancamento.iExercicio = tLancamento_Sort.iExercicio
'        tProcessa_Lancamento.tLancamento.iLote = tLancamento_Sort.iLote
'        tProcessa_Lancamento.tLancamento.iPeriodoLan = tLancamento_Sort.iPeriodoLan
'        tProcessa_Lancamento.tLancamento.iPeriodoLote = tLancamento_Sort.iPeriodoLote
'        tProcessa_Lancamento.tLancamento.iSeq = tLancamento_Sort.iSeq
'        tProcessa_Lancamento.tLancamento.lDoc = tLancamento_Sort.lDoc
'        tProcessa_Lancamento.tLancamento.sCcl = StringZ(tLancamento_Sort.sCcl)
'        tProcessa_Lancamento.tLancamento.sConta = StringZ(tLancamento_Sort.sConta)
'        tProcessa_Lancamento.tLancamento.sHistorico = StringZ(tLancamento_Sort.sHistorico)
'        tProcessa_Lancamento.tLancamento.sOrigem = StringZ(tLancamento_Sort.sOrigem)
'
'        Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE
'
'            sConta = tProcessa_Lancamento.tLancamento.sConta
'            sCcl = tProcessa_Lancamento.tLancamento.sCcl
'
'            'processa os lancamentos
'            lErro = Processa_Lancamento_CclMes1(tProcessa_Lancamento)
'            If lErro <> SUCESSO Then Error 5053
'
'            sPeriodo = tProcessa_Lancamento.sPeriodo
'
'            'seleciona os totais de debito e credito do ccl em questão
'            lErro = Comando_ExecutarPos(lComando5, "SELECT Exercicio FROM MvPerCcl WHERE FilialEmpresa=? AND Exercicio = ? AND Conta = ? AND Ccl = ?", 0, iExercicio, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, sConta, sCcl)
'            If lErro <> AD_SQL_SUCESSO Then Error 5350
'
'            lErro = Comando_BuscarPrimeiro(lComando5)
'            If lErro <> AD_SQL_SUCESSO Then Error 5376
'
'            'atualiza os totais de debito e credito do ccl e periodo em questão
'            lErro = Comando_ExecutarPos(lComando6, "UPDATE MvPerCcl SET Deb" + sPeriodo + "= Deb" + sPeriodo + " + ?, Cre" + sPeriodo + "= Cre" + sPeriodo + " + ?", lComando5, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
'            If lErro <> AD_SQL_SUCESSO Then Error 5054
'
'        Loop
'
'    End If
'
'    Call Comando_Fechar(lComando3)
'    Call Comando_Fechar(lComando4)
'    Call Comando_Fechar(lComando5)
'    Call Comando_Fechar(lComando6)
'    Call Comando_Fechar(lComando7)
'
'    Processa_Lancamento_CclMes = SUCESSO
'
'    Exit Function
'
'Erro_Processa_Lancamento_CclMes:
'
'    Processa_Lancamento_CclMes = Err
'
'    Select Case Err
'
'        Case 5049, 5050, 5349, 5351, 5377
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 5051
'            Call Rotina_Erro(vbOKOnly, "ERRO_CLASSIFICAR_ARQUIVO_SORT", Err)
'
'        Case 5052
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQUIVO_TEMP", Err)
'
'        Case 5053
'
'        Case 5054
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL", Err, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.iExercicio, sCcl, sConta)
'
'        Case 5350, 5376
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL1", Err, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.iExercicio, sCcl, sConta)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154876)
'
'    End Select
'
'    Call Comando_Fechar(lComando3)
'    Call Comando_Fechar(lComando4)
'    Call Comando_Fechar(lComando5)
'    Call Comando_Fechar(lComando6)
'    Call Comando_Fechar(lComando7)
'
'    Exit Function
'
'End Function

'Function Processa_Lancamento_CclMes1(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'
'Dim dDebito1 As Double
'Dim dCredito1 As Double
'Dim sConta As String
'Dim sCcl As String
'Dim lErro As Long
'Dim sPeriodo As String
'Dim dtData As Date
'Dim lPosicao As Long
'Dim tLancamento_Sort As typeLancamento_Sort
'Dim dtData1 As Date
'
'On Error GoTo Erro_Processa_Lancamento_CclMes1
'
'    dtData1 = tProcessa_Lancamento.tLancamento.dtData 'apenas p/inicializar com um valor valido c/data
'
'    'inicializa os acumuladores de debito e credito por mes
'    tProcessa_Lancamento.dDebito = 0
'    tProcessa_Lancamento.dCredito = 0
'
'    'guarda o número da conta e ccl que está sendo processada
'    sConta = tProcessa_Lancamento.tLancamento.sConta
'    sCcl = tProcessa_Lancamento.tLancamento.sCcl
'
'    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sConta = tProcessa_Lancamento.tLancamento.sConta And sCcl = tProcessa_Lancamento.tLancamento.sCcl
'
'        'inicializa os acumuladores de debito e credito por dia
'        dDebito1 = 0
'        dCredito1 = 0
'
'        'guarda a data que está sendo processado
'        dtData = tProcessa_Lancamento.tLancamento.dtData
'
'        Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sConta = tProcessa_Lancamento.tLancamento.sConta And sCcl = tProcessa_Lancamento.tLancamento.sCcl And dtData = tProcessa_Lancamento.tLancamento.dtData
'
'            'acumula debito e credito
'            If tProcessa_Lancamento.tLancamento.dValor > 0 Then
'                dCredito1 = dCredito1 + (tProcessa_Lancamento.tLancamento.dValor * tProcessa_Lancamento.iOperacao)
'            Else
'                dDebito1 = dDebito1 - (tProcessa_Lancamento.tLancamento.dValor * tProcessa_Lancamento.iOperacao)
'            End If
'
'            'le a chave do proximo lancamento ordenado por conta
'            tProcessa_Lancamento.iFim_de_Arquivo = Sort_Ler(tProcessa_Lancamento.lID_Arq_Sort1, lPosicao)
'
'            DoEvents
'
'            If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then
'
'                'le o proximo lancamento
'                lErro = Arq_Temp_Ler(tProcessa_Lancamento.lID_Arq_Temp, tLancamento_Sort, lPosicao)
'                If lErro <> SUCESSO Then Error 5055
'
'                tProcessa_Lancamento.tLancamento.iFilialEmpresa = tLancamento_Sort.iFilialEmpresa
'                tProcessa_Lancamento.tLancamento.dtData = tLancamento_Sort.dtData
'                tProcessa_Lancamento.tLancamento.dValor = tLancamento_Sort.dValor
'                tProcessa_Lancamento.tLancamento.iExercicio = tLancamento_Sort.iExercicio
'                tProcessa_Lancamento.tLancamento.iLote = tLancamento_Sort.iLote
'                tProcessa_Lancamento.tLancamento.iPeriodoLan = tLancamento_Sort.iPeriodoLan
'                tProcessa_Lancamento.tLancamento.iPeriodoLote = tLancamento_Sort.iPeriodoLote
'                tProcessa_Lancamento.tLancamento.iSeq = tLancamento_Sort.iSeq
'                tProcessa_Lancamento.tLancamento.lDoc = tLancamento_Sort.lDoc
'                tProcessa_Lancamento.tLancamento.sCcl = StringZ(tLancamento_Sort.sCcl)
'                tProcessa_Lancamento.tLancamento.sConta = StringZ(tLancamento_Sort.sConta)
'                tProcessa_Lancamento.tLancamento.sHistorico = StringZ(tLancamento_Sort.sHistorico)
'                tProcessa_Lancamento.tLancamento.sOrigem = StringZ(tLancamento_Sort.sOrigem)
'
'
'            End If
'
'        Loop
'
'        'acumula debitos e creditos
'        tProcessa_Lancamento.dDebito = tProcessa_Lancamento.dDebito + dDebito1
'        tProcessa_Lancamento.dCredito = tProcessa_Lancamento.dCredito + dCredito1
'
'        'seleciona o total de credito e debito do ccl, conta, dia em questão
'        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando4, "SELECT Data FROM MvDiaCcl WHERE FilialEmpresa = ? AND Ccl = ? AND Conta = ? AND Data = ?", 0, dtData1, tProcessa_Lancamento.iFilialEmpresa, sCcl, sConta, dtData)
'        If lErro <> AD_SQL_SUCESSO Then Error 5352
'
'        lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando4)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5379
'
'        'se nao tiver o registro contendo o total do ccl,conta ===> incluir
'        If lErro = AD_SQL_SEM_DADOS Then
'
'            'inserir o registro de total de ccl, conta
'            lErro = Comando_Executar(tProcessa_Lancamento.lComando7, "INSERT INTO MvDiaCcl (FilialEmpresa, Ccl, Conta, Data, Deb, Cre) VALUES (?,?,?,?,?,?)", tProcessa_Lancamento.iFilialEmpresa, sCcl, sConta, dtData, dDebito1, dCredito1)
'            If lErro <> AD_SQL_SUCESSO Then Error 5378
'
'        Else
'
'            'atualiza  o total de credito e debito para ccl, conta e dia selecionado
'            lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando3, "UPDATE MvDiaCcl SET Deb = Deb + ?, Cre = Cre + ?", tProcessa_Lancamento.lComando4, dDebito1, dCredito1)
'            If lErro <> AD_SQL_SUCESSO Then Error 5056
'
'        End If
'
'    Loop
'
'    Processa_Lancamento_CclMes1 = SUCESSO
'
'    Exit Function
'
'Erro_Processa_Lancamento_CclMes1:
'
'    Processa_Lancamento_CclMes1 = Err
'
'    Select Case Err
'
'        Case 5055
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQUIVO_TEMP", Err)
'
'        Case 5056
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACCL", Err, tProcessa_Lancamento.iFilialEmpresa, sCcl, sConta, CStr(dtData))
'
'        Case 5352, 5379
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACCL1", Err, tProcessa_Lancamento.iFilialEmpresa, sCcl, sConta, CStr(dtData))
'
'        Case 5378
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVDIACCL", Err, tProcessa_Lancamento.iFilialEmpresa, sCcl, sConta, CStr(dtData))
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154877)
'
'    End Select
'
'    Exit Function
'
'End Function

Function Processa_Lancamento_ContaDia(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'processa a atualização dos saldos de conta diário

Dim alComando(1 To 19) As Long
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Processa_Lancamento_ContaDia

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 5025
        
    Next
    
    For iIndice = LBound(tProcessa_Lancamento.alComando) To UBound(tProcessa_Lancamento.alComando)
    
        tProcessa_Lancamento.alComando(iIndice) = Comando_Abrir()
        If tProcessa_Lancamento.alComando(iIndice) = 0 Then Error 20793
        
    Next
    
    For iIndice = LBound(tProcessa_Lancamento.alComando1) To UBound(tProcessa_Lancamento.alComando1)
    
        tProcessa_Lancamento.alComando1(iIndice) = Comando_Abrir()
        If tProcessa_Lancamento.alComando1(iIndice) = 0 Then Error 20816
        
    Next
    
    tProcessa_Lancamento.lComando3 = alComando(1)
    tProcessa_Lancamento.lComando4 = alComando(2)
    tProcessa_Lancamento.lComando5 = alComando(3)
    tProcessa_Lancamento.lComando6 = alComando(4)
    tProcessa_Lancamento.lComando7 = alComando(5)
    tProcessa_Lancamento.lComando8 = alComando(6)
    tProcessa_Lancamento.lComando9 = alComando(7)
    tProcessa_Lancamento.lComando10 = alComando(8)
    tProcessa_Lancamento.lComando11 = alComando(9)
    tProcessa_Lancamento.lComando12 = alComando(10)
    tProcessa_Lancamento.lComando13 = alComando(11)
    tProcessa_Lancamento.lComando14 = alComando(12)
    tProcessa_Lancamento.lComando15 = alComando(13)
    tProcessa_Lancamento.lComando16 = alComando(14)
    tProcessa_Lancamento.lComando17 = alComando(15)
    tProcessa_Lancamento.lComando18 = alComando(16)
    tProcessa_Lancamento.lComando19 = alComando(17)
    tProcessa_Lancamento.lComando20 = alComando(18)
    tProcessa_Lancamento.lComando21 = alComando(19)
    
    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE

        'Processa a atualização da tabela MvDiaCta
        lErro = Processa_Lancamento_ContaDia1(1, tProcessa_Lancamento)
        If lErro <> SUCESSO Then Error 5032

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    For iIndice = LBound(tProcessa_Lancamento.alComando) To UBound(tProcessa_Lancamento.alComando)
        Call Comando_Fechar(tProcessa_Lancamento.alComando(iIndice))
    Next

    For iIndice = LBound(tProcessa_Lancamento.alComando1) To UBound(tProcessa_Lancamento.alComando1)
        Call Comando_Fechar(tProcessa_Lancamento.alComando1(iIndice))
    Next

    Processa_Lancamento_ContaDia = SUCESSO

    Exit Function

Erro_Processa_Lancamento_ContaDia:

    Processa_Lancamento_ContaDia = Err

    Select Case Err

        Case 5025, 5335, 5342, 5343, 5373, 19413, 20793, 20816
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5032

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154878)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    For iIndice = LBound(tProcessa_Lancamento.alComando) To UBound(tProcessa_Lancamento.alComando)
        Call Comando_Fechar(tProcessa_Lancamento.alComando(iIndice))
    Next
    
    For iIndice = LBound(tProcessa_Lancamento.alComando1) To UBound(tProcessa_Lancamento.alComando1)
        Call Comando_Fechar(tProcessa_Lancamento.alComando1(iIndice))
    Next
    
    Exit Function

End Function

Function Processa_Lancamento_CclDia(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'processa a atualização dos saldos de centro de custo/lucro diário

Dim lComando3 As Long
Dim lComando4 As Long
Dim lComando5 As Long
Dim lComando6 As Long
Dim lComando7 As Long
Dim lErro As Long
Dim lPosicao As Long
Dim tLancamento_Sort As typeLancamento_Sort

On Error GoTo Erro_Processa_Lancamento_CclDia

    lComando3 = 0
    lComando4 = 0
    lComando5 = 0
    lComando6 = 0
    lComando7 = 0

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10630

    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 10631

    lComando5 = Comando_Abrir()
    If lComando5 = 0 Then Error 10632

    lComando6 = Comando_Abrir()
    If lComando6 = 0 Then Error 10633

    lComando7 = Comando_Abrir()
    If lComando7 = 0 Then Error 10634

    tProcessa_Lancamento.lComando3 = lComando3
    tProcessa_Lancamento.lComando4 = lComando4
    tProcessa_Lancamento.lComando5 = lComando5
    tProcessa_Lancamento.lComando6 = lComando6
    tProcessa_Lancamento.lComando7 = lComando7
    
    'classifica os lançamentos por ccl+conta
    lErro = Sort_Classificar(tProcessa_Lancamento.lID_Arq_Sort1)
    If lErro <> AD_BOOL_TRUE Then Error 55833

    'le a chave do primeiro lancamento
    tProcessa_Lancamento.iFim_de_Arquivo = Sort_Ler(tProcessa_Lancamento.lID_Arq_Sort1, lPosicao)

    If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then

        'le o primeiro lancamento
        lErro = Arq_Temp_Ler(tProcessa_Lancamento.lID_Arq_Temp, tLancamento_Sort, lPosicao)
        If lErro <> SUCESSO Then Error 55834

        tProcessa_Lancamento.tLancamento.iFilialEmpresa = tLancamento_Sort.iFilialEmpresa
        tProcessa_Lancamento.tLancamento.dtData = tLancamento_Sort.dtData
        tProcessa_Lancamento.tLancamento.dValor = tLancamento_Sort.dValor
        tProcessa_Lancamento.tLancamento.iExercicio = tLancamento_Sort.iExercicio
        tProcessa_Lancamento.tLancamento.iLote = tLancamento_Sort.iLote
        tProcessa_Lancamento.tLancamento.iPeriodoLan = tLancamento_Sort.iPeriodoLan
        tProcessa_Lancamento.tLancamento.iPeriodoLote = tLancamento_Sort.iPeriodoLote
        tProcessa_Lancamento.tLancamento.iSeq = tLancamento_Sort.iSeq
        tProcessa_Lancamento.tLancamento.lDoc = tLancamento_Sort.lDoc
        tProcessa_Lancamento.tLancamento.sCcl = StringZ(tLancamento_Sort.sCcl)
        tProcessa_Lancamento.tLancamento.sConta = StringZ(tLancamento_Sort.sConta)
        tProcessa_Lancamento.tLancamento.sHistorico = StringZ(tLancamento_Sort.sHistorico)
        tProcessa_Lancamento.tLancamento.sOrigem = StringZ(tLancamento_Sort.sOrigem)
        tProcessa_Lancamento.tLancamento.iCredDeb = tLancamento_Sort.iCredDeb
        tProcessa_Lancamento.tLancamento.iGerencial = tLancamento_Sort.iGerencial

        Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE
    
            'Processa a atualização da tabela MvDiaCta
            lErro = Processa_Lancamento_CclDia1(1, tProcessa_Lancamento)
            If lErro <> SUCESSO Then Error 10635
    
        Loop

    End If

    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    Call Comando_Fechar(lComando5)
    Call Comando_Fechar(lComando6)
    Call Comando_Fechar(lComando7)

    Processa_Lancamento_CclDia = SUCESSO

    Exit Function

Erro_Processa_Lancamento_CclDia:

    Processa_Lancamento_CclDia = Err

    Select Case Err

        Case 10630, 10631, 10632, 10633, 10634
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10635

        Case 55833
            Call Rotina_Erro(vbOKOnly, "ERRO_CLASSIFICAR_ARQUIVO_SORT", Err)

        Case 55834
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQUIVO_TEMP", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154879)

    End Select

    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    Call Comando_Fechar(lComando5)
    Call Comando_Fechar(lComando6)
    Call Comando_Fechar(lComando7)
    
    Exit Function

End Function

Function Processa_Lancamento_ContaDia1(ByVal iNivel As Integer, tProcessa_Lancamento As typeProcessa_Lancamento) As Long

Dim dDebito1 As Double
Dim dCredito1 As Double
Dim sConta As String
Dim iExiste_Proximo_Nivel As Integer
Dim dtData As Date
Dim lErro As Long
Dim sConta1 As String
Dim dtData1 As Date

On Error GoTo Erro_Processa_Lancamento_ContaDia1

    'inicializa acumuladores de debito e credito
    dDebito1 = 0
    dCredito1 = 0

    sConta = String(STRING_CONTA, 0)

    'guarda o número da conta que está sendo processada
    lErro = Mascara_RetornaContaNoNivel(iNivel, tProcessa_Lancamento.tLancamento.sConta, sConta)
    If lErro <> SUCESSO Then Error 9365
    
    sConta1 = sConta

    'guarda a data que está sendo processado
    dtData = tProcessa_Lancamento.tLancamento.dtData

    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sConta1 = sConta And dtData = tProcessa_Lancamento.tLancamento.dtData

        'verifica se a conta possui um segmento mais "profundo"
        iExiste_Proximo_Nivel = Mascara_ExisteProxNivelCta(iNivel, tProcessa_Lancamento.tLancamento.sConta)

        If iExiste_Proximo_Nivel = SUCESSO Then

            'se existe um nível mais "profundo" da conta, processe-o
            lErro = Processa_Lancamento_ContaDia1(iNivel + 1, tProcessa_Lancamento)
            If lErro <> SUCESSO Then Error 5039

            'acumula debito e credito para a conta e data em questão
            dDebito1 = dDebito1 + tProcessa_Lancamento.dDebito
            dCredito1 = dCredito1 + tProcessa_Lancamento.dCredito


            If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then

                'verifica se a conta possui o nivel em questão. Se não possuir sair do loop
                iExiste_Proximo_Nivel = Mascara_ExisteProxNivelCta(iNivel - 1, tProcessa_Lancamento.tLancamento.sConta)
                If iExiste_Proximo_Nivel <> SUCESSO Then Exit Do
                
                sConta = String(STRING_CONTA, 0)
            
                'guarda o número da conta para o nível em questão
                lErro = Mascara_RetornaContaNoNivel(iNivel, tProcessa_Lancamento.tLancamento.sConta, sConta)
                If lErro <> SUCESSO Then Error 9366
                
            End If

        Else
        
            'se este é o nível mais profundo da conta, processe-o
            lErro = Processa_Lancamento_Analitico_ContaDia(tProcessa_Lancamento, dtData, sConta)
            If lErro <> SUCESSO Then Error 5040
            
            dDebito1 = tProcessa_Lancamento.dDebito
            dCredito1 = tProcessa_Lancamento.dCredito
            
            Exit Do

        End If

    Loop

    tProcessa_Lancamento.dDebito = dDebito1
    tProcessa_Lancamento.dCredito = dCredito1

    lErro = Atualiza_ContaDia(tProcessa_Lancamento, sConta1, dtData)
    If lErro <> SUCESSO Then Error 10621

    Processa_Lancamento_ContaDia1 = SUCESSO

    Exit Function

Erro_Processa_Lancamento_ContaDia1:

    Processa_Lancamento_ContaDia1 = Err

    Select Case Err

        Case 5039, 5040, 10621, 20497

        Case 9365, 9366
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, tProcessa_Lancamento.tLancamento.sConta, iNivel)
                                                                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154880)

    End Select

    Exit Function

End Function

Function Processa_Lancamento_CclDia1(ByVal iNivel As Integer, tProcessa_Lancamento As typeProcessa_Lancamento) As Long

Dim dDebito1 As Double
Dim dCredito1 As Double
Dim sCcl As String
Dim iExiste_Proximo_Nivel As Integer
Dim dtData As Date
Dim lErro As Long
Dim sCcl1 As String
Dim dtData1 As Date

On Error GoTo Erro_Processa_Lancamento_CclDia1

    
    'inicializa acumuladores de debito e credito
    dDebito1 = 0
    dCredito1 = 0

    sCcl = String(STRING_CCL, 0)

    'guarda o número do centro de custo/lucro que está sendo processado
    lErro = Mascara_RetornaCclNoNivel(iNivel, tProcessa_Lancamento.tLancamento.sCcl, sCcl)
    If lErro <> SUCESSO Then Error 10636
    
    sCcl1 = sCcl

    'guarda a data que está sendo processado
    dtData = tProcessa_Lancamento.tLancamento.dtData

    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sCcl1 = sCcl And dtData = tProcessa_Lancamento.tLancamento.dtData

        'verifica se o centro de custo/lucro possui um segmento mais "profundo"
        iExiste_Proximo_Nivel = Mascara_ExisteProxNivelCcl(iNivel, tProcessa_Lancamento.tLancamento.sCcl)

        If iExiste_Proximo_Nivel = SUCESSO Then

            'se existe um nível mais "profundo" do centro de custo/lucro, processe-o
            lErro = Processa_Lancamento_CclDia1(iNivel + 1, tProcessa_Lancamento)
            If lErro <> SUCESSO Then Error 10637

            'acumula debito e credito para o centro de custo/lucro e data em questão
            dDebito1 = dDebito1 + tProcessa_Lancamento.dDebito
            dCredito1 = dCredito1 + tProcessa_Lancamento.dCredito


            If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then

                'verifica se o centro de custo/lucro possui o nivel em questão. Se não possuir sair do loop
                iExiste_Proximo_Nivel = Mascara_ExisteProxNivelCcl(iNivel - 1, tProcessa_Lancamento.tLancamento.sCcl)
                If iExiste_Proximo_Nivel <> SUCESSO Then Exit Do
                
                sCcl = String(STRING_CCL, 0)
            
                'guarda o número do centro de custo/lucro para o nível em questão
                lErro = Mascara_RetornaCclNoNivel(iNivel, tProcessa_Lancamento.tLancamento.sCcl, sCcl)
                If lErro <> SUCESSO Then Error 10638
                
            End If

        Else
        
            'se este é o nível mais profundo do centro de custo/lucro, processe-o
            lErro = Processa_Lancamento_Analitico_CclDia(tProcessa_Lancamento, sCcl, dtData)
            If lErro <> SUCESSO Then Error 10639
            
            dDebito1 = tProcessa_Lancamento.dDebito
            dCredito1 = tProcessa_Lancamento.dCredito
            
            Exit Do

        End If

    Loop

    tProcessa_Lancamento.dDebito = dDebito1
    tProcessa_Lancamento.dCredito = dCredito1

    lErro = Atualiza_CclDia(tProcessa_Lancamento, sCcl1, "", dtData)
    If lErro <> SUCESSO Then Error 10640

    Processa_Lancamento_CclDia1 = SUCESSO

    Exit Function

Erro_Processa_Lancamento_CclDia1:

    Processa_Lancamento_CclDia1 = Err

    Select Case Err

        Case 10637, 10639, 10640

        Case 10636, 10638
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclNoNivel", Err, tProcessa_Lancamento.tLancamento.sCcl, iNivel)
                                                                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154881)

    End Select

    Exit Function

End Function

Function Atualiza_ContaDia(tProcessa_Lancamento As typeProcessa_Lancamento, sConta1 As String, dtData As Date) As Long

Dim lErro As Long
Dim dtData1 As Date

On Error GoTo Erro_Atualiza_ContaDia

    'seleciona os totais de debito e credito do dia do lancamento para serem atualizados
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando4, "SELECT Data FROM MvDiaCta WHERE FilialEmpresa = ? AND Conta = ? AND Data = ?", 0, dtData1, tProcessa_Lancamento.iFilialEmpresa, sConta1, dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 5336

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando4)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5381
    
    'se nao encontrou o registro com os totais ==> cadastra o registro
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo de dia
        lErro = Comando_Executar(tProcessa_Lancamento.lComando7, "INSERT INTO MvDiaCta (FilialEmpresa,Conta,Data,Deb,Cre) VALUES (?,?,?,?,?)", tProcessa_Lancamento.iFilialEmpresa, sConta1, dtData, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then Error 5380

    Else

        'atualiza os totais de debito e credito do dia do lancamento
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando3, "UPDATE MvDiaCta SET Deb = Deb + ?, Cre = Cre + ?", tProcessa_Lancamento.lComando4, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then Error 5014

    End If

    'seleciona os totais de debito e credito do dia do lancamento para serem atualizados no ambito empresa
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando4, "SELECT Data FROM MvDiaCta WHERE FilialEmpresa = ? AND Conta = ? AND Data = ?", 0, dtData1, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta1, dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 10617
        
    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando4)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10618
    
    'se nao encontrou o registro com os totais ==> cadastra o registro no ambito empresa
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo de dia no ambito empresa
        lErro = Comando_Executar(tProcessa_Lancamento.lComando7, "INSERT INTO MvDiaCta (FilialEmpresa,Conta,Data,Deb,Cre) VALUES (?,?,?,?,?)", IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta1, dtData, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then Error 10619
    
    Else

        'atualiza os totais de debito e credito do dia do lancamento no ambito empresa
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando3, "UPDATE MvDiaCta SET Deb = Deb + ?, Cre = Cre + ?", tProcessa_Lancamento.lComando4, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then Error 10620

    End If
    
    Atualiza_ContaDia = SUCESSO
    
    Exit Function
    
Erro_Atualiza_ContaDia:

    Atualiza_ContaDia = Err
    
    Select Case Err
    
        Case 5014
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACTA", Err, tProcessa_Lancamento.iFilialEmpresa, sConta1, CStr(dtData))

        Case 5336, 5381
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACTA1", Err, tProcessa_Lancamento.iFilialEmpresa, sConta1, CStr(dtData))

        Case 5380
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVDIACTA", Err, tProcessa_Lancamento.iFilialEmpresa, sConta1, CStr(dtData))
            
        Case 10617, 10618
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACTA1", Err, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta1, CStr(dtData))

        Case 10619
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVDIACTA", Err, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta1, CStr(dtData))

        Case 10620
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACTA", Err, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta1, CStr(dtData))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154882)

    End Select

    Exit Function
    
End Function

Function Atualiza_CclDia(tProcessa_Lancamento As typeProcessa_Lancamento, sCcl1 As String, sConta1 As String, dtData As Date) As Long

Dim lErro As Long
Dim dtData1 As Date


On Error GoTo Erro_Atualiza_CclDia

    'seleciona os totais de debito e credito do dia do lancamento para serem atualizados
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando4, "SELECT Data FROM MvDiaCcl WHERE FilialEmpresa = ? AND Ccl = ? AND Conta = ? AND Data = ?", 0, dtData1, tProcessa_Lancamento.iFilialEmpresa, sCcl1, sConta1, dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 10643

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando4)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10644
    
    'se nao encontrou o registro com os totais ==> cadastra o registro
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo de dia
        lErro = Comando_Executar(tProcessa_Lancamento.lComando7, "INSERT INTO MvDiaCcl (FilialEmpresa,Ccl,Conta,Data,Deb,Cre) VALUES (?,?,?,?,?,?)", tProcessa_Lancamento.iFilialEmpresa, sCcl1, sConta1, dtData, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then Error 10645

    Else

        'atualiza os totais de debito e credito do dia do lancamento
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando3, "UPDATE MvDiaCcl SET Deb = Deb + ?, Cre = Cre + ?", tProcessa_Lancamento.lComando4, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then Error 10646

    End If

    'seleciona os totais de debito e credito do dia do lancamento para serem atualizados no ambito empresa
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando4, "SELECT Data FROM MvDiaCcl WHERE FilialEmpresa = ? AND Ccl = ? AND Conta = ? AND Data = ?", 0, dtData1, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl1, sConta1, dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 10647

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando4)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10648
    
    'se nao encontrou o registro com os totais ==> cadastra o registro no ambito empresa
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo de dia no ambito empresa
        lErro = Comando_Executar(tProcessa_Lancamento.lComando7, "INSERT INTO MvDiaCcl (FilialEmpresa,Ccl,Conta,Data,Deb,Cre) VALUES (?,?,?,?,?,?)", IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl1, sConta1, dtData, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then Error 10649

    Else

        'atualiza os totais de debito e credito do dia do lancamento no ambito empresa
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando3, "UPDATE MvDiaCcl SET Deb = Deb + ?, Cre = Cre + ?", tProcessa_Lancamento.lComando4, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then Error 10650

    End If
    
    Atualiza_CclDia = SUCESSO
    
    Exit Function
    
Erro_Atualiza_CclDia:

    Atualiza_CclDia = Err
    
    Select Case Err
    
        Case 10643, 10644
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACCL1", Err, tProcessa_Lancamento.iFilialEmpresa, sCcl1, sConta1, CStr(dtData))

        Case 10645
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVDIACCL", Err, tProcessa_Lancamento.iFilialEmpresa, sCcl1, sConta1, CStr(dtData))
            
        Case 10646
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACCL", Err, tProcessa_Lancamento.iFilialEmpresa, sCcl1, sConta1, CStr(dtData))

        Case 10647, 10648
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACCL1", Err, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl1, sConta1, CStr(dtData))

        Case 10649
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVDIACCL", Err, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl1, sConta1, CStr(dtData))
            
        Case 10650
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACCL", Err, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl1, sConta1, CStr(dtData))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154883)

    End Select

    Exit Function
    
End Function


Function Processa_Lancamento_ContaMes(tProcessa_Lancamento As typeProcessa_Lancamento) As Long

Dim lErro As Long
Dim lPosicao As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim tLancamento_Sort As typeLancamento_Sort

On Error GoTo Erro_Processa_Lancamento_ContaMes

    lComando3 = 0
    lComando4 = 0

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5031
    
    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 5346

    tProcessa_Lancamento.lComando3 = lComando3
    tProcessa_Lancamento.lComando4 = lComando4

    'classifica os lançamentos por ccl+conta
    lErro = Sort_Classificar(tProcessa_Lancamento.lID_Arq_Sort)
    If lErro <> AD_BOOL_TRUE Then Error 5027

    'le a chave do primeiro lancamento
    tProcessa_Lancamento.iFim_de_Arquivo = Sort_Ler(tProcessa_Lancamento.lID_Arq_Sort, lPosicao)

    If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then

        'le o primeiro lancamento
        lErro = Arq_Temp_Ler(tProcessa_Lancamento.lID_Arq_Temp, tLancamento_Sort, lPosicao)
        If lErro <> SUCESSO Then Error 5028

        tProcessa_Lancamento.tLancamento.iFilialEmpresa = tLancamento_Sort.iFilialEmpresa
        tProcessa_Lancamento.tLancamento.dtData = tLancamento_Sort.dtData
        tProcessa_Lancamento.tLancamento.dValor = tLancamento_Sort.dValor
        tProcessa_Lancamento.tLancamento.iExercicio = tLancamento_Sort.iExercicio
        tProcessa_Lancamento.tLancamento.iLote = tLancamento_Sort.iLote
        tProcessa_Lancamento.tLancamento.iPeriodoLan = tLancamento_Sort.iPeriodoLan
        tProcessa_Lancamento.tLancamento.iPeriodoLote = tLancamento_Sort.iPeriodoLote
        tProcessa_Lancamento.tLancamento.iSeq = tLancamento_Sort.iSeq
        tProcessa_Lancamento.tLancamento.lDoc = tLancamento_Sort.lDoc
        tProcessa_Lancamento.tLancamento.sCcl = StringZ(tLancamento_Sort.sCcl)
        tProcessa_Lancamento.tLancamento.sConta = StringZ(tLancamento_Sort.sConta)
        tProcessa_Lancamento.tLancamento.sHistorico = StringZ(tLancamento_Sort.sHistorico)
        tProcessa_Lancamento.tLancamento.sOrigem = StringZ(tLancamento_Sort.sOrigem)
        tProcessa_Lancamento.tLancamento.iCredDeb = tLancamento_Sort.iCredDeb
        tProcessa_Lancamento.tLancamento.iGerencial = tLancamento_Sort.iGerencial

        Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE

            'processa os lancamentos
            lErro = Processa_Lancamento_ContaMes1(1, tProcessa_Lancamento)
            If lErro <> SUCESSO Then Error 10524

        Loop

    End If

    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)

    Processa_Lancamento_ContaMes = SUCESSO

    Exit Function

Erro_Processa_Lancamento_ContaMes:

    Processa_Lancamento_ContaMes = Err

    Select Case Err

        Case 5027
            Call Rotina_Erro(vbOKOnly, "ERRO_CLASSIFICAR_ARQUIVO_SORT", Err)
        
        Case 5028
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQUIVO_TEMP", Err)
        
        Case 5031
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5346
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 10524

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154884)

    End Select

    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)

    Exit Function

End Function

Function Processa_Lancamento_CclMes(tProcessa_Lancamento As typeProcessa_Lancamento) As Long

Dim lErro As Long
Dim lPosicao As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim tLancamento_Sort As typeLancamento_Sort

On Error GoTo Erro_Processa_Lancamento_CclMes

    lComando3 = 0
    lComando4 = 0

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10519
    
    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 10520

    tProcessa_Lancamento.lComando3 = lComando3
    tProcessa_Lancamento.lComando4 = lComando4

    'classifica os lançamentos por ccl+conta
    lErro = Sort_Classificar(tProcessa_Lancamento.lID_Arq_Sort2)
    If lErro <> AD_BOOL_TRUE Then Error 10521

    'le a chave do primeiro lancamento
    tProcessa_Lancamento.iFim_de_Arquivo = Sort_Ler(tProcessa_Lancamento.lID_Arq_Sort2, lPosicao)

    If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then

        'le o primeiro lancamento
        lErro = Arq_Temp_Ler(tProcessa_Lancamento.lID_Arq_Temp, tLancamento_Sort, lPosicao)
        If lErro <> SUCESSO Then Error 10522

        tProcessa_Lancamento.tLancamento.iFilialEmpresa = tLancamento_Sort.iFilialEmpresa
        tProcessa_Lancamento.tLancamento.dtData = tLancamento_Sort.dtData
        tProcessa_Lancamento.tLancamento.dValor = tLancamento_Sort.dValor
        tProcessa_Lancamento.tLancamento.iExercicio = tLancamento_Sort.iExercicio
        tProcessa_Lancamento.tLancamento.iLote = tLancamento_Sort.iLote
        tProcessa_Lancamento.tLancamento.iPeriodoLan = tLancamento_Sort.iPeriodoLan
        tProcessa_Lancamento.tLancamento.iPeriodoLote = tLancamento_Sort.iPeriodoLote
        tProcessa_Lancamento.tLancamento.iSeq = tLancamento_Sort.iSeq
        tProcessa_Lancamento.tLancamento.lDoc = tLancamento_Sort.lDoc
        tProcessa_Lancamento.tLancamento.sCcl = StringZ(tLancamento_Sort.sCcl)
        tProcessa_Lancamento.tLancamento.sConta = StringZ(tLancamento_Sort.sConta)
        tProcessa_Lancamento.tLancamento.sHistorico = StringZ(tLancamento_Sort.sHistorico)
        tProcessa_Lancamento.tLancamento.sOrigem = StringZ(tLancamento_Sort.sOrigem)
        tProcessa_Lancamento.tLancamento.iCredDeb = tLancamento_Sort.iCredDeb
        tProcessa_Lancamento.tLancamento.iGerencial = tLancamento_Sort.iGerencial

        Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE

            'processa os lancamentos
            lErro = Processa_Lancamento_CclMes1(1, tProcessa_Lancamento)
            If lErro <> SUCESSO Then Error 10523

        Loop

    End If

    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)

    Processa_Lancamento_CclMes = SUCESSO

    Exit Function

Erro_Processa_Lancamento_CclMes:

    Processa_Lancamento_CclMes = Err

    Select Case Err

        Case 10519, 10520
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 10521
            Call Rotina_Erro(vbOKOnly, "ERRO_CLASSIFICAR_ARQUIVO_SORT", Err)
        
        Case 10522
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ARQUIVO_TEMP", Err)
        
        Case 10523
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154885)

    End Select

    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)

    Exit Function

End Function

Function Processa_Lancamento_CclMes1(ByVal iNivel As Integer, tProcessa_Lancamento As typeProcessa_Lancamento) As Long

Dim dDebito1 As Double
Dim dCredito1 As Double
Dim sCcl As String
Dim sCcl1 As String
Dim sCcl2 As String
Dim iExiste_Proximo_Nivel As Integer
Dim lErro As Long
Dim sPeriodo As String
Dim iExercicio As Integer

On Error GoTo Erro_Processa_Lancamento_CclMes1

    'inicializa os acumuladores de debito e credito
    dDebito1 = 0
    dCredito1 = 0

    sCcl = String(STRING_CCL, 0)
    
    'guarda o centro de custo/lucro que está sendo processada
    lErro = Mascara_RetornaCclNoNivel(iNivel, tProcessa_Lancamento.tLancamento.sCcl, sCcl)
    If lErro <> SUCESSO Then Error 10525
    
    sCcl1 = sCcl

    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sCcl = sCcl1

        'verifica se o centro de custo/lucro possui um nivel mais "profundo"
        iExiste_Proximo_Nivel = Mascara_ExisteProxNivelCcl(iNivel, tProcessa_Lancamento.tLancamento.sCcl)

        If iExiste_Proximo_Nivel = SUCESSO Then

            'se existe um nivel mais profundo, processa-o
            lErro = Processa_Lancamento_CclMes1(iNivel + 1, tProcessa_Lancamento)
            If lErro <> SUCESSO Then Error 10526

            'acumula debitos e creditos
            dDebito1 = dDebito1 + tProcessa_Lancamento.dDebito
            dCredito1 = dCredito1 + tProcessa_Lancamento.dCredito
    
            If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then
                    
                'verifica se o centro de custo/lucro possui o nivel em questão. Se não possuir sair do loop
                iExiste_Proximo_Nivel = Mascara_ExisteProxNivelCcl(iNivel - 1, tProcessa_Lancamento.tLancamento.sCcl)
                If iExiste_Proximo_Nivel <> SUCESSO Then Exit Do
    
                sCcl = String(STRING_CCL, 0)
    
                'armazena o centro de custo/lucro do nivel em questão
                lErro = Mascara_RetornaCclNoNivel(iNivel, tProcessa_Lancamento.tLancamento.sCcl, sCcl)
                If lErro <> SUCESSO Then Error 10527
    
            End If

        Else

            'se este é o nivel mais profundo de centro de custo/lucro, processa-o
            lErro = Processa_Lancamento_Analitico_CclMes(tProcessa_Lancamento, sCcl)
            If lErro <> SUCESSO Then Error 10528

            dDebito1 = tProcessa_Lancamento.dDebito
            dCredito1 = tProcessa_Lancamento.dCredito
            
            Exit Do

        End If

    Loop

    tProcessa_Lancamento.dDebito = dDebito1
    tProcessa_Lancamento.dCredito = dCredito1

    lErro = Atualiza_CclMes(tProcessa_Lancamento, sCcl1, "")
    If lErro <> SUCESSO Then Error 10625

    Processa_Lancamento_CclMes1 = SUCESSO

    Exit Function

Erro_Processa_Lancamento_CclMes1:

    Processa_Lancamento_CclMes1 = Err

    Select Case Err

        Case 10525, 10527
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCclNoNivel", Err, tProcessa_Lancamento.tLancamento.sCcl, iNivel)
        
        Case 10526, 10528, 10625
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154886)

    End Select

    Exit Function

End Function

Function Atualiza_CclMes(tProcessa_Lancamento As typeProcessa_Lancamento, sCcl1 As String, sConta1 As String) As Long

Dim lErro As Long
Dim sPeriodo As String
Dim iExercicio As Integer
Dim alComando(1 To 2)  As Long
Dim iIndice As Integer
Dim dSldIni As Double

On Error GoTo Erro_Atualiza_CclMes

    'Abre o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 188303
    Next

    sPeriodo = tProcessa_Lancamento.sPeriodo

    'seleciona o saldo de centros de custo/lucro mensal para fazer sua atualizacao
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando4, "SELECT Exercicio FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Ccl = ? AND Conta = ?", 0, iExercicio, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, sCcl1, sConta1)
    If lErro <> AD_SQL_SUCESSO Then gError 10529

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando4)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 10530

    If lErro = AD_SQL_SEM_DADOS Then

        'Pesquisa o saldo inicial da conta x centro de custo/lucro em questão
        lErro = Comando_Executar(alComando(1), "SELECT SldIni FROM SaldoInicialContaCcl WHERE FilialEmpresa=? AND Conta=? AND Ccl=?", dSldIni, tProcessa_Lancamento.iFilialEmpresa, sConta1, sCcl1)
        If lErro <> AD_SQL_SUCESSO Then gError 188304

        'le o saldo
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188305

        If lErro = AD_SQL_SEM_DADOS Then dSldIni = 0

        'Insere o registro de Saldo de Ccl
        lErro = Comando_Executar(alComando(2), "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Ccl, Conta, SldIni, Deb" + sPeriodo + ", Cre" + sPeriodo + ") VALUES (?, ?, ?, ?, ?, ?, ?)", tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, sCcl1, sConta1, dSldIni, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 188306


    Else

        'atualiza o total de credito e debito da conta e periodo especificados
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando3, "UPDATE MvPerCcl SET Deb" + sPeriodo + "= Deb" + sPeriodo + " + ?, Cre" + sPeriodo + "= Cre" + sPeriodo + " + ?", tProcessa_Lancamento.lComando4, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 10531

    End If

    'seleciona o saldo de centros de custo/lucro mensal para fazer sua atualizacao, ambito empresa
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando4, "SELECT Exercicio FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Ccl = ? AND Conta = ?", 0, iExercicio, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, sCcl1, sConta1)
    If lErro <> AD_SQL_SUCESSO Then gError 10626

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando4)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 10627

    If lErro = AD_SQL_SEM_DADOS Then

        'Pesquisa o saldo inicial da conta x centro de custo/lucro em questão
        lErro = Comando_Executar(alComando(1), "SELECT SldIni FROM SaldoInicialContaCcl WHERE FilialEmpresa=? AND Conta=? AND Ccl=?", dSldIni, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta1, sCcl1)
        If lErro <> AD_SQL_SUCESSO Then gError 188307

        'le o saldo
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 188308

        If lErro = AD_SQL_SEM_DADOS Then dSldIni = 0

        'Insere o registro de Saldo de Ccl
        lErro = Comando_Executar(alComando(2), "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Ccl, Conta, SldIni, Deb" + sPeriodo + ", Cre" + sPeriodo + ") VALUES (?, ?, ?, ?, ?, ?, ?)", IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, sCcl1, sConta1, dSldIni, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 188309


    Else

        'atualiza o total de credito e debito da conta e periodo especificados
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando3, "UPDATE MvPerCcl SET Deb" + sPeriodo + "= Deb" + sPeriodo + " + ?, Cre" + sPeriodo + "= Cre" + sPeriodo + " + ?", tProcessa_Lancamento.lComando4, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 10628

    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Atualiza_CclMes = SUCESSO
    
    Exit Function
    
Erro_Atualiza_CclMes:

    Atualiza_CclMes = gErr
    
    Select Case gErr
    
        Case 10529, 10530
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL1", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, sCcl1, sConta1)

        Case 10531
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, sCcl1, sConta1)

        Case 10626, 10627
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL1", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, sCcl1, sConta1)

        Case 10628
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, sCcl1, sConta1)

        Case 188303
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 188304, 188305
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL3", gErr, tProcessa_Lancamento.iFilialEmpresa, sConta1, sCcl1)

        Case 188306
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, sCcl1, sConta1)

        Case 188307, 188308
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SALDOINICIALCONTACCL3", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta1, sCcl1)

        Case 188309
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, sCcl1, sConta1)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154887)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Processa_Lancamento_ContaMes1(ByVal iNivel As Integer, tProcessa_Lancamento As typeProcessa_Lancamento) As Long

Dim dDebito1 As Double
Dim dCredito1 As Double
Dim sConta As String
Dim sConta1 As String
Dim sConta2 As String
Dim iExiste_Proximo_Nivel As Integer
Dim lErro As Long

On Error GoTo Erro_Processa_Lancamento_ContaMes1

    'inicializa os acumuladores de debito e credito
    dDebito1 = 0
    dCredito1 = 0

    sConta = String(STRING_CONTA, 0)
    
    'guarda o número da conta que está sendo processada
    lErro = Mascara_RetornaContaNoNivel(iNivel, tProcessa_Lancamento.tLancamento.sConta, sConta)
    If lErro <> SUCESSO Then Error 9406
    
    sConta1 = sConta

    Do While tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE And sConta = sConta1

        'verifica se a conta possui um nivel mais "profundo"
        iExiste_Proximo_Nivel = Mascara_ExisteProxNivelCta(iNivel, tProcessa_Lancamento.tLancamento.sConta)

        If iExiste_Proximo_Nivel = SUCESSO Then

            'se existe um nivel mais profundo, processa-o
            lErro = Processa_Lancamento_ContaMes1(iNivel + 1, tProcessa_Lancamento)
            If lErro <> SUCESSO Then Error 5041

            'acumula debitos e creditos
            dDebito1 = dDebito1 + tProcessa_Lancamento.dDebito
            dCredito1 = dCredito1 + tProcessa_Lancamento.dCredito
    
            If tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE Then
                    
                'verifica se a conta possui o nivel em questão. Se não possuir sair do loop
                iExiste_Proximo_Nivel = Mascara_ExisteProxNivelCta(iNivel - 1, tProcessa_Lancamento.tLancamento.sConta)
                If iExiste_Proximo_Nivel <> SUCESSO Then Exit Do
    
                sConta = String(STRING_CONTA, 0)
    
                'armazena a conta do nivel em questão
                lErro = Mascara_RetornaContaNoNivel(iNivel, tProcessa_Lancamento.tLancamento.sConta, sConta)
                If lErro <> SUCESSO Then Error 9407
    
            End If

        Else

            'se este é o nivel mais profundo da conta, processa-o
            lErro = Processa_Lancamento_Analitico_ContaMes(tProcessa_Lancamento, sConta)
            If lErro <> SUCESSO Then Error 5042

            dDebito1 = tProcessa_Lancamento.dDebito
            dCredito1 = tProcessa_Lancamento.dCredito
            
            Exit Do

        End If

    Loop

    tProcessa_Lancamento.dDebito = dDebito1
    tProcessa_Lancamento.dCredito = dCredito1

    lErro = Atualiza_ContaMes(tProcessa_Lancamento, sConta1)
    If lErro <> SUCESSO Then Error 10624

    Processa_Lancamento_ContaMes1 = SUCESSO

    Exit Function

Erro_Processa_Lancamento_ContaMes1:

    Processa_Lancamento_ContaMes1 = Err

    Select Case Err

        Case 5029
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCTA", Err, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, sConta)

        Case 5041, 5042, 10624
    
        Case 5347, 5382
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA1", Err, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, sConta)

        Case 9406, 9407
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaContaNoNivel", Err, tProcessa_Lancamento.tLancamento.sConta, iNivel)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154888)

    End Select

    Exit Function

End Function

Function Atualiza_ContaMes(tProcessa_Lancamento As typeProcessa_Lancamento, sConta1 As String) As Long

Dim sPeriodo As String
Dim iExercicio As Integer
Dim lErro As Long

On Error GoTo Erro_Atualiza_ContaMes

    sPeriodo = tProcessa_Lancamento.sPeriodo

    'seleciona o saldo de contas mensal para fazer sua atualizacao
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando4, "SELECT Exercicio FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ?", 0, iExercicio, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, sConta1)
    If lErro <> AD_SQL_SUCESSO Then Error 5347

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando4)
    If lErro <> AD_SQL_SUCESSO Then Error 5382

    'atualiza o total de credito e debito da conta e periodo especificados
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando3, "UPDATE MvPerCta SET Deb" + sPeriodo + "= Deb" + sPeriodo + " + ?, Cre" + sPeriodo + "= Cre" + sPeriodo + " + ?", tProcessa_Lancamento.lComando4, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
    If lErro <> AD_SQL_SUCESSO Then Error 5029

    'seleciona o saldo de contas mensal para fazer sua atualizacao no ambito empresa
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando4, "SELECT Exercicio FROM MvPerCta WHERE FilialEmpresa = ? AND Exercicio = ? AND Conta = ?", 0, iExercicio, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, sConta1)
    If lErro <> AD_SQL_SUCESSO Then Error 10622

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando4)
    If lErro <> AD_SQL_SUCESSO Then Error 10623

    'atualiza o total de credito e debito da conta e periodo especificados
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando3, "UPDATE MvPerCta SET Deb" + sPeriodo + "= Deb" + sPeriodo + " + ?, Cre" + sPeriodo + "= Cre" + sPeriodo + " + ?", tProcessa_Lancamento.lComando4, tProcessa_Lancamento.dDebito, tProcessa_Lancamento.dCredito)
    If lErro <> AD_SQL_SUCESSO Then Error 10624

    Atualiza_ContaMes = SUCESSO
    
    Exit Function
    
Erro_Atualiza_ContaMes:

    Atualiza_ContaMes = Err
    
    Select Case Err
    
        Case 5029, 10624
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCTA", Err, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, sConta1)

        Case 5347, 5382, 10622, 10623
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA1", Err, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, sConta1)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154889)

    End Select

    Exit Function

End Function

Function Processa_Lancamentos(tLote As typeLote_batch, ByVal iOperacao As Integer, ByVal iUsoCcl As Integer, ByVal iOperacao1 As Integer) As Long
'recebe os dados de um lote e comando a atualização/desatualização (dependendo do conteudo de iOperacao) dos seus lançamentos

Dim lErro As Long
Dim lPosicao As Long
Dim tLancamento_Sort As typeLancamento_Sort
Dim lComando2 As Long
Dim tProcessa_Lancamento As typeProcessa_Lancamento

On Error GoTo Erro_Processa_Lancamentos

    lComando2 = 0

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 5011

    tProcessa_Lancamento.iFilialEmpresa = tLote.iFilialEmpresa
    tProcessa_Lancamento.iOperacao = iOperacao
    tProcessa_Lancamento.iPeriodo = tLote.iPeriodo
    tProcessa_Lancamento.sPeriodo = Format(tLote.iPeriodo, "00")
    tProcessa_Lancamento.iExercicio = tLote.iExercicio
    tProcessa_Lancamento.iUsoCcl = iUsoCcl
    tProcessa_Lancamento.iOperacao1 = iOperacao1
    tProcessa_Lancamento.sCclAglutinado = "***"

    lErro = Verifica_Aglutina_Lancam_Por_Dia(tLote.sOrigem, tProcessa_Lancamento.iAglutinaLancamPorDia)
    If lErro <> SUCESSO Then gError 20493
    
    tProcessa_Lancamento.tLancamento.sConta = String(STRING_CONTA, 0)
    tProcessa_Lancamento.tLancamento.sCcl = String(STRING_CCL, 0)
    tProcessa_Lancamento.tLancamento.sHistorico = String(STRING_HISTORICO, 0)
    tProcessa_Lancamento.tLancamento.sProduto = String(STRING_PRODUTO, 0)
    tProcessa_Lancamento.tLancamento.sDocOrigem = String(STRING_DOCORIGEM, 0)
    tProcessa_Lancamento.tLancamento.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    tProcessa_Lancamento.tLancamento.sUsuario = String(STRING_USUARIO_NOMEREDUZIDO, 0)

    'Pesquisa os lançamentos pertencentes ao lote
    lErro = Comando_ExecutarPos(lComando2, "SELECT PeriodoLan, Doc, Seq, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, Aglutina, Produto, ApropriaCRProd, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Quantidade, DataEstoque, Status, Modelo, Gerencial, SubTipo, Usuario FROM LanPendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLote = ? AND Lote = ? ORDER BY Conta, Data, Ccl", _
        0, tProcessa_Lancamento.tLancamento.iPeriodoLan, tProcessa_Lancamento.tLancamento.lDoc, tProcessa_Lancamento.tLancamento.iSeq, tProcessa_Lancamento.tLancamento.dtData, tProcessa_Lancamento.tLancamento.sConta, tProcessa_Lancamento.tLancamento.sCcl, tProcessa_Lancamento.tLancamento.sHistorico, tProcessa_Lancamento.tLancamento.dValor, tProcessa_Lancamento.tLancamento.lNumIntDoc, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iTransacao, tProcessa_Lancamento.tLancamento.iAglutina, tProcessa_Lancamento.tLancamento.sProduto, tProcessa_Lancamento.tLancamento.iApropriaCRProd, _
                                    tProcessa_Lancamento.tLancamento.lContaSimples, tProcessa_Lancamento.tLancamento.iSeqContraPartida, tProcessa_Lancamento.tLancamento.iEscaninho_Custo, tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar, tProcessa_Lancamento.tLancamento.iClienteFornecedor, tProcessa_Lancamento.tLancamento.sDocOrigem, tProcessa_Lancamento.tLancamento.dQuantidade, tProcessa_Lancamento.tLancamento.dtDataEstoque, tProcessa_Lancamento.tLancamento.iStatus, tProcessa_Lancamento.tLancamento.sModelo, tProcessa_Lancamento.tLancamento.iGerencial, tProcessa_Lancamento.tLancamento.iSubTipo, tProcessa_Lancamento.tLancamento.sUsuario, tLote.iFilialEmpresa, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then gError 5012

    tProcessa_Lancamento.tLancamento.iFilialEmpresa = tLote.iFilialEmpresa
    tProcessa_Lancamento.tLancamento.sOrigem = tLote.sOrigem
    tProcessa_Lancamento.tLancamento.iExercicio = tLote.iExercicio
    tProcessa_Lancamento.tLancamento.iLote = tLote.iLote
    tProcessa_Lancamento.tLancamento.iPeriodoLote = tLote.iPeriodo
    
    tProcessa_Lancamento.lComando2 = lComando2

    'Le o primeiro lançamento pertencente ao lote
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9364
    
    If lErro = AD_SQL_SUCESSO Then

        lErro = Processa_Lancamentos_1(tProcessa_Lancamento)
        If lErro <> SUCESSO Then gError 83543

    End If

    Call Comando_Fechar(lComando2)
    
    tProcessa_Lancamento.lComando2 = 0
    
    Processa_Lancamentos = SUCESSO

    Exit Function

Erro_Processa_Lancamentos:

    Processa_Lancamentos = gErr

    Select Case gErr

        Case 5011
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 5012, 9364
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS", gErr, tLote.iFilialEmpresa, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tLote.iLote)

        Case 20493, 83543
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154890)

    End Select

    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Public Function Carrega_ColFiliais() As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_ColFiliais

    Set gcolFiliais = New Collection
    
    lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, gcolFiliais)
    If lErro <> SUCESSO Then Error 55838

    Carrega_ColFiliais = SUCESSO
    
    Exit Function

Erro_Carrega_ColFiliais:

    Carrega_ColFiliais = Err

    Select Case Err
    
        Case 55838
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 154891)

    End Select

    Exit Function

End Function

'Transportar para ClassSelect
Function CMProdApurado_Escaninho_Le_Mes(iFilialEmpresa As Integer, sProduto As String, iMes As Integer, iAno As Integer, dCMPAtual As Double, iEscaninho_Custo As Integer, alComando() As Long) As Long
'Calcula Custo Médio de Produção Apurado do Produto ou o custo medio do Escaninho passado como parametro
'Se o mês  passado ainda não foi apurado retorna ZERO

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_CMProdApurado_Escaninho_Le_Mes

    objEstoqueMes.iFilialEmpresa = iFilialEmpresa
    
    lErro = CF("EstoqueMes_Le_Apurado", objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 46225 Then gError 71554
        
    'Se já existe um Mes apurado
    If lErro = SUCESSO Then
        
        'Se o Ano o Mês passado for Menor ou igual ao ultimo Ano-Mes apurado
        If (iAno < objEstoqueMes.iAno) Or (iAno = objEstoqueMes.iAno And iMes <= objEstoqueMes.iMes) Then
        
            If iEscaninho_Custo = ESCANINHO_NOSSO Then
        
                'calcula o custo medio de produção do produto em questão
                lErro = CF("Calcula_CustoMedioProducao", iFilialEmpresa, sProduto, iAno, iMes, dCMPAtual)
                If lErro <> SUCESSO And lErro <> 25433 And lErro <> 55052 Then gError 71555
                
                'Se não encontrou o SldMesEst
                If lErro = 25433 Then gError 71556
                
                'Se não encontrou o SldMesEst
                If lErro = 55052 Then gError 71557
                
            ElseIf iEscaninho_Custo = ESCANINHO_NOSSO_EM_CONSIGNACAO Then
        
                lErro = CF("SldMesEst2_Le_CustoMedioConsig", alComando(1), iFilialEmpresa, iAno, sProduto, iMes, dCMPAtual)
                If lErro <> SUCESSO Then gError 71558
                
            ElseIf iEscaninho_Custo = ESCANINHO_NOSSO_EM_CONSERTO Then
        
                lErro = CF("SldMesEst2_Le_CustoMedioConserto", alComando(2), iFilialEmpresa, iAno, sProduto, iMes, dCMPAtual)
                If lErro <> SUCESSO Then gError 71559

            ElseIf iEscaninho_Custo = ESCANINHO_NOSSO_EM_DEMO Then
        
                lErro = CF("SldMesEst2_Le_CustoMedioDemo", alComando(3), iFilialEmpresa, iAno, sProduto, iMes, dCMPAtual)
                If lErro <> SUCESSO Then gError 71560
            
            ElseIf iEscaninho_Custo = ESCANINHO_NOSSO_EM_OUTROS Then
        
                lErro = CF("SldMesEst2_Le_CustoMedioOutros", alComando(4), iFilialEmpresa, iAno, sProduto, iMes, dCMPAtual)
                If lErro <> SUCESSO Then gError 71561
            
            ElseIf iEscaninho_Custo = ESCANINHO_NOSSO_EM_BENEF Then
        
                lErro = CF("SldMesEst2_Le_CustoMedioBenef", alComando(5), iFilialEmpresa, iAno, sProduto, iMes, dCMPAtual)
                If lErro <> SUCESSO Then gError 71562
        
            ElseIf iEscaninho_Custo = ESCANINHO_3_EM_CONSIGNACAO Then
        
                lErro = CF("SldMesEst1_Le_CustoMedioConsig3", alComando(6), iFilialEmpresa, iAno, sProduto, iMes, dCMPAtual)
                If lErro <> SUCESSO Then gError 71563
        
            ElseIf iEscaninho_Custo = ESCANINHO_3_EM_CONSERTO Then
        
                lErro = CF("SldMesEst1_Le_CustoMedioConserto3", alComando(7), iFilialEmpresa, iAno, sProduto, iMes, dCMPAtual)
                If lErro <> SUCESSO Then gError 71564

            ElseIf iEscaninho_Custo = ESCANINHO_3_EM_DEMO Then
        
                lErro = CF("SldMesEst1_Le_CustoMedioDemo3", alComando(8), iFilialEmpresa, iAno, sProduto, iMes, dCMPAtual)
                If lErro <> SUCESSO Then gError 71565
            
            ElseIf iEscaninho_Custo = ESCANINHO_3_EM_OUTROS Then
        
                lErro = CF("SldMesEst1_Le_CustoMedioOutros3", alComando(9), iFilialEmpresa, iAno, sProduto, iMes, dCMPAtual)
                If lErro <> SUCESSO Then gError 71566
            
            ElseIf iEscaninho_Custo = ESCANINHO_3_EM_BENEF Then
        
                lErro = CF("SldMesEst1_Le_CustoMedioBenef3", alComando(10), iFilialEmpresa, iAno, sProduto, iMes, dCMPAtual)
                If lErro <> SUCESSO Then gError 71567
            
            
            End If
            
        End If
    
    End If
    
    CMProdApurado_Escaninho_Le_Mes = SUCESSO

    Exit Function

Erro_CMProdApurado_Escaninho_Le_Mes:

    CMProdApurado_Escaninho_Le_Mes = gErr

    Select Case gErr

        Case 71554, 71555, 71558, 71559, 71560, 71561, 71562, 71563, 71564, 71565, 71566, 71567
        
        Case 71556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO", gErr, iFilialEmpresa, iAno, sProduto)
        
        Case 71557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO1", gErr, iFilialEmpresa, sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154892)

    End Select

    Exit Function

End Function

Private Function Atualiza_CliForn(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'atualiza os cadastros que sintetizam os dados de cliente e fornecedor para os livros (diario/razao) auxiliares

Dim lErro As Long

On Error GoTo Erro_Atualiza_CliForn

    If tProcessa_Lancamento.tLancamento.iClienteFornecedor = CLIENTE_ORIGEM_DOCUMENTO Then

        lErro = Atualiza_MvDiaCli(tProcessa_Lancamento)
        If lErro <> SUCESSO Then gError 71654

        lErro = Atualiza_MvPerCli(tProcessa_Lancamento)
        If lErro <> SUCESSO Then gError 71655

    ElseIf tProcessa_Lancamento.tLancamento.iClienteFornecedor = FORNECEDOR_ORIGEM_DOCUMENTO Then

        lErro = Atualiza_MvDiaForn(tProcessa_Lancamento)
        If lErro <> SUCESSO Then gError 71656

        lErro = Atualiza_MvPerForn(tProcessa_Lancamento)
        If lErro <> SUCESSO Then gError 71657

    End If

    Atualiza_CliForn = SUCESSO
    
    Exit Function
    
Erro_Atualiza_CliForn:

    Atualiza_CliForn = gErr
    
    Select Case gErr
    
        Case 71654, 71655, 71656, 71657
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154893)

    End Select

    Exit Function
    
End Function

Private Function Atualiza_MvDiaCli(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'atualiza o arquivo de movimentos diários de cliente

Dim lErro As Long
Dim dtData1 As Date
Dim dCredito As Double
Dim dDebito As Double

On Error GoTo Erro_Atualiza_MvDiaCli

    'seleciona os totais de debito e credito do dia do lancamento para serem atualizados
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando10, "SELECT Data FROM MvDiaCli WHERE FilialEmpresa = ? AND Cliente = ? AND Filial = ? AND Data = ?", 0, dtData1, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 71658

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando10)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71659
    
    If tProcessa_Lancamento.tLancamento.iCredDebLivroAuxiliar = CONTA_CREDITO Then
        dCredito = tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar
    Else
        dDebito = -tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar
    End If
    
    'se nao encontrou o registro com os totais ==> cadastra o registro
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo de dia
        lErro = Comando_Executar(tProcessa_Lancamento.lComando11, "INSERT INTO MvDiaCli (FilialEmpresa,Cliente, Filial,Data,Deb,Cre) VALUES (?,?,?,?,?,?)", tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71660

    Else

        'atualiza os totais de debito e credito do dia do lancamento
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando12, "UPDATE MvDiaCli SET Deb = Deb + ?, Cre = Cre + ?", tProcessa_Lancamento.lComando10, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71661

    End If

    'seleciona os totais de debito e credito do dia do lancamento para serem atualizados no ambito empresa
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando10, "SELECT Data FROM MvDiaCli WHERE FilialEmpresa = ? AND Cliente = ? AND Filial = ? AND Data = ?", 0, dtData1, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 71662

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando10)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71663
    
    'se nao encontrou o registro com os totais ==> cadastra o registro
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo de dia
        lErro = Comando_Executar(tProcessa_Lancamento.lComando11, "INSERT INTO MvDiaCli (FilialEmpresa,Cliente, Filial,Data,Deb,Cre) VALUES (?,?,?,?,?,?)", IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71664

    Else

        'atualiza os totais de debito e credito do dia do lancamento
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando12, "UPDATE MvDiaCli SET Deb = Deb + ?, Cre = Cre + ?", tProcessa_Lancamento.lComando10, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71665

    End If

    Atualiza_MvDiaCli = SUCESSO
    
    Exit Function
    
Erro_Atualiza_MvDiaCli:

    Select Case gErr
    
        Case 71658, 71659
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACLI1", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)

        Case 71660
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVDIACLI", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)
            
        Case 71661
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACLI", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)

        Case 71662, 71663
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACLI1", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)

        Case 71664
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVDIACLI", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)
            
        Case 71665
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACLI", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154894)

    End Select

    Exit Function

End Function

Private Function Atualiza_MvDiaForn(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'atualiza o arquivo de movimentos diários de fornecedor

Dim lErro As Long
Dim dtData1 As Date
Dim dCredito As Double
Dim dDebito As Double

On Error GoTo Erro_Atualiza_MvDiaForn

    'seleciona os totais de debito e credito do dia do lancamento para serem atualizados
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando13, "SELECT Data FROM MvDiaForn WHERE FilialEmpresa = ? AND Fornecedor = ? AND Filial = ? AND Data = ?", 0, dtData1, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 71666

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando13)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71667
    
    If tProcessa_Lancamento.tLancamento.iCredDebLivroAuxiliar = CONTA_CREDITO Then
        dCredito = tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar
    Else
        dDebito = -tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar
    End If
    
    'se nao encontrou o registro com os totais ==> cadastra o registro
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo de dia
        lErro = Comando_Executar(tProcessa_Lancamento.lComando14, "INSERT INTO MvDiaForn (FilialEmpresa,Fornecedor, Filial,Data,Deb,Cre) VALUES (?,?,?,?,?,?)", tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71668

    Else

        'atualiza os totais de debito e credito do dia do lancamento
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando15, "UPDATE MvDiaForn SET Deb = Deb + ?, Cre = Cre + ?", tProcessa_Lancamento.lComando13, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71669

    End If

    'seleciona os totais de debito e credito do dia do lancamento para serem atualizados no ambito empresa
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando13, "SELECT Data FROM MvDiaForn WHERE FilialEmpresa = ? AND Fornecedor = ? AND Filial = ? AND Data = ?", 0, dtData1, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 71670

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando13)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71671
    
    'se nao encontrou o registro com os totais ==> cadastra o registro
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo de dia
        lErro = Comando_Executar(tProcessa_Lancamento.lComando14, "INSERT INTO MvDiaForn (FilialEmpresa,Fornecedor, Filial,Data,Deb,Cre) VALUES (?,?,?,?,?,?)", IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71672

    Else

        'atualiza os totais de debito e credito do dia do lancamento
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando15, "UPDATE MvDiaForn SET Deb = Deb + ?, Cre = Cre + ?", tProcessa_Lancamento.lComando13, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71673

    End If

    Atualiza_MvDiaForn = SUCESSO
    
    Exit Function
    
Erro_Atualiza_MvDiaForn:

    Select Case gErr
    
        Case 71666, 71667
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIAFORN1", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)

        Case 71668
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVDIAFORN", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)
            
        Case 71669
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIAFORN", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)

        Case 71670, 71671
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIAFORN1", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)

        Case 71672
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVDIAFORN", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)
            
        Case 71673
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIAFORN", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.dtData)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154895)

    End Select

    Exit Function

End Function

Private Function Atualiza_MvPerCli(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'atualiza o arquivo de saldos mensais por cliente

Dim sPeriodo As String
Dim iExercicio As Integer
Dim lErro As Long
Dim dCredito As Double
Dim dDebito As Double

On Error GoTo Erro_Atualiza_MvPerCli

    sPeriodo = tProcessa_Lancamento.sPeriodo

    'seleciona o saldo de contas mensal para fazer sua atualizacao
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando16, "SELECT Exercicio FROM MvPerCli WHERE FilialEmpresa = ? AND Exercicio = ? AND Cliente = ? AND Filial =?", 0, iExercicio, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)
    If lErro <> AD_SQL_SUCESSO Then gError 71674

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando16)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71675

    If tProcessa_Lancamento.tLancamento.iCredDebLivroAuxiliar = CONTA_CREDITO Then
        dCredito = tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar
    Else
        dDebito = -tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar
    End If
    
    'se nao encontrou o registro com os totais ==> cadastra o registro
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo do mes
        lErro = Comando_Executar(tProcessa_Lancamento.lComando17, "INSERT INTO MvPerCli (FilialEmpresa, Exercicio, Cliente, Filial, Deb" + sPeriodo + " , Cre" + sPeriodo + ") VALUES (?,?,?,?,?,?)", tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71676

    Else
    
        'atualiza o saldo do mes
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando18, "UPDATE MvPerCli SET Deb" + sPeriodo + "= Deb" + sPeriodo + " + ?, Cre" + sPeriodo + "= Cre" + sPeriodo + " + ?", tProcessa_Lancamento.lComando16, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71677

    End If

    'seleciona o saldo de contas mensal da EMPRESA_TODA para fazer sua atualizacao
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando16, "SELECT Exercicio FROM MvPerCli WHERE FilialEmpresa = ? AND Exercicio = ? AND Cliente = ? AND Filial =?", 0, iExercicio, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)
    If lErro <> AD_SQL_SUCESSO Then gError 71678

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando16)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71679

    'se nao encontrou o registro com os totais ==> cadastra o registro
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo do mes
        lErro = Comando_Executar(tProcessa_Lancamento.lComando17, "INSERT INTO MvPerCli (FilialEmpresa, Exercicio, Cliente, Filial, Deb" + sPeriodo + " , Cre" + sPeriodo + ") VALUES (?,?,?,?,?,?)", IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71680

    Else
    
        'atualiza o saldo do mes
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando18, "UPDATE MvPerCli SET Deb" + sPeriodo + "= Deb" + sPeriodo + " + ?, Cre" + sPeriodo + "= Cre" + sPeriodo + " + ?", tProcessa_Lancamento.lComando16, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71681

    End If

    Atualiza_MvPerCli = SUCESSO
    
    Exit Function
    
Erro_Atualiza_MvPerCli:

    Atualiza_MvPerCli = Err
    
    Select Case gErr
    
        Case 71674, 71675
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCLI", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)

        Case 71676
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCLI", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)
            
        Case 71677
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCLI", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)

        Case 71678, 71679
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCLI", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)

        Case 71680
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCLI", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)
            
        Case 71681
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCLI", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154896)

    End Select

    Exit Function

End Function

Private Function Atualiza_MvPerForn(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'atualiza o arquivo de saldos mensais por Fornecedor

Dim sPeriodo As String
Dim iExercicio As Integer
Dim lErro As Long
Dim dCredito As Double
Dim dDebito As Double

On Error GoTo Erro_Atualiza_MvPerForn

    sPeriodo = tProcessa_Lancamento.sPeriodo

    'seleciona o saldo de contas mensal para fazer sua atualizacao
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando19, "SELECT Exercicio FROM MvPerForn WHERE FilialEmpresa = ? AND Exercicio = ? AND Fornecedor = ? AND Filial =?", 0, iExercicio, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)
    If lErro <> AD_SQL_SUCESSO Then gError 71682

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando19)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71683

    If tProcessa_Lancamento.tLancamento.iCredDebLivroAuxiliar = CONTA_CREDITO Then
        dCredito = tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar
    Else
        dDebito = -tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar
    End If
    
    'se nao encontrou o registro com os totais ==> cadastra o registro
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo do mes
        lErro = Comando_Executar(tProcessa_Lancamento.lComando20, "INSERT INTO MvPerForn (FilialEmpresa, Exercicio, Fornecedor, Filial, Deb" + sPeriodo + " , Cre" + sPeriodo + ") VALUES (?,?,?,?,?,?)", tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71684

    Else
    
        'atualiza o saldo do mes
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando21, "UPDATE MvPerForn SET Deb" + sPeriodo + "= Deb" + sPeriodo + " + ?, Cre" + sPeriodo + "= Cre" + sPeriodo + " + ?", tProcessa_Lancamento.lComando19, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71685

    End If

    'seleciona o saldo de contas mensal da EMPRESA_TODA para fazer sua atualizacao
    lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando19, "SELECT Exercicio FROM MvPerForn WHERE FilialEmpresa = ? AND Exercicio = ? AND Fornecedor = ? AND Filial =?", 0, iExercicio, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)
    If lErro <> AD_SQL_SUCESSO Then gError 71686

    lErro = Comando_BuscarPrimeiro(tProcessa_Lancamento.lComando19)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71687

    'se nao encontrou o registro com os totais ==> cadastra o registro
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o saldo do mes
        lErro = Comando_Executar(tProcessa_Lancamento.lComando20, "INSERT INTO MvPerForn (FilialEmpresa, Exercicio, Fornecedor, Filial, Deb" + sPeriodo + " , Cre" + sPeriodo + ") VALUES (?,?,?,?,?,?)", IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71688

    Else
    
        'atualiza o saldo do mes
        lErro = Comando_ExecutarPos(tProcessa_Lancamento.lComando21, "UPDATE MvPerForn SET Deb" + sPeriodo + "= Deb" + sPeriodo + " + ?, Cre" + sPeriodo + "= Cre" + sPeriodo + " + ?", tProcessa_Lancamento.lComando19, dDebito, dCredito)
        If lErro <> AD_SQL_SUCESSO Then gError 71689

    End If

    Atualiza_MvPerForn = SUCESSO
    
    Exit Function
    
Erro_Atualiza_MvPerForn:

    Atualiza_MvPerForn = Err
    
    Select Case gErr
    
        Case 71682, 71683
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERFORN", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)

        Case 71684
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERFORN", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)
            
        Case 71685
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERFORN", gErr, tProcessa_Lancamento.iFilialEmpresa, tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)

        Case 71686, 71687
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERFORN", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)

        Case 71688
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERFORN", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)
            
        Case 71689
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERFORN", gErr, IIf(tProcessa_Lancamento.iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), tProcessa_Lancamento.iExercicio, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iFilialCliForn)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154897)

    End Select

    Exit Function

End Function

Function gcolModulo_sOrigemAglutina(sOrigem As String) As String

    Select Case sOrigem
    
        Case "CP", "BCP", "CCP"
            gcolModulo_sOrigemAglutina = "ACP"
        
        Case "CR", "BCR", "CCR"
            gcolModulo_sOrigemAglutina = "ACR"

        Case "COM", "CCM"
            gcolModulo_sOrigemAglutina = "ACM"
            
        Case "EST", "CES"
            gcolModulo_sOrigemAglutina = "AES"

        Case "FAT", "CFT"
            gcolModulo_sOrigemAglutina = "AFT"
            
        Case "TES", "CTS"
            gcolModulo_sOrigemAglutina = "ATS"
    
        Case Else
            gcolModulo_sOrigemAglutina = ""
            
    End Select
    
End Function

Function Processa_Lancamentos_1(tProcessa_Lancamento As typeProcessa_Lancamento) As Long
'continuação de Processa_Lancamentos

Dim lErro As Long
Dim tLancamento_Sort As typeLancamento_Sort
Dim lPosicao As Long

On Error GoTo Erro_Processa_Lancamentos_1

    tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE

    'Abre o arquivo que vai armazenar temporariamente os lancamentos do lote
    tProcessa_Lancamento.lID_Arq_Temp = Arq_Temp_Criar(Len(tLancamento_Sort))
    If tProcessa_Lancamento.lID_Arq_Temp = 0 Then gError 83533

    'Abre o arquivo 1 de sort
    tProcessa_Lancamento.lID_Arq_Sort = Sort_Abrir(Len(lPosicao), 1)
    If tProcessa_Lancamento.lID_Arq_Sort = 0 Then gError 83534

    'Abre o arquivo 2 de sort
    tProcessa_Lancamento.lID_Arq_Sort1 = Sort_Abrir(Len(lPosicao), 3)
    If tProcessa_Lancamento.lID_Arq_Sort1 = 0 Then gError 83535

    'Abre o arquivo 3 de sort
    tProcessa_Lancamento.lID_Arq_Sort2 = Sort_Abrir(Len(lPosicao), 2)
    If tProcessa_Lancamento.lID_Arq_Sort2 = 0 Then gError 83536

    'Processa a atualização dos Saldos de Contas Por Dia
    lErro = Processa_Lancamento_ContaDia(tProcessa_Lancamento)
    If lErro <> SUCESSO Then gError 83537

    'Prepara o arquivo temporário para leitura
    lErro = Arq_Temp_Preparar(tProcessa_Lancamento.lID_Arq_Temp)
    If lErro <> AD_BOOL_TRUE Then gError 83538

    'Processa a atualização dos Saldos de Contas Por Mes
    lErro = Processa_Lancamento_ContaMes(tProcessa_Lancamento)
    If lErro <> SUCESSO Then gError 83539

    'se o sistema usa centro de custo/lucro contabil ou extra-contábil, computa os saldos correspondentes
    If tProcessa_Lancamento.iUsoCcl = CCL_USA_CONTABIL Or tProcessa_Lancamento.iUsoCcl = CCL_USA_EXTRACONTABIL Then
        
        'Processa a atualização dos Saldos de Centro de Custo/Lucro Por Mes
        lErro = Processa_Lancamento_CclMes(tProcessa_Lancamento)
        If lErro <> SUCESSO Then gError 83540
        
        lErro = Processa_Lancamento_CclDia(tProcessa_Lancamento)
        If lErro <> SUCESSO Then gError 83541
        
    End If

    Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)
    Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort)
    Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort1)
    Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort2)

    Processa_Lancamentos_1 = SUCESSO

    Exit Function

Erro_Processa_Lancamentos_1:

    Processa_Lancamentos_1 = gErr

    Select Case gErr

        Case 83533
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_ARQUIVO_TEMPORARIO", gErr)

        Case 83534
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_ARQUIVO_SORT", gErr)
            Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)

        Case 83535
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_ARQUIVO_SORT", gErr)
            Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort)

        Case 83536
            Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort1)
            
        Case 83537, 83539, 83540, 83541
            Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort1)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort2)

        Case 83538
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_ARQUIVO_TEMP", gErr)
            Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort1)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort2)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154898)

    End Select

    Exit Function

End Function

Private Function Verifica_Exercicio_Fechado(colExercicio As Collection, ByVal iExercicio As Integer, ByVal iFilialEmpresa As Integer) As Long
'verifica se o exercicio está habilitado a receber lançamentos, ou seja, está presente em colExercicio senão verifica se o exercicio está fechado,
'se não estiver coloca-o como aberto e na colecao colExercicio

Dim iAchou As Integer
Dim lErro As Long
Dim iIndice As Integer
Dim alComando1(1 To 3) As Long

On Error GoTo Erro_Verifica_Exercicio_Fechado

    For iIndice = 1 To colExercicio.Count
    
        If iExercicio = colExercicio.Item(iIndice) Then
            iAchou = 1
            Exit For
        End If
        
    Next

    If iAchou = 0 Then

        For iIndice = LBound(alComando1) To UBound(alComando1)
            alComando1(iIndice) = Comando_Abrir()
            If alComando1(iIndice) = 0 Then gError 83583
        Next
        
        lErro = Atualiza_ExercicioFilial_Nao_Apurado(alComando1(), iExercicio, iFilialEmpresa)
        If lErro <> SUCESSO Then gError 83584
        
        colExercicio.Add iExercicio
        
    End If
    
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    
    Verifica_Exercicio_Fechado = SUCESSO
    
    Exit Function
    
Erro_Verifica_Exercicio_Fechado:

    Verifica_Exercicio_Fechado = gErr

    Select Case gErr
    
        Case 83583
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 83584
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154899)
        
    End Select
    
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    
    Exit Function

End Function

Private Function Lancamento_Excluir(lComando As Long, lComando1 As Long, tLancamento As typeLancamento) As Long
'exclui o lançamento passado como parametro do bd

Dim iSeq As Integer
Dim lErro As Long

On Error GoTo Erro_Lancamento_Excluir

    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM Lancamentos", lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 83616
    
    Lancamento_Excluir = SUCESSO
    
    Exit Function
    
Erro_Lancamento_Excluir:

    Lancamento_Excluir = gErr
    
    Select Case gErr
    
        Case 83613, 83614
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS6", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
        Case 83615
            Call Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_NAO_CADASTRADO", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
        Case 83616
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LANCAMENTO", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154900)
        
    End Select
    
    Exit Function
    
End Function

Function Processa_Exclusao_LanPendente(iFilialEmpresa As Integer, ByVal iOrigemLcto As Integer, lNumIntDocOrigem As Long, colExercicio As Collection) As Long
'exclui os lançamentos pendentes associados ao par OrigemLcto/NumIntDocOrigem

Dim lComando As Long
Dim lErro As Long
Dim iIndice As Integer
Dim tLancamento As typeLancamento

On Error GoTo Erro_Processa_Exclusao_LanPendente

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 83600

    tLancamento.sOrigem = String(STRING_ORIGEM, 0)

    'Pesquisa os lançamentos pendentes pertencentes ao NumIntDoc/OrigemLcto
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq FROM LanPendente WHERE FilialEmpresa = ? AND NumIntDoc = ? AND Lancamentos.Transacao = TransacaoCTBCodigo.Codigo AND TransacaoCTBCodigo.OrigemLcto = ? ", _
    0, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq, iFilialEmpresa, lNumIntDocOrigem, iOrigemLcto)
    If lErro <> AD_SQL_SUCESSO Then gError 83605

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83606
    
    Do While lErro = AD_SQL_SUCESSO
    
        'verifica se o exercicio está habilitado a receber lançamentos, ou seja, está presente em colExercicio senão verifica se o exercicio está fechado se estiver ==> erro,
        'se não estiver coloca-o como aberto e na colecao colExercicio
        lErro = Verifica_Exercicio_Fechado(colExercicio, tLancamento.iExercicio, tLancamento.iFilialEmpresa)
        If lErro <> SUCESSO Then gError 83626
        
        'exclui o lançamento pendente em questão
        lErro = LanPendente_Excluir1(tLancamento)
        If lErro <> SUCESSO Then gError 83633
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83608
    
    Loop

    Call Comando_Fechar(lComando)

    Processa_Exclusao_LanPendente = SUCESSO
    
    Exit Function
    
Erro_Processa_Exclusao_LanPendente:

    Processa_Exclusao_LanPendente = gErr

    Select Case gErr
    
        Case 83600
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 83605, 83606, 83608
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS10", gErr, lNumIntDocOrigem, iOrigemLcto)

        Case 83626, 83633
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154901)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function LanPendente_Excluir1(tLancamento As typeLancamento) As Long
'exclui o lançamento pendente passado como parametro

Dim alComando1(1 To 2) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim iSeq As Integer

On Error GoTo Erro_LanPendente_Excluir1

    For iIndice = LBound(alComando1) To UBound(alComando1)
        alComando1(iIndice) = Comando_Abrir()
        If alComando1(iIndice) = 0 Then gError 83627
    Next

    'Pesquisa os lançamentos pendentes pertencentes a Filial/Origem/Exercicio/Periodo/Doc/Seq passados como parametro
    lErro = Comando_ExecutarPos(alComando1(1), "SELECT Seq FROM LanPendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ? AND Seq =? ", _
    0, iSeq, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
    If lErro <> AD_SQL_SUCESSO Then gError 83628

    lErro = Comando_BuscarPrimeiro(alComando1(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83629
    
    'se o lancamento pendente não estiver cadastrado ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 83632
    
    'excluir o lancamento pendente
    lErro = Comando_ExecutarPos(alComando1(2), "DELETE FROM LanPendente", alComando1(1))
    If lErro <> AD_SQL_SUCESSO Then gError 83630
    
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    
    LanPendente_Excluir1 = SUCESSO
    
    Exit Function
    
Erro_LanPendente_Excluir1:

    LanPendente_Excluir1 = gErr
    
    Select Case gErr

        Case 83627
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 83628, 83629
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANPENDENTE6", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
        Case 83630
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LANPENDENTE", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
        Case 83632
            Call Rotina_Erro(vbOKOnly, "ERRO_LANPENDENTE_NAO_CADASTRADO", gErr, tLancamento.iFilialEmpresa, tLancamento.sOrigem, tLancamento.iExercicio, tLancamento.iPeriodoLan, tLancamento.lDoc, tLancamento.iSeq)
        
    End Select
    
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    
    Exit Function

End Function

Function Processa_Exclusao_Lancamentos(ByVal iFilialEmpresa As Integer, ByVal iOrigemLcto As Integer, lNumIntDocOrigem As Long, colExercicio As Collection) As Long
'exclui os lançamentos já contabilizados  associados a Filial/OrigemLcto/lNumIntDocOrigem se houverem

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim tProcessa_Lancamento As typeProcessa_Lancamento
Dim iUsoCcl As Integer

On Error GoTo Erro_Processa_Exclusao_Lancamentos

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 83617

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 83618

    tProcessa_Lancamento.lComando2 = lComando2

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 83827

    lErro = Comando_Executar(lComando1, "SELECT UsoCcl FROM Configuracao", iUsoCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 83619
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 83620

    tProcessa_Lancamento.iFilialEmpresa = iFilialEmpresa
    tProcessa_Lancamento.iOperacao = ATUALIZACAO
    tProcessa_Lancamento.iUsoCcl = iUsoCcl
    tProcessa_Lancamento.iOperacao1 = ROTINA_EXCLUSAO_LANCAMENTOS
    tProcessa_Lancamento.sCclAglutinado = "***"
    tProcessa_Lancamento.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA

    tProcessa_Lancamento.tLancamento.sConta = String(STRING_CONTA, 0)
    tProcessa_Lancamento.tLancamento.sCcl = String(STRING_CCL, 0)
    tProcessa_Lancamento.tLancamento.sHistorico = String(STRING_HISTORICO, 0)
    tProcessa_Lancamento.tLancamento.sProduto = String(STRING_PRODUTO, 0)
    tProcessa_Lancamento.tLancamento.sDocOrigem = String(STRING_DOCORIGEM, 0)
    tProcessa_Lancamento.tLancamento.sModelo = String(STRING_PADRAOCONTAB_MODELO, 0)
    Set tProcessa_Lancamento.colExercicio = colExercicio

    'Pesquisa os lançamentos pertencentes ao NumIntDoc/OrigemLcto
    lErro = Comando_Executar(lComando3, "SELECT DISTINCT FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc FROM Lancamentos, TransacaoCTBCodigo WHERE FilialEmpresa = ? AND NumIntDoc = ? AND Lancamentos.Transacao = TransacaoCTBCodigo.Codigo AND TransacaoCTBCodigo.OrigemLcto = ? ORDER BY FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc", _
    tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sOrigem, tProcessa_Lancamento.tLancamento.iExercicio, tProcessa_Lancamento.tLancamento.iPeriodoLan, tProcessa_Lancamento.tLancamento.lDoc, iFilialEmpresa, lNumIntDocOrigem, iOrigemLcto)
    If lErro <> AD_SQL_SUCESSO Then gError 83828

    'Le o primeiro lançamento pertencente ao NumIntDoc/OrigemLcto
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83829
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Pesquisa os lançamentos pertencentes a FilialEmpresa/Origem/Exercicio/PeriodoLan/Doc e que tenham o campo produto preenchido
        lErro = Comando_ExecutarPos(lComando2, "SELECT Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor, NumIntDoc, FilialCliForn, CliForn, Transacao, DocAglutinado, SeqAglutinado, Aglutinado, Produto, ApropriaCRProd, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Quantidade, DataEstoque, Status, Modelo, Gerencial FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLan = ? AND Doc = ? ORDER BY Conta, Data, Ccl", 0, _
        tProcessa_Lancamento.tLancamento.iSeq, tProcessa_Lancamento.tLancamento.iLote, tProcessa_Lancamento.tLancamento.iPeriodoLote, tProcessa_Lancamento.tLancamento.dtData, tProcessa_Lancamento.tLancamento.sConta, tProcessa_Lancamento.tLancamento.sCcl, tProcessa_Lancamento.tLancamento.sHistorico, tProcessa_Lancamento.tLancamento.dValor, tProcessa_Lancamento.tLancamento.lNumIntDoc, tProcessa_Lancamento.tLancamento.iFilialCliForn, tProcessa_Lancamento.tLancamento.lCliForn, tProcessa_Lancamento.tLancamento.iTransacao, tProcessa_Lancamento.tLancamento.lDocAglutinado, tProcessa_Lancamento.tLancamento.iSeqAglutinado, tProcessa_Lancamento.tLancamento.iAglutina, tProcessa_Lancamento.tLancamento.sProduto, tProcessa_Lancamento.tLancamento.iApropriaCRProd, _
                                        tProcessa_Lancamento.tLancamento.lContaSimples, tProcessa_Lancamento.tLancamento.iSeqContraPartida, tProcessa_Lancamento.tLancamento.iEscaninho_Custo, tProcessa_Lancamento.tLancamento.dValorLivroAuxiliar, tProcessa_Lancamento.tLancamento.iClienteFornecedor, tProcessa_Lancamento.tLancamento.sDocOrigem, tProcessa_Lancamento.tLancamento.dQuantidade, tProcessa_Lancamento.tLancamento.dtDataEstoque, tProcessa_Lancamento.tLancamento.iStatus, tProcessa_Lancamento.tLancamento.sModelo, tProcessa_Lancamento.tLancamento.iGerencial, tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sOrigem, tProcessa_Lancamento.tLancamento.iExercicio, tProcessa_Lancamento.tLancamento.iPeriodoLan, tProcessa_Lancamento.tLancamento.lDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 83621
        
        'Le o primeiro lançamento pertencente ao lote
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83622
        
        tProcessa_Lancamento.iPeriodo = tProcessa_Lancamento.tLancamento.iPeriodoLan
        tProcessa_Lancamento.sPeriodo = Format(tProcessa_Lancamento.tLancamento.iPeriodoLan, "00")
        tProcessa_Lancamento.iExercicio = tProcessa_Lancamento.tLancamento.iExercicio
        
        If lErro = AD_SQL_SUCESSO Then
    
            lErro = Processa_Lancamentos_1(tProcessa_Lancamento)
            If lErro <> SUCESSO Then gError 83623
    
        End If

        'Le o proximo lançamento pertencente ao NumIntDoc/OrigemLcto
        lErro = Comando_BuscarProximo(lComando3)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83830

    Loop

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    tProcessa_Lancamento.lComando2 = 0
    
    Processa_Exclusao_Lancamentos = SUCESSO

    Exit Function

Erro_Processa_Exclusao_Lancamentos:

    Processa_Exclusao_Lancamentos = gErr

    Select Case gErr

        Case 83617, 83618
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 83619, 83620
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", gErr)

        Case 83828, 83829, 83830
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS9", gErr, lNumIntDocOrigem, iOrigemLcto)

        Case 83621, 83622
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS8", gErr, tProcessa_Lancamento.tLancamento.iFilialEmpresa, tProcessa_Lancamento.tLancamento.sOrigem, tProcessa_Lancamento.tLancamento.iExercicio, tProcessa_Lancamento.tLancamento.iPeriodoLan, tProcessa_Lancamento.tLancamento.lDoc)

        Case 83623
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154902)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function DataContabil_Valida(ByVal dtData As Date, ByVal sOrigem As String) As Long
'Tratamento do validate do campo DataContabil

Dim lErro As Long
Dim objPeriodo As New ClassPeriodo
Dim objExercicio As New ClassExercicio
Dim objPeriodosFilial As New ClassPeriodosFilial

On Error GoTo Erro_DataContabil_Valida

    lErro = CF("Periodo_Le", dtData, objPeriodo)
    If lErro <> SUCESSO Then gError 185066

    'Verifica se Exercicio está fechado
    lErro = CF("Exercicio_Le", objPeriodo.iExercicio, objExercicio)
    If lErro <> SUCESSO And lErro <> 10083 Then gError 185067

    'Exercicio não cadastrado
    If lErro = 10083 Then gError 185068

    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If objExercicio.iStatus = EXERCICIO_FECHADO Then gError 185069
    End If

    objPeriodosFilial.iFilialEmpresa = giFilialEmpresa
    objPeriodosFilial.iExercicio = objPeriodo.iExercicio
    objPeriodosFilial.iPeriodo = objPeriodo.iPeriodo
    objPeriodosFilial.sOrigem = sOrigem

    lErro = CF("PeriodosFilial_Le", objPeriodosFilial)
    If lErro <> SUCESSO Then gError 185070

    If giDesconsideraFechamentoPeriodo <> MARCADO Then
        If objPeriodosFilial.iFechado = PERIODO_FECHADO Then gError 185071
    End If
    
    DataContabil_Valida = SUCESSO

    Exit Function

Erro_DataContabil_Valida:
    
    DataContabil_Valida = gErr

    Select Case gErr

        Case 185066, 185067, 185070

        Case 185068
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_CADASTRADO", gErr, objPeriodo.iExercicio)

        Case 185069
            'Não é possível fazer lançamentos em exercício fechado
            Call Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_EXERCICIO_FECHADO", gErr, objPeriodo.iExercicio)

        Case 185071
            Call Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTOS_PERIODO_FECHADO", gErr, objPeriodosFilial.iExercicio, objPeriodosFilial.iPeriodo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185072)

    End Select

    Exit Function

End Function
