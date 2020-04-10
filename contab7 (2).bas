Attribute VB_Name = "Module7"
'Reprocessamento

Option Explicit

Dim giFilialOrigem As Integer

Function Rotina_Reprocessamento_Int(iFilialEmpresa As Integer, iExercicio As Integer, iPeriodo As Integer) As Long

Dim objFiliais As AdmFiliais
Dim lTransacao As Long
Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim tFaixaPeriodo As typeFaixaPeriodo

On Error GoTo Erro_Rotina_Reprocessamento_Int

'     For iPeriodo = 2 To 12

        lTransacao = 0
    
       'Inicia a transação
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 10693
    
        For iIndice = LBound(alComando) To UBound(alComando)
            
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 5272
            
        Next
    
        giFilialOrigem = iFilialEmpresa
    
        If iFilialEmpresa = EMPRESA_TODA Or iFilialEmpresa = Abs(giFilialAuxiliar) Then
    
            'Pesquisa o periodo em questão
            lErro = Comando_ExecutarPos(alComando(1), "SELECT DataInicio, DataFim FROM Periodo WHERE Exercicio = ? AND Periodo = ?", 0, tFaixaPeriodo.dtDataInicio, tFaixaPeriodo.dtDataFim, iExercicio, iPeriodo)
            If lErro <> AD_SQL_SUCESSO Then gError 5278
        
            'Le o periodo
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO Then gError 5279
        
            'se estiver trabalhando a nivel de empresa toda ==> zera os saldos de conta e ccl neste nivel
            lErro = Zera_Saldos_Empresa_Toda(iExercicio, iPeriodo, tFaixaPeriodo)
            If lErro <> SUCESSO Then gError 188337
    
            'se tiver selecionado a empresa, executa o reprocessamento para cada filial
            For Each objFiliais In gcolFiliais
        
                If objFiliais.iCodFilial <> EMPRESA_TODA And objFiliais.iCodFilial <> Abs(giFilialAuxiliar) Then
        
        
        
                    lErro = Rotina_Reprocessamento0(objFiliais.iCodFilial, iExercicio, iPeriodo)
                    If lErro <> SUCESSO Then gError 10694
                    
                End If
            
            Next
        
        Else
        
            'se tiver decidido reprocessar somente uma filial
            lErro = Rotina_Reprocessamento0(iFilialEmpresa, iExercicio, iPeriodo)
            If lErro <> SUCESSO Then gError 10695
        
        End If
    
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
       'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 10696
    
'    Next
    
    'Alteração Daniel em 09/05/02
    Call Rotina_Aviso(vbOKOnly, "AVISO_REPROCESSAMENTO_EXECUTADO")

    Rotina_Reprocessamento_Int = SUCESSO

    Exit Function

Erro_Rotina_Reprocessamento_Int:

    Rotina_Reprocessamento_Int = Err

    Select Case Err

        Case 10693
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 10694, 10695, 188337
            
        Case 10696
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154936)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Function Rotina_Reprocessamento0(iFilialEmpresa As Integer, iExercicio As Integer, iPeriodo As Integer) As Long

Dim alComando(1 To 7) As Long
Dim lErro As Long
Dim iStatus As Integer
Dim tFaixaPeriodo As typeFaixaPeriodo
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo
Dim iIndice As Integer
Dim iApurado As Integer

On Error GoTo Erro_Rotina_Reprocessamento0

    For iIndice = LBound(alComando) To UBound(alComando)
        
        alComando(iIndice) = 0
        
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 5272
        
    Next
    
    'Pesquisa o Exercicio em questão
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Status FROM Exercicios WHERE Exercicio = ?", 0, iStatus, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5273

    'le o Exercicio
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 5274

    'não permite a mudança no status do exercicio
    lErro = Comando_LockShared(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 5276

    'Se o Exercicio estiver fechado ==> erro
    If iStatus = EXERCICIO_FECHADO Then Error 5275

    'Pesquisa o ExercicioFilial em questão
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Status FROM ExerciciosFilial WHERE FilialEmpresa=? AND Exercicio = ?", 0, iStatus, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 10697

    'le o ExercicioFilial
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 10698

    'não permite a mudança no status do exercicioFilial
    lErro = Comando_LockExclusive(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 10699
    
    'Pesquisa o periodo em questão
    lErro = Comando_ExecutarPos(alComando(3), "SELECT DataInicio, DataFim FROM Periodo WHERE Exercicio = ? AND Periodo = ?", 0, tFaixaPeriodo.dtDataInicio, tFaixaPeriodo.dtDataFim, iExercicio, iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 5278

    'Le o periodo
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then Error 5279

    'Lock do Periodo
    lErro = Comando_LockExclusive(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then Error 5280

    'Pesquisa o periodoFilial em questão
    lErro = Comando_ExecutarPos(alComando(4), "SELECT Apurado FROM PeriodosFilial WHERE FilialEmpresa = ? AND Exercicio = ? AND Periodo = ?", 0, iApurado, iFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 10700

    'Le o periodoFilial
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO Then Error 10701

    'Lock do PeriodoFilial
    lErro = Comando_LockExclusive(alComando(4))
    If lErro <> AD_SQL_SUCESSO Then Error 10702

    'Acerta o valor de lancamentos aglutinadores
    lErro = LctosAglutinadores_AcertaValor(iFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 5297
    
    lErro = LctosAglutinados_Sem_Aglutinador(iFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 20848
    
    'Reinicializar saldos do periodo
    lErro = Reinicializa_Saldos_Reprocessa(iFilialEmpresa, iExercicio, iPeriodo, tFaixaPeriodo)
    If lErro <> SUCESSO Then Error 5297

    'Processa os lançamentos do periodo
    lErro = Reprocessa_Lancamentos(iFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 5282

    'Atualiza o periodofilial indicando que não está apurado
    lErro = Comando_ExecutarPos(alComando(5), "UPDATE PeriodosFilial SET Apurado = ?", alComando(4), PERIODO_NAO_APURADO)
    If lErro <> AD_SQL_SUCESSO Then Error 9522

    'Atualiza o ExercicioFilial indicando que está não está apurado
    lErro = Comando_ExecutarPos(alComando(7), "UPDATE ExerciciosFilial SET Status = ?", alComando(2), EXERCICIO_ABERTO)
    If lErro <> AD_SQL_SUCESSO Then Error 10703
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Rotina_Reprocessamento0 = SUCESSO

    Exit Function

Erro_Rotina_Reprocessamento0:

    Rotina_Reprocessamento0 = Err

    Select Case Err

        Case 5272
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", lErro)

        Case 5273, 5274
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO", lErro, iExercicio)

        Case 5275
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_FECHADO", lErro, iExercicio)

        Case 5276
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIO", lErro, iExercicio)

        Case 5278, 5279
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOS", lErro, iExercicio, iPeriodo)

        Case 5280
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODO", lErro, iExercicio, iPeriodo)

        Case 5282, 5297

        Case 5283
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", lErro)

        Case 9522
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PERIODOSFILIAL", lErro, iPeriodo, iExercicio, iFilialEmpresa)

        Case 10697, 10698
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL", lErro, iFilialEmpresa, iExercicio)

        Case 10699
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOSFILIAL", lErro, iFilialEmpresa, iExercicio)

        Case 10700, 10701
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL", lErro, iFilialEmpresa, iExercicio, iPeriodo)

        Case 10702
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODOSFILIAL", lErro, iFilialEmpresa, iExercicio, iPeriodo)

        Case 10703
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOSFILIAL", lErro, iExercicio, iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 154937)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Inicializa_Movimento_Dia_Ccl(ByVal iFilialEmpresa As Integer, tFaixaPeriodo As typeFaixaPeriodo) As Long

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim sConta As String
Dim sCcl As String
Dim dtData As Date
Dim dtData1 As Date
Dim dDebito As Double
Dim dCredito As Double

On Error GoTo Erro_Inicializa_Movimento_Dia_Ccl

    lComando = 0
    lComando1 = 0
    lComando2 = 0
    lComando3 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5327

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5328

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10719

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10720

    sConta = String(STRING_CONTA, 0)
    sCcl = String(STRING_CCL, 0)

    'Pesquisa todos os saldos do ccl
    lErro = Comando_ExecutarPos(lComando, "SELECT Ccl, Conta, Data, Deb, Cre FROM MvDiaCcl WHERE FilialEmpresa=? AND Data >= ? AND Data <= ? ", 0, sCcl, sConta, dtData, dDebito, dCredito, iFilialEmpresa, tFaixaPeriodo.dtDataInicio, tFaixaPeriodo.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then Error 5329

    'Le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9510

    Do While lErro = AD_SQL_SUCESSO

        'zera o saldo do ccl no dia que será reprocessado
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MvDiaCcl SET Deb = 0, Cre = 0", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 5331

        If giFilialOrigem <> EMPRESA_TODA And giFilialOrigem <> Abs(giFilialAuxiliar) Then

            'Pesquisa o saldo de centro de custo diario no ambito empresa
            lErro = Comando_ExecutarPos(lComando2, "SELECT Data FROM MvDiaCcl WHERE FilialEmpresa=? AND Ccl=? AND Conta=? AND Data = ?", 0, dtData1, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl, sConta, dtData)
            If lErro <> AD_SQL_SUCESSO Then Error 10721
    
            'Le o saldo diario no ambito empresa
            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10722
    
            'atualiza o saldo diario do ccl no ambito empresa
            lErro = Comando_ExecutarPos(lComando3, "UPDATE MvDiaCcl SET Deb = Deb - ?, Cre = Cre - ?", lComando2, dDebito, dCredito)
            If lErro <> AD_SQL_SUCESSO Then Error 10723

        End If

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9511

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Inicializa_Movimento_Dia_Ccl = SUCESSO

    Exit Function

Erro_Inicializa_Movimento_Dia_Ccl:

    Inicializa_Movimento_Dia_Ccl = Err

    Select Case Err

        Case 5327, 5328, 10719, 10720
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5329, 9510, 9511
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACCL", Err)

        Case 5331
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACCL", Err, iFilialEmpresa, sCcl, sConta, dtData)

        Case 10721, 10722
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACCL1", Err, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl, sConta, dtData)

        Case 10723
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACCL", Err, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sCcl, sConta, dtData)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154938)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Inicializa_Movimento_Dia_Conta(ByVal iFilialEmpresa As Integer, tFaixaPeriodo As typeFaixaPeriodo) As Long

Dim lErro As Long
Dim lComando As Long, lComando1 As Long
Dim lComando2 As Long, lComando3 As Long
Dim sConta As String
Dim dtData As Date
Dim dtData1 As Date
Dim dDebito As Double
Dim dCredito As Double

On Error GoTo Erro_Inicializa_Movimento_Dia_Conta

    lComando = 0
    lComando1 = 0
    lComando2 = 0
    lComando3 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5321

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5322

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10714

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10715

    sConta = String(STRING_CONTA, 0)
    
    'Pesquisa todos os saldos de conta
    lErro = Comando_ExecutarPos(lComando, "SELECT Conta, Data, Deb, Cre FROM MvDiaCta WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? ", 0, sConta, dtData, dDebito, dCredito, iFilialEmpresa, tFaixaPeriodo.dtDataInicio, tFaixaPeriodo.dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then Error 5323

    'Le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9512

    Do While lErro = AD_SQL_SUCESSO

        'zera o saldo da conta no periodo que será reprocessado
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MvDiaCta SET Deb = 0, Cre = 0", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 5325

        If giFilialOrigem <> EMPRESA_TODA And giFilialOrigem <> Abs(giFilialAuxiliar) Then

            'Pesquisa o saldo diario da conta no ambito empresa
            lErro = Comando_ExecutarPos(lComando2, "SELECT Data FROM MvDiaCta WHERE FilialEmpresa = ? AND Conta = ? AND Data = ?", 0, dtData1, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta, dtData)
            If lErro <> AD_SQL_SUCESSO Then Error 10716
    
            'Le o saldo diario no ambito empresa
            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10717
    
            'atualiza o saldo diario da conta no ambito empresa
            lErro = Comando_ExecutarPos(lComando3, "UPDATE MvDiaCta SET Deb = Deb - ?, Cre = Cre - ?", lComando2, dDebito, dCredito)
            If lErro <> AD_SQL_SUCESSO Then Error 10718

        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9513

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Inicializa_Movimento_Dia_Conta = SUCESSO

    Exit Function

Erro_Inicializa_Movimento_Dia_Conta:

    Inicializa_Movimento_Dia_Conta = Err

    Select Case Err

        Case 5321, 5322, 10714, 10715
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5323, 9512, 9513
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACTA", Err)

        Case 5325
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACTA", Err, iFilialEmpresa, sConta, CStr(dtData))

        Case 10716, 10717
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACTA1", Err, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta, CStr(dtData))

        Case 10718
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACTA", Err, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), sConta, CStr(dtData))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154939)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Inicializa_Movimento_Periodo_Ccl(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer) As Long

Dim lErro As Long
Dim sPeriodo As String
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim sConta As String
Dim sCcl As String
Dim dDebito As Double
Dim dCredito As Double
Dim iExercicio1 As Integer

On Error GoTo Erro_Inicializa_Movimento_Periodo_Ccl

    lComando = 0
    lComando1 = 0
    lComando2 = 0
    lComando3 = 0


    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5309

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5310

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10709

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10710

    sPeriodo = Format(iPeriodo, "00")
    
    sConta = String(STRING_CONTA, 0)
    sCcl = String(STRING_CCL, 0)

    'Pesquisa todos os saldos de centro de custo/lucro
    lErro = Comando_ExecutarPos(lComando, "SELECT Ccl, Conta, Deb" + sPeriodo + ", Cre" + sPeriodo + "  FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND (Deb" + sPeriodo + " <> 0 OR Cre" + sPeriodo + " <> 0)", 0, sCcl, sConta, dDebito, dCredito, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5311

    'Le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9514

    Do While lErro = AD_SQL_SUCESSO

        'zera o saldo da centro de custo no periodo que será reprocessado
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MvPerCcl SET Deb" + sPeriodo + " = 0, Cre" + sPeriodo + " = 0", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 5313

        If giFilialOrigem <> EMPRESA_TODA And giFilialOrigem <> Abs(giFilialAuxiliar) Then

            'Pesquisa o saldo de centro de custo da empresa
            lErro = Comando_ExecutarPos(lComando2, "SELECT Exercicio FROM MvPerCcl WHERE FilialEmpresa =? AND Exercicio = ? AND Ccl = ? AND Conta = ?", 0, iExercicio1, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio, sCcl, sConta)
            If lErro <> AD_SQL_SUCESSO Then Error 10711
    
            'Le o saldo da empresa
            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10712
    
            'atualiza o saldo do centro de custo no ambito empresa
            lErro = Comando_ExecutarPos(lComando3, "UPDATE MvPerCcl SET Deb" + sPeriodo + " = Deb" + sPeriodo + " - ?, Cre" + sPeriodo + " = Cre" + sPeriodo + " - ?", lComando2, dDebito, dCredito)
            If lErro <> AD_SQL_SUCESSO Then Error 10713
    
        End If
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9515

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Inicializa_Movimento_Periodo_Ccl = SUCESSO

    Exit Function

Erro_Inicializa_Movimento_Periodo_Ccl:

    Inicializa_Movimento_Periodo_Ccl = Err

    Select Case Err

        Case 5309, 5310, 10709, 10710
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5311, 9514, 9515
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", Err)

        Case 5313
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL", Err, iFilialEmpresa, iExercicio, sCcl, sConta)

        Case 10711, 10712
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL1", Err, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio, sCcl, sConta)

        Case 10713
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL", Err, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio, sCcl, sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154940)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Inicializa_Movimento_Periodo_Conta(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer) As Long

Dim lErro As Long
Dim sPeriodo As String
Dim lComando As Long, lComando1 As Long
Dim lComando2 As Long, lComando3 As Long
Dim sConta As String
Dim dCredito As Double
Dim dDebito As Double
Dim iExercicio1 As Integer

On Error GoTo Erro_Inicializa_Movimento_Periodo_Conta

    lComando = 0
    lComando1 = 0
    lComando2 = 0
    lComando3 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5303

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5304

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10704

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10705

    sPeriodo = Format(iPeriodo, "00")

    sConta = String(STRING_CONTA, 0)

    'Pesquisa todos os saldos de conta
    lErro = Comando_ExecutarPos(lComando, "SELECT Conta, Deb" + sPeriodo + ", Cre" + sPeriodo + " FROM MvPerCta WHERE FilialEmpresa =? AND Exercicio = ?", 0, sConta, dDebito, dCredito, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5305

    'Le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9516

    Do While lErro = AD_SQL_SUCESSO

        'zera o saldo da conta no periodo que será reprocessado
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MvPerCta SET Deb" + sPeriodo + " = 0, Cre" + sPeriodo + " = 0", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 5307

        If giFilialOrigem <> EMPRESA_TODA And giFilialOrigem <> Abs(giFilialAuxiliar) Then

            'Pesquisa o saldo de conta da empresa
            lErro = Comando_ExecutarPos(lComando2, "SELECT Exercicio FROM MvPerCta WHERE FilialEmpresa =? AND Exercicio = ? And Conta = ?", 0, iExercicio1, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio, sConta)
            If lErro <> AD_SQL_SUCESSO Then Error 10706
    
            'Le o saldo da empresa
            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10707
    
            'atualiza o saldo da conta no ambito empresa
            lErro = Comando_ExecutarPos(lComando3, "UPDATE MvPerCta SET Deb" + sPeriodo + " = Deb" + sPeriodo + " - ?, Cre" + sPeriodo + " = Cre" + sPeriodo + " - ?", lComando2, dDebito, dCredito)
            If lErro <> AD_SQL_SUCESSO Then Error 10708

        End If

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9517

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Inicializa_Movimento_Periodo_Conta = SUCESSO

    Exit Function

Erro_Inicializa_Movimento_Periodo_Conta:

    Inicializa_Movimento_Periodo_Conta = Err

    Select Case Err

        Case 5303, 5304, 10704, 10705
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5305, 9516, 9517
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", Err)

        Case 5307
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCTA", Err, iFilialEmpresa, iExercicio, sConta)

        Case 10706, 10707
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA1", Err, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio, sConta)

        Case 10708
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCTA", Err, IIf(iFilialEmpresa > Abs(giFilialAuxiliar), Abs(giFilialAuxiliar), EMPRESA_TODA), iExercicio, sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154941)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Reinicializa_Saldos_Reprocessa(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer, tFaixaPeriodo As typeFaixaPeriodo) As Long

Dim lErro As Long

On Error GoTo Erro_Reinicializa_Saldos_Reprocessa

    lErro = Inicializa_Movimento_Periodo_Conta(iFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> SUCESSO Then gError 5298

    lErro = Inicializa_Movimento_Periodo_Ccl(iFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> SUCESSO Then gError 5299

    lErro = Inicializa_Movimento_Dia_Conta(iFilialEmpresa, tFaixaPeriodo)
    If lErro <> SUCESSO Then gError 5301

    lErro = Inicializa_Movimento_Dia_Ccl(iFilialEmpresa, tFaixaPeriodo)
    If lErro <> SUCESSO Then gError 5302

    Reinicializa_Saldos_Reprocessa = SUCESSO

    Exit Function

Erro_Reinicializa_Saldos_Reprocessa:

    Reinicializa_Saldos_Reprocessa = gErr

    Select Case gErr

        Case 5298, 5299, 5301, 5302

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154942)

    End Select

    Exit Function

End Function

Function Reprocessa_Lancamentos(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer) As Long

Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_Reprocessa_Lancamentos

    lComando = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5284

'    'Cria um indice na tabela de Lancamento cuja chave seja (Exercicio, Periodo, Conta, Data)
'    lErro = Comando_Executar(lComando, "CREATE INDEX indice_lancamento1 ON Lancamentos (FilialEmpresa, Exercicio,PeriodoLan,Conta,Data)")
'    If lErro <> AD_SQL_SUCESSO Then Error 5285
    
    lErro = Reprocessa_Lancamentos1(iFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> SUCESSO Then Error 5286

'    'Remove o indice
'    lErro = Comando_Executar(lComando, "DROP INDEX Lancamentos.indice_lancamento1")
'    If lErro <> AD_SQL_SUCESSO Then Error 5296

    Call Comando_Fechar(lComando)

    Reprocessa_Lancamentos = SUCESSO

    Exit Function

Erro_Reprocessa_Lancamentos:

    Reprocessa_Lancamentos = Err

    Select Case Err

        Case 5284, 5315, 5318
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5285, 5316, 5319
            Call Rotina_Erro(vbOKOnly, "ERRO_CRIACAO_INDICE", Err)

        Case 5286

        Case 5296, 5317, 5320
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_INDICE", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154943)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Reprocessa_Lancamentos1(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer) As Long

Dim lErro As Long
Dim lPosicao As Long
Dim tLancamento As typeLancamento
Dim lComando1 As Long
Dim lComando2 As Long
Dim tProcessa_Lancamento As typeProcessa_Lancamento
Dim tLancamento_Sort As typeLancamento_Sort

On Error GoTo Erro_Reprocessa_Lancamentos1

    lComando1 = 0
    lComando2 = 0

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9528

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5287

    lErro = Comando_Executar(lComando1, "SELECT UsoCcl FROM Configuracao", tProcessa_Lancamento.iUsoCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9529
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 9530

    tProcessa_Lancamento.iFilialEmpresa = iFilialEmpresa
    tProcessa_Lancamento.iOperacao = ATUALIZACAO
    tProcessa_Lancamento.iOperacao1 = ROTINA_REPROCESSAMENTO_BATCH
    tProcessa_Lancamento.iPeriodo = iPeriodo
    tProcessa_Lancamento.sPeriodo = Format(iPeriodo, "00")
    tProcessa_Lancamento.iExercicio = iExercicio
    tProcessa_Lancamento.tLancamento.sConta = String(STRING_CONTA, 0)
    tProcessa_Lancamento.tLancamento.sCcl = String(STRING_CCL, 0)

    'Pesquisa os lançamentos pertencentes ao periodo reprocessado
    lErro = Comando_Executar(lComando2, "SELECT Conta, Ccl, Data, Valor FROM Lancamentos WHERE FilialEmpresa=? AND Exercicio = ? AND PeriodoLan = ? AND (Aglutinado = 0 OR Aglutinado = 1) ORDER BY Conta, Data", tProcessa_Lancamento.tLancamento.sConta, tProcessa_Lancamento.tLancamento.sCcl, tProcessa_Lancamento.tLancamento.dtData, tProcessa_Lancamento.tLancamento.dValor, iFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 5288

    tProcessa_Lancamento.lComando2 = lComando2

    'Le o primeiro lançamento pertencente ao lote
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9518

    If lErro = AD_SQL_SUCESSO Then

        tProcessa_Lancamento.iFim_de_Arquivo = AD_BOOL_TRUE

        'Abre o arquivo que vai armazenar temporariamente os lancamentos do lote
        tProcessa_Lancamento.lID_Arq_Temp = Arq_Temp_Criar(Len(tLancamento_Sort))
        If tProcessa_Lancamento.lID_Arq_Temp = 0 Then Error 5289

        'Abre o arquivo 1 de sort
        tProcessa_Lancamento.lID_Arq_Sort = Sort_Abrir(Len(lPosicao), 1)
        If tProcessa_Lancamento.lID_Arq_Sort = 0 Then Error 5290

        'Abre o arquivo 2 de sort
        tProcessa_Lancamento.lID_Arq_Sort1 = Sort_Abrir(Len(lPosicao), 3)
        If tProcessa_Lancamento.lID_Arq_Sort1 = 0 Then Error 5291

        'Abre o arquivo 3 de sort
        tProcessa_Lancamento.lID_Arq_Sort2 = Sort_Abrir(Len(lPosicao), 2)
        If tProcessa_Lancamento.lID_Arq_Sort2 = 0 Then Error 10724

        'Processa a atualização dos Saldos de Contas Por Dia
        lErro = Processa_Lancamento_ContaDia(tProcessa_Lancamento)
        If lErro <> SUCESSO Then Error 5293

        'Prepara o arquivo temporário para leitura
        lErro = Arq_Temp_Preparar(tProcessa_Lancamento.lID_Arq_Temp)
        If lErro <> AD_BOOL_TRUE Then Error 5292

        'Processa a atualização dos Saldos de Contas Por Mes
        lErro = Processa_Lancamento_ContaMes(tProcessa_Lancamento)
        If lErro <> SUCESSO Then Error 5294

        'se o sistema usa centro de custo/lucro contabil ou extra-contábil, computa os saldos correspondentes
        If tProcessa_Lancamento.iUsoCcl = CCL_USA_CONTABIL Or tProcessa_Lancamento.iUsoCcl = CCL_USA_EXTRACONTABIL Then
            
            'Processa a atualização dos Saldos de Ccl Por Mes
            lErro = Processa_Lancamento_CclMes(tProcessa_Lancamento)
            If lErro <> SUCESSO Then Error 5295
            
            lErro = Processa_Lancamento_CclDia(tProcessa_Lancamento)
            If lErro <> SUCESSO Then Error 10725
            
        End If

        Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)
        Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort)
        Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort1)
        Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort2)

    End If

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    tProcessa_Lancamento.lComando2 = 0
    Reprocessa_Lancamentos1 = SUCESSO

    Exit Function

Erro_Reprocessa_Lancamentos1:

    Reprocessa_Lancamentos1 = Err

    Select Case Err

        Case 5287
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5288, 9518
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS1", Err, iExercicio, iPeriodo)

        Case 5289
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_ARQUIVO_TEMPORARIO", Err)

        Case 5290
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_ARQUIVO_SORT", Err)
            Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)

        Case 5291
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_ARQUIVO_SORT", Err)
            Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort)

        Case 5292
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_ARQUIVO_TEMP", Err)
            Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort1)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort2)

        Case 5293, 5294, 5295, 10725
            Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort1)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort2)

        Case 9529, 9530
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
        
        Case 10724
            Call Arq_Temp_Destruir(tProcessa_Lancamento.lID_Arq_Temp)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort)
            Call Sort_Destruir(tProcessa_Lancamento.lID_Arq_Sort1)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154944)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function LctosAglutinadores_AcertaValor(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer) As Long
'onde for necessario, acerta o valor dos lancamentos aglutinadores de forma que seu valor seja igual à soma dos valores dos lancamentos aglutinados à ele associados

Dim lErro As Long, iIndice As Integer, dValor As Double, dValorCalc As Double
Dim alComando(1 To 3) As Long, sOrigem As String, lDoc As Long, iSeq As Integer

On Error GoTo Erro_LctosAglutinadores_AcertaValor

    For iIndice = LBound(alComando) To UBound(alComando)
        
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 106581
        
    Next
    
    sOrigem = String(STRING_ORIGEM, 0)
    
    'pesquisa lancamentos aglutinadores com erro de valor
    lErro = Comando_Executar(alComando(1), "SELECT Origem, Doc, Seq, ValorCalc FROM LctoAglutConfAux2 WHERE FilialEmpresa = ? AND Exercicio = ? AND PeriodoLan = ?", _
        sOrigem, lDoc, iSeq, dValorCalc, iFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then gError 106582
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106583
    
    Do While lErro = AD_SQL_SUCESSO
        
        lErro = Comando_ExecutarPos(alComando(2), "SELECT Valor FROM Lancamentos WHERE FilialEmpresa = ? AND Exercicio = ? AND PeriodoLan = ? AND Origem = ? AND Doc = ? AND Seq = ?", 0, _
            dValor, iFilialEmpresa, iExercicio, iPeriodo, sOrigem, lDoc, iSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 106585
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106586
        
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 106587
        
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE Lancamentos SET Valor = ?", alComando(2), dValorCalc)
        If lErro <> AD_SQL_SUCESSO Then gError 106588
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106584
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    LctosAglutinadores_AcertaValor = SUCESSO
     
    Exit Function
    
Erro_LctosAglutinadores_AcertaValor:

    LctosAglutinadores_AcertaValor = gErr
     
    Select Case gErr
          
        Case 106581
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 106582, 106583, 106584
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANC_AGLUT_COM_ERRO", gErr)
        
        Case 106585, 106586
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)
        
        Case 106587
            Call Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIO_LANCAMENTO_CTB", gErr)
        
        Case 106588
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LANCAMENTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154945)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Zera_Saldos_Empresa_Toda(ByVal iExercicio As Integer, ByVal iPeriodo As Integer, tFaixaPeriodo As typeFaixaPeriodo) As Long
'se estiver trabalhando a nivel de empresa toda ==> zera os saldos de conta e ccl neste nivel

Dim lErro As Long
Dim sPeriodo As String
Dim lComando As Long, lComando1 As Long
Dim lComando2 As Long, lComando3 As Long
Dim sConta As String
Dim dCredito As Double
Dim dDebito As Double
Dim iExercicio1 As Integer

On Error GoTo Erro_Zera_Saldos_Empresa_Toda

    If giFilialOrigem = EMPRESA_TODA Or giFilialOrigem = Abs(giFilialAuxiliar) Then

        lComando = Comando_Abrir()
        If lComando = 0 Then gError 188331
    
        sPeriodo = Format(iPeriodo, "00")
    
        sConta = String(STRING_CONTA, 0)
    
        'zera o saldo da conta no periodo que será reprocessado
        lErro = Comando_Executar(lComando, "UPDATE MvPerCta SET Deb" + sPeriodo + " = 0, Cre" + sPeriodo + " = 0 WHERE (FilialEmpresa=? Or FilialEmpresa=?) AND Exercicio=?  AND (Deb" + sPeriodo + " <> 0 OR Cre" + sPeriodo + " <> 0)", EMPRESA_TODA, Abs(giFilialAuxiliar), iExercicio)
        If lErro <> AD_SQL_SUCESSO Then gError 188332
    
        'zera o saldo da ccl no periodo que será reprocessado
        lErro = Comando_Executar(lComando, "UPDATE MvPerCcl SET Deb" + sPeriodo + " = 0, Cre" + sPeriodo + " = 0 WHERE (FilialEmpresa=? Or FilialEmpresa=?) AND Exercicio=?  AND (Deb" + sPeriodo + " <> 0 OR Cre" + sPeriodo + " <> 0)", EMPRESA_TODA, Abs(giFilialAuxiliar), iExercicio)
        If lErro <> AD_SQL_SUCESSO Then gError 188333
    
        lErro = Comando_Executar(lComando, "UPDATE MvDiaCta SET Deb = 0, Cre = 0 WHERE (FilialEmpresa=? Or FilialEmpresa=?) AND Data >= ? AND Data <= ?", EMPRESA_TODA, Abs(giFilialAuxiliar), tFaixaPeriodo.dtDataInicio, tFaixaPeriodo.dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 188334
    
        'zera o saldo do ccl no dia que será reprocessado
        lErro = Comando_Executar(lComando, "UPDATE MvDiaCcl SET Deb = 0, Cre = 0 WHERE (FilialEmpresa=? Or FilialEmpresa=?) AND Data >= ? AND Data <= ?", EMPRESA_TODA, Abs(giFilialAuxiliar), tFaixaPeriodo.dtDataInicio, tFaixaPeriodo.dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 188335
    
        Call Comando_Fechar(lComando)
    
    End If
    
    Zera_Saldos_Empresa_Toda = SUCESSO

    Exit Function

Erro_Zera_Saldos_Empresa_Toda:

    Zera_Saldos_Empresa_Toda = gErr

    Select Case gErr

        Case 188331
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 188332
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCTA", gErr, EMPRESA_TODA, iExercicio)

        Case 188333
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL", Err, EMPRESA_TODA, iExercicio)

        Case 188334
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACTA", Err, EMPRESA_TODA)

        Case 188335
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVDIACCL", Err, EMPRESA_TODA)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188336)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Private Function LctosAglutinados_Sem_Aglutinador(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo As Integer) As Long
'onde for necessario, acerta o valor dos lancamentos aglutinadores de forma que seu valor seja igual à soma dos valores dos lancamentos aglutinados à ele associados

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 6) As Long

Dim tLancamento As typeLancamento
Dim dtData As Date
Dim sOrigem As String
Dim sConta As String
Dim sCcl As String
Dim lDocAglutinado As Long
Dim objLancamento As New ClassLancamentos
Dim iGerencial As Integer
Dim dValorAglutinado As Double
Dim iUltSeqAglutinado As Integer
Dim lDocAux As Long, lDoc As Long

On Error GoTo Erro_LctosAglutinados_Sem_Aglutinador

    For iIndice = LBound(alComando) To UBound(alComando)
        
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 203269
        
    Next
    
    'Preenche o Histórico do Lançamento Aglutinado
    lErro = CF("Preenche_Historico_Lanc_Aglutinado", objLancamento)
    If lErro <> SUCESSO Then gError 203270
    
    tLancamento.sOrigem = String(STRING_ORIGEM, 0)
    tLancamento.sConta = String(STRING_CONTA, 0)
    tLancamento.sCcl = String(STRING_CCL, 0)
    
    'pesquisa lancamentos aglutinados sem lancamento aglutinador associado
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Origem, Conta, Ccl, Data, Valor, Gerencial, Doc FROM Lancamentos WHERE FilialEmpresa = ? AND Exercicio = ? AND PeriodoLan = ? AND Aglutinado = 1 AND DocAglutinado = -1 AND Conta liKE '118010134%' ORDER BY Origem, Data, Conta, Ccl, Gerencial ", 0, _
        tLancamento.sOrigem, tLancamento.sConta, tLancamento.sCcl, tLancamento.dtData, tLancamento.dValor, tLancamento.iGerencial, tLancamento.lDoc, iFilialEmpresa, iExercicio, iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then gError 203271
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203272
    
    dtData = tLancamento.dtData
    sOrigem = tLancamento.sOrigem
    sConta = tLancamento.sConta
    sCcl = tLancamento.sCcl
    iGerencial = tLancamento.iGerencial
    lDoc = tLancamento.lDoc
    
    iUltSeqAglutinado = 1
    dValorAglutinado = 0
    
    If lErro = AD_SQL_SUCESSO Then
    
        'pega o número do proximo voucher(documento) disponível
        lErro = CF("Voucher_Automatico1", iFilialEmpresa, iExercicio, iPeriodo, gcolModulo_sOrigemAglutina(tLancamento.sOrigem), lDocAglutinado)
        If lErro <> SUCESSO Then gError 203273
    
    End If
    
    Do While lErro = AD_SQL_SUCESSO
        
        
        If dtData <> tLancamento.dtData Or sOrigem <> tLancamento.sOrigem Or tLancamento.sConta <> sConta Or tLancamento.sCcl <> sCcl Or iGerencial <> tLancamento.iGerencial Then
        
            'Se o lançamento aglutinado não existe ==> insere o lançamento aglutinado
            lErro = Comando_Executar(alComando(2), "INSERT INTO Lancamentos (FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Seq,Lote,PeriodoLote,Data,Conta,Ccl,Valor,DocAglutinado, SeqAglutinado, Aglutinado, Historico, Gerencial) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            iFilialEmpresa, gcolModulo_sOrigemAglutina(sOrigem), iExercicio, iPeriodo, lDocAglutinado, iUltSeqAglutinado, 0, iPeriodo, dtData, sConta, sCcl, dValorAglutinado, lDocAglutinado, iUltSeqAglutinado, LANCAMENTO_AGLUTINADO, objLancamento.sHistorico, iGerencial)
            If lErro <> AD_SQL_SUCESSO Then gError 203274
        
            dValorAglutinado = 0
    
        
            If dtData <> tLancamento.dtData Or gcolModulo_sOrigemAglutina(sOrigem) <> gcolModulo_sOrigemAglutina(tLancamento.sOrigem) Then
                    
                'Pesquisa o lançamento aglutinado da data em questao com o maior sequencial
                lErro = Comando_ExecutarPos(alComando(4), "SELECT Doc, Seq FROM Lancamentos WHERE FilialEmpresa = ? AND Data = ? AND Origem = ? AND Aglutinado =2 ORDER BY Seq DESC", 0, lDocAglutinado, iUltSeqAglutinado, iFilialEmpresa, tLancamento.dtData, gcolModulo_sOrigemAglutina(tLancamento.sOrigem))
                If lErro <> AD_SQL_SUCESSO Then gError 203280
                
                'Le o lançamento aglutinado com sequencial maior
                lErro = Comando_BuscarPrimeiro(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203281
    
            
                If lErro = AD_SQL_SEM_DADOS Then
            
                    'mostra número do proximo voucher(documento) disponível
                    lErro = CF("Voucher_Automatico1", iFilialEmpresa, iExercicio, iPeriodo, gcolModulo_sOrigemAglutina(tLancamento.sOrigem), lDocAglutinado)
                    If lErro <> SUCESSO Then gError 203275
                
                    iUltSeqAglutinado = 0
                    
                End If
                
                dtData = tLancamento.dtData
                sOrigem = tLancamento.sOrigem
                
            End If
        
            sConta = tLancamento.sConta
            sCcl = tLancamento.sCcl
            iUltSeqAglutinado = iUltSeqAglutinado + 1
            iGerencial = tLancamento.iGerencial
        
        
        End If
        
        dValorAglutinado = dValorAglutinado + tLancamento.dValor
    
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE Lancamentos SET DocAglutinado = ?, SeqAglutinado = ?", alComando(1), lDocAglutinado, iUltSeqAglutinado)
        If lErro <> AD_SQL_SUCESSO Then gError 203276
        
        'pesquisa os lancamentos nao aglutinados do doc em questao para colocar nos sequenciais seguintes do documento aglutinador
        lErro = Comando_ExecutarPos(alComando(5), "SELECT Doc FROM Lancamentos AS L WHERE L.FilialEmpresa = ? AND L.Exercicio = ? AND L.PeriodoLan = ? AND L.Doc = ? AND L.Aglutinado = 0 AND Origem = ? ORDER BY Seq ", 0, _
            lDocAux, iFilialEmpresa, iExercicio, iPeriodo, lDoc, sOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 203271
        
        lErro = Comando_BuscarPrimeiro(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203272
        
        Do While lErro = AD_SQL_SUCESSO
        
            lErro = Comando_ExecutarPos(alComando(6), "UPDATE Lancamentos SET DocAglutinado = ?, SeqAglutinado = ?", alComando(5), lDocAglutinado, iUltSeqAglutinado)
            If lErro <> AD_SQL_SUCESSO Then gError 203276
        
            iUltSeqAglutinado = iUltSeqAglutinado + 1
        
            lErro = Comando_BuscarProximo(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203272
        
        Loop
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 203277
    
    Loop
    
    If Len(Trim(sConta)) > 0 Then
    
        'Se o lançamento aglutinado não existe ==> insere o lançamento aglutinado
        lErro = Comando_Executar(alComando(2), "INSERT INTO Lancamentos (FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Seq,Lote,PeriodoLote,Data,Conta,Ccl,Valor,DocAglutinado, SeqAglutinado, Aglutinado, Historico, Gerencial, Usuario,DataRegistro,HoraRegistro,SubTipo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        iFilialEmpresa, gcolModulo_sOrigemAglutina(sOrigem), iExercicio, iPeriodo, lDocAglutinado, iUltSeqAglutinado, 0, iPeriodo, dtData, sConta, sCcl, dValorAglutinado, lDocAglutinado, iUltSeqAglutinado, LANCAMENTO_AGLUTINADO, objLancamento.sHistorico, iGerencial, gsUsuario, Date, CDbl(Time), 0)
        If lErro <> AD_SQL_SUCESSO Then gError 203278
    
    End If
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    LctosAglutinados_Sem_Aglutinador = SUCESSO
     
    Exit Function
    
Erro_LctosAglutinados_Sem_Aglutinador:

    LctosAglutinados_Sem_Aglutinador = gErr
     
    Select Case gErr
          
        Case 203269
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 203270, 203273, 203275, 203278, 203280, 203281
        
        Case 203271, 203272, 203277
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS3", gErr)
        
        Case 203274
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTO", gErr, iFilialEmpresa, gcolModulo_sOrigemAglutina(sOrigem), iExercicio, iPeriodo, lDocAglutinado, iUltSeqAglutinado)
        
        Case 203276
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LANCAMENTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203279)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function


