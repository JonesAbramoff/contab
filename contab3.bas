Attribute VB_Name = "Module3"
'Fechamento de Exercicio

Option Explicit
  
Function Rotina_Fechamento_Exercicio_Int(ByVal iExercicio As Integer, sConta_Ativo_Inicial As String, sConta_Ativo_Final As String, sConta_Passivo_Inicial As String, sConta_Passivo_Final As String) As Long

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim iStatus As Integer
Dim lTransacao As Long
Dim iNumPeriodos As Integer
Dim tContas As typeContas_Fechamento

On Error GoTo Erro_Rotina_Fechamento_Exercicio_Int

    lComando = 0
    lComando1 = 0
    lComando2 = 0
    lComando3 = 0
    lTransacao = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5112

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5355

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5125
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10692
    
    tContas.sConta_Ativo_Inicial = sConta_Ativo_Inicial
    tContas.sConta_Ativo_Final = sConta_Ativo_Final
    tContas.sConta_Passivo_Inicial = sConta_Passivo_Inicial
    tContas.sConta_Passivo_Final = sConta_Passivo_Final

   'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 5123

    'Pesquisa o Exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT NumPeriodos, Status FROM Exercicios WHERE Exercicio = ?", 0, iNumPeriodos, iStatus, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5113

    'le o Exercicio em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5114

    'Se o Exercicio em questão estiver fechado ==> erro
    If iStatus = EXERCICIO_FECHADO Then Error 5115

    'loca o Exercicio em questão
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5116

    'Pesquisa o Exercicio anterior ao em questão
    lErro = Comando_Executar(lComando1, "SELECT Status FROM Exercicios WHERE Exercicio = ?", iStatus, iExercicio - 1)
    If lErro <> AD_SQL_SUCESSO Then Error 5117

    'le o Exercicio anterior
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5118

    'Se existir o exercicio anterior
    If lErro = AD_SQL_SUCESSO Then
        'Se o status do Exercicio anterior for diferente de fechado ==> erro
        If iStatus <> EXERCICIO_FECHADO Then Error 5119
        
    End If

    'Pesquisa o Exercicio posterior ao em questão
    lErro = Comando_Executar(lComando1, "SELECT Status FROM Exercicios WHERE Exercicio = ?", iStatus, iExercicio + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 5120

    'le o Exercicio posterior
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5121

    'se o exercicio posterior ainda não tiver sido criado, solicitar sua criação antes do fechamento do exercicio em questão
    If lErro = AD_SQL_SEM_DADOS Then Error 9501

    'Se existir exercicio posterior
    If lErro = AD_SQL_SUCESSO Then
        'Se o status do Exercicio posterior for fechado ==> erro
        If iStatus = EXERCICIO_FECHADO Then Error 5122
        
    End If

    'verifica se o exercicio foi apurado para todas as filiais da empresa
    lErro = Verifica_Apuracao_Exercicio(lComando3, iExercicio)
    If lErro <> SUCESSO Then Error 10685

    'descobre o total dos registros a serem processados
    lErro = Saldos_Total_Registros(iExercicio)
    If lErro <> SUCESSO Then Error 20353

    'transfere os saldos de conta para o proximo exercicio
    lErro = Transfere_Saldos_Conta_Fechamento(iExercicio, tContas, iNumPeriodos)
    If lErro <> SUCESSO Then Error 5124

    'transfere os saldos de ccl para o proximo exercicio
    lErro = Transfere_Saldos_Ccl_Fechamento(iExercicio, tContas, iNumPeriodos)
    If lErro <> SUCESSO Then Error 5128

    'Atualiza o Exercicio indicando que foi fechado
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Exercicios SET Status = ?", lComando, EXERCICIO_FECHADO)
    If lErro <> AD_SQL_SUCESSO Then Error 5126

    'Coloca o status dos ExerciciosFilial de todas as filiais deste exercicio com  status fechado
    lErro = Rotina_Fechamento_ExerciciosFilial(iExercicio)
    If lErro <> SUCESSO Then Error 55850

   'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 5127

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Rotina_Fechamento_Exercicio_Int = SUCESSO
    
    'Alteracao Daniel em 07/05/02
    Call Rotina_Aviso(vbOKOnly, "AVISO_FECHAMENTO_EXERCICIO_EXECUTADO", iExercicio)

    Exit Function

Erro_Rotina_Fechamento_Exercicio_Int:

    Rotina_Fechamento_Exercicio_Int = Err

    Select Case Err

        Case 5112, 10692
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5113, 5114
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO", Err, iExercicio)

        Case 5115
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_FECHADO", Err, iExercicio)

        Case 5116
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCACAO_EXERCICIO", Err, iExercicio)

        Case 5117, 5118
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO", Err, iExercicio - 1)

        Case 5119
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_FECHADO", Err, iExercicio - 1)
        
        Case 5120, 5121
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO", Err, iExercicio + 1)

        Case 5122
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_FECHADO", Err, iExercicio + 1)

        Case 5123
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 5124, 5128, 10685, 20353, 55850

        Case 5125
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5126
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOS", Err, iExercicio)

        Case 5127
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 5355
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 9501
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_POSTERIOR_INEXISTENTE", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154914)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function
    
Function Saldos_Total_Registros(ByVal iExercicio As Integer) As Long

Dim lComando As Long, lComando1 As Long
Dim lNumReg As Long
Dim lNumReg1 As Long
Dim lErro As Long

On Error GoTo Erro_Saldos_Total_Registros

    lComando = 0
    lComando1 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 20347

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 20348

    'Seleciona os centros de custo do exercicio em questao
    lErro = Comando_Executar(lComando, "SELECT Count(*) FROM MvPerCcl WHERE Exercicio = ?", lNumReg, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 20349

    'Le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 20350

    'Seleciona as contas  do exercicio em questao
    lErro = Comando_Executar(lComando1, "SELECT Count(*) FROM MvPerCta WHERE Exercicio = ? ", lNumReg1, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 20351

    'Le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 20352

    TelaAcompanhaBatch.dValorTotal = lNumReg + lNumReg1

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Saldos_Total_Registros = SUCESSO

    Exit Function

Erro_Saldos_Total_Registros:

    Saldos_Total_Registros = Err

    Select Case Err

        Case 20347, 20348
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 20349, 20350
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", Err)

        Case 20351, 20352
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154915)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function


Function Transfere_Saldos_Ccl_Fechamento(ByVal iExercicio As Integer, tContas As typeContas_Fechamento, ByVal iNumPeriodos As Integer) As Long

Dim lComando As Long, lComando1 As Long, lComando2 As Long, lComando3 As Long
Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim iPeriodo As Integer
Dim dSldIni As Double
Dim sConta As String
Dim sCcl As String
Dim iExercicio1 As Integer, lErro As Long
Dim curSaldo As Currency
Dim iFilialEmpresa As Integer
Dim vbMesRes As VbMsgBoxResult

On Error GoTo Erro_Transfere_Saldos_Ccl_Fechamento

    lComando = 0
    lComando1 = 0
    lComando2 = 0
    lComando3 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 5139

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 5138

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 5356

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 5356

    sConta = String(STRING_CONTA, 0)
    sCcl = String(STRING_CCL, 0)



    'Seleciona os centros de custo do exercicio em questao
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Ccl, Conta, SldIni, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cre01, Cre02, Cre03, Cre04, Cre05, Cre06, Cre07, Cre08, Cre09, Cre10, Cre11, Cre12 FROM MvPerCcl WHERE Exercicio = ?", iFilialEmpresa, sCcl, sConta, dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), iExercicio)
    If lErro <> AD_SQL_SUCESSO Then gError 5140

    'Le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9421

    Do While lErro = AD_SQL_SUCESSO

        curSaldo = CCur(dSldIni)

        For iPeriodo = 1 To iNumPeriodos
                
            curSaldo = curSaldo + dCredito(iPeriodo)
            curSaldo = curSaldo - dDebito(iPeriodo)

        Next

        If curSaldo <> 0 Then

            'Seleciona o centro de custo do exercicio seguinte para transferencia dos saldos
            lErro = Comando_ExecutarPos(lComando2, "SELECT Exercicio FROM MvPerCcl WHERE FilialEmpresa=? AND Exercicio = ? AND Ccl = ? AND Conta = ?", 0, iExercicio1, iFilialEmpresa, iExercicio + 1, sCcl, sConta)
            If lErro <> AD_SQL_SUCESSO Then gError 5357

            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 5358

            If lErro = AD_SQL_SEM_DADOS Then

                'Insere o registro de Saldo de Ccl
                lErro = Comando_Executar(lComando3, "INSERT INTO MvPerCcl (FilialEmpresa, Exercicio, Ccl, Conta, SldIni) VALUES (?, ?, ?, ?, ?)", iFilialEmpresa, iExercicio + 1, sCcl, sConta, CDbl(curSaldo))
                If lErro <> AD_SQL_SUCESSO Then gError 188310

            Else
        
                'transfere o saldo para o exercicio seguinte
                lErro = Comando_ExecutarPos(lComando1, "UPDATE MvPerCcl SET SldIni = ?", lComando2, CDbl(curSaldo))
                If lErro <> AD_SQL_SUCESSO Then gError 5141
            
            End If
            
        End If


        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 9422
        
        DoEvents
        
        TelaAcompanhaBatch.dValorAtual = TelaAcompanhaBatch.dValorAtual + 1
        
        TelaAcompanhaBatch.ProgressBar1.Value = CInt((TelaAcompanhaBatch.dValorAtual / TelaAcompanhaBatch.dValorTotal) * 100)

        TelaAcompanhaBatch.TotReg.Caption = CStr(CLng(TelaAcompanhaBatch.TotReg.Caption) + 1)
            
        If TelaAcompanhaBatch.iCancelaBatch = CANCELA_BATCH Then
        
            vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_ATUALIZACAO_LOTES")
            
            If vbMesRes = vbYes Then gError 20354
                
            TelaAcompanhaBatch.iCancelaBatch = 0
                
        End If

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Transfere_Saldos_Ccl_Fechamento = SUCESSO

    Exit Function

Erro_Transfere_Saldos_Ccl_Fechamento:

    Transfere_Saldos_Ccl_Fechamento = gErr

    Select Case gErr

        Case 5138
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 5139
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 5140, 9421, 9422
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", gErr)

        Case 5141
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL", gErr, iFilialEmpresa, iExercicio + 1, sCcl, sConta)

        Case 5356
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 5357, 5358
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL1", gErr, iFilialEmpresa, iExercicio + 1, sCcl, sConta)
            
        Case 20354
            
        Case 188310
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MVPERCCL", gErr, iFilialEmpresa, iExercicio + 1, sCcl, sConta)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154916)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Transfere_Saldos_Conta_Fechamento(ByVal iExercicio As Integer, tContas As typeContas_Fechamento, ByVal iNumPeriodos As Integer) As Long

Dim lComando As Long, lComando1 As Long, lComando2 As Long
Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim iPeriodo As Integer
Dim dSldIni As Double
Dim sConta As String
Dim iExercicio1 As Integer, lErro As Long
Dim curSaldo As Currency
Dim iFilialEmpresa As Integer
Dim vbMesRes As VbMsgBoxResult

On Error GoTo Erro_Transfere_Saldos_Conta_Fechamento

    lComando = 0
    lComando1 = 0
    lComando2 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5134

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5137

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5359
    
    sConta = String(STRING_CONTA, 0)

    'Seleciona as contas  do exercicio em questao
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Conta, SldIni, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cre01, Cre02, Cre03, Cre04, Cre05, Cre06, Cre07, Cre08, Cre09, Cre10, Cre11, Cre12 FROM MvPerCta WHERE Exercicio = ? ", iFilialEmpresa, sConta, dSldIni, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5135

    'Le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9423

    Do While lErro = AD_SQL_SUCESSO

        curSaldo = CCur(dSldIni)

        For iPeriodo = 1 To iNumPeriodos
                
            curSaldo = curSaldo + dCredito(iPeriodo)
            curSaldo = curSaldo - dDebito(iPeriodo)

        Next

        If curSaldo <> 0 Then

            lErro = Comando_ExecutarPos(lComando2, "SELECT Exercicio FROM MvPerCta WHERE FilialEmpresa=? AND Exercicio = ? AND Conta = ?", 0, iExercicio1, iFilialEmpresa, iExercicio + 1, sConta)
            If lErro <> AD_SQL_SUCESSO Then Error 5360

            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO Then Error 5361
        
            'transfere o saldo para o exercicio seguinte
            lErro = Comando_ExecutarPos(lComando1, "UPDATE MvPerCta SET SldIni = ?", lComando2, CDbl(curSaldo))
            If lErro <> AD_SQL_SUCESSO Then Error 5136
            
        End If

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9424

        DoEvents
        
        TelaAcompanhaBatch.dValorAtual = TelaAcompanhaBatch.dValorAtual + 1
        
        TelaAcompanhaBatch.ProgressBar1.Value = CInt((TelaAcompanhaBatch.dValorAtual / TelaAcompanhaBatch.dValorTotal) * 100)

        TelaAcompanhaBatch.TotReg.Caption = CStr(CLng(TelaAcompanhaBatch.TotReg.Caption) + 1)
            
        If TelaAcompanhaBatch.iCancelaBatch = CANCELA_BATCH Then
        
            vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_ATUALIZACAO_LOTES")
            
            If vbMesRes = vbYes Then Error 20355
                
            TelaAcompanhaBatch.iCancelaBatch = 0
                
        End If

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Transfere_Saldos_Conta_Fechamento = SUCESSO

    Exit Function

Erro_Transfere_Saldos_Conta_Fechamento:

    Transfere_Saldos_Conta_Fechamento = Err

    Select Case Err

        Case 5134
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5135, 9423, 9424
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", Err)

        Case 5136
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCTA", Err, iFilialEmpresa, iExercicio + 1, sConta)
            
        Case 5137
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5359
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5360, 5361
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA1", Err, iFilialEmpresa, iExercicio + 1, sConta)

        Case 20355

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154917)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function


Function Verifica_Apuracao_Exercicio(lComando As Long, iExercicio As Integer) As Long
'verifica se o exercicio foi apurado para todas as filiais da empresa
'TEM QUE SER EXECUTADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim iStatus As Integer
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Verifica_Apuracao_Exercicio

    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa, Status FROM ExerciciosFilial WHERE Exercicio = ?", 0, iFilialEmpresa, iStatus, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 10687

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10688
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 10689
        
        If iFilialEmpresa <> EMPRESA_TODA And iFilialEmpresa <> Abs(giFilialAuxiliar) Then
        
            'se o exercicio não estiver apurado ==> erro
            If iStatus <> EXERCICIO_APURADO Then Error 10690
        
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10691
    
    Loop

    Verifica_Apuracao_Exercicio = SUCESSO
    
    Exit Function
    
Erro_Verifica_Apuracao_Exercicio:

    Verifica_Apuracao_Exercicio = Err
    
    Select Case Err
    
        Case 10687, 10688, 10691
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL1", Err, iExercicio)
    
        Case 10689
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOSFILIAL", Err, iFilialEmpresa, iExercicio)
    
        Case 10690
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIOSFILIAL_NAO_APURADO", Err, iFilialEmpresa, iExercicio)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154918)
        
    End Select
    
    Exit Function
    
End Function

Private Function Rotina_Fechamento_ExerciciosFilial(ByVal iExercicio As Integer) As Long
'Coloca o status dos ExerciciosFilial de todas as filiais deste exercicio com  status fechado

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Rotina_Fechamento_ExerciciosFilial

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55844

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 55845

    'Pesquisa os ExerciciosFilial do Exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa FROM ExerciciosFilial WHERE Exercicio = ?", 0, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 55846

    'le o Exercicio em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55847

    Do While lErro = AD_SQL_SUCESSO

        'Atualiza o Exercicio indicando que foi fechado
        lErro = Comando_ExecutarPos(lComando1, "UPDATE ExerciciosFilial SET Status = ?", lComando, EXERCICIO_FECHADO)
        If lErro <> AD_SQL_SUCESSO Then Error 55848

        'le o Exercicio em questão
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55849

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Rotina_Fechamento_ExerciciosFilial = SUCESSO

    Exit Function

Erro_Rotina_Fechamento_ExerciciosFilial:

    Rotina_Fechamento_ExerciciosFilial = Err

    Select Case Err

        Case 55844, 55845
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55846, 55847, 55849
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL1", Err, iExercicio)

        Case 55848
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOSFILIAL", Err, iExercicio, iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154919)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

