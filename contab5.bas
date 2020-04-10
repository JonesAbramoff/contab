Attribute VB_Name = "Module5"
'Reabertura de Exercicio

Option Explicit

Function Rotina_Reabertura_Exercicio_Int(ByVal iExercicio As Integer, sConta_Ativo_Inicial As String, sConta_Ativo_Final As String, sConta_Passivo_Inicial As String, sConta_Passivo_Final As String) As Long

Dim lErro As Long
Dim lComando As Long, lComando1 As Long, lComando2 As Long
Dim iStatus As Integer
Dim lTransacao As Long
Dim tContas As typeContas_Fechamento
Dim vbMesRes As VbMsgBoxResult

On Error GoTo Erro_Rotina_Reabertura_Exercicio_Int

    lComando = 0
    lComando1 = 0
    lComando2 = 0
    lTransacao = 0

    tContas.sConta_Ativo_Inicial = sConta_Ativo_Inicial
    tContas.sConta_Ativo_Final = sConta_Ativo_Final
    tContas.sConta_Passivo_Inicial = sConta_Passivo_Inicial
    tContas.sConta_Passivo_Final = sConta_Passivo_Final

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5158

   'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 5166

    'Pesquisa o Exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT  Status FROM Exercicios WHERE Exercicio = ?", 0, iStatus, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5159

    'le o Exercicio
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5160

    'Se o Exercicio não estiver fechado ==> erro
    If iStatus <> EXERCICIO_FECHADO Then Error 5161

    'loca o Exercicio
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5162

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5362

    'Pesquisa o Exercicio posterior ao em questão e verifica se não está fechado
    lErro = Comando_Executar(lComando1, "SELECT Status FROM Exercicios WHERE Exercicio = ?", iStatus, iExercicio + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 5163

    'le o Exercicio posterior
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 5164

    'se existe um exercicio posterior ao em questão ==> ele não pode estar fechado
    If lErro = AD_SQL_SUCESSO Then

        'Se o Exercicio posterior ao em questao estiver fechado ==> erro
        If iStatus = EXERCICIO_FECHADO Then Error 5165

    End If

    'descobre o total dos registros a serem processados
    lErro = Saldos_Total_Registros(iExercicio + 1)
    If lErro <> SUCESSO Then Error 20357

    'zera os saldos iniciais das contas do exercicio posterior ao em questão
    lErro = Reabertura_Saldos(iExercicio + 1, tContas)
    If lErro <> SUCESSO Then Error 5167

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5169

    'Atualiza o Exercicio indicando que está aberto
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Exercicios SET Status = ?", lComando, EXERCICIO_ABERTO)
    If lErro <> AD_SQL_SUCESSO Then Error 5170

    'Coloca o status dos ExerciciosFilial de todas as filiais deste exercicio com  status aberto
    lErro = Rotina_Reabertura_ExerciciosFilial(iExercicio)
    If lErro <> SUCESSO Then Error 55857

   'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 5171

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Rotina_Reabertura_Exercicio_Int = SUCESSO

    Exit Function

Erro_Rotina_Reabertura_Exercicio_Int:

    Rotina_Reabertura_Exercicio_Int = Err

    Select Case Err

        Case 5158
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5159, 5160
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO", Err, iExercicio)

        Case 5161
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_NAO_FECHADO", Err, iExercicio)

        Case 5162
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCACAO_EXERCICIO", Err, iExercicio)

        Case 5163, 5164
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO", Err, iExercicio + 1)

        Case 5165
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_FECHADO", Err, iExercicio + 1)
        
        Case 5166
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 5167, 20357, 55857

        Case 5169
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5170
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOS", Err, iExercicio)

        Case 5171
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 5362
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154920)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Reabertura_Saldos(ByVal iExercicio As Integer, tContas As typeContas_Fechamento) As Long

Dim lComando As Long, lComando1 As Long, lComando2 As Long, lComando3 As Long
Dim iFilialEmpresa As Integer
Dim lErro As Long
Dim sConta As String
Dim sCcl As String
Dim vbMesRes As VbMsgBoxResult

On Error GoTo Erro_Reabertura_Saldos

    lComando = 0
    lComando1 = 0
    lComando2 = 0
    lComando3 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5172

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5173

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5363

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5364

    sConta = String(STRING_CONTA, 0)

    'seleciona os saldos de conta que serao inicializados
    lErro = Comando_ExecutarPos(lComando2, "SELECT FilialEmpresa, Conta FROM MvPerCta WHERE Exercicio = ?", 0, iFilialEmpresa, sConta, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5365

    'acessa o primeiro saldo de conta
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9425

    Do While lErro = AD_SQL_SUCESSO

        'Zera os saldos de conta
        lErro = Comando_ExecutarPos(lComando, "UPDATE MvPerCta SET SldIni = 0", lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 5174

        'acessa o proximo saldo de conta
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9426

        DoEvents
        
        TelaAcompanhaBatch.dValorAtual = TelaAcompanhaBatch.dValorAtual + 1
        
        TelaAcompanhaBatch.ProgressBar1.Value = CInt((TelaAcompanhaBatch.dValorAtual / TelaAcompanhaBatch.dValorTotal) * 100)

        TelaAcompanhaBatch.TotReg.Caption = CStr(CLng(TelaAcompanhaBatch.TotReg.Caption) + 1)
            
        If TelaAcompanhaBatch.iCancelaBatch = CANCELA_BATCH Then
        
            vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_ATUALIZACAO_LOTES")
            
            If vbMesRes = vbYes Then Error 20358
                
            TelaAcompanhaBatch.iCancelaBatch = 0
                
        End If

    Loop


    sCcl = String(STRING_CCL, 0)
    sConta = String(STRING_CONTA, 0)

    'seleciona os saldos de centro de custo que serao inicializados
    lErro = Comando_ExecutarPos(lComando3, "SELECT FilialEmpresa, Ccl, Conta FROM MvPerCcl WHERE Exercicio = ?", 0, iFilialEmpresa, sCcl, sConta, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5367

    'acessa o primeiro saldo de centro de custo
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9427

    Do While lErro = AD_SQL_SUCESSO

        'Zera os saldos de ccl, conta
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MvPerCcl SET SldIni = 0", lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 5175

        'acessa o proximo saldo de centro de custo
        lErro = Comando_BuscarProximo(lComando3)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9428

        DoEvents
        
        TelaAcompanhaBatch.dValorAtual = TelaAcompanhaBatch.dValorAtual + 1
        
        TelaAcompanhaBatch.ProgressBar1.Value = CInt((TelaAcompanhaBatch.dValorAtual / TelaAcompanhaBatch.dValorTotal) * 100)

        TelaAcompanhaBatch.TotReg.Caption = CStr(CLng(TelaAcompanhaBatch.TotReg.Caption) + 1)
            
        If TelaAcompanhaBatch.iCancelaBatch = CANCELA_BATCH Then
        
            vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_ATUALIZACAO_LOTES")
            
            If vbMesRes = vbYes Then Error 20359
                
            TelaAcompanhaBatch.iCancelaBatch = 0
                
        End If

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Reabertura_Saldos = SUCESSO

    Exit Function

Erro_Reabertura_Saldos:

    Reabertura_Saldos = Err

    Select Case Err

        Case 5172
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5173
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5174
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCTA", Err, iFilialEmpresa, iExercicio, sConta)

        Case 5175
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MVPERCCL", Err, iFilialEmpresa, iExercicio, sCcl, sConta)

        Case 5363
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5364
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5365, 9425, 9426
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA", Err)

        Case 5367, 9427, 9428
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", Err)
            
        Case 20358, 20359

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154921)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Private Function Rotina_Reabertura_ExerciciosFilial(ByVal iExercicio As Integer) As Long
'Coloca o status dos ExerciciosFilial de todas as filiais deste exercicio com  status aberto

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Rotina_Reabertura_ExerciciosFilial

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55851

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 55852

    'Pesquisa os ExerciciosFilial do Exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa FROM ExerciciosFilial WHERE Exercicio = ?", 0, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 55853

    'le o Exercicio em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55854

    Do While lErro = AD_SQL_SUCESSO

        'Atualiza o Exercicio indicando que foi fechado
        lErro = Comando_ExecutarPos(lComando1, "UPDATE ExerciciosFilial SET Status = ?", lComando, EXERCICIO_ABERTO)
        If lErro <> AD_SQL_SUCESSO Then Error 55855

        'le o Exercicio em questão
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55856

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Rotina_Reabertura_ExerciciosFilial = SUCESSO

    Exit Function

Erro_Rotina_Reabertura_ExerciciosFilial:

    Rotina_Reabertura_ExerciciosFilial = Err

    Select Case Err

        Case 55851, 55852
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55853, 55854, 55856
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL1", Err, iExercicio)

        Case 55855
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOSFILIAL", Err, iExercicio, iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154922)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

