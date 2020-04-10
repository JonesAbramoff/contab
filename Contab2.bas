Attribute VB_Name = "Module2"
'Apuração de Exercicio

Option Explicit

Function Rotina_Apura_Exercicio_Int(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iLote As Integer, sHistorico As String, sConta_Resultado As String, colContasApuracao As Collection) As Long
'realiza a apuração do exercicio iExercicio para as receitas e despesas passadas como parametro e gera um lote contendo a conta resultado da apuracao.

Dim objFiliais As AdmFiliais
Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_Rotina_Apura_Exercicio_Int

    lTransacao = 0

   'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 5078

    'Atualiza a informacao sobre o lote que está sendo processado para a Empresa Toda
    lErro = Atualiza_Lote_Empresa(iExercicio, iLote)
    If lErro <> SUCESSO Then Error 20420

    If iFilialEmpresa = EMPRESA_TODA Then

        TelaAcompanhaBatch.dValorTotal = gcolFiliais.Count
    
        'se tiver selecionado a empresa, executa a apuracao para cada filial
        For Each objFiliais In gcolFiliais
    
            If objFiliais.iCodFilial <> EMPRESA_TODA And objFiliais.iCodFilial <> Abs(giFilialAuxiliar) Then
    
                lErro = Rotina_Apura_Exercicio0(objFiliais.iCodFilial, iExercicio, iLote, sHistorico, sConta_Resultado, colContasApuracao)
                If lErro <> SUCESSO Then Error 10668
                
                TelaAcompanhaBatch.dValorAtual = TelaAcompanhaBatch.dValorAtual + 1
        
                TelaAcompanhaBatch.ProgressBar1.Value = CInt((TelaAcompanhaBatch.dValorAtual / TelaAcompanhaBatch.dValorTotal) * 100)
                
            End If
        
        Next
    
    Else
    
        'se tiver decidido apurar somente uma filial
        lErro = Rotina_Apura_Exercicio0(iFilialEmpresa, iExercicio, iLote, sHistorico, sConta_Resultado, colContasApuracao)
        If lErro <> SUCESSO Then Error 10669
    
        TelaAcompanhaBatch.ProgressBar1.Value = 100
        
    End If

   'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 5086

    Rotina_Apura_Exercicio_Int = SUCESSO

    Exit Function

Erro_Rotina_Apura_Exercicio_Int:

    Rotina_Apura_Exercicio_Int = Err

    Select Case Err

        Case 5078
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 5086
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 10668, 10669, 20420
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154903)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Atualiza_Lote_Empresa(ByVal iExercicio As Integer, ByVal iLote As Integer) As Long
'Atualiza a informacao sobre o lote que está sendo processado para a Empresa Toda

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim iLote1 As Integer

On Error GoTo Erro_Atualiza_Lote_Empresa

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 20421

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 20422

    'Pesquisa o ExercicioFilial da EMPRESA TODA
    lErro = Comando_ExecutarPos(lComando, "SELECT LoteApuracao FROM ExerciciosFilial WHERE Exercicio = ? AND (FilialEmpresa=? Or FilialEmpresa=?)", 0, iLote1, iExercicio, EMPRESA_TODA, Abs(giFilialAuxiliar))
    If lErro <> AD_SQL_SUCESSO Then Error 20423

    'le o ExercicioFilial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 20424

    'se o lote desta apuracao for menor ou igual ao cadastrado no BD ==> erro
    If iLote <= iLote1 Then Error 20427

    'lock do ExercicioFilial
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 20425

    'Atualiza o ultimo lote de apuracao da EMPRESA TODA
    lErro = Comando_ExecutarPos(lComando1, "UPDATE ExerciciosFilial SET LoteApuracao = ?", lComando, iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 20426

    'le o ExercicioFilial
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20424

    If lErro = AD_SQL_SUCESSO Then

        'lock do ExercicioFilial
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 20425
    
        'Atualiza o ultimo lote de apuracao da EMPRESA TODA
        lErro = Comando_ExecutarPos(lComando1, "UPDATE ExerciciosFilial SET LoteApuracao = ?", lComando, iLote)
        If lErro <> AD_SQL_SUCESSO Then Error 20426

    End If
    
    Atualiza_Lote_Empresa = SUCESSO

    Exit Function

Erro_Atualiza_Lote_Empresa:

    Atualiza_Lote_Empresa = Err

    Select Case Err

        Case 20421, 20422
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 20423, 20424
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL", Err, EMPRESA_TODA, iExercicio)
            
        Case 20425
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOSFILIAL", Err, EMPRESA_TODA, iExercicio)

        Case 20426
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOSFILIAL", Err, iExercicio, EMPRESA_TODA)

        Case 20427
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTEAPURACAO_JA_UTILIZADO", Err, iLote, iExercicio)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154904)

    End Select

    Exit Function

End Function

Function Rotina_Apura_Exercicio0(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iLote As Integer, sHistorico As String, sConta_Resultado As String, colContasApuracao As Collection) As Long
'realiza a apuração do exercicio iExercicio para as receitas e despesas passadas como parametro e gera um lote contendo a conta resultado da apuracao.

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim tApuracao As typeApuracao
Dim lErro As Long
Dim iStatus As Integer

On Error GoTo Erro_Rotina_Apura_Exercicio0

    tApuracao.iFilialEmpresa = iFilialEmpresa
    tApuracao.iExercicio = iExercicio
    tApuracao.sConta_Resultado = sConta_Resultado
    tApuracao.sHistorico = sHistorico
    Set tApuracao.colContasApuracao = colContasApuracao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5070

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5102

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 9397

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10667

    'Pesquisa o uso do centro de custo/lucro
    lErro = Comando_Executar(lComando2, "SELECT UsoCcl FROM Configuracao", tApuracao.iUsoCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9398
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 9399
    
    'Pesquisa o Exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Status, NumPeriodos, DataFim FROM Exercicios WHERE Exercicio = ?", 0, iStatus, tApuracao.iNumPeriodos, tApuracao.dtData, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5071

    'le o Exercicio
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5072

    'não permite a mudança no status do exercicio
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5074
    
    'Se o Exercicio estiver fechado ==> erro
    If iStatus = EXERCICIO_FECHADO Then Error 5073

    'Pesquisa o ExercicioFilial em questão
    lErro = Comando_ExecutarPos(lComando3, "SELECT Status, LoteApuracao, DocApuracao, ExisteLoteApuracao FROM ExerciciosFilial WHERE Exercicio = ? AND FilialEmpresa=?", 0, tApuracao.iStatusExercicioFilial, tApuracao.iLote, tApuracao.lDoc, tApuracao.iExisteLoteApuracao, iExercicio, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 10664

    'le o ExercicioFilial
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO Then Error 10665

    'lock do ExercicioFilial
    lErro = Comando_LockExclusive(lComando3)
    If lErro <> AD_SQL_SUCESSO Then Error 10666
    
    tApuracao.iPeriodo = tApuracao.iNumPeriodos

    lErro = Rotina_Apura_Exercicio1(tApuracao, iLote)
    If lErro <> SUCESSO Then Error 9477

    'Atualiza o exercicio indicando que foi apurado e armazenando o número do lote referente à ultima apuração
    lErro = Comando_ExecutarPos(lComando1, "UPDATE ExerciciosFilial SET Status = ?, LoteApuracao = ?, DataApuracao=?, DocApuracao=?, ExisteLoteApuracao=?", lComando3, EXERCICIO_APURADO, tApuracao.iLote, Date, tApuracao.lDoc, EXISTE_LOTE_APURACAO_EXERCICIO)
    If lErro <> AD_SQL_SUCESSO Then Error 5103

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Rotina_Apura_Exercicio0 = SUCESSO

    Exit Function

Erro_Rotina_Apura_Exercicio0:

    Rotina_Apura_Exercicio0 = Err

    Select Case Err

        Case 5070, 5102, 9397, 10667
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5071, 5072
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO", Err, iExercicio)

        Case 5073
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_FECHADO", Err, iExercicio)

        Case 5074
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCACAO_EXERCICIO", Err, iExercicio)

        Case 5103
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOSFILIAL", Err, iExercicio, iFilialEmpresa)

        Case 9398, 9399
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
            
        Case 9477
            
        Case 10664, 10665
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL", Err, iFilialEmpresa, iExercicio)
            
        Case 10666
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOSFILIAL", Err, iFilialEmpresa, iExercicio)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154905)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Private Function Rotina_Apura_Exercicio1(tApuracao As typeApuracao, ByVal iLote As Integer) As Long

Dim lComando As Long
Dim lErro As Long
Dim iTipoConta As Integer

On Error GoTo Erro_Rotina_Apura_Exercicio1

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9485

    'le a conta resultado
    lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta FROM PlanoConta WHERE Conta=?", iTipoConta, tApuracao.sConta_Resultado)
    If lErro <> AD_SQL_SUCESSO Then Error 9478
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9479
    
    'lock da conta resultado
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9480

    'se existia um lote anterior de apuração para o exercicio em questão ===> estorna-o
    If tApuracao.iExisteLoteApuracao = EXISTE_LOTE_APURACAO_EXERCICIO Then

        lErro = Estorno_Apuracao_Exercicio(tApuracao)
        If lErro <> SUCESSO Then Error 5229

    End If

    'verifica se os periodos possuem lotes de apuracao.
    'Se possuirem, estorna-os
    lErro = Periodos_Apuracao_Exercicio(tApuracao)
    If lErro <> SUCESSO Then Error 5230

    'guarda o número do lote que vai ser criado
    tApuracao.iLote = iLote

    'processa a criacao do lote de apuracao e sua contabilizacao
    lErro = Processa_Apuracao_Exercicio(tApuracao)
    If lErro <> SUCESSO Then Error 5239

    Call Comando_Fechar(lComando)
    
    Rotina_Apura_Exercicio1 = SUCESSO

    Exit Function

Erro_Rotina_Apura_Exercicio1:

    Rotina_Apura_Exercicio1 = Err

    Select Case Err

        Case 5229, 5230, 5239

        Case 9478, 9479
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)
        
        Case 9480
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, tApuracao.sConta_Resultado)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Atualiza_Lote_Apuracao(tApuracao As typeApuracao) As Long

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tLote As typeLote_batch
Dim tLote1 As typeLote
Dim objExercicio As New ClassExercicio
Dim objPeriodo As New ClassPeriodo

On Error GoTo Erro_Atualiza_Lote_Apuracao

    lComando = 0
    lComando1 = 0
    lComando2 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5105

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9460

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5252

    tLote1.sIdOriginal = String(STRING_IDORIGINAL, 0)

    'Pesquisa o lote no banco de dados
    lErro = Comando_ExecutarPos(lComando, "SELECT TotCre, TotDeb, TotInf, Status, IdOriginal, NumDocInf, NumDocAtual, IdAtualizacao FROM LotePendente WHERE FilialEmpresa=? AND Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote = ?", 0, tLote1.dTotCre, tLote1.dTotDeb, tLote1.dTotInf, tLote1.iStatus, tLote1.sIdOriginal, tLote1.iNumDocInf, tLote1.iNumDocAtual, tLote1.iIDAtualizacao, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 5106
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5107
    
    'loca o lote de apuracao
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5108

    tLote.iFilialEmpresa = tApuracao.iFilialEmpresa
    tLote.sOrigem = tApuracao.sOrigem_Apuracao
    tLote.iExercicio = tApuracao.iExercicio
    tLote.iPeriodo = tApuracao.iPeriodo
    tLote.iLote = tApuracao.iLote
    
    tLote1.iFilialEmpresa = tLote.iFilialEmpresa
    tLote1.sOrigem = tLote.sOrigem
    tLote1.iExercicio = tLote.iExercicio
    tLote1.iPeriodo = tLote.iPeriodo
    tLote1.iLote = tLote.iLote

    'Processa os lançamentos do lote
    lErro = Processa_Lancamentos(tLote, ATUALIZACAO, tApuracao.iUsoCcl, 0)
    If lErro <> SUCESSO Then Error 5251

    'exclui o lote da tabela de lotes pendentes
    lErro = Comando_ExecutarPos(lComando2, "DELETE From LotePendente", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9461

    lErro = Comando_Executar(lComando1, "INSERT INTO Lote (FilialEmpresa, Origem, Exercicio, Periodo, Lote, TotCre, TotDeb, TotInf, Status, IdOriginal, NumDocInf, NumDocAtual, IdAtualizacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", tLote1.iFilialEmpresa, tLote1.sOrigem, tLote1.iExercicio, tLote1.iPeriodo, tLote1.iLote, tLote1.dTotCre, tLote1.dTotDeb, tLote1.dTotInf, LOTE_ATUALIZADO, tLote1.sIdOriginal, tLote1.iNumDocInf, tLote1.iNumDocAtual, tLote1.iIDAtualizacao)
    If lErro <> AD_SQL_SUCESSO Then Error 9462

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Atualiza_Lote_Apuracao = SUCESSO

    Exit Function

Erro_Atualiza_Lote_Apuracao:

    Atualiza_Lote_Apuracao = Err
    lErro = Err

    Call CF("Exercicio_Le", tApuracao.iExercicio, objExercicio)
    Call CF("Periodo_Le_ExercicioPeriodo", tApuracao.iExercicio, tApuracao.iPeriodo, objPeriodo)

    Select Case lErro

        Case 5105, 5252, 9460
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5106, 5107
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE", Err, tApuracao.sOrigem_Apuracao, objExercicio.sNomeExterno, objPeriodo.sNomeExterno, tApuracao.iLote)

        Case 5108
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LOTE", Err, tApuracao.sOrigem_Apuracao, objExercicio.sNomeExterno, objPeriodo.sNomeExterno, tApuracao.iLote)

        Case 5251

        Case 9461
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LOTEPENDENTE", lErro, tApuracao.sOrigem_Apuracao, objExercicio.sNomeExterno, objPeriodo.sNomeExterno, tApuracao.iLote)

        Case 9462
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOTE", lErro, tApuracao.sOrigem_Apuracao, objExercicio.sNomeExterno, objPeriodo.sNomeExterno, tApuracao.iLote)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154906)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Private Function Cria_Lancamentos_Apuracao(tApuracao As typeApuracao) As Long

Dim lComando As Long
Dim lErro As Long
Dim dSaldo As Double
Dim sCcl As String
Dim sConta As String
Dim objClass2batch As New Class2batch

On Error GoTo Erro_Cria_Lancamentos_Apuracao

    lComando = 0
    tApuracao.lComando1_Ccl = 0
    tApuracao.lComando2_Ccl = 0
    tApuracao.lComando1_Conta = 0
    tApuracao.lComando2_Conta = 0
    
    tApuracao.lDoc = tApuracao.lDoc + 1

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5081

    tApuracao.lComando1_Ccl = Comando_Abrir()
    If tApuracao.lComando1_Ccl = 0 Then Error 5092

    tApuracao.lComando2_Ccl = Comando_Abrir()
    If tApuracao.lComando2_Ccl = 0 Then Error 5093

    tApuracao.lComando1_Conta = Comando_Abrir()
    If tApuracao.lComando1_Conta = 0 Then Error 5098

    tApuracao.lComando2_Conta = Comando_Abrir()
    If tApuracao.lComando2_Conta = 0 Then Error 5099

    tApuracao.sConta = String(STRING_CONTA, 0)

    'descobre o centro de custo/lucro da conta resultado
    If tApuracao.iUsoCcl = CCL_USA_CONTABIL Then

        sCcl = String(STRING_CCL, 0)
    
        lErro = Mascara_RetornaCcl(tApuracao.sConta_Resultado, tApuracao.sCcl_ContaPonte)
        If lErro <> AD_SQL_SUCESSO Then Error 9488

    Else
        tApuracao.sCcl_ContaPonte = ""
    End If


    lErro = objClass2batch.Apuracao_Exercicio_Comando_SQL(tApuracao.sSQL, tApuracao.colContasApuracao)
    If lErro <> SUCESSO Then Error 9775
    
    lErro = objClass2batch.Apuracao_Exercicio_Executa_SQL(tApuracao.sSQL, tApuracao.sConta, tApuracao.asConta, lComando, tApuracao.colContasApuracao)
    If lErro <> SUCESSO Then Error 9776

    'Le o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9410
    
    Do While lErro = AD_SQL_SUCESSO

        'guarda o saldo que vai ser apurado para o centro de custo.
        dSaldo = 0
        
        If tApuracao.iUsoCcl = CCL_USA_EXTRACONTABIL Then
        
            lErro = Processa_Ccl_Apuracao(tApuracao, dSaldo)
            If lErro <> SUCESSO And lErro <> CONTA_SEM_CCL Then Error 5083
            
        End If
        
        lErro = Processa_Conta_Apuracao(tApuracao, -dSaldo)
        If lErro <> SUCESSO Then Error 5100

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9411
        
    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(tApuracao.lComando1_Ccl)
    Call Comando_Fechar(tApuracao.lComando2_Ccl)
    Call Comando_Fechar(tApuracao.lComando1_Conta)
    Call Comando_Fechar(tApuracao.lComando2_Conta)

    Cria_Lancamentos_Apuracao = SUCESSO

    Exit Function

Erro_Cria_Lancamentos_Apuracao:

    Cria_Lancamentos_Apuracao = Err

    Select Case Err

        Case 5081, 5092, 5093, 5098, 5099
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 9410, 9411, 9775, 9776
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)

        Case 5083, 5100, 9775

        Case 5097
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTO", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 1)

        Case 9488
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCcl", Err, tApuracao.sConta_Resultado)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154907)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(tApuracao.lComando1_Ccl)
    Call Comando_Fechar(tApuracao.lComando2_Ccl)
    Call Comando_Fechar(tApuracao.lComando1_Conta)
    Call Comando_Fechar(tApuracao.lComando2_Conta)

    Exit Function

End Function

Function Estorno_Apuracao_Exercicio(tApuracao As typeApuracao) As Long

Dim lErro As Long

On Error GoTo Erro_Estorno_Apuracao_Exercicio


    tApuracao.sOrigem_Apuracao = "APE"
    tApuracao.sOrigem_Estorno = "EAE"

    'gera o lote de estorno de apuração para o exercicio em questão caso exista
    lErro = Gera_Lote_Estorno_Apura(tApuracao)
    If lErro <> SUCESSO Then Error 5222

    'atualiza o lote de estorno
    lErro = Atualiza_Lote_Estorno_Apura(tApuracao)
    If lErro <> SUCESSO Then Error 5223

    Estorno_Apuracao_Exercicio = SUCESSO

    Exit Function

Erro_Estorno_Apuracao_Exercicio:

    Estorno_Apuracao_Exercicio = Err

    Select Case Err

        Case 5222, 5223

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154908)

    End Select

    Exit Function

End Function

Function Gera_Lote_Apuracao(tApuracao As typeApuracao) As Long

Dim lComando As Long
Dim lTransacao As Long
Dim lErro As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim iLote As Integer

On Error GoTo Erro_Gera_Lote_Apuracao

    lComando = 0
    lComando2 = 0
    lComando3 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5079

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5237
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 5353

    'Inserir o lote
    lErro = Comando_Executar(lComando, "INSERT INTO LotePendente (FilialEmpresa, Origem, Exercicio, Periodo, Lote) VALUES (?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 5088

    tApuracao.iNumLanc = 0
    tApuracao.dSaldo = 0
    tApuracao.dTotCre = 0
    tApuracao.dTotDeb = 0

    lErro = Cria_Lancamentos_Apuracao(tApuracao)
    If lErro <> SUCESSO Then Error 5080

    'Selecionar a capa de lote
    lErro = Comando_ExecutarPos(lComando3, "SELECT Lote FROM LotePendente  WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote = ?", 0, iLote, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 5354

    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO Then Error 5383

    'Atualizar a capa de lote
    lErro = Comando_ExecutarPos(lComando2, "UPDATE LotePendente SET TotCre = ?, TotDeb = ?, NumDocAtual = ?", lComando3, tApuracao.dTotCre, tApuracao.dTotDeb, tApuracao.iNumLanc)
    If lErro <> AD_SQL_SUCESSO Then Error 5238

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Gera_Lote_Apuracao = SUCESSO

    Exit Function

Erro_Gera_Lote_Apuracao:

    Gera_Lote_Apuracao = Err

    Select Case Err

        Case 5079, 5237, 5353
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5080

        Case 5088
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOTE", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)

        Case 5238
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LOTE", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)

        Case 5354, 5383
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154909)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function Periodos_Apuracao_Exercicio(tApuracao As typeApuracao) As Long
'verifica se os periodos possuem apuração. Se possuirem, estorna.

Dim lComando1 As Long, lComando2 As Long
Dim lErro As Long
Dim iApurado As Integer

On Error GoTo Erro_Periodos_Apuracao_Exercicio

    lComando1 = 0
    lComando2 = 0

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5075

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5102

    'Pesquisa os periodos que serão estornados
    lErro = Comando_ExecutarPos(lComando1, "SELECT Periodo, Lote, ExisteApuracaoPeriodo FROM PeriodosFilial WHERE FilialEmpresa=? AND Exercicio = ? ", 0, tApuracao.iPeriodo, tApuracao.iLote, iApurado, tApuracao.iFilialEmpresa, tApuracao.iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5076

    'Le o primeiro periodo
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9413

    tApuracao.sOrigem_Apuracao = "APP"
    tApuracao.sOrigem_Estorno = "EAP"

    Do While lErro = AD_SQL_SUCESSO
    
        'loca o Periodo que será apurado
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 5077
        
        'se existia um lote anterior de apuração para o periodo em questão ===> estorna-o
        If iApurado = EXISTE_LOTE_APURACAO_PERIODO Then

            'gera o lote de estorno de apuração para o periodo em questão caso exista
            lErro = Gera_Lote_Estorno_Apura(tApuracao)
            If lErro <> SUCESSO Then Error 5231

            'atualiza o lote de estorno
            lErro = Atualiza_Lote_Estorno_Apura(tApuracao)
            If lErro <> SUCESSO Then Error 5232

        End If

        'Atualiza o periodo indicando que foi estornado a apuracao
        lErro = Comando_ExecutarPos(lComando2, "UPDATE PeriodosFilial SET Apurado = ?, ExisteApuracaoPeriodo = ?", lComando1, PERIODO_NAO_APURADO, NAO_EXISTE_LOTE_APURACAO_PERIODO)
        If lErro <> AD_SQL_SUCESSO Then Error 5103

        'Le o proximo periodo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9414

    Loop

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Periodos_Apuracao_Exercicio = SUCESSO

    Exit Function

Erro_Periodos_Apuracao_Exercicio:

    Periodos_Apuracao_Exercicio = Err

    Select Case Err

        Case 5075, 5102
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5076, 9413, 9414
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL1", Err)

        Case 5077
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODOSFILIAL", Err, tApuracao.iFilialEmpresa, tApuracao.iExercicio, tApuracao.iPeriodo)

        Case 5103
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PERIODOSFILIAL", Err, tApuracao.iPeriodo, tApuracao.iExercicio, tApuracao.iFilialEmpresa)

        Case 5231, 5232

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154910)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Processa_Apuracao_Exercicio(tApuracao As typeApuracao) As Long

Dim lErro As Long

On Error GoTo Erro_Processa_Apuracao_Exercicio

    'siglas de lote de apuração de exercicio e lote de estorno de apuração de exercicio
    tApuracao.sOrigem_Apuracao = "APE"
    tApuracao.sOrigem_Estorno = "EAE"

    'gera o lote de apuração do exercicio
    lErro = Gera_Lote_Apuracao(tApuracao)
    If lErro <> SUCESSO Then Error 5234

    'contabiliza o lote de apuração do exercicio
    lErro = Atualiza_Lote_Apuracao(tApuracao)
    If lErro <> SUCESSO Then Error 5235

    Processa_Apuracao_Exercicio = SUCESSO

    Exit Function

Erro_Processa_Apuracao_Exercicio:

    Processa_Apuracao_Exercicio = Err

    Select Case Err

        Case 5234, 5235

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154911)

    End Select

    Exit Function

End Function

Private Function Processa_Ccl_Apuracao(tApuracao As typeApuracao, dSaldoConta As Double) As Long

Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim iPeriodo As Integer
Dim dSaldo As Double
Dim lErro As Long
Dim sCcl As String

On Error GoTo Erro_Processa_Ccl_Apuracao

    sCcl = String(STRING_CCL, 0)

    'Seleciona os ccl da tabela que contém os saldos mensais por ccl
    lErro = Comando_Executar(tApuracao.lComando1_Ccl, "SELECT Ccl, Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cre01, Cre02, Cre03, Cre04, Cre05, Cre06, Cre07, Cre08, Cre09, Cre10, Cre11, Cre12 FROM MvPerCcl WHERE FilialEmpresa=? AND Exercicio = ? AND Conta = ?", sCcl, dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), tApuracao.iFilialEmpresa, tApuracao.iExercicio, tApuracao.sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5087

    'Le o primeiro registro para a conta em questão
    lErro = Comando_BuscarPrimeiro(tApuracao.lComando1_Ccl)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9419
    
    If lErro = AD_SQL_SEM_DADOS Then Error 5085
    
    Do While lErro = AD_SQL_SUCESSO
    
        dSaldo = 0

        For iPeriodo = 1 To tApuracao.iNumPeriodos
            
            dSaldo = dSaldo + dCredito(iPeriodo) - dDebito(iPeriodo)
 
        Next


        If dSaldo <> 0 Then
        
            'insere o lancamento para zerar a conta/ccl
            lErro = Comando_Executar(tApuracao.lComando2_Ccl, "INSERT INTO LanPendente (FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor) Values (?,?,?,?,?,?,?,?,?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 1, tApuracao.iLote, tApuracao.iPeriodo, tApuracao.dtData, tApuracao.sConta, sCcl, tApuracao.sHistorico, -dSaldo)
            If lErro <> AD_SQL_SUCESSO Then Error 5090
    
            'insere o lancamento da Conta Resultado
            lErro = Comando_Executar(tApuracao.lComando2_Ccl, "INSERT INTO LanPendente (FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor) Values (?,?,?,?,?,?,?,?,?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 2, tApuracao.iLote, tApuracao.iPeriodo, tApuracao.dtData, tApuracao.sConta_Resultado, tApuracao.sCcl_ContaPonte, tApuracao.sHistorico, dSaldo)
            If lErro <> AD_SQL_SUCESSO Then Error 55830
    
            dSaldoConta = dSaldoConta + dSaldo
    
            If dSaldo < 0 Then dSaldo = -dSaldo
            
            tApuracao.dTotDeb = tApuracao.dTotDeb + dSaldo
            tApuracao.dTotCre = tApuracao.dTotCre + dSaldo
            
            tApuracao.iNumLanc = tApuracao.iNumLanc + 1
            tApuracao.lDoc = tApuracao.lDoc + 1
            
            
        End If
        
        lErro = Comando_BuscarProximo(tApuracao.lComando1_Ccl)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9420
        
    Loop

    Processa_Ccl_Apuracao = SUCESSO

    Exit Function

Erro_Processa_Ccl_Apuracao:

    Processa_Ccl_Apuracao = Err

    Select Case Err
        
        Case 5085
            Processa_Ccl_Apuracao = CONTA_SEM_CCL

        Case 5087, 9419, 9420
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL", Err)

        Case 5090
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTO", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 1)
        
        Case 55830
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTO", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 2)
        
        Case 9476
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_MVPERCCL", Err, tApuracao.iFilialEmpresa, tApuracao.iExercicio, sCcl, tApuracao.sConta)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154912)

    End Select

    Exit Function

End Function

Private Function Processa_Conta_Apuracao(tApuracao As typeApuracao, dSaldoConta As Double) As Long

Dim dCredito(NUM_MAX_PERIODOS) As Double
Dim dDebito(NUM_MAX_PERIODOS) As Double
Dim iPeriodo As Integer
Dim dSaldo As Double
Dim lErro As Long
Dim sCcl As String

On Error GoTo Erro_Processa_Conta_Apuracao

    'Seleciona a conta passada como parametro da tabela de Saldos por Periodo
    lErro = Comando_Executar(tApuracao.lComando1_Conta, "SELECT Deb01, Deb02, Deb03, Deb04, Deb05, Deb06, Deb07, Deb08, Deb09, Deb10, Deb11, Deb12, Cre01, Cre02, Cre03, Cre04, Cre05, Cre06, Cre07, Cre08, Cre09, Cre10, Cre11, Cre12 FROM MvPerCta WHERE FilialEmpresa=? AND Exercicio = ? AND Conta = ?", dDebito(1), dDebito(2), dDebito(3), dDebito(4), dDebito(5), dDebito(6), dDebito(7), dDebito(8), dDebito(9), dDebito(10), dDebito(11), dDebito(12), dCredito(1), dCredito(2), dCredito(3), dCredito(4), dCredito(5), dCredito(6), dCredito(7), dCredito(8), dCredito(9), dCredito(10), dCredito(11), dCredito(12), tApuracao.iFilialEmpresa, tApuracao.iExercicio, tApuracao.sConta)
    If lErro <> AD_SQL_SUCESSO Then Error 5094

    'Le a conta em questão
    lErro = Comando_BuscarPrimeiro(tApuracao.lComando1_Conta)
    If lErro <> AD_SQL_SUCESSO Then Error 5095

    'inicializa o saldo com o saldo oriundo do centro de custo
    dSaldo = dSaldoConta

    For iPeriodo = 1 To tApuracao.iNumPeriodos
                
        dSaldo = dSaldo + dCredito(iPeriodo) - dDebito(iPeriodo)

    Next

    If dSaldo <> 0 Then
    
    
        'descobre o centro de custo/lucro da conta
        If tApuracao.iUsoCcl = CCL_USA_CONTABIL Then
    
            sCcl = String(STRING_CCL, 0)
        
            lErro = Mascara_RetornaCcl(tApuracao.sConta, sCcl)
            If lErro <> AD_SQL_SUCESSO Then Error 9487
    
        Else
            sCcl = ""
        End If
    
        'insere o lancamento para zerar a conta
        lErro = Comando_Executar(tApuracao.lComando2_Conta, "INSERT INTO LanPendente (FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor) Values (?,?,?,?,?,?,?,?,?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 1, tApuracao.iLote, tApuracao.iPeriodo, tApuracao.dtData, tApuracao.sConta, sCcl, tApuracao.sHistorico, -dSaldo)
        If lErro <> AD_SQL_SUCESSO Then Error 55832
    
        'insere o lancamento na conta-resultado
        lErro = Comando_Executar(tApuracao.lComando2_Conta, "INSERT INTO LanPendente (FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor) Values (?,?,?,?,?,?,?,?,?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 2, tApuracao.iLote, tApuracao.iPeriodo, tApuracao.dtData, tApuracao.sConta_Resultado, tApuracao.sCcl_ContaPonte, tApuracao.sHistorico, dSaldo)
        If lErro <> AD_SQL_SUCESSO Then Error 55831
    
        If dSaldo < 0 Then dSaldo = -dSaldo
        
        tApuracao.dTotCre = tApuracao.dTotCre + dSaldo
        tApuracao.dTotDeb = tApuracao.dTotDeb + dSaldo
    
        tApuracao.iNumLanc = tApuracao.iNumLanc + 1
    
        tApuracao.lDoc = tApuracao.lDoc + 1
        
    End If

    Processa_Conta_Apuracao = SUCESSO

    Exit Function

Erro_Processa_Conta_Apuracao:

    Processa_Conta_Apuracao = Err

    Select Case Err

        Case 5094, 5095
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA1", Err, tApuracao.iFilialEmpresa, tApuracao.iExercicio, tApuracao.sConta)

        Case 55831
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTO", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 2)
            
        Case 55832
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTO", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 1)
            
        Case 9487
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCcl", Err, tApuracao.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154913)

    End Select

    Exit Function

End Function

Function Rotina_Desapura_Exercicio_Int(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iLote As Integer, sHistorico As String, sConta_Resultado As String, colContasApuracao As Collection) As Long
'realiza a desapuração do exercicio iExercicio para as receitas e despesas passadas como parametro e gera um lote contendo a conta resultado da apuracao.

Dim objFiliais As AdmFiliais
Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_Rotina_Desapura_Exercicio_Int

    lTransacao = 0

   'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 188361

    If iFilialEmpresa = EMPRESA_TODA Then

        TelaAcompanhaBatch.dValorTotal = gcolFiliais.Count
    
        'se tiver selecionado a empresa, executa a apuracao para cada filial
        For Each objFiliais In gcolFiliais
    
            If objFiliais.iCodFilial <> EMPRESA_TODA And objFiliais.iCodFilial <> Abs(giFilialAuxiliar) Then
    
                lErro = Rotina_Desapura_Exercicio0(objFiliais.iCodFilial, iExercicio, iLote, sHistorico, sConta_Resultado, colContasApuracao)
                If lErro <> SUCESSO Then gError 188362
                
                TelaAcompanhaBatch.dValorAtual = TelaAcompanhaBatch.dValorAtual + 1
        
                TelaAcompanhaBatch.ProgressBar1.Value = CInt((TelaAcompanhaBatch.dValorAtual / TelaAcompanhaBatch.dValorTotal) * 100)
                
            End If
        
        Next
    
    Else
    
        'se tiver decidido apurar somente uma filial
        lErro = Rotina_Desapura_Exercicio0(iFilialEmpresa, iExercicio, iLote, sHistorico, sConta_Resultado, colContasApuracao)
        If lErro <> SUCESSO Then gError 188363
    
        TelaAcompanhaBatch.ProgressBar1.Value = 100
        
    End If

   'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 188364

    Rotina_Desapura_Exercicio_Int = SUCESSO

    Exit Function

Erro_Rotina_Desapura_Exercicio_Int:

    Rotina_Desapura_Exercicio_Int = gErr

    Select Case gErr

        Case 188361
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 188362, 188363
            
        Case 188364
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188365)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Rotina_Desapura_Exercicio0(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iLote As Integer, sHistorico As String, sConta_Resultado As String, colContasApuracao As Collection) As Long
'realiza a desapuração do exercicio iExercicio para as receitas e despesas passadas como parametro e gera um lote contendo a conta resultado da apuracao.

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim tApuracao As typeApuracao
Dim lErro As Long
Dim iStatus As Integer

On Error GoTo Erro_Rotina_Desapura_Exercicio0

    tApuracao.iFilialEmpresa = iFilialEmpresa
    tApuracao.iExercicio = iExercicio
    tApuracao.sConta_Resultado = sConta_Resultado
    tApuracao.sHistorico = sHistorico
    Set tApuracao.colContasApuracao = colContasApuracao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 188366

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 188367

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 188368

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 188369

    'Pesquisa o uso do centro de custo/lucro
    lErro = Comando_Executar(lComando2, "SELECT UsoCcl FROM Configuracao", tApuracao.iUsoCcl)
    If lErro <> AD_SQL_SUCESSO Then gError 188370
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO Then gError 188371
    
    'Pesquisa o Exercicio em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Status, NumPeriodos, DataFim FROM Exercicios WHERE Exercicio = ?", 0, iStatus, tApuracao.iNumPeriodos, tApuracao.dtData, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then gError 188372

    'le o Exercicio
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 188373

    'não permite a mudança no status do exercicio
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 188374
    
    'Se o Exercicio estiver fechado ==> erro
    If iStatus = EXERCICIO_FECHADO Then gError 188375

    'Pesquisa o ExercicioFilial em questão
    lErro = Comando_ExecutarPos(lComando3, "SELECT Status, LoteApuracao, DocApuracao, ExisteLoteApuracao FROM ExerciciosFilial WHERE Exercicio = ? AND FilialEmpresa=?", 0, tApuracao.iStatusExercicioFilial, tApuracao.iLote, tApuracao.lDoc, tApuracao.iExisteLoteApuracao, iExercicio, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 188376

    'le o ExercicioFilial
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO Then gError 188377

    'lock do ExercicioFilial
    lErro = Comando_LockExclusive(lComando3)
    If lErro <> AD_SQL_SUCESSO Then gError 188378
    
    tApuracao.iPeriodo = tApuracao.iNumPeriodos

    lErro = Rotina_Desapura_Exercicio1(tApuracao, iLote)
    If lErro <> SUCESSO Then gError 188379

    'Atualiza o exercicio indicando que foi apurado e armazenando o número do lote referente à ultima apuração
    lErro = Comando_ExecutarPos(lComando1, "UPDATE ExerciciosFilial SET Status = ?, ExisteLoteApuracao=?", lComando3, EXERCICIO_ABERTO, NAO_EXISTE_LOTE_APURACAO_EXERCICIO)
    If lErro <> AD_SQL_SUCESSO Then gError 188380

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Rotina_Desapura_Exercicio0 = SUCESSO

    Exit Function

Erro_Rotina_Desapura_Exercicio0:

    Rotina_Desapura_Exercicio0 = gErr

    Select Case gErr

        Case 188366 To 188369
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 188370, 188371
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", gErr)
            
        Case 188372, 188373
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO", gErr, iExercicio)

        Case 188374
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCACAO_EXERCICIO", gErr, iExercicio)

        Case 188375
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_FECHADO", gErr, iExercicio)

        Case 188376, 188377
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL", gErr, iFilialEmpresa, iExercicio)
            
        Case 188378
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOSFILIAL", gErr, iFilialEmpresa, iExercicio)
            
        Case 188379

        Case 188380
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOSFILIAL", gErr, iExercicio, iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188381)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Private Function Rotina_Desapura_Exercicio1(tApuracao As typeApuracao, ByVal iLote As Integer) As Long

Dim lComando As Long
Dim lErro As Long
Dim iTipoConta As Integer

On Error GoTo Erro_Rotina_Desapura_Exercicio1

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 188382

'    'le a conta resultado
'    lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta FROM PlanoConta WHERE Conta=?", iTipoConta, tApuracao.sConta_Resultado)
'    If lErro <> AD_SQL_SUCESSO Then gError 188383
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO Then gError 188384
'
'    'lock da conta resultado
'    lErro = Comando_LockShared(lComando)
'    If lErro <> AD_SQL_SUCESSO Then gError 188385

    'se existia um lote anterior de apuração para o exercicio em questão ===> estorna-o
    If tApuracao.iExisteLoteApuracao = EXISTE_LOTE_APURACAO_EXERCICIO Then

        lErro = Estorno_Apuracao_Exercicio(tApuracao)
        If lErro <> SUCESSO Then gError 188386

    End If

    'verifica se os periodos possuem lotes de apuracao.
    'Se possuirem, estorna-os
    lErro = Periodos_Apuracao_Exercicio(tApuracao)
    If lErro <> SUCESSO Then gError 188387

    Call Comando_Fechar(lComando)
    
    Rotina_Desapura_Exercicio1 = SUCESSO

    Exit Function

Erro_Rotina_Desapura_Exercicio1:

    Rotina_Desapura_Exercicio1 = gErr

    Select Case gErr

        Case 188382
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 188383, 188384
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", gErr)
        
        Case 188385
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", gErr, tApuracao.sConta_Resultado)

        Case 188386, 188387

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188388)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

