Attribute VB_Name = "Module6"
'Apuração de Periodos

Option Explicit


Function Atualiza_Lote_Apura(tApuracao As typeApuracao) As Long

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tLote As typeLote_batch
Dim tLote1 As typeLote

On Error GoTo Erro_Atualiza_Lote_Apura

    lComando = 0
    lComando1 = 0
    lComando2 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5211

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9465

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5257

    tLote1.sIdOriginal = String(STRING_IDORIGINAL, 0)

    'Pesquisa o lote em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT TotCre, TotDeb, TotInf, Status, IdOriginal, NumDocInf, NumDocAtual, IdAtualizacao FROM LotePendente WHERE FilialEmpresa=? AND Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote = ?", 0, tLote1.dTotCre, tLote1.dTotDeb, tLote1.dTotInf, tLote1.iStatus, tLote1.sIdOriginal, tLote1.iNumDocInf, tLote1.iNumDocAtual, tLote1.iIDAtualizacao, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 5212

    'Le o lote
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5213

    'loca o lote de apuracao
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5214

    tLote.iFilialEmpresa = tApuracao.iFilialEmpresa
    tLote.sOrigem = tApuracao.sOrigem_Apuracao
    tLote.iExercicio = tApuracao.iExercicio
    tLote.iPeriodo = tApuracao.iPeriodo
    tLote.iLote = tApuracao.iLote + 1
    
    tLote1.iFilialEmpresa = tLote.iFilialEmpresa
    tLote1.sOrigem = tLote.sOrigem
    tLote1.iExercicio = tLote.iExercicio
    tLote1.iPeriodo = tLote.iPeriodo
    tLote1.iLote = tLote.iLote

    'Processa os lançamentos do lote
    lErro = Processa_Lancamentos(tLote, ATUALIZACAO, tApuracao.iUsoCcl, 0)
    If lErro <> SUCESSO Then Error 5256

    'exclui o lote da tabela de lotes pendentes
    lErro = Comando_ExecutarPos(lComando2, "DELETE From LotePendente", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9463

    'insere o lote na a tabela de lotes atualizados
    lErro = Comando_Executar(lComando1, "INSERT INTO Lote (FilialEmpresa, Origem, Exercicio, Periodo, Lote, TotCre, TotDeb, TotInf, Status, IdOriginal, NumDocInf, NumDocAtual, IdAtualizacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", tLote1.iFilialEmpresa, tLote1.sOrigem, tLote1.iExercicio, tLote1.iPeriodo, tLote1.iLote, tLote1.dTotCre, tLote1.dTotDeb, tLote1.dTotInf, LOTE_ATUALIZADO, tLote1.sIdOriginal, tLote1.iNumDocInf, tLote1.iNumDocAtual, tLote1.iIDAtualizacao)
    If lErro <> AD_SQL_SUCESSO Then Error 9464

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Atualiza_Lote_Apura = SUCESSO

    Exit Function

Erro_Atualiza_Lote_Apura:

    Atualiza_Lote_Apura = Err

    Select Case Err

        Case 5211, 5257, 9465
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5212, 5213
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote + 1)

        Case 5214
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LOTE", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote + 1)

        Case 5256

        Case 9463
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LOTEPENDENTE", lErro, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote + 1)

        Case 9464
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOTE", lErro, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote + 1)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154923)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Atualiza_Lote_Estorno_Apura(tApuracao As typeApuracao) As Long

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tLote As typeLote_batch
Dim tLote1 As typeLote

On Error GoTo Erro_Atualiza_Lote_Estorno_Apura

    lComando = 0
    lComando1 = 0
    lComando2 = 0
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5224

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9457

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5262

    tLote1.sIdOriginal = String(STRING_IDORIGINAL, 0)

    'Pesquisa o lote no banco de dados
    lErro = Comando_ExecutarPos(lComando, "SELECT TotCre, TotDeb, TotInf, Status, IdOriginal, NumDocInf, NumDocAtual, IdAtualizacao FROM LotePendente WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote = ?", 0, tLote1.dTotCre, tLote1.dTotDeb, tLote1.dTotInf, tLote1.iStatus, tLote1.sIdOriginal, tLote1.iNumDocInf, tLote1.iNumDocAtual, tLote1.iIDAtualizacao, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Estorno, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 5225

    tLote.iFilialEmpresa = tApuracao.iFilialEmpresa
    tLote.sOrigem = tApuracao.sOrigem_Estorno
    tLote.iExercicio = tApuracao.iExercicio
    tLote.iPeriodo = tApuracao.iPeriodo
    tLote.iLote = tApuracao.iLote
    
    tLote1.iFilialEmpresa = tLote.iFilialEmpresa
    tLote1.sOrigem = tLote.sOrigem
    tLote1.iExercicio = tLote.iExercicio
    tLote1.iPeriodo = tLote.iPeriodo
    tLote1.iLote = tLote.iLote

    'Le o lote
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5226

    'loca o lote de apuracao
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5227

    'Processa os lançamentos do lote
    lErro = Processa_Lancamentos(tLote, ATUALIZACAO, tApuracao.iUsoCcl, 0)
    If lErro <> SUCESSO Then Error 5261

    'marca o lote como atualizado
    lErro = Comando_ExecutarPos(lComando2, "DELETE From LotePendente", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 9459

    lErro = Comando_Executar(lComando1, "INSERT INTO Lote (FilialEmpresa, Origem, Exercicio, Periodo, Lote, TotCre, TotDeb, TotInf, Status, IdOriginal, NumDocInf, NumDocAtual, IdAtualizacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", tLote1.iFilialEmpresa, tLote1.sOrigem, tLote1.iExercicio, tLote1.iPeriodo, tLote1.iLote, tLote1.dTotCre, tLote1.dTotDeb, tLote1.dTotInf, LOTE_ATUALIZADO, tLote1.sIdOriginal, tLote1.iNumDocInf, tLote1.iNumDocAtual, tLote1.iIDAtualizacao)
    If lErro <> AD_SQL_SUCESSO Then Error 9458

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Atualiza_Lote_Estorno_Apura = SUCESSO

    Exit Function

Erro_Atualiza_Lote_Estorno_Apura:

    Atualiza_Lote_Estorno_Apura = Err

    Select Case Err

        Case 5224, 5262, 9457
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5225, 5226
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Estorno, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)

        Case 5227
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_LOTE", Err, tLote.iFilialEmpresa, tLote.sOrigem, tLote.iExercicio, tLote.iPeriodo, tApuracao.iLote)

        Case 5261
        
        Case 9458
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOTE", lErro, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Estorno, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)
            
        Case 9459
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_LOTEPENDENTE", lErro, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Estorno, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154924)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Gera_Lote_Apura(tApuracao As typeApuracao) As Long

Dim lComando As Long, lComando1 As Long, lComando2 As Long
Dim lErro As Long
Dim iLote As Integer

On Error GoTo Erro_Gera_Lote_Apura

    lComando = 0
    lComando1 = 0
    lComando2 = 0
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5195

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5368

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5369
    
    tApuracao.iNumLanc = 0
    tApuracao.dTotCre = 0
    tApuracao.dTotDeb = 0
    
    'Inserir o lote
    lErro = Comando_Executar(lComando, "INSERT INTO LotePendente (FilialEmpresa, Origem, Exercicio, Periodo, Lote) VALUES (?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 5196

    'gera os lançamentos para o lote de apuração
    lErro = Processa_Lancamentos_Apura(tApuracao)
    If lErro <> SUCESSO Then Error 5197

    'Seleciona o lote que acabou de ser inserido para atualiza-lo
    lErro = Comando_ExecutarPos(lComando2, "SELECT Lote FROM LotePendente WHERE FilialEmpresa=? AND Origem  = ? AND Exercicio = ? AND Periodo = ? AND Lote = ?", 0, iLote, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 5370

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 5371

    'Atualizar os valores do lote
    lErro = Comando_ExecutarPos(lComando1, "UPDATE LotePendente SET TotCre = ?, TotDeb = ?, NumDocAtual = ?", lComando2, tApuracao.dTotCre, tApuracao.dTotDeb, tApuracao.iNumLanc)
    If lErro <> AD_SQL_SUCESSO Then Error 5198

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Gera_Lote_Apura = SUCESSO

    Exit Function

Erro_Gera_Lote_Apura:

    Gera_Lote_Apura = Err

    Select Case Err

        Case 5195, 5368, 5369
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", lErro)

        Case 5196
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOTE", lErro, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote + 1)

        Case 5197

        Case 5198
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_LOTE", lErro, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote + 1)

        Case 5370, 5371
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE", lErro, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote + 1)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 154925)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Gera_Lote_Estorno_Apura(tApuracao As typeApuracao) As Long

Dim lComando As Long, lComando1 As Long
Dim lErro As Long
Dim dTotCre As Double, dTotDeb As Double
Dim iNumLancAtual As Integer

On Error GoTo Erro_Gera_Lote_Estorno_Apura

    lComando = 0
    lComando1 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5219

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5216

    'Selecionar o lote a ser estornado
    lErro = Comando_Executar(lComando1, "SELECT TotCre, TotDeb, NumDocAtual FROM Lote WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND Periodo = ? AND Lote = ?", dTotCre, dTotDeb, iNumLancAtual, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 5217

    'le o lote a ser estornado
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 5218

    'Insere o lote de estorno
    lErro = Comando_Executar(lComando, "INSERT INTO LotePendente (FilialEmpresa, Origem, Exercicio, Periodo, Lote, TotCre, TotDeb, NumDocAtual) VALUES (?,?,?,?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Estorno, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote, dTotDeb, dTotCre, iNumLancAtual)
    If lErro <> AD_SQL_SUCESSO Then Error 5220

    'gera os lançamentos para o lote de estorno da apuração
    lErro = Processa_Lancamentos_Estorno_Apura(tApuracao)
    If lErro <> SUCESSO Then Error 5221

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Gera_Lote_Estorno_Apura = SUCESSO

    Exit Function

Erro_Gera_Lote_Estorno_Apura:

    Gera_Lote_Estorno_Apura = Err

    Select Case Err

        Case 5216, 5219
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5217, 5218
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTE", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)
        
        Case 5220
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOTE", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Estorno, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)

        Case 5221

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154926)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Processa_Faixa_Contas_Apura(tApuracao As typeApuracao) As Long

Dim lErro As Long
Dim dDebito As Double
Dim dCredito As Double
Dim objClass2batch As New Class2batch
'Dim sCcl As String

On Error GoTo Erro_Processa_Faixa_Contas_Apura

    If tApuracao.iZeraRD = MARCADO Then
    
        tApuracao.lDoc = tApuracao.lDoc + 1
    
'        If tApuracao.iUsoCcl = CCL_USA_CONTABIL Then
'
'            sCcl = String(STRING_CCL, 0)
'
'            lErro = Mascara_RetornaCcl(tApuracao.sConta_Resultado, sCcl)
'            If lErro <> AD_SQL_SUCESSO Then Error 9811
'
'        Else
'            sCcl = ""
'        End If
        
    End If

    tApuracao.sConta = String(STRING_CONTA, 0)

    'Seleciona as contas analiticas da faixa de contas selecionada
    lErro = objClass2batch.Apuracao_Exercicio_Executa_SQL(tApuracao.sSQL, tApuracao.sConta, tApuracao.asConta, tApuracao.lComando1, tApuracao.colContasApuracao)
    If lErro <> SUCESSO Then Error 9811

    'Le a primeira conta
    lErro = Comando_BuscarPrimeiro(tApuracao.lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9439

    Do While lErro = AD_SQL_SUCESSO

        'Seleciona os saldos da conta em questão
        lErro = Comando_Executar(tApuracao.lComando2, "SELECT Deb" + tApuracao.sPeriodo + ", Cre" + tApuracao.sPeriodo + " FROM MvPerCta WHERE FilialEmpresa=? AND Exercicio = ? AND Conta = ?", dDebito, dCredito, tApuracao.iFilialEmpresa, tApuracao.iExercicio, tApuracao.sConta)
        If lErro <> AD_SQL_SUCESSO Then Error 5207

        'Le os saldos da conta em questão
        lErro = Comando_BuscarPrimeiro(tApuracao.lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 5208
        
        If tApuracao.iZeraRD = MARCADO And (dCredito - dDebito) <> 0 Then
            
'            'insere o resultado e a contra partida
'            lErro = Comando_Executar(tApuracao.lComando3, "INSERT INTO LanPendente (FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, tApuracao.iLote + 1, tApuracao.iPeriodo, tApuracao.dtData, tApuracao.sConta_Resultado, sCcl, tApuracao.sHistorico, (dCredito - dDebito))
'            If lErro <> AD_SQL_SUCESSO Then Error 9811
        
            tApuracao.iSeq = tApuracao.iSeq + 1
        
            'insere o resultado e a contra partida
            lErro = Comando_Executar(tApuracao.lComando3, "INSERT INTO LanPendente (FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, tApuracao.iSeq, tApuracao.iLote + 1, tApuracao.iPeriodo, tApuracao.dtData, tApuracao.sConta, tApuracao.sCcl_ContaPonte, tApuracao.sHistorico, -(dCredito - dDebito))
            If lErro <> AD_SQL_SUCESSO Then Error 9811
            
        End If

        tApuracao.dSaldo = tApuracao.dSaldo + dCredito - dDebito

        'le a proxima conta
        lErro = Comando_BuscarProximo(tApuracao.lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9440

    Loop

    Processa_Faixa_Contas_Apura = SUCESSO

    Exit Function

Erro_Processa_Faixa_Contas_Apura:

    Processa_Faixa_Contas_Apura = Err

    Select Case Err

        Case 5203, 9439, 9440, 9811
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)
        
        Case 5207, 5208
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCTA1", Err, tApuracao.iFilialEmpresa, tApuracao.iExercicio, tApuracao.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154927)

    End Select

    Exit Function

End Function

Function Processa_Lancamentos_Apura(tApuracao As typeApuracao) As Long

'Dim lComando As Long
Dim lErro As Long
Dim lErro1 As Integer
Dim sConta_Resultado As String
Dim sCcl As String
Dim objClass2batch As New Class2batch

On Error GoTo Erro_Processa_Lancamentos_Apura

    tApuracao.lComando1 = 0
    tApuracao.lComando2 = 0
    tApuracao.lComando3 = 0

    tApuracao.lComando1 = Comando_Abrir()
    If tApuracao.lComando1 = 0 Then Error 5200

    tApuracao.lComando2 = Comando_Abrir()
    If tApuracao.lComando2 = 0 Then Error 5205

    tApuracao.lComando3 = Comando_Abrir()
    If tApuracao.lComando3 = 0 Then Error 5206
    
    tApuracao.dSaldo = 0

    lErro = objClass2batch.Apuracao_Exercicio_Comando_SQL(tApuracao.sSQL, tApuracao.colContasApuracao)
    If lErro <> SUCESSO Then Error 9810

    'processa a faixa de contas selecionada
    lErro = Processa_Faixa_Contas_Apura(tApuracao)
    If lErro <> SUCESSO Then Error 5202

    lErro = Processa_Lancamentos_Apura1(tApuracao)
    If lErro <> SUCESSO Then Error 9493

    Call Comando_Fechar(tApuracao.lComando1)
    Call Comando_Fechar(tApuracao.lComando2)
    Call Comando_Fechar(tApuracao.lComando3)

    Processa_Lancamentos_Apura = SUCESSO

    Exit Function

Erro_Processa_Lancamentos_Apura:

    Processa_Lancamentos_Apura = Err

    Select Case Err

        Case 5199, 5200, 5205, 5206
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5201, 9441, 9442
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESULTADO", Err)

        Case 5202, 9493, 9810

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154928)

    End Select

'    Call Comando_Fechar(lComando)
    Call Comando_Fechar(tApuracao.lComando1)
    Call Comando_Fechar(tApuracao.lComando2)
    Call Comando_Fechar(tApuracao.lComando3)

    Exit Function

End Function

Function Processa_Lancamentos_Apura1(tApuracao As typeApuracao) As Long

Dim lErro As Long
Dim sCcl As String
Dim iTipoConta As Integer

On Error GoTo Erro_Processa_Lancamentos_Apura1
        
    'le a conta resultado
    lErro = Comando_ExecutarLockado(tApuracao.lComando4, "SELECT TipoConta FROM PlanoConta WHERE Conta=?", iTipoConta, tApuracao.sConta_Resultado)
    If lErro <> AD_SQL_SUCESSO Then Error 9490
    
    lErro = Comando_BuscarPrimeiro(tApuracao.lComando4)
    If lErro <> AD_SQL_SUCESSO Then Error 9491
    
    'lock da conta resultado
    lErro = Comando_LockShared(tApuracao.lComando4)
    If lErro <> AD_SQL_SUCESSO Then Error 9492
    
    If tApuracao.iZeraRD = DESMARCADO Then tApuracao.lDoc = tApuracao.lDoc + 1
    
    If tApuracao.iUsoCcl = CCL_USA_CONTABIL Then
    
        sCcl = String(STRING_CCL, 0)
        
        lErro = Mascara_RetornaCcl(tApuracao.sConta_Resultado, sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 9474
    
    Else
        sCcl = ""
    End If
    
    'insere o resultado e a contra partida
    lErro = Comando_Executar(tApuracao.lComando3, "INSERT INTO LanPendente (FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 1, tApuracao.iLote + 1, tApuracao.iPeriodo, tApuracao.dtData, tApuracao.sConta_Resultado, sCcl, tApuracao.sHistorico, tApuracao.dSaldo)
    If lErro <> AD_SQL_SUCESSO Then Error 5209
    
    If tApuracao.iZeraRD = DESMARCADO Then
        'insere o resultado e a contra partida
        lErro = Comando_Executar(tApuracao.lComando3, "INSERT INTO LanPendente (FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 2, tApuracao.iLote + 1, tApuracao.iPeriodo, tApuracao.dtData, tApuracao.sConta_Ponte, tApuracao.sCcl_ContaPonte, tApuracao.sHistorico, -tApuracao.dSaldo)
        If lErro <> AD_SQL_SUCESSO Then Error 5210
    End If
    
    If tApuracao.dSaldo < 0 Then
        tApuracao.dTotCre = tApuracao.dTotCre - tApuracao.dSaldo
        tApuracao.dTotDeb = tApuracao.dTotDeb - tApuracao.dSaldo
    Else
        tApuracao.dTotCre = tApuracao.dTotCre + tApuracao.dSaldo
        tApuracao.dTotDeb = tApuracao.dTotDeb + tApuracao.dSaldo
    End If
    
    tApuracao.iNumLanc = tApuracao.iNumLanc + 1

    Processa_Lancamentos_Apura1 = SUCESSO
    
    Exit Function

Erro_Processa_Lancamentos_Apura1:

    Processa_Lancamentos_Apura1 = Err

    Select Case Err

        Case 5209
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTO", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 1)

        Case 5210
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTO", Err, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.lDoc, 2)
            
        Case 9474
            lErro = Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCcl", Err, tApuracao.sConta_Resultado)

        Case 9490, 9491
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)
        
        Case 9492
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, tApuracao.sConta_Resultado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154929)
            
        End Select
        
        Exit Function

End Function

Function Processa_Lancamentos_Estorno_Apura(tApuracao As typeApuracao) As Long

Dim lComando As Long, lComando1 As Long
Dim lErro As Long
Dim lDoc As Long
Dim iSeq As Integer
Dim dValor As Double
Dim sConta As String
Dim sHistorico As String
Dim sCcl As String
Dim dtData As Date

On Error GoTo Erro_Processa_Lancamentos_Estorno_Apura

    lComando = 0
    lComando1 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5222

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5223

    sCcl = String(STRING_CCL, 0)
    sConta = String(STRING_CONTA, 0)
    sHistorico = String(STRING_HISTORICO, 0)

    'seleciona os lancamentos a serem estornados
    lErro = Comando_Executar(lComando, "SELECT Doc, Seq, Conta, Ccl, Historico, Valor, Data FROM Lancamentos WHERE FilialEmpresa = ? AND Origem = ? AND Exercicio = ? AND PeriodoLote = ? AND Lote = ?", lDoc, iSeq, sConta, sCcl, sHistorico, dValor, dtData, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 5224

    'Le o primeiro lancamento a ser estornado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9443
    
    Do While lErro = AD_SQL_SUCESSO
        
        'insere o lancamento de estorno
        lErro = Comando_Executar(lComando1, "INSERT INTO LanPendente (FilialEmpresa, Origem, Exercicio, PeriodoLan, Doc, Seq, Lote, PeriodoLote, Data, Conta, Ccl, Historico, Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", tApuracao.iFilialEmpresa, tApuracao.sOrigem_Estorno, tApuracao.iExercicio, tApuracao.iPeriodo, lDoc, iSeq, tApuracao.iLote, tApuracao.iPeriodo, dtData, sConta, sCcl, sHistorico, -dValor)
        If lErro <> AD_SQL_SUCESSO Then Error 5225
            
        'Le o proximo lancamento a ser estornado
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9444

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Processa_Lancamentos_Estorno_Apura = SUCESSO

    Exit Function

Erro_Processa_Lancamentos_Estorno_Apura:

    Processa_Lancamentos_Estorno_Apura = Err
    
    Select Case Err

        Case 5222, 5223
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", lErro)
        
        Case 5224, 9443, 9444
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS", lErro, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Apuracao, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.iLote)

        Case 5225
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LANCAMENTO", lErro, tApuracao.iFilialEmpresa, tApuracao.sOrigem_Estorno, tApuracao.iExercicio, tApuracao.iPeriodo, lDoc, iSeq)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 154930)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Rotina_Apura_Periodos_Int(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo_Inicial As Integer, ByVal iPeriodo_Final As Integer, sContaResultado As String, sContaPonte As String, colContasApuracao As Collection, sHistorico As String, Optional ByVal iZeraRD As Integer = 0) As Long

Dim objFiliais As AdmFiliais
Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_Rotina_Apura_Periodos_Int

    lTransacao = 0

   'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 10670

    If iFilialEmpresa = EMPRESA_TODA Or iFilialEmpresa = Abs(giFilialAuxiliar) Then

        TelaAcompanhaBatch.dValorTotal = gcolFiliais.Count

        'se tiver selecionado a empresa, executa a apuracao para cada filial
        For Each objFiliais In gcolFiliais
    
            If objFiliais.iCodFilial <> EMPRESA_TODA And objFiliais.iCodFilial <> Abs(giFilialAuxiliar) Then
    
                lErro = Rotina_Apura_Periodos0(objFiliais.iCodFilial, iExercicio, iPeriodo_Inicial, iPeriodo_Final, sContaResultado, sContaPonte, colContasApuracao, sHistorico, iZeraRD)
                If lErro <> SUCESSO Then Error 10671
                
                TelaAcompanhaBatch.dValorAtual = TelaAcompanhaBatch.dValorAtual + 1
        
                TelaAcompanhaBatch.ProgressBar1.Value = CInt((TelaAcompanhaBatch.dValorAtual / TelaAcompanhaBatch.dValorTotal) * 100)
                
            End If
        
        Next
    
    Else
    
        'se tiver decidido apurar somente uma filial
        lErro = Rotina_Apura_Periodos0(iFilialEmpresa, iExercicio, iPeriodo_Inicial, iPeriodo_Final, sContaResultado, sContaPonte, colContasApuracao, sHistorico, iZeraRD)
        If lErro <> SUCESSO Then Error 10672
    
        TelaAcompanhaBatch.ProgressBar1.Value = 100
    
    End If

   'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 10673

    Rotina_Apura_Periodos_Int = SUCESSO
    
    'Alteracao Daniel em 07/05/02
    Call Rotina_Aviso(vbOKOnly, "AVISO_APURACAOPERIODO_EXECUTADO_SUCESSO")

    Exit Function

Erro_Rotina_Apura_Periodos_Int:

    Rotina_Apura_Periodos_Int = Err

    Select Case Err

        Case 10670
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 10671, 10672
            
        Case 10673
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154931)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Rotina_Apura_Periodos0(ByVal iFilialEmpresa As Integer, ByVal iExercicio As Integer, ByVal iPeriodo_Inicial As Integer, ByVal iPeriodo_Final As Integer, sContaResultado As String, sContaPonte As String, colContasApuracao As Collection, sHistorico As String, Optional ByVal iZeraRD As Integer = 0) As Long

Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim iExercicio1 As Integer
Dim iPeriodo1 As Integer
Dim tApuracao As typeApuracao
Dim iStatus As Integer
Dim lErro As Long

On Error GoTo Erro_Rotina_Apura_Periodos0

    tApuracao.iFilialEmpresa = iFilialEmpresa
    tApuracao.iExercicio = iExercicio
    tApuracao.iPeriodo_Inicial = iPeriodo_Inicial
    tApuracao.iPeriodo_Final = iPeriodo_Final
    tApuracao.sConta_Ponte = sContaPonte
    tApuracao.sHistorico = sHistorico
    tApuracao.sOrigem_Estorno = "EAP"
    tApuracao.sOrigem_Apuracao = "APP"
    tApuracao.sConta_Resultado = sContaResultado
    tApuracao.iZeraRD = iZeraRD
    tApuracao.iSeq = 1
    Set tApuracao.colContasApuracao = colContasApuracao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 5176

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 9400

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 10677

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 20456

    lErro = Comando_Executar(lComando1, "SELECT UsoCcl FROM Configuracao", tApuracao.iUsoCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 9401
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 9402

    'Pesquisa o Exercicio em questão
    lErro = Comando_ExecutarLockado(lComando, "SELECT Status, NumPeriodos, DataFim FROM Exercicios WHERE Exercicio = ?", iStatus, tApuracao.iNumPeriodos, tApuracao.dtData, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 5177

    'le o Exercicio
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5178

    'não permite a mudança no status do exercicio
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 5180

    'Se o Exercicio estiver fechado ==> erro
    If iStatus = EXERCICIO_FECHADO Then Error 5179
    
    'Pesquisa o ExercicioFilial em questão
    lErro = Comando_ExecutarPos(lComando2, "SELECT Status, LoteApuracao, ExisteLoteApuracao FROM ExerciciosFilial WHERE FilialEmpresa = ? AND Exercicio = ?", 0, tApuracao.iStatusExercicioFilial, tApuracao.iLote, tApuracao.iExisteLoteApuracao, iFilialEmpresa, iExercicio)
    If lErro <> AD_SQL_SUCESSO Then Error 10674

    'le o ExercicioFilial
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 10675

    'lock do ExercicioFilial
    lErro = Comando_LockExclusive(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 10676
    
    lErro = Rotina_Apura_Periodos1(tApuracao)
    If lErro <> SUCESSO Then Error 9481
    
    'Atualiza o exercicio indicando que nao esta mais apurado
    lErro = Comando_ExecutarPos(lComando3, "UPDATE ExerciciosFilial SET Status = ?, ExisteLoteApuracao=?", lComando2, EXERCICIO_ABERTO, NAO_EXISTE_LOTE_APURACAO_EXERCICIO)
    If lErro <> AD_SQL_SUCESSO Then Error 20457
    
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Rotina_Apura_Periodos0 = SUCESSO

    Exit Function

Erro_Rotina_Apura_Periodos0:

    Rotina_Apura_Periodos0 = Err

    Select Case Err

        Case 5176, 9400, 10677, 20456
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 5177, 5178
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIO", Err, iExercicio)

        Case 5179
            Call Rotina_Erro(vbOKOnly, "ERRO_EXERCICIO_FECHADO", Err, iExercicio)

        Case 5180
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCACAO_EXERCICIO", Err, iExercicio)

        Case 9401, 9402
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", Err)
            
        Case 9481

        Case 10674, 10675
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EXERCICIOSFILIAL", Err, iFilialEmpresa, iExercicio)
            
        Case 10676
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_EXERCICIOSFILIAL", Err, iFilialEmpresa, iExercicio)

        Case 20457
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EXERCICIOSFILIAL", Err, iExercicio, iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154932)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Private Function Rotina_Apura_Periodos1(tApuracao As typeApuracao) As Long

Dim lErro As Long
Dim iTipoConta As Integer
Dim lComando As Long

On Error GoTo Erro_Rotina_Apura_Periodos1

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 9486

    If tApuracao.iZeraRD = DESMARCADO Then

        'le a conta ponte
        lErro = Comando_ExecutarLockado(lComando, "SELECT TipoConta FROM PlanoConta WHERE Conta=?", iTipoConta, tApuracao.sConta_Ponte)
        If lErro <> AD_SQL_SUCESSO Then Error 9482
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 9483
        
        'lock da conta ponte
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 9484
    
        'guarda o centro de custo/lucro da conta ponte
        If tApuracao.iUsoCcl = CCL_USA_CONTABIL Then
        
            tApuracao.sCcl_ContaPonte = String(STRING_CCL, 0)
            
            lErro = Mascara_RetornaCcl(tApuracao.sConta_Ponte, tApuracao.sCcl_ContaPonte)
            If lErro <> AD_SQL_SUCESSO Then Error 9475
        
        Else
            tApuracao.sCcl_ContaPonte = ""
        End If
    
    End If
    
    'se o exericicio estava apurado ===> estorna-o
    If tApuracao.iExisteLoteApuracao = EXISTE_LOTE_APURACAO_EXERCICIO Then
    
        tApuracao.iPeriodo = tApuracao.iNumPeriodos

        lErro = Estorno_Apuracao_Exercicio(tApuracao)
        If lErro <> SUCESSO Then Error 5229

        tApuracao.sOrigem_Estorno = "EAP"
        tApuracao.sOrigem_Apuracao = "APP"

    End If
    
    
    'Apura a faixa de periodos passada como parametro
    lErro = Processa_Periodos_Apura(tApuracao)
    If lErro <> SUCESSO Then Error 5181
    
    Call Comando_Fechar(lComando)
    
    Rotina_Apura_Periodos1 = SUCESSO

    Exit Function

Erro_Rotina_Apura_Periodos1:

    Rotina_Apura_Periodos1 = Err

    Select Case Err

        Case 5181

        Case 9475
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_RetornaCcl", Err, tApuracao.sConta_Ponte)

        Case 9482, 9483
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANOCONTA", Err)
        
        Case 9484
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANOCONTA", Err, tApuracao.sConta_Ponte)
            
        Case 9486
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    

End Function

Function Processa_Periodos_Apura(tApuracao As typeApuracao) As Long

Dim lComando1 As Long, lComando2 As Long
Dim lComando3 As Long
Dim lErro As Long
Dim iApurado As Integer
Dim lDoc As Long

On Error GoTo Erro_Processa_Periodos_Apura

    lComando1 = 0
    lComando2 = 0
    lComando3 = 0

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5182

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 5183
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 10678
    
    tApuracao.lComando4 = Comando_Abrir()
    If tApuracao.lComando4 = 0 Then Error 9489

    'Pesquisa os periodos que serão apurados
    lErro = Comando_ExecutarPos(lComando1, "SELECT Periodo, DataFim FROM Periodo WHERE Exercicio = ? AND (Periodo >= ? AND Periodo <= ?)", 0, tApuracao.iPeriodo, tApuracao.dtData, tApuracao.iExercicio, tApuracao.iPeriodo_Inicial, tApuracao.iPeriodo_Final)
    If lErro <> AD_SQL_SUCESSO Then Error 5184

    'Le o primeiro periodo
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9445

    Do While lErro = AD_SQL_SUCESSO
    
        'loca o Periodo que será apurado
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 5186

        'Pesquisa o periodoFilial
        lErro = Comando_ExecutarPos(lComando3, "SELECT Lote, DocApuracao, ExisteApuracaoPeriodo FROM PeriodosFilial WHERE FilialEmpresa=? AND Exercicio = ? AND Periodo = ?", 0, tApuracao.iLote, tApuracao.lDoc, iApurado, tApuracao.iFilialEmpresa, tApuracao.iExercicio, tApuracao.iPeriodo)
        If lErro <> AD_SQL_SUCESSO Then Error 10679

        'Le o periodoFilial
        lErro = Comando_BuscarPrimeiro(lComando3)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 10680

        'loca o Periodo que será apurado
        lErro = Comando_LockExclusive(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 10681

        'converte o periodo para string
        tApuracao.sPeriodo = Format(tApuracao.iPeriodo, "00")
        
        'se existia um lote anterior de apuração para o periodo em questão ===> estorna-o
        If iApurado = EXISTE_LOTE_APURACAO_PERIODO Then

            'gera o lote de estorno de apuração para o periodo em questão caso exista
            lErro = Gera_Lote_Estorno_Apura(tApuracao)
            If lErro <> SUCESSO Then Error 5222

            'atualiza o lote de estorno
            lErro = Atualiza_Lote_Estorno_Apura(tApuracao)
            If lErro <> SUCESSO Then Error 5223

        End If
        
        '######################################################
        'Inserido por Wagner 31/03/2006
        lErro = CF("Voucher_Automatico1", tApuracao.iFilialEmpresa, tApuracao.iExercicio, tApuracao.iPeriodo, tApuracao.sOrigem_Apuracao, lDoc)
        If lErro <> SUCESSO Then Error 5223
        
        tApuracao.lDoc = lDoc - 1
        '######################################################

        'gera o lote de apuração para o periodo em questão
        lErro = Gera_Lote_Apura(tApuracao)
        If lErro <> SUCESSO Then Error 5187

        'atualiza o lote de apuração para o periodo em questão
        lErro = Atualiza_Lote_Apura(tApuracao)
        If lErro <> SUCESSO Then Error 5188

        'Atualiza o periodo indicando que foi apurado
        lErro = Comando_ExecutarPos(lComando2, "UPDATE PeriodosFilial SET Apurado = ?, Lote = Lote + 1, DataApuracao = ?, DocApuracao = ?, ExisteApuracaoPeriodo=? ", lComando3, PERIODO_APURADO, Date, tApuracao.lDoc, EXISTE_LOTE_APURACAO_PERIODO)
        If lErro <> AD_SQL_SUCESSO Then Error 5189

        'libera o Periodo apurado
        lErro = Comando_Unlock(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 5190

        'libera o PeriodoFilial apurado
        lErro = Comando_Unlock(lComando3)
        If lErro <> AD_SQL_SUCESSO Then Error 10682

        'Le o proximo periodo
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9446

    Loop
    
'    lErro = Resultado_Exclui_CodigoApuracao(tApuracao.lCodigoApuracao)
'    If lErro <> SUCESSO Then Error 9500

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(tApuracao.lComando4)

    Processa_Periodos_Apura = SUCESSO

    Exit Function

Erro_Processa_Periodos_Apura:

    Processa_Periodos_Apura = Err

    Select Case Err

        Case 5182, 5183, 9489, 10678
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", lErro)
        
        Case 5184, 9445, 9446
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO1", lErro)

        Case 5186
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODO", lErro, tApuracao.iExercicio, tApuracao.iPeriodo)

        Case 5187, 5188, 5222, 5223, 9500
        
        Case 5189
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PERIODOSFILIAL", lErro, tApuracao.iPeriodo, tApuracao.iExercicio, tApuracao.iFilialEmpresa)

        Case 5190
            Call Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_PERIODO", lErro, tApuracao.iPeriodo, tApuracao.iExercicio)

        Case 10679, 10680
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODOSFILIAL", tApuracao.iFilialEmpresa, tApuracao.iExercicio, tApuracao.iPeriodo)

        Case 10681
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PERIODOSFILIAL", lErro, tApuracao.iFilialEmpresa, tApuracao.iExercicio, tApuracao.iPeriodo)

        Case 10682
            Call Rotina_Erro(vbOKOnly, "ERRO_UNLOCK_PERIODOSFILIAL", lErro, tApuracao.iFilialEmpresa, tApuracao.iExercicio, tApuracao.iPeriodo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 154933)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(tApuracao.lComando4)

    Exit Function

End Function

'Function Resultado_Exclui_CodigoApuracao(lCodigoApuracao As Long) As Long
'
'Dim lErro As Long
'Dim lComando1 As Long
'Dim lComando2 As Long
'Dim lCodigoApuracao1 As Long
'
'On Error GoTo Erro_Resultado_Exclui_CodigoApuracao
'
'    lComando1 = Comando_Abrir()
'    If lComando1 = 0 Then Error 9494
'
'    lComando2 = Comando_Abrir()
'    If lComando2 = 0 Then Error 9495
'
'    'Pesquisa a tabela Resultado
'    lErro = Comando_ExecutarPos(lComando1, "SELECT CodigoApuracao FROM Resultado WHERE CodigoApuracao = ?", 0, lCodigoApuracao1, lCodigoApuracao)
'    If lErro <> AD_SQL_SUCESSO Then Error 9496
'
'    'Le o primeiro registro
'    lErro = Comando_BuscarPrimeiro(lComando1)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9497
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        'Exclui o registro em questão
'        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM Resultado", lComando1)
'        If lErro <> AD_SQL_SUCESSO Then Error 9498
'
'        'Le o proximo registro
'        lErro = Comando_BuscarProximo(lComando1)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9499
'
'    Loop
'
'    Call Comando_Fechar(lComando1)
'    Call Comando_Fechar(lComando2)
'
'    Resultado_Exclui_CodigoApuracao = SUCESSO
'
'    Exit Function
'
'Erro_Resultado_Exclui_CodigoApuracao:
'
'    Resultado_Exclui_CodigoApuracao = Err
'
'    Select Case Err
'
'        Case 9494, 9495
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 9496, 9497, 9499
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESULTADO", Err, lCodigoApuracao)
'
'        Case 9498
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RESULTADO", Err, lCodigoApuracao)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154934)
'
'    End Select
'
'    Call Comando_Fechar(lComando1)
'    Call Comando_Fechar(lComando2)
'
'    Exit Function
'
'End Function

Function Retorna_Ultimo_Dia_Periodo(ByVal iExercicio As Integer, ByVal iPeriodo As Integer, dtData As Date) As Long

Dim lComando1 As Long
Dim lErro As Long

On Error GoTo Erro_Retorna_Ultimo_Dia_Periodo

    lComando1 = 0
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 5268

    'Pesquisa o periodo em questão
    lErro = Comando_Executar(lComando1, "SELECT DataFim FROM Periodo WHERE Exercicio = ? AND Periodo = ?", dtData, iExercicio, iPeriodo)
    If lErro <> AD_SQL_SUCESSO Then Error 5269

    'Le o primeiro periodo
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 5270

    Call Comando_Fechar(lComando1)

    Retorna_Ultimo_Dia_Periodo = SUCESSO

    Exit Function

Erro_Retorna_Ultimo_Dia_Periodo:

    Retorna_Ultimo_Dia_Periodo = Err

    Select Case Err

        Case 5268
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 5269, 5270
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PERIODO", Err, iExercicio, iPeriodo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154935)

    End Select

    Call Comando_Fechar(lComando1)

    Exit Function

End Function

