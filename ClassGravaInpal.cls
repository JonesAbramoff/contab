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
Dim iApropriacao As Integer, vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Lancamento_Grava_BD

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 6012
    Next
    
    'Inicia Loop de GRAVACAO com Lançamentos do GRID
    For Each objLancamento_Detalhe In colLancamento_Detalhe
            
        'Alteração Daniel em 30/08/2002
        If Len(Trim(objLancamento_Detalhe.sHistorico)) = 0 Then Error 32279
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

        'William
        'INPAL: se conta comeca com 3 ou 4 entao exigir ccl
        If Len(Trim(objLancamento_Detalhe.sConta)) > 0 Then
            If Trim(left(objLancamento_Detalhe.sConta, 1)) = "3" Or Trim(left(objLancamento_Detalhe.sConta, 1)) = "4" Then
                If Len(Trim(objLancamento_Detalhe.sCcl)) = 0 Then Error 32254
            End If
        End If
        
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
            If lErro <> SUCESSO Then Error 32311
            
        End If
            
        objLancamento_Detalhe.dtData = objLancamento_Cabecalho.dtData
        objLancamento_Detalhe.iFilialEmpresa = objLancamento_Cabecalho.iFilialEmpresa
            
        'trata a parte de custo que pode ser tratado neste momento (Custo Medio e Custo Standard)
        lErro = CF("LanPendente_Trata_Produto", alComando(), objLancamento_Detalhe, iApropriacao)
        If lErro <> SUCESSO Then Error 20530
        
        If objLancamento_Detalhe.dtDataEstoque = 0 Then objLancamento_Detalhe.dtDataEstoque = DATA_NULA
            
        'Faz inserção do Lançamento na Tabela de Lançamentos Pendentes
        'reprocessamento (incluido Quantidade e DataEstoque )
        lErro = Comando_Executar(alComando(1), "INSERT INTO LanPendente (FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Lote,PeriodoLote,Data,Seq,Conta,Ccl,Valor,Historico,NumIntDoc,FilialCliForn,CliForn,Transacao,Aglutina,Produto, ApropriaCRProd, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Quantidade, DataEstoque, Status, Modelo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc, objLancamento_Cabecalho.iLote, objLancamento_Cabecalho.iPeriodoLote, objLancamento_Cabecalho.dtData, objLancamento_Detalhe.iSeq, objLancamento_Detalhe.sConta, objLancamento_Detalhe.sCcl, objLancamento_Detalhe.dValor, objLancamento_Detalhe.sHistorico, objLancamento_Detalhe.lNumIntDoc, objLancamento_Detalhe.iFilialCliForn, objLancamento_Detalhe.lCliForn, objLancamento_Detalhe.iTransacao, objLancamento_Detalhe.iAglutina, objLancamento_Detalhe.sProduto, objLancamento_Cabecalho.iCustoRealProducao, _
                                 objLancamento_Detalhe.lContaSimples, objLancamento_Detalhe.iSeqContraPartida, objLancamento_Detalhe.iEscaninho_Custo, objLancamento_Detalhe.dValorLivroAuxiliar, objLancamento_Detalhe.iClienteFornecedor, objLancamento_Detalhe.sDocOrigem, objLancamento_Detalhe.dQuantidade, objLancamento_Detalhe.dtDataEstoque, objLancamento_Detalhe.iStatus, objLancamento_Detalhe.sModelo)
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

'    'INPAL
'    'exigir lote dependendo de origem
'    If objLancamento_Cabecalho.iLote = 0 Then
'
'        Select Case objLancamento_Cabecalho.sOrigem
'
'            Case "TES", "FAT", "EST", "CR", "CP", "CTB"
'
'                vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_FALTOU_LOTE_CONTABIL")
'                If vbMsgBox <> vbYes Then
'
'                    vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_GRAVA_SEM_LOTE_CONTABIL")
'                    If vbMsgBox <> vbYes Then Error 32555
'
'                Else
'
'                    Error 32555
'
'                End If
'
'        End Select
'
'    End If

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
    
        Case 5790, 20530, 32555, 32311
    
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
    
        Case 32254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_OBRIGATORIA", Err, objLancamento_Detalhe.sConta)
    
    'Alteracao Daniel em 30/08/2002
        Case 32279
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_SEM_HISTORICO", Err)
    'Fim da Alteração Daniel em 30/08/02
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
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
Dim objHistPadrao As New ClassHistPadrao, vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Lancamento_Atualiza_BD

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 11552
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 11553
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 11554
    
    'Inicia Loop de GRAVACAO com Lançamentos do GRID
    For Each objLancamento_Detalhe In colLancamento_Detalhe
            
        'Alteração Daniel em 30/08/2002
        If Len(Trim(objLancamento_Detalhe.sHistorico)) = 0 Then Error 32279
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
        
        'William
        'INPAL: se conta comeca com 3 ou 4 entao exigir ccl
        If objLancamento_Detalhe.iStatus = 0 And Len(Trim(objLancamento_Detalhe.sConta)) > 0 Then
            If Trim(left(objLancamento_Detalhe.sConta, 1)) = "3" Or Trim(left(objLancamento_Detalhe.sConta, 1)) = "4" Then
                If Len(Trim(objLancamento_Detalhe.sCcl)) = 0 Then Error 32255
            End If
        End If

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
            If lErro <> SUCESSO Then Error 32311
            
        End If
        
        If giSetupUsoCcl = CCL_USA_CONTABIL Then
        
            sCcl = String(STRING_CCL, 0)
            
            lErro = Mascara_RetornaCcl(objLancamento_Detalhe.sConta, sCcl)
            If lErro <> SUCESSO Then Error 11565
            
            objLancamento_Detalhe.sCcl = sCcl
            
            lErro = CF("Ccl_Critica_Analitica", sCcl)
            If lErro <> SUCESSO Then Error 32311
            
        End If
            
        If objLancamento_Detalhe.dtDataEstoque = 0 Then objLancamento_Detalhe.dtDataEstoque = DATA_NULA
            
        'Faz inserção do Lançamento na Tabela de Lançamentos Pendentes
        lErro = Comando_Executar(lComando1, "INSERT INTO LanPendente (FilialEmpresa,Origem,Exercicio,PeriodoLan,Doc,Data,Seq,PeriodoLote,Lote,Conta,Ccl,Valor,Historico,NumIntDoc,FilialCliForn,CliForn,Transacao,Aglutina, Produto, ContaSimples, SeqContraPartida, EscaninhoCusto, ValorLivroAuxiliar, ClienteFornecedor, DocOrigem, Quantidade, DataEstoque, ApropriaCRProd, Status, Modelo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objLancamento_Cabecalho.iFilialEmpresa, objLancamento_Cabecalho.sOrigem, objLancamento_Cabecalho.iExercicio, objLancamento_Cabecalho.iPeriodoLan, objLancamento_Cabecalho.lDoc, objLancamento_Cabecalho.dtData, objLancamento_Detalhe.iSeq, objLancamento_Cabecalho.iPeriodoLan, 0, objLancamento_Detalhe.sConta, objLancamento_Detalhe.sCcl, objLancamento_Detalhe.dValor, objLancamento_Detalhe.sHistorico, objLancamento_Detalhe.lNumIntDoc, objLancamento_Detalhe.iFilialCliForn, objLancamento_Detalhe.lCliForn, objLancamento_Detalhe.iTransacao, objLancamento_Detalhe.iAglutina, objLancamento_Detalhe.sProduto, objLancamento_Detalhe.lContaSimples, objLancamento_Detalhe.iSeqContraPartida, objLancamento_Detalhe.iEscaninho_Custo, objLancamento_Detalhe.dValorLivroAuxiliar, objLancamento_Detalhe.iClienteFornecedor, objLancamento_Detalhe.sDocOrigem, objLancamento_Detalhe.dQuantidade, objLancamento_Detalhe.dtDataEstoque, objLancamento_Detalhe.iApropriaCRProd, objLancamento_Detalhe.iStatus, objLancamento_Detalhe.sModelo)
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

'    'INPAL
'    'exigir lote dependendo de origem
'    If objLancamento_Cabecalho.iLote = 0 Then
'
'        Select Case objLancamento_Cabecalho.sOrigem
'
'            Case "TES", "FAT", "EST", "CR", "CP", "CTB"
'
'                vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_FALTOU_LOTE_CONTABIL")
'                If vbMsgBox <> vbYes Then
'
'                    vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_GRAVA_SEM_LOTE_CONTABIL")
'                    If vbMsgBox <> vbYes Then Error 32555
'
'                Else
'
'                    Error 32555
'
'                End If
'
'        End Select
'
'    End If
    
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
    
        Case 32311
        
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
        
        Case 32255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_OBRIGATORIA", Err, objLancamento_Detalhe.sConta)
                        
        Case 32555
        
    'Alteracao Daniel em 30/08/2002
        Case 32279
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LANCAMENTO_SEM_HISTORICO", Err)
    'Fim da Alteração Daniel em 30/08/02
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Public Sub Atualiza_Lancamentos_MovimentoEstoque()
'Atualiza as quantidades em Lançamentos e Lançamentos pendentes, de acordo com os Movimentos de Estoque

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lTransacao As Long
Dim iTransacao As Integer
Dim lNumIntDoc As Long
Dim sProduto As String
Dim dCustoUnitario As Double
Dim dQuantidade As Double

On Error GoTo Erro_Atualiza_Lancamentos_MovimentoEstoque

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 5000

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 5001

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 5002

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 5003

    'inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 5004
    
    Open "Lancamentos.txt" For Output As #1
    
    sProduto = String(STRING_PRODUTO, 0)

    'pesquisa nome reduzido no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT Quantidade, Transacao, NumIntDoc, Produto FROM Lancamentos WHERE Valor = 0 AND Transacao <> 0 AND Produto NOT LIKE 'H%' AND Data <= '11-04-2002' AND Produto <>'' ORDER BY Produto", 0, dQuantidade, iTransacao, lNumIntDoc, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 5005

    'verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 5006

    Do While lErro = AD_SQL_SUCESSO

        lErro = CF("Retorna_CustoUnitario", iTransacao, lNumIntDoc, sProduto, dCustoUnitario)
        If lErro <> AD_SQL_SUCESSO Then gError 5014
        
        If dCustoUnitario = 0 Then
            Print #1, "Custo:" & dCustoUnitario & " Prod:" & sProduto & " NumInt:" & lNumIntDoc & " Trans:" & iTransacao
        End If
        
        'pesquisa Canal no BD
        lErro = Comando_ExecutarPos(lComando1, "UPDATE Lancamentos SET Valor = ?", lComando, dQuantidade * dCustoUnitario)
        If lErro <> AD_SQL_SUCESSO Then gError 5007

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 5008

    Loop

    sProduto = String(STRING_PRODUTO, 0)

    'pesquisa nome reduzido no BD
    lErro = Comando_ExecutarPos(lComando2, "SELECT Quantidade, Transacao, NumIntDoc, Produto FROM LanPendente WHERE Valor = 0 AND Transacao <> 0 AND Produto NOT LIKE 'H%' AND Data <= '11-04-2002' AND Produto <>'' ORDER BY Produto", 0, dQuantidade, iTransacao, lNumIntDoc, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 5009

    'verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 5010

    Do While lErro = AD_SQL_SUCESSO
        
        'o NumIntDoc 53684 está com problema de integridade no BD e por isso não está sendo levado em consideração
        If lNumIntDoc <> 53684 And iTransacao = 19 And sProduto = "Q000039" Then
        
            lErro = CF("Retorna_CustoUnitario", iTransacao, lNumIntDoc, sProduto, dCustoUnitario)
            If lErro <> AD_SQL_SUCESSO Then gError 5015
            
            If dCustoUnitario = 0 Then
                Print #1, "Custo:" & dCustoUnitario & " Prod:" & sProduto & " NumInt:" & lNumIntDoc & " Trans:" & iTransacao
            End If
            
            'pesquisa Canal no BD
            lErro = Comando_ExecutarPos(lComando3, "UPDATE LanPendente SET Valor = ?", lComando2, dQuantidade * dCustoUnitario)
            If lErro <> AD_SQL_SUCESSO Then gError 5011
        
        End If
            
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 5012
    
    Loop
 
    'confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 5013
    
    lErro = Rotina_Aviso(vbOKOnly, "Atualização executada com Sucesso!")
    
    Close #1
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando2)

    Exit Sub

Erro_Atualiza_Lancamentos_MovimentoEstoque:
        
        Close #1
        
        Select Case gErr

        Case 5000, 5001, 5002, 5003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 5004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 5005, 5006, 5008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS", gErr)

        Case 5007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANCAMENTO", gErr)

        Case 5009, 5010, 5012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LANCAMENTOS_PENDENTES", gErr)

        Case 5011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_LANPENDENTE", gErr)

        Case 5013
             lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 5014, 5015

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Sub

End Sub

Function Processa_Lancamento_Analitico_ContaDia_Cust(ByVal sConta As String) As Long

Dim lErro As Long

On Error GoTo Erro_Processa_Lancamento_Analitico_ContaDia_Cust

    If left(sConta, 2) = "14" Or left(sConta, 2) = "24" Then gError 12345

    Processa_Lancamento_Analitico_ContaDia_Cust = SUCESSO
    
    Exit Function
    
Erro_Processa_Lancamento_Analitico_ContaDia_Cust:

    Processa_Lancamento_Analitico_ContaDia_Cust = gErr

    Select Case gErr
    
        Case 12345
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CTA_COMPENSACAO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Exit Function
    
End Function

Public Function ImportCtb_Valida_Emp(ByVal sNomeArq As String, ByRef bValidaEmp As Boolean) As Long
    
    'If objFilial.sCgc <> Mid(sRegistro, 3, 14) And InStr(UCase(sNomeArq), "SP") = 0 Then gError 184470 'o arquivo de sp serve para as filiais PR e av brasil
    If InStr(UCase(sNomeArq), "SP") <> 0 Then
        bValidaEmp = False
    Else
        bValidaEmp = True
    End If
    ImportCtb_Valida_Emp = SUCESSO
End Function

Public Function ImportCtb_Insere_Lcto(ByVal sNomeArq As String, ByVal iFilialEmpresa As Integer, ByVal sCcl As String, ByRef bImportaLcto As Boolean) As Long
Dim bArqSP As Boolean
    bArqSP = (InStr(UCase(sNomeArq), "SP") <> 0)
    If (Not bArqSP) Or (iFilialEmpresa = 3 And left(sCcl, 1) = "3") Or (iFilialEmpresa = 1 And left(sCcl, 1) <> "3") Then
        bImportaLcto = True
    Else
        bImportaLcto = False
    End If
    ImportCtb_Insere_Lcto = SUCESSO
End Function

Public Function ImportCtb_Trata_Conta_Ccl(ByRef sConta As String, ByRef sCcl As String) As Long
    If left(sConta, 1) = "1" Or left(sConta, 1) = "2" Then sCcl = ""
    ImportCtb_Trata_Conta_Ccl = SUCESSO
End Function
