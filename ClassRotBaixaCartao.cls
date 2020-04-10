VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotBaixaCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function AdmExtFin_Obtem_Datas(ByVal lNumIntDocMov As Long, ByVal dtData As Date, dtDataMov As Date, dtDataBaixa As Date, dtDataCTB As Date) As Long
    dtDataMov = dtData
    dtDataBaixa = dtData
    dtDataCTB = dtData
    AdmExtFin_Obtem_Datas = SUCESSO
End Function

Private Function AdmExtFin_AtualizarExtratos2(ByVal objMov As ClassAdmExtFinMov, alComando() As Long, ByVal lCliente As Long, ByVal iFilialCliente As Integer) As Long
'atualiza um movto de conta corrente

Dim lErro As Long, objMovDet As ClassAdmExtFinMovDet, objMovCCI As New ClassMovContaCorrente
Dim dtData As Date, dValor As Double, iCodConta As Integer, iFilialEmpresa As Integer
Dim objContabAutomatica As ClassContabAutomatica, objRotBxCartaoCtb As ClassRotBxCartaoCtb
Dim objBaixaReceber As New ClassBaixaReceber, iIndice As Integer
Dim dValorCredito As Double
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim colInfoComissao As New colInfoComissao
Dim objContabil As ClassContabil
Dim lCodigo As Long
Dim sObservacao As String
Dim dtDataBaixa As Date
Dim dtDataCTB As Date
Dim dtDataMov As Date

On Error GoTo Erro_AdmExtFin_AtualizarExtratos2

    dValorCredito = 0
    sObservacao = ""
    For Each objMovDet In objMov.colMovDetCred
        dValorCredito = dValorCredito + objMovDet.dSaldoParcRec
        sObservacao = sObservacao & IIf(sObservacao <> "", "\", "") & "CC *" & right(objMovDet.sNumCartao, 6) & " R$" & Format(objMovDet.dSaldoParcRec, "STANDARD")
    Next
    
    sObservacao = left(sObservacao, STRING_DEB_REC_OBS)

    lErro = Comando_ExecutarPos(alComando(4), "SELECT Data, ValorLiq, CodConta, FilialEmpresa FROM AdmExtFinMov WHERE NumIntDoc = ?", 0, _
        dtData, dValor, iCodConta, iFilialEmpresa, objMov.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 202672

    lErro = Comando_BuscarProximo(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202673
    If lErro <> AD_SQL_SUCESSO Then gError 202674
    
    lErro = CF("AdmExtFin_Obtem_Datas", objMov.lNumIntDoc, dtData, dtDataMov, dtDataBaixa, dtDataCTB)
    If lErro <> SUCESSO Then gError 202675
    
    'preencher os dados do movimento da conta corrente
    With objMovCCI
        .dtDataMovimento = dtDataMov
        .dtDataContabil = dtDataCTB
        .dtDataBaixa = dtDataBaixa
        .dValor = dValor '- dValorCredito
        .iFilialEmpresa = iFilialEmpresa
        .iCodConta = iCodConta
        .iPortador = 0
        .iTipo = MOVCCI_EXTRATO_CARTAO_CRED
        .iTipoMeioPagto = DINHEIRO
        .lNumero = 0
        .lNumRefInterna = objMov.lNumIntDoc
        .sHistorico = "Recebimento de adm. de cartoes de credito"
        .iConciliado = NAO_CONCILIADO
        .iExcluido = MOVCONTACORRENTE_NAO_EXCLUIDO
    End With

    objMov.iCodConta = iCodConta
    objMov.dValorLiq = dValor
    
    If objMovCCI.dValor < -DELTA_VALORMONETARIO Then gError 209804
    
    'If objMovCCI.dValor > DELTA_VALORMONETARIO Then ' And objMov.colMovDet.Count > 0 Then
        'grava o movto de débito
        lErro = CF("MovCCI_Grava_PagRec", objMovCCI)
        If lErro <> SUCESSO Then gError 202675
    'End If
    
    objMov.lNumMovto = objMovCCI.lNumMovto
    
    If dValorCredito > DELTA_VALORMONETARIO And objMov.colMovDetCred.Count > 0 Then
    
        lErro = CF("Config_ObterAutomatico_EmTrans", "CRConfig", "NUM_PROX_CREDITO", "DebitosRecCli", "NumTitulo", lCodigo)
        If lErro <> SUCESSO Then gError 202675
            
        objDebitoRecCli.iFilialEmpresa = giFilialEmpresa
        objDebitoRecCli.lCliente = lCliente
        objDebitoRecCli.iFilial = iFilialCliente
        objDebitoRecCli.sSiglaDocumento = "NCR" 'Nota de crédito
        objDebitoRecCli.lNumTitulo = lCodigo
        objDebitoRecCli.dtDataEmissao = dtDataMov
        objDebitoRecCli.dValorTotal = dValorCredito
        objDebitoRecCli.dValorProdutos = dValorCredito
        objDebitoRecCli.sObservacao = sObservacao
        
        lErro = CF("DebitoRecCli_Grava_EmTrans", objDebitoRecCli, colInfoComissao, objContabil)
        If lErro <> SUCESSO Then gError 202675
        
    End If

    lErro = Comando_ExecutarPos(alComando(5), "UPDATE AdmExtFinMov SET NumMovto = ?, NumIntDebRecCli  = ?", alComando(4), objMovCCI.lNumMovto, objDebitoRecCli.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 202676

    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then

        Set objContabAutomatica = New ClassContabAutomatica
        Set objRotBxCartaoCtb = New ClassRotBxCartaoCtb
        Set objRotBxCartaoCtb.objMov = objMov
        Set objRotBxCartaoCtb.objDebitoRecCli = objDebitoRecCli

        lErro = objContabAutomatica.Inicializa_Contab(objRotBxCartaoCtb, MODULO_CONTASARECEBER, "BCR", objMovCCI.dtDataContabil, DATA_NULA)
        If lErro <> SUCESSO Then gError 202677

    End If
    
    objBaixaReceber.dtData = objMovCCI.dtDataBaixa
    objBaixaReceber.dtDataContabil = objMovCCI.dtDataContabil
    Set objBaixaReceber.colBaixaParcRec = New colBaixaParcRec

    For Each objMovDet In objMov.colMovDet

        'Gera uma baixaparcrec p\ a parcela a receber em questão
        objBaixaReceber.colBaixaParcRec.Add 0, 0, objMovDet.lNumIntParcRec, 0, STATUS_LANCADO, 0, 0, 0, objMovDet.dSaldoParcRec, objMovDet.dSaldoParcRec, COBRADOR_PROPRIA_EMPRESA

    Next

    'Grava a Baixa a receber
    lErro = CF("BaixaReceber_Grava_Recebimento", objBaixaReceber, objBaixaReceber.colBaixaParcRec, objMovCCI, Nothing, 0, 0, objContabAutomatica)
    If lErro <> SUCESSO Then gError 202678

    For iIndice = 1 To objMov.colMovDet.Count
        objMov.colMovDet.Item(iIndice).lNumIntBaixaParcRec = objBaixaReceber.colBaixaParcRec.Item(iIndice).lNumIntDoc
        
        lErro = Comando_Executar(alComando(6), "UPDATE AdmExtFinMovDet SET NumIntBaixaParcRec = ? WHERE NumIntDoc = ?", objMov.colMovDet.Item(iIndice).lNumIntBaixaParcRec, objMov.colMovDet.Item(iIndice).lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 202679
        
    Next
    
    If Not (objContabAutomatica Is Nothing) Then

        lErro = objContabAutomatica.Finaliza_Contab()
        If lErro <> SUCESSO Then gError 202680

    End If

    AdmExtFin_AtualizarExtratos2 = SUCESSO

    Exit Function

Erro_AdmExtFin_AtualizarExtratos2:

    AdmExtFin_AtualizarExtratos2 = gErr

    Select Case gErr
    
        Case 202672 To 202674
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMEXTFINMOV", gErr)
            
        Case 202675, 202677, 202678, 202680
        
        Case 202676
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ADMEXTFINMOV", gErr)
            
        Case 202679
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ADMEXTFINMOVDET", gErr)
            
        Case 209804 'dtData, dValor, iCodConta, iFilialEmpresa, objMov.lNumIntDoc
            Call Rotina_Erro(vbOKOnly, "ERRO_ADMEXTFINMOV_VALOR_NEGATIVO", gErr, objMov.lNumIntDoc, iCodConta, dtData, dValor)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202681)

    End Select

    Exit Function

End Function

Private Function AdmExtFin_AtualizarExtratos1(ByVal lNumIntArq As Long, ByVal iBandeira As Integer, alComando() As Long, objFrmAguarde As ClassFrmAguarde, objFrmAguardeTela As Object) As Long
'atualiza um arquivo

Dim lErro As Long
Dim dValorBruto As Double, dValorComissao As Double, sNumCartao As String, sAutorizacao As String, iNumParcela As Integer, dValor As Double
Dim dSaldo As Double, dValorLiq As Double, iQtdeEncontrada As Integer, iCodErro As Integer, lNumIntMovDet As Long
Dim iIgnorarErros As Integer, lNumIntMov As Long, lNumIntMovAnt As Long
Dim objMov As ClassAdmExtFinMov, objMovDet As ClassAdmExtFinMovDet, lNumIntParcRec As Long
Dim lCliente As Long, colParcUtilizadas As New Collection, vValorAux As Variant, bParcLoc As Boolean
Dim iFilialCliente As Integer, iBandeiraAux As Integer
Dim sEstabelecimento As String
Dim dtDataCompra As Date, lNumIntParcRecAux As Long
Dim dSaldoAux As Double, sNumCartaoAux1 As String, sNumCartaoAux2 As String

On Error GoTo Erro_AdmExtFin_AtualizarExtratos1
    
    lNumIntMovAnt = -1
    sNumCartao = String(STRING_NUMERO_CARTAO, 0)
    sAutorizacao = String(STRING_APROVACAO_CARTAO, 0)
    sEstabelecimento = String(STRING_REDE_ESTABELECIMENTO, 0)
    
    lErro = Comando_Executar(alComando(1), "SELECT MOV.Bandeira, MOV.Estabelecimento, MOV.NumIntDoc, MOV.IgnorarErros, MOV.ValorBruto, MOV.ValorComissao, MOVDET.NumCartao, MOVDET.NumParcela, MOVDET.Autorizacao, MOVDET.Valor, MOVDET.NumIntDoc, MOVDET.DataCompra FROM AdmExtFinMov MOV, AdmExtFinMovDet MOVDET WHERE MOV.NumIntDoc = MOVDET.NumIntMov AND MOV.NumMovto = 0 AND NumIntArq = ? ORDER BY MOV.NumIntDoc, MOVDET.NumIntDoc", _
        iBandeiraAux, sEstabelecimento, lNumIntMov, iIgnorarErros, dValorBruto, dValorComissao, sNumCartao, iNumParcela, sAutorizacao, dValor, lNumIntMovDet, dtDataCompra, lNumIntArq)
    If lErro <> AD_SQL_SUCESSO Then gError 202661

    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202662

    Do While lErro = AD_SQL_SUCESSO
    
        If iBandeira = 0 Then iBandeira = iBandeiraAux

        'OBTER CLIENTE
        lErro = Comando_Executar(alComando(7), "SELECT Cliente, FilialCli FROM Redes WHERE Bandeira = ? AND Estabelecimento = ? AND Cliente <> 0", lCliente, iFilialCliente, iBandeira, sEstabelecimento)
        If lErro <> AD_SQL_SUCESSO Then gError 202706
    
        lErro = Comando_BuscarProximo(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202707

        If lNumIntMovAnt <> lNumIntMov Then

            If lNumIntMovAnt <> -1 Then

                If objMov.colMovDet.Count <> 0 Or objMov.colMovDetCred.Count <> 0 Then

                    'grava o movto com as baixas
                    lErro = AdmExtFin_AtualizarExtratos2(objMov, alComando, lCliente, iFilialCliente)
                    If lErro <> SUCESSO Then gError 202663

                End If

            End If

            Set objMov = New ClassAdmExtFinMov
            objMov.lNumIntDoc = lNumIntMov

            lNumIntMovAnt = lNumIntMov

        End If

        'procura a parcela
        iQtdeEncontrada = 0
        
        If Abs(dValorBruto) > DELTA_VALORMONETARIO Then
            dValorLiq = Arredonda_Moeda((1 + (dValorComissao / dValorBruto)) * dValor / 100)
        Else
            dValorLiq = 0
        End If

        If left(sNumCartao, 5) <> "00000" Then
        
            sNumCartaoAux1 = Replace(sNumCartao, "*", "_") & "%"
            sNumCartaoAux2 = Mid(sNumCartaoAux1, 1, Len(sNumCartaoAux1) - 5) & "X" & Mid(sNumCartaoAux1, Len(sNumCartaoAux1) - 3)

            lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Saldo FROM ParcelasRec WHERE (NumeroCartao LIKE ? OR NumeroCartao LIKE ?) AND (AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR ? = '') AND DataTransacaoCartao BETWEEN ? AND ? AND Valor BETWEEN ? AND ? AND NumParcela = ? AND Status = ?", _
                lNumIntParcRecAux, dSaldoAux, sNumCartaoAux1, sNumCartaoAux2, sAutorizacao, "0" & sAutorizacao, "00" & sAutorizacao, "000" & sAutorizacao, "0000" & sAutorizacao, sAutorizacao, DateAdd("d", -28, dtDataCompra), DateAdd("d", 28, dtDataCompra), dValor * 0.0094, dValor * 0.0106, IIf(iNumParcela = 0, 1, iNumParcela), STATUS_ABERTO)
        Else
            lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Saldo FROM ParcelasRec WHERE (AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR ? = '') AND DataTransacaoCartao BETWEEN ? AND ? AND Valor BETWEEN ? AND ? AND NumParcela = ? AND Status = ?", _
                lNumIntParcRecAux, dSaldoAux, sAutorizacao, "0" & sAutorizacao, "00" & sAutorizacao, "000" & sAutorizacao, "0000" & sAutorizacao, sAutorizacao, DateAdd("d", -14, dtDataCompra), DateAdd("d", 14, dtDataCompra), dValor * 0.0094, dValor * 0.0106, IIf(iNumParcela = 0, 1, iNumParcela), STATUS_ABERTO)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 202664

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202665

        lNumIntParcRec = 0
        dSaldo = 0
        Do While lErro = AD_SQL_SUCESSO

            If dValorLiq > DELTA_VALORMONETARIO Then
                If Abs(dSaldoAux - dValorLiq) < 1 Or Abs((dSaldoAux - dValorLiq) / dValorLiq) < 0.01 Then
                    bParcLoc = False
                    For Each vValorAux In colParcUtilizadas
                        If lNumIntParcRecAux = vValorAux Then
                            bParcLoc = True 'Essa parcela já está contida em outro item desse arquivo e
                            'por isso não pode ser utilizada novamente (baixa duplicada)
                            Exit For
                        End If
                    Next
                    If Not bParcLoc Then
                        iQtdeEncontrada = iQtdeEncontrada + 1
                        lNumIntParcRec = lNumIntParcRecAux
                        dSaldo = dSaldoAux
                        colParcUtilizadas.Add lNumIntParcRec
                    End If
                End If
            End If
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202666

        Loop

        iCodErro = 0
        If iQtdeEncontrada = 0 Then iCodErro = ADMEXTFIN_ERRO_CV_SEMPARCELA
        If iQtdeEncontrada > 1 Then iCodErro = ADMEXTFIN_ERRO_CV_MUITASPARCELAS

        lErro = Comando_Executar(alComando(3), "UPDATE AdmExtFinMovDet SET CodErro = ?, NumIntParc = ? WHERE NumIntDoc = ?", iCodErro, lNumIntParcRec, lNumIntMovDet)
        If lErro <> AD_SQL_SUCESSO Then gError 202667

        If iCodErro = 0 Then
            Set objMovDet = New ClassAdmExtFinMovDet
            objMovDet.lNumIntDoc = lNumIntMovDet
            objMovDet.lNumIntParcRec = lNumIntParcRec
            objMovDet.dSaldoParcRec = dSaldo

            objMov.colMovDet.Add objMovDet
        Else
            Set objMovDet = New ClassAdmExtFinMovDet
            objMovDet.lNumIntDoc = lNumIntMovDet
            objMovDet.dSaldoParcRec = dValorLiq
            objMovDet.sNumCartao = sNumCartao
            objMovDet.sAutorizacao = sAutorizacao

            objMov.colMovDetCred.Add objMovDet
            'If iIgnorarErros = 0 Then gError 201052
        End If
        
        lNumIntParcRec = 0
        
        Call objFrmAguardeTela.ProcessouItem
        If objFrmAguarde.iCancelar = MARCADO Then Exit Do

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202668

    Loop

'    Set objFrmAguardeTela = Nothing
    If objFrmAguarde.iCancelar = MARCADO Then gError 202669

    If Not (objMov Is Nothing) Then
    
        If objMov.colMovDet.Count <> 0 Or objMov.colMovDetCred.Count <> 0 Then
    
            'grava o movto com as baixas
            lErro = AdmExtFin_AtualizarExtratos2(objMov, alComando, lCliente, iFilialCliente)
            If lErro <> SUCESSO Then gError 202670
    
        End If

    End If

    AdmExtFin_AtualizarExtratos1 = SUCESSO

    Exit Function

Erro_AdmExtFin_AtualizarExtratos1:

    AdmExtFin_AtualizarExtratos1 = gErr

    Select Case gErr
    
        Case 202661, 202662, 202668
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMEXTFINMOV", gErr)
            
        Case 202663, 202669, 202670
        
        Case 202664, 202665, 202666
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", gErr)
            
        Case 202667
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ADMEXTFINMOV", gErr)
        
        Case 202706, 202707
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REDES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202671)

    End Select

    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If
    
    Exit Function

End Function

Public Function AdmExtFin_AtualizarExtratos(ByVal objBaixaCartao As ClassBaixaCartao) As Long

Dim lErro As Long, lTransacao As Long, lNumIntArq As Long, iIndice As Integer
Dim alComando(0 To 7) As Long 'já usados de 0 a
Dim objArq As ClassAdmExtFinArqsLidos
Dim dtData As Date
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde

On Error GoTo Erro_AdmExtFin_AtualizarExtratos

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 202655

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202656
    Next
    
    dtData = DATA_NULA

    objFrmAguarde.iTotalItens = objBaixaCartao.iTotalReg
    Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)

    For Each objArq In objBaixaCartao.colArq
    
        If objArq.iNaoAtualizar = DESMARCADO Then

            'atualizar o arquivo
            lErro = AdmExtFin_AtualizarExtratos1(objArq.lNumIntDoc, objArq.iBandeira, alComando, objFrmAguarde, objFrmAguardeTela)
            If lErro <> SUCESSO Then gError 202657
            
            dtData = Date
            
        Else
            For iIndice = 1 To objArq.iTotalReg
                Call objFrmAguardeTela.ProcessouItem
                If objFrmAguarde.iCancelar = MARCADO Then gError 202657
            Next
        End If
        
        lErro = Comando_Executar(alComando(0), "UPDATE AdmExtFinArqsLidos SET NaoAtualizar = ?, DataAtualizado = ? WHERE NumIntDoc = ? ", _
            objArq.iNaoAtualizar, dtData, objArq.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 202658

    Next
    
    Set objFrmAguardeTela = Nothing

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 202659
    
    AdmExtFin_AtualizarExtratos = SUCESSO

    Exit Function

Erro_AdmExtFin_AtualizarExtratos:

    AdmExtFin_AtualizarExtratos = gErr

    Select Case gErr
           
        Case 202655
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 202656
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 202657
        
        Case 202658
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ADMEXTFINARQSLIDOS", gErr)

        Case 202659
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202660)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If

    Exit Function

End Function

Private Function AdmExtFin_ValidarExtratos1(ByVal lNumIntArq As Long, lNumErros As Long, alComando() As Long) As Long

Dim lErro As Long, iStatus As Integer
Dim dValorBruto As Double, dValorComissao As Double, sNumCartao As String, sAutorizacao As String, iNumParcela As Integer, dValor As Double
Dim dSaldo As Double, dValorLiq As Double, iQtdeEncontrada As Integer, iCodErro As Integer, lNumIntMovDet As Long
Dim iQtdeNaoAberta As Integer, sNumCartaoParc As String, bNumCartaoNaoBate As Boolean
Dim lNumIntParc As Long, colParcUtilizadas As New Collection, vValorAux As Variant, bParcLoc As Boolean
Dim dtDataCompra As Date, lNumIntParcAux As Long, sNumCartaoAux1 As String, sNumCartaoAux2 As String

On Error GoTo Erro_AdmExtFin_ValidarExtratos1

    sNumCartao = String(STRING_NUMERO_CARTAO, 0)
    sAutorizacao = String(STRING_APROVACAO_CARTAO, 0)
    lErro = Comando_Executar(alComando(1), "SELECT MOV.ValorBruto, MOV.ValorComissao, MOVDET.NumCartao, MOVDET.NumParcela, MOVDET.Autorizacao, MOVDET.Valor, MOVDET.NumIntDoc, MOVDET.DataCompra FROM AdmExtFinMov MOV, AdmExtFinMovDet MOVDET WHERE MOV.NumIntDoc = MOVDET.NumIntMov AND MOV.NumMovto = 0 AND NumIntArq = ? ORDER BY MOVDET.NumIntDoc", _
        dValorBruto, dValorComissao, sNumCartao, iNumParcela, sAutorizacao, dValor, lNumIntMovDet, dtDataCompra, lNumIntArq)
    If lErro <> AD_SQL_SUCESSO Then gError 202681

    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202682

    Do While lErro = AD_SQL_SUCESSO
            
        iQtdeEncontrada = 0
        iQtdeNaoAberta = 0
        bNumCartaoNaoBate = False
        
        If Abs(dValorBruto) > DELTA_VALORMONETARIO Then
            dValorLiq = Arredonda_Moeda((1 + (dValorComissao / dValorBruto)) * dValor / 100)
        Else
            dValorLiq = 0
        End If
        
        sNumCartaoParc = String(STRING_NUMERO_CARTAO, 0)
        If left(sNumCartao, 5) <> "00000" Then
        
            sNumCartaoAux1 = Replace(sNumCartao, "*", "_") & "%"
            sNumCartaoAux2 = Mid(sNumCartaoAux1, 1, Len(sNumCartaoAux1) - 5) & "X" & Mid(sNumCartaoAux1, Len(sNumCartaoAux1) - 3)
        
            lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Saldo, Status, NumeroCartao FROM ParcelasRecTodas WHERE (NumeroCartao LIKE ? OR NumeroCartao LIKE ?) AND (AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR ? = '') AND DataTransacaoCartao BETWEEN ? AND ? AND Valor BETWEEN ? AND ? AND NumParcela = ?", _
                lNumIntParcAux, dSaldo, iStatus, sNumCartaoParc, sNumCartaoAux1, sNumCartaoAux2, sAutorizacao, "0" & sAutorizacao, "00" & sAutorizacao, "000" & sAutorizacao, "0000" & sAutorizacao, sAutorizacao, DateAdd("d", -28, dtDataCompra), DateAdd("d", 28, dtDataCompra), dValor * 0.0094, dValor * 0.0106, IIf(iNumParcela = 0, 1, iNumParcela))
        Else
            lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Saldo, Status, NumeroCartao FROM ParcelasRecTodas WHERE (AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR AprovacaoCartao = ? OR ? = '') AND DataTransacaoCartao BETWEEN ? AND ? AND Valor BETWEEN ? AND ? AND NumParcela = ?", _
                lNumIntParcAux, dSaldo, iStatus, sNumCartaoParc, sAutorizacao, "0" & sAutorizacao, "00" & sAutorizacao, "000" & sAutorizacao, "0000" & sAutorizacao, sAutorizacao, DateAdd("d", -14, dtDataCompra), DateAdd("d", 14, dtDataCompra), dValor * 0.0094, dValor * 0.0106, IIf(iNumParcela = 0, 1, iNumParcela))
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 202683

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202684

        lNumIntParc = 0
        Do While lErro = AD_SQL_SUCESSO
        
            'dValorLiq = Arredonda_Moeda((1 + (dValorComissao / dValorBruto)) * dValor / 100)

            If dValorLiq > DELTA_VALORMONETARIO Then
                If (InStr(sNumCartaoParc, sNumCartao) <> 0 Or left(sNumCartao, 5) = "00000") And Abs(dSaldo - dValorLiq) < 1 Or Abs((dSaldo - dValorLiq) / dValorLiq) < 0.01 And iStatus = STATUS_ABERTO Then
                    bParcLoc = False
                    For Each vValorAux In colParcUtilizadas
                        If lNumIntParcAux = vValorAux Then
                            bParcLoc = True 'Essa parcela já está contida em outro item desse arquivo e
                            'por isso não pode ser utilizada novamente (baixa duplicada)
                            Exit For
                        End If
                    Next
                    If Not bParcLoc Then
                        iQtdeEncontrada = iQtdeEncontrada + 1
                        lNumIntParc = lNumIntParcAux
                        colParcUtilizadas.Add lNumIntParc
                    End If
                End If
            End If
            
            If iStatus <> STATUS_ABERTO Then iQtdeNaoAberta = iQtdeNaoAberta + 1
            If InStr(sNumCartaoParc, sNumCartao) = 0 And left(sNumCartao, 5) <> "00000" Then bNumCartaoNaoBate = True
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202685

        Loop

        iCodErro = 0
        If iQtdeEncontrada = 0 Then iCodErro = ADMEXTFIN_ERRO_CV_SEMPARCELA
        If iQtdeEncontrada > 1 Then iCodErro = ADMEXTFIN_ERRO_CV_MUITASPARCELAS
        If iQtdeEncontrada = 0 And iQtdeNaoAberta <> 0 Then iCodErro = ADMEXTFIN_ERRO_CV_PARCNAOABERTA
        If iQtdeEncontrada = 0 And iQtdeNaoAberta = 0 And bNumCartaoNaoBate Then iCodErro = ADMEXTFIN_ERRO_CV_DIFNUMCARTAO

        lErro = Comando_Executar(alComando(3), "UPDATE AdmExtFinMovDet SET CodErro = ?, NumIntParc = ? WHERE NumIntDoc = ?", iCodErro, lNumIntParc, lNumIntMovDet)
        If lErro <> AD_SQL_SUCESSO Then gError 202686
        
        lNumIntParc = 0

        If iCodErro <> 0 Then lNumErros = lNumErros + 1

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202687

    Loop

    AdmExtFin_ValidarExtratos1 = SUCESSO

    Exit Function

Erro_AdmExtFin_ValidarExtratos1:

    AdmExtFin_ValidarExtratos1 = gErr

    Select Case gErr

        Case 202681, 202682, 202687
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMEXTFINMOV", gErr)
        
        Case 202683, 202684, 202685
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", gErr)
            
        Case 202686
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ADMEXTFINMOV", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202688)

    End Select

    Exit Function

End Function

Public Function AdmExtFin_ValidarExtratos(ByVal iFilialEmpresa As Integer, ByVal iBandeira As Integer, ByVal dtImportacaoIni As Date, ByVal dtImportacaoFim As Date) As Long

Dim lErro As Long, lTransacao As Long, lNumIntArq As Long, lNumErros As Long, iIndice As Integer
Dim alComando(0 To 20) As Long 'já usados de 0 a

On Error GoTo Erro_AdmExtFin_ValidarExtratos

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 202689

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202690
    Next

    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc FROM AdmExtFinArqsLidos WHERE NaoAtualizar = 0 AND FilialEmpresa = ? AND Bandeira = ? AND DataAtualizado = ? AND (DataImportacao >= ? OR ?=?) AND (DataImportacao <= ? OR ?=?) ORDER BY NumIntDoc", _
    lNumIntArq, iFilialEmpresa, iBandeira, DATA_NULA, dtImportacaoIni, dtImportacaoIni, DATA_NULA, dtImportacaoFim, dtImportacaoFim, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 202691

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202692

    Do While lErro = AD_SQL_SUCESSO

        'valida o arquivo
        lNumErros = 0
        lErro = AdmExtFin_ValidarExtratos1(lNumIntArq, lNumErros, alComando)
        If lErro <> SUCESSO Then gError 202693

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202694

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 202695

    AdmExtFin_ValidarExtratos = SUCESSO

    Exit Function

Erro_AdmExtFin_ValidarExtratos:

    AdmExtFin_ValidarExtratos = gErr

    Select Case gErr

        Case 202689
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 202690
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 202691, 202692, 202694
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMEXTFINARQSLIDOS", gErr)
        
        Case 202693

        Case 202695
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202696)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function AdmExtFin_ImportarExtratos(ByVal iFilialEmpresa As Integer, ByVal iBandeira As Integer, ByVal objTela As Object) As Long

Dim lErro As Long, sDiretorio As String
Dim objFSO As New FileSystemObject, iIndice As Integer
Dim objFolder As Folder, lTransacao As Long, lNumIntArq As Long
Dim objFile As File, sFileName As String, lNumIntDoc As Long
Dim ts As TextStream, bArqAberto As Boolean, sRegistro As String
Dim objRotBaixaCartaoRede As ClassRotBxCartRede
Dim alComando(0 To 20) As Long 'já usados de 0 a 8

On Error GoTo Erro_AdmExtFin_ImportarExtratos

    bArqAberto = False

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 201046

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201047
    Next

    'ler a tabela redes para obter o diretorio de onde importar os arquivos
    sDiretorio = String(STRING_REDE_DIRIMPORTACAOEXTRATO, 0)
    lErro = Comando_Executar(alComando(0), "SELECT DirImportacaoExtrato FROM Redes WHERE FilialEmpresa = ? AND Bandeira = ? AND Estabelecimento <> ?", sDiretorio, iFilialEmpresa, iBandeira, "")
    If lErro <> AD_SQL_SUCESSO Then gError 201048

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201049
    If lErro <> AD_SQL_SUCESSO Then gError 201050

    If sDiretorio = "" Then gError 201099
    
    'Pega todos os aquivos da pasta
    Set objFolder = objFSO.GetFolder(sDiretorio)

    'Para cada arquivo
    For Each objFile In objFolder.Files

        'Pega o Nome
        sFileName = objFile.Name

        'verifica se o arquivo já foi importado anteriormente
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc FROM AdmExtFinArqsLidos WHERE NomeArq = ? AND Bandeira = ? AND NaoAtualizar <> 1 ", lNumIntDoc, sFileName, iBandeira)
        If lErro <> AD_SQL_SUCESSO Then gError 201052

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201053

        If lErro <> AD_SQL_SUCESSO Then

            Select Case iBandeira
            
                Case ADMEXTFIN_VARIAS 'Arquivo com múltiplas banderiras da Cielo
                    Set objRotBaixaCartaoRede = New ClassRotBxCartCielo

                Case ADMEXTFIN_AMEX
                    Set objRotBaixaCartaoRede = New ClassRotBxCartAmex

                Case ADMEXTFIN_VISANET
                    Set objRotBaixaCartaoRede = New ClassRotBxCartVisanet

                Case ADMEXTFIN_REDECARD
                    Set objRotBaixaCartaoRede = New ClassRotBxCartRedecard

                Case Else
                    gError 201051

            End Select

            bArqAberto = False
            'abrir arquivo texto
            Set ts = objFSO.OpenTextFile(sDiretorio & IIf(right(sDiretorio, 1) = "\", "", "\") & sFileName, 1, 0)
            bArqAberto = True

            With objRotBaixaCartaoRede
                .objAux.sNomeArq = sFileName
                .objAux.sDiretorio = sDiretorio
                .objAux.iBandeira = iBandeira
                .objAux.iFilialEmpresa = iFilialEmpresa
                .objAux.bPularArquivo = False
                Set .objAux.colMov = New Collection
            End With

            lErro = objRotBaixaCartaoRede.ImportarArqInicio(alComando)
            If lErro <> SUCESSO Then gError 201061

            If objRotBaixaCartaoRede.objAux.bPularArquivo = False Then

                'obter um numint para o arquivo (lNumIntArq)
                lErro = CF("Config_ObterNumInt", "CRConfig", "NUM_PROX_ROTBXCARTAO", lNumIntArq)
                If lErro <> SUCESSO Then gError 201070

                objRotBaixaCartaoRede.objAux.lNumIntArq = lNumIntArq

                'Até chegar ao fim do arquivo
                Do While Not ts.AtEndOfStream

                    'Busca o próximo registro do arquivo
                    sRegistro = ts.ReadLine

                    If Len(Trim(sRegistro)) <> 0 Then
                    
                        lErro = objRotBaixaCartaoRede.ImportarArqTrataRegistro(sRegistro, alComando)
                        If lErro <> SUCESSO Then gError 201062

                    End If

                Loop

                lErro = objRotBaixaCartaoRede.ImportarArqFim(alComando)
                If lErro <> SUCESSO Then gError 201063

                'inserir registro em AdmExtFinArqsLidos
                lErro = Comando_Executar(alComando(2), "INSERT INTO AdmExtFinArqsLidos (NumIntDoc, NomeArq, DataImportacao, HoraImportacao, Usuario, DataAtualizado, FilialEmpresa, Bandeira) VALUES (?,?,?,?,?,?,?,?)", _
                    lNumIntArq, sFileName, gdtDataHoje, CDbl(Time), gsUsuario, DATA_NULA, giFilialEmpresa, iBandeira)
                If lErro <> AD_SQL_SUCESSO Then gError 201071

            End If

            'fechar arquivo texto
            ts.Close
            bArqAberto = False

        End If

        lErro = objTela.Processa_Registro
        If lErro <> SUCESSO Then gError 201063

    Next

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 201064

    'Para cada arquivo
    For Each objFile In objFolder.Files
        Call Copia_Arquivo(sDiretorio & IIf(right(sDiretorio, 1) = "\", "", "\") & objFile.Name, sDiretorio & IIf(right(sDiretorio, 1) = "\", "", "\") & "Lidos\" & objFile.Name)
        Kill sDiretorio & IIf(right(sDiretorio, 1) = "\", "", "\") & objFile.Name
    Next

    AdmExtFin_ImportarExtratos = SUCESSO

    Exit Function

Erro_AdmExtFin_ImportarExtratos:

    AdmExtFin_ImportarExtratos = gErr

    Select Case gErr

        Case 201046
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 201047
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 201048, 201049, 201050
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIRETORIO_REDE", gErr, iBandeira, iFilialEmpresa)
       
        Case 201051
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_REDE_NAO_IMPLEMENTADA", gErr, iBandeira)

        Case 201052, 201053
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMEXTFINARQSLIDOS", gErr, iBandeira, iFilialEmpresa)

        Case 201061 To 201063
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTACAO_ARQUIVO", gErr, sFileName)
        
        Case 201070, 201071
        
        Case 201099
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_REDE_NAO_CONFIG", gErr, iBandeira, iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201040)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function AdmExtFin_Le(ByVal iFilialEmpresa As Integer, ByVal iBandeira As Integer, iQtd As Integer) As Long

Dim lErro As Long, sDiretorio As String
Dim objFSO As New FileSystemObject, iIndice As Integer
Dim objFolder As Folder
Dim alComando(0 To 0) As Long
Dim lCodigo As Long

On Error GoTo Erro_AdmExtFin_Le

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201047
    Next

    'ler a tabela redes para obter o diretorio de onde importar os arquivos
    sDiretorio = String(STRING_REDE_DIRIMPORTACAOEXTRATO, 0)
    lErro = Comando_Executar(alComando(0), "SELECT DirImportacaoExtrato, Codigo FROM Redes WHERE FilialEmpresa = ? AND Bandeira = ? AND Estabelecimento <> ?", sDiretorio, lCodigo, iFilialEmpresa, iBandeira, "")
    If lErro <> AD_SQL_SUCESSO Then gError 201048

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201049
    If lErro <> AD_SQL_SUCESSO Then gError 201050
    
    If Len(Trim(sDiretorio)) = 0 Then gError 202250

    If Len(Trim(Dir(sDiretorio, vbDirectory))) = 0 Then gError 202251

    'Pega todos os aquivos da pasta
    Set objFolder = objFSO.GetFolder(sDiretorio)
    
    iQtd = objFolder.Files.Count

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    AdmExtFin_Le = SUCESSO

    Exit Function

Erro_AdmExtFin_Le:

    AdmExtFin_Le = gErr

    Select Case gErr
 
        Case 201047
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 201048, 201049, 201050
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIRETORIO_REDE", gErr, iBandeira, iFilialEmpresa)

        Case 202250
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO_REDE", gErr, lCodigo, iFilialEmpresa)

        Case 202251
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, sDiretorio)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201040)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function BaixaCartao_Le(ByVal objBaixaCartao As ClassBaixaCartao) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 3) As Long
Dim objArq As ClassAdmExtFinArqsLidos
Dim objMov As ClassAdmExtFinMov
Dim objMovDet As ClassAdmExtFinMovDet
Dim tArq As typeAdmExtFinArqsLidos
Dim tMov As typeAdmExtFinMov
Dim tMovDet As typeAdmExtFinMovDet
Dim objCtaCorr As ClassContasCorrentesInternas
Dim colCtaCorr As New Collection
Dim bAchou As Boolean
Dim dSaldoParc As Double
Dim lNumTitulo As Long
Dim lNumIntTitulo As Long

On Error GoTo Erro_BaixaCartao_Le

    lErro = AdmExtFin_ValidarExtratos(objBaixaCartao.iFilialEmpresa, objBaixaCartao.iBandeira, objBaixaCartao.dtDataDe, objBaixaCartao.dtDataAte)
    If lErro <> SUCESSO Then gError 202590

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202591
    Next
    
    tArq.sNomeArq = String(255, 0)

    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc, NomeArq, DataImportacao, HoraImportacao FROM AdmExtFinArqsLidos WHERE NaoAtualizar = 0 AND FilialEmpresa = ? AND Bandeira = ? AND DataAtualizado = ? AND (DataImportacao >= ? OR ?=?) AND (DataImportacao <= ? OR ?=?) ", _
        tArq.lNumIntDoc, tArq.sNomeArq, tArq.dtDataImportacao, tArq.dHoraImportacao, objBaixaCartao.iFilialEmpresa, objBaixaCartao.iBandeira, DATA_NULA, objBaixaCartao.dtDataDe, objBaixaCartao.dtDataDe, DATA_NULA, objBaixaCartao.dtDataAte, objBaixaCartao.dtDataAte, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 202592

    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202593

    Do While lErro = AD_SQL_SUCESSO

        Set objArq = New ClassAdmExtFinArqsLidos

        objArq.lNumIntDoc = tArq.lNumIntDoc
        objArq.sNomeArq = tArq.sNomeArq
        objArq.dtDataImportacao = tArq.dtDataImportacao
        objArq.dHoraImportacao = tArq.dHoraImportacao
        objArq.iBandeira = objBaixaCartao.iBandeira
        objArq.iFilialEmpresa = objBaixaCartao.iFilialEmpresa
        
        'Alocação de espaço no buffer
        tMov.sEstabelecimento = String(255, 0)
    
        'Le a tabelaAdmExtFinMov
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, NumIntArq, Estabelecimento, Tipo, CodConta, Data, " & _
                    "ValorBruto, ValorComissao, ValorRejeitado, ValorLiq, IgnorarErros, FilialEmpresa, NumMovto, Bandeira FROM AdmExtFinMov WHERE NumIntArq= ? AND NumMovto = 0", _
                    tMov.lNumIntDoc, tMov.lNumIntArq, tMov.sEstabelecimento, tMov.iTipo, _
                    tMov.iCodConta, tMov.dtData, tMov.dValorBruto, tMov.dValorComissao, tMov.dValorRejeitado, _
                    tMov.dValorLiq, tMov.iIgnorarErros, tMov.iFilialEmpresa, tMov.lNumMovto, _
                    tMov.iBandeira, objArq.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 202594
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202595
    
        Do While lErro <> AD_SQL_SEM_DADOS
    
            Set objMov = New ClassAdmExtFinMov
            
            objMov.lNumIntDoc = tMov.lNumIntDoc
            objMov.lNumIntArq = tMov.lNumIntArq
            objMov.sEstabelecimento = tMov.sEstabelecimento
            objMov.iTipo = tMov.iTipo
            objMov.iCodConta = tMov.iCodConta
            objMov.dtData = tMov.dtData
            objMov.dValorBruto = tMov.dValorBruto
            objMov.dValorComissao = tMov.dValorComissao
            objMov.dValorRejeitado = tMov.dValorRejeitado
            objMov.dValorLiq = tMov.dValorLiq
            objMov.iFilialEmpresa = tMov.iFilialEmpresa
            objMov.lNumMovto = tMov.lNumMovto
            objMov.iBandeira = tMov.iBandeira
        
            bAchou = False
            For Each objCtaCorr In colCtaCorr
                If objCtaCorr.iCodigo = tMov.iCodConta Then
                    bAchou = True
                    Exit For
                End If
            Next
        
            If Not bAchou Then
                Set objCtaCorr = New ClassContasCorrentesInternas
                lErro = CF("ContaCorrenteInt_Le", tMov.iCodConta, objCtaCorr)
                If lErro <> SUCESSO Then gError 202596
            End If
            objMov.sNomeCtaCorrente = objCtaCorr.sNomeReduzido
        
            'Alocação de espaço no buffer
            tMovDet.sNumCartao = String(255, 0)
            tMovDet.sAutorizacao = String(255, 0)
            tMovDet.sRO = String(255, 0)
            tMovDet.sNSU = String(255, 0)
            tMovDet.sObservacao = String(255, 0)
        
            'Le a tabelaAdmExtFinMovDet
            lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, NumIntMov, Tipo, Valor, NumCartao, DataCompra, " & _
                        "NumParcela, Autorizacao, RO, NSU, NumIntParc, CodErro, Observacao, " & _
                        "NumIntBaixaParcRec FROM AdmExtFinMovDet WHERE NumIntMov= ? ", _
                        tMovDet.lNumIntDoc, tMovDet.lNumIntMov, tMovDet.iTipo, tMovDet.dValor, _
                        tMovDet.sNumCartao, tMovDet.dtDataCompra, tMovDet.iNumParcela, tMovDet.sAutorizacao, tMovDet.sRO, _
                        tMovDet.sNSU, tMovDet.lNumIntParc, tMovDet.iCodErro, tMovDet.sObservacao, tMovDet.lNumIntBaixaParcRec, _
                        objMov.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 202597
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202598
        
            Do While lErro <> AD_SQL_SEM_DADOS
        
                Set objMovDet = New ClassAdmExtFinMovDet
                
                tMovDet.dValor = tMovDet.dValor / 100
        
                objMovDet.lNumIntDoc = tMovDet.lNumIntDoc
                objMovDet.lNumIntMov = tMovDet.lNumIntMov
                objMovDet.iTipo = tMovDet.iTipo
                objMovDet.dValor = tMovDet.dValor
                objMovDet.sNumCartao = tMovDet.sNumCartao
                objMovDet.dtDataCompra = tMovDet.dtDataCompra
                objMovDet.iNumParcela = tMovDet.iNumParcela
                objMovDet.sAutorizacao = tMovDet.sAutorizacao
                objMovDet.sRO = tMovDet.sRO
                objMovDet.sNSU = tMovDet.sNSU
                objMovDet.lNumIntParc = tMovDet.lNumIntParc
                objMovDet.iCodErro = tMovDet.iCodErro
                objMovDet.lNumIntBaixaParcRec = tMovDet.lNumIntBaixaParcRec
                
                objMovDet.iQtdParc = 1
                objMovDet.dVlrParc = objMovDet.dValor
                
                If objMovDet.lNumIntParc <> 0 And objMovDet.iCodErro = 0 Then
                    lErro = Comando_Executar(alComando(3), "SELECT P.Saldo, T.NumTitulo, T.NumIntDoc FROM ParcelasRecTodas AS P, TitulosRecTodos AS T WHERE T.NumIntDoc = P.NumIntTitulo AND P.NumIntDoc = ? ", dSaldoParc, lNumTitulo, lNumIntTitulo, objMovDet.lNumIntParc)
                    If lErro <> AD_SQL_SUCESSO Then gError 202599
                
                    'Busca Primeiro
                    lErro = Comando_BuscarPrimeiro(alComando(3))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202600
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objMovDet.dVlrParcEnc = dSaldoParc
                        objMovDet.lNumTitulo = lNumTitulo
                        objMovDet.lNumIntTitulo = lNumIntTitulo
                        objMovDet.iQtdParcEnc = 1
                    End If

                End If
                
                objMov.dVlrParc = objMov.dVlrParc + objMovDet.dVlrParc
                objMov.dVlrParcEnc = objMov.dVlrParcEnc + objMovDet.dVlrParcEnc
                objMov.iQtdParc = objMov.iQtdParc + objMovDet.iQtdParc
                objMov.iQtdParcEnc = objMov.iQtdParcEnc + objMovDet.iQtdParcEnc
                objMov.iTotalReg = objMov.iTotalReg + objMovDet.iQtdParc
                
                objMov.colMovDet.Add objMovDet
        
                'Busca Próximo
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202601
       
            Loop

            objMov.dVlrDep = objMov.dValorLiq
            objMov.iQtdDep = 1

            objArq.dVlrParc = objMov.dVlrParc + objArq.dVlrParc
            objArq.dVlrParcEnc = objMov.dVlrParcEnc + objArq.dVlrParcEnc
            objArq.iQtdParc = objMov.iQtdParc + objArq.iQtdParc
            objArq.iQtdParcEnc = objMov.iQtdParcEnc + objArq.iQtdParcEnc
            objArq.dVlrDep = objMov.dVlrDep + objArq.dVlrDep
            objArq.iQtdDep = objMov.iQtdDep + objArq.iQtdDep
            objArq.iTotalReg = objMov.iTotalReg + objArq.iTotalReg
        
            objArq.colMov.Add objMov
        
            'Busca Próximo
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202602
        
        Loop
        
        objBaixaCartao.iTotalReg = objBaixaCartao.iTotalReg + objArq.iTotalReg
        
        objBaixaCartao.colArq.Add objArq

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202603

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    BaixaCartao_Le = SUCESSO

    Exit Function

Erro_BaixaCartao_Le:

    BaixaCartao_Le = gErr

    Select Case gErr
    
        Case 202590

        Case 202591
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 202592, 202593, 202603
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMEXTFINARQSLIDOS", gErr)

        Case 202594, 202595, 202602
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMEXTFINMOV", gErr)
            
        Case 202596

        Case 202597, 202598, 202601
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADMEXTFINMOVDET", gErr)

        Case 202599, 202600
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202604)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Copia_Arquivo(ByVal sDir As String, ByVal sDirDestino As String) As Long

On Error GoTo Erro_Copia_Arquivo

    FileCopy sDir, sDirDestino

    Exit Function
    
Erro_Copia_Arquivo:

    Copia_Arquivo = gErr
     
    Select Case gErr
          
        Case 53
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182654)
     
    End Select
     
    Exit Function

End Function
