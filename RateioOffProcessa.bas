Attribute VB_Name = "RateioOffProcessa"
Option Explicit

Function Rotina_RateioOff_Int(objRateioOffBatch As ClassRateioOffBatch) As Long

Dim lErro As Long
Dim vCodigo As Variant
Dim lCodigo As Long
Dim vbMesRes As VbMsgBoxResult

On Error GoTo Erro_Rotina_RateioOff_Int

    TelaAcompanhaBatch.dValorTotal = objRateioOffBatch.colRateios.Count
    
    For Each vCodigo In objRateioOffBatch.colRateios
    
        lCodigo = vCodigo

        lErro = Processa_RateioOff(lCodigo, objRateioOffBatch)
        If lErro <> SUCESSO Then Error 36804
        
        If lErro = SUCESSO Then Call Rotina_Erro(vbOKOnly, "ERRO_RATEIOOFF_PROCESSADO", Err, lCodigo)
        
        lErro = DoEvents()
        
        TelaAcompanhaBatch.dValorAtual = TelaAcompanhaBatch.dValorAtual + 1
        
        TelaAcompanhaBatch.TotReg.Caption = CStr(TelaAcompanhaBatch.dValorAtual)
        
        TelaAcompanhaBatch.ProgressBar1.Value = CInt((TelaAcompanhaBatch.dValorAtual / TelaAcompanhaBatch.dValorTotal) * 100)

        If TelaAcompanhaBatch.iCancelaBatch = CANCELA_BATCH Then
        
            vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_APURACAO_RATEIOS")
            
            If vbMesRes = vbYes Then Error 41537
                
            TelaAcompanhaBatch.iCancelaBatch = 0
                
        End If

    Next

    Rotina_RateioOff_Int = SUCESSO
    
    Exit Function
    
Erro_Rotina_RateioOff_Int:

    Rotina_RateioOff_Int = Err
    
    Select Case Err
    
        Case 36804
            Call Rotina_Erro(vbOKOnly, "ERRO_RATEIOOFF_BATCH", Err, lCodigo)
            Resume Next
        
        Case 41537
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 166154)

    End Select

    Exit Function
            
End Function

Private Function Processa_RateioOff(lCodigo As Long, objRateioOffBatch As ClassRateioOffBatch) As Long
'Processa um rateio Offline

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim objRateioOff As New ClassRateioOff
Dim colRateioOff As New Collection
Dim colContas As New Collection
Dim objPeriodo As New ClassPeriodo

On Error GoTo Erro_Processa_RateioOff

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 36812
    Next

    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 36813

     objRateioOff.lCodigo = lCodigo
    
    'le os rateios com o codigo em questão
    lErro = CF("RateioOff_Le_Doc1", alComando(1), objRateioOff, colRateioOff)
    If lErro <> SUCESSO And lErro <> 36816 Then Error 36818
    
    'se não encontrou rateio com este codigo
    If lErro = 36816 Then Error 36819
    
    'le as contas de origem relativos ao Rateio passado como parametro e coloca-os em colContas
    lErro = CF("RateioOffContas_Le_Doc1", alComando(4), objRateioOff, colContas)
    If lErro <> SUCESSO Then Error 55824
    
    lErro = CF("Periodo_Le", objRateioOffBatch.dtData, objPeriodo)
    If lErro <> SUCESSO Then Error 36821
    
    'processa o rateio
    lErro = Processa_RateioOff1(alComando(2), alComando(3), colRateioOff, objRateioOffBatch, objPeriodo, colContas)
    If lErro <> SUCESSO Then Error 36820
        
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 36821
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
        
    Processa_RateioOff = SUCESSO

    Exit Function

Erro_Processa_RateioOff:

    Processa_RateioOff = Err

    Select Case Err

        Case 36812
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", lErro)

        Case 36813
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", lErro)

        Case 36818, 36820, 55824
        
        Case 36819
            Call Rotina_Erro(vbOKOnly, "ERRO_RATEIOOFF_NAO_CADASTRADO1", lCodigo)

        Case 36821
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", lErro)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 166155)

    End Select

    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function
            
End Function

Private Function Processa_RateioOff1(lComando As Long, lComando1 As Long, colRateioOff As Collection, objRateioOffBatch As ClassRateioOffBatch, objPeriodo As ClassPeriodo, colContas As Collection) As Long
'Processa o rateio

Dim lErro As Long
Dim iIndice As Integer
Dim objRateioOff As ClassRateioOff

On Error GoTo Erro_Processa_RateioOff1

    Set objRateioOff = colRateioOff.Item(1)

    Select Case objRateioOff.iTipo
    
        Case TIPO_RATEIOOFF_MENSAL
        
            'faz o rateio do tipo ccl mensal
            lErro = Processa_RateioOff_Mensal(lComando, colRateioOff, objRateioOffBatch, objPeriodo, colContas)
            If lErro <> SUCESSO Then Error 36822
            
        Case TIPO_RATEIOOFF_ACUMULADO
        
            'faz o rateio to tipo ccl em periodos acumulados
            lErro = Processa_RateioOff_Acumulado(lComando1, colRateioOff, objRateioOffBatch, objPeriodo, colContas)
            If lErro <> SUCESSO Then Error 36823
            
        Case Else
            Error 36824
            
    End Select
        
    Processa_RateioOff1 = SUCESSO

    Exit Function

Erro_Processa_RateioOff1:

    Processa_RateioOff1 = Err

    Select Case Err

        Case 36822, 36823

        Case 36824
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_RATEIOOFF_INVALIDO", objRateioOff.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 166156)

    End Select

    Exit Function

End Function

Private Function Processa_RateioOff_Mensal(lComando As Long, colRateioOff As Collection, objRateioOffBatch As ClassRateioOffBatch, objPeriodo As ClassPeriodo, colContas As Collection) As Long
'faz o rateio do tipo ccl mensal

Dim lErro As Long
Dim dTotalCcl As Double
Dim objRateioOff As ClassRateioOff
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim colLancamento_Detalhe As New Collection
Dim sSQL As String

On Error GoTo Erro_Processa_RateioOff_Mensal

    Set objRateioOff = colRateioOff.Item(1)

    sSQL = "SELECT SUM(Cre" + Format(objPeriodo.iPeriodo, "00") + " - Deb" + Format(objPeriodo.iPeriodo, "00") + ") AS TOT FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Ccl = ?"

    Call Adiciona_Contas_Origem(sSQL, colContas)

    'totaliza o saldo do centro de custo para o periodo em questão
    lErro = Comando_Executar(lComando, sSQL, dTotalCcl, objRateioOffBatch.iFilialEmpresa, objPeriodo.iExercicio, objRateioOff.sCclOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 36826
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36827
    
    If dTotalCcl <> 0 Then
    
        lErro = Preenche_Lancamento_Cabecalho(objRateioOffBatch, objPeriodo, objLancamento_Cabecalho)
        If lErro <> SUCESSO Then Error 36829
        
        lErro = Preenche_Lancamento_Detalhe(objRateioOffBatch, objLancamento_Cabecalho, colLancamento_Detalhe, dTotalCcl, colRateioOff)
        If lErro <> SUCESSO Then Error 36830
        
        'gera os lançamentos pendentes pois estão dentro de um lote
        lErro = CF("Lancamento_Grava0", objLancamento_Cabecalho, colLancamento_Detalhe)
        If lErro <> SUCESSO Then Error 36831

    Else
    
            Call Rotina_Erro(vbOKOnly, "ERRO_RATEIOOFF_CCL_ZERADO", objRateioOff.lCodigo)
    
    End If

    Processa_RateioOff_Mensal = SUCESSO

    Exit Function

Erro_Processa_RateioOff_Mensal:

    Processa_RateioOff_Mensal = Err

    Select Case Err

        Case 36826, 36827
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL3", objRateioOff.sCclOrigem)

        Case 36829, 36830, 36831

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 166157)

    End Select

    Exit Function

End Function

Private Function Processa_RateioOff_Acumulado(lComando As Long, colRateioOff As Collection, objRateioOffBatch As ClassRateioOffBatch, objPeriodo As ClassPeriodo, colContas As Collection) As Long
'faz o rateio do tipo ccl mensal

Dim lErro As Long
Dim dTotalCcl As Double
Dim objRateioOff As ClassRateioOff
Dim objLancamento_Cabecalho As New ClassLancamento_Cabecalho
Dim colLancamento_Detalhe As New Collection
Dim sSQL As String
Dim iPeriodo As Integer

On Error GoTo Erro_Processa_RateioOff_Acumulado

    Set objRateioOff = colRateioOff.Item(1)

    sSQL = "SELECT SUM("
    
    For iPeriodo = objRateioOffBatch.iPeriodoInicial To objRateioOffBatch.iPeriodoFinal
        sSQL = sSQL + "Cre" + Format(objPeriodo.iPeriodo, "00") + " - Deb" + Format(objPeriodo.iPeriodo, "00") + " + "
    Next
    
    sSQL = Mid(sSQL, 1, Len(sSQL) - 3)
    
    sSQL = sSQL + ") AS TOT FROM MvPerCcl WHERE FilialEmpresa = ? AND Exercicio = ? AND Ccl = ?"
    
    Call Adiciona_Contas_Origem(sSQL, colContas)
    
    'totaliza o saldo do centro de custo para os periodos em questão
    lErro = Comando_Executar(lComando, sSQL, dTotalCcl, objRateioOffBatch.iFilialEmpresa, objPeriodo.iExercicio, objRateioOff.sCclOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 36838
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36839
    
    lErro = Preenche_Lancamento_Cabecalho(objRateioOffBatch, objPeriodo, objLancamento_Cabecalho)
    If lErro <> SUCESSO Then Error 36840
    
    lErro = Preenche_Lancamento_Detalhe(objRateioOffBatch, objLancamento_Cabecalho, colLancamento_Detalhe, dTotalCcl, colRateioOff)
    If lErro <> SUCESSO Then Error 36841
    
    'gera os lançamentos pendentes pois estão dentro de um lote
    lErro = CF("Lancamento_Grava0", objLancamento_Cabecalho, colLancamento_Detalhe)
    If lErro <> SUCESSO Then Error 36842

    Processa_RateioOff_Acumulado = SUCESSO

    Exit Function

Erro_Processa_RateioOff_Acumulado:

    Processa_RateioOff_Acumulado = Err

    Select Case Err

        Case 36838, 36839
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCCL3", objRateioOff.sCclOrigem)

        Case 36840, 36841, 36842

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 166158)

    End Select

    Exit Function

End Function

Private Function Preenche_Lancamento_Cabecalho(objRateioOffBatch As ClassRateioOffBatch, objPeriodo As ClassPeriodo, objLancamento_Cabecalho As ClassLancamento_Cabecalho) As Long

Dim lErro As Long
Dim lDoc As Long

On Error GoTo Erro_Preenche_Lancamento_Cabecalho

    objLancamento_Cabecalho.iFilialEmpresa = objRateioOffBatch.iFilialEmpresa
    objLancamento_Cabecalho.sOrigem = MODULO_CONTABILIDADE
    objLancamento_Cabecalho.iLote = objRateioOffBatch.iLote
    objLancamento_Cabecalho.dtData = objRateioOffBatch.dtData
    objLancamento_Cabecalho.iExercicio = objPeriodo.iExercicio
    objLancamento_Cabecalho.iPeriodoLan = objPeriodo.iPeriodo
    objLancamento_Cabecalho.iPeriodoLote = objPeriodo.iPeriodo
    
    lErro = CF("Voucher_Automatico1", objLancamento_Cabecalho.iFilialEmpresa, objPeriodo.iExercicio, objPeriodo.iPeriodo, objLancamento_Cabecalho.sOrigem, lDoc)
    If lErro <> SUCESSO Then Error 36828
            
    objLancamento_Cabecalho.lDoc = lDoc

    Preenche_Lancamento_Cabecalho = SUCESSO

    Exit Function

Erro_Preenche_Lancamento_Cabecalho:

    Preenche_Lancamento_Cabecalho = Err

    Select Case Err

        Case 36828

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 166159)

    End Select

    Exit Function

End Function

Private Function Preenche_Lancamento_Detalhe(objRateioOffBatch As ClassRateioOffBatch, objLancamento_Cabecalho As ClassLancamento_Cabecalho, colLancamento_Detalhe As Collection, ByVal dValor As Double, colRateioOff As Collection)

Dim lErro As Long
Dim iIndice1 As Integer
Dim objLancamento_Detalhe As New ClassLancamento_Detalhe
Dim objRateioOff As ClassRateioOff
Dim dTotalLanc As Double

On Error GoTo Erro_Preenche_Lancamento_Detalhe

    Set objRateioOff = colRateioOff.Item(1)

    dValor = Round(dValor, 2)

    'cria o lançamento que registra a contra-partida pelo valor total
    objLancamento_Detalhe.iSeq = 1
    objLancamento_Detalhe.sConta = objRateioOff.sContaCre
    objLancamento_Detalhe.dValor = -dValor
    objLancamento_Detalhe.sHistorico = objRateioOffBatch.sHistorico
    objLancamento_Detalhe.sCcl = objRateioOff.sCclOrigem
    objLancamento_Detalhe.sOrigem = ""
    objLancamento_Detalhe.sProduto = ""
    
    'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
    colLancamento_Detalhe.Add objLancamento_Detalhe
    
    iIndice1 = 1
    
    For Each objRateioOff In colRateioOff
    
        iIndice1 = iIndice1 + 1
    
        Set objLancamento_Detalhe = New ClassLancamento_Detalhe
        
        objLancamento_Detalhe.iSeq = iIndice1
            
        objLancamento_Detalhe.sConta = objRateioOff.sConta
        objLancamento_Detalhe.sCcl = objRateioOff.sCcl
        objLancamento_Detalhe.dValor = Round(dValor * objRateioOff.dPercentual, 2)
        dTotalLanc = dTotalLanc + objLancamento_Detalhe.dValor
        objLancamento_Detalhe.sHistorico = objRateioOffBatch.sHistorico
        objLancamento_Detalhe.sOrigem = ""
        objLancamento_Detalhe.sProduto = ""
            
        'Armazena o objeto objLancamento_Detalhe na coleção colLancamento_Detalhe
        colLancamento_Detalhe.Add objLancamento_Detalhe
        
    Next

    'se o valor total rateado for diferente da soma das parcelas ==> ajustar o ultimo lançamento da coleção
    If dTotalLanc <> dValor Then objLancamento_Detalhe.dValor = objLancamento_Detalhe.dValor + (dValor - dTotalLanc)

    Preenche_Lancamento_Detalhe = SUCESSO

    Exit Function

Erro_Preenche_Lancamento_Detalhe:

    Preenche_Lancamento_Detalhe = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", lErro, Error$, 166160)

    End Select

    Exit Function

End Function

Private Sub Adiciona_Contas_Origem(sSQL As String, colContas As Collection)

Dim objRateioOffContas As ClassRateioOffContas

    sSQL = sSQL & " AND ("

    If colContas.Count > 0 Then

       For Each objRateioOffContas In colContas
       
           sSQL = sSQL & " (Conta >= '" & objRateioOffContas.sContaInicio & "' AND Conta <= '" & objRateioOffContas.sContaInicio & "') OR "
    
       Next
       
       If colContas.Count > 0 Then
           sSQL = Left(sSQL, Len(sSQL) - 4)
           sSQL = sSQL & ")"
       End If
    
    Else
    
        sSQL = sSQL & " Conta = '' )"
        
    End If
    
End Sub
