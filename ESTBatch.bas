Attribute VB_Name = "ESTBatch"
'Option Explicit
''**** Quando iMes for janeiro e houver Dezembro do ano anterior,
''**** transportar só valor (saldo do ano) para ValorInicial
'
''???? Tranferir para ErrosMAT
'Const ERRO_CMP_E_CST_ZERADOS = 0 'Parametro: sProduto, iMes
''Atenção, o produto %s está com o custo médio de produção e o custo standard do mes %i zerados. Favor colocar o custo standard.
'
'Dim giReprocessamento As Integer 'indica se esta executando a rotina de reprocessamento dos movimentos de estoque
'
'Function Rotina_CustoMedioProducao_Int(ByVal iFilialEmpresa As Integer, iAno As Integer, iMes As Integer) As Long
''calcula o custo médio de produção para mes/ano passados e valora os movimentos de estoque
'
'Dim lTransacao As Long
'Dim alComando(1 To 28) As Long
'Dim sComandoSQL(1 To 9) As String
'Dim iIndice As Integer
'Dim lErro As Long
'Dim lTotalProdutos As Long 'nº de produtos que participam processo
'Dim tMovEstoque As typeItemMovEstoque
'Dim tMovEstoque2 As typeItemMovEstoque
'Dim tSldMesEst As typeSldMesEst
'Dim tSldDiaEst As typeSldDiaEst
'Dim tSldDiaEstAlm As typeSldDiaEstAlm
'Dim dCPAtual As Double 'Custo Producao do mês atual (iMes)
'Dim tProduto As typeProduto '???Vai sair quando tiver quantidades em UMEstoque nos movimentos
'Dim dCMPAtual As Double 'Custo Medio Producao do mês atual (iMes)
'Dim colAlmoxInfo As Collection
'Dim tTipoMovEst As typeTipoMovEst
'Dim lTotalProdutos2 As Long
'Dim tSldMesEst2 As typeSldMesEst2
'Dim tSldMesEst1 As typeSldMesEst1
'
'On Error GoTo Erro_Rotina_CustoMedioProducao_Int
'
'    giReprocessamento = 0 'indica que nao se trata de um reprocessamento
'
'    'Abre comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 25234
'    Next
'
'    'Inicia a transação
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 25235
'
'    'Critica se EstoqueMes ñ tem CMP apurado e está fechado. Faz lock exclusive.
'    lErro = CF("Rotina_CMP_EstoqueMes_CriticaLock",alComando(8), iFilialEmpresa, iAno, iMes)
'    If lErro <> SUCESSO And lErro <> 25287 And lErro <> 25289 And lErro <> 25290 And lErro <> 25290 Then gError 25236
'
'    'se o estoquemes nao estiver cadastrado ==> erro
'    If lErro = 25287 Then gError 83756
'
'    'se o estoquemes nao estiver aberto ==> erro
'    If lErro = 25289 Then gError 83757
'
'    'se o mes nao tiver o custo de producao apurado ==> erro
'    If lErro = 25290 Then gError 83758
'
'    'Retorna o número de Produtos produzidos que tiveram Movtos nesta FilialEmpresa
'    lErro = Rotina_CMP_TotalMovEstoque(iFilialEmpresa, iAno, iMes, lTotalProdutos)
'    If lErro <> SUCESSO Then gError 25237
'
'    If iMes = 12 Then
'
'        'Retorna o número total de Produtos a serem transferidos para o proximo Ano
'        lErro = Rotina_CMP_TotalTransfereValorInicial(iFilialEmpresa, iAno, iMes, lTotalProdutos2)
'        If lErro <> SUCESSO Then gError 69021
'
'    End If
'
'    'Tela acompanhamento Batch inicializa dValorTotal
'    TelaAcompanhaBatchEST.dValorTotal = lTotalProdutos + (lTotalProdutos2 * 2)
'
'    'Se houve movimentos de estoque de Produtos produzidos,
'    If lTotalProdutos > 0 Then
'
'        'Monta 6 comandos SQL -> leitura MovEstoque, atualizacao MovEstoque, atualizacao SaldoMesEst, atualizacao SaldoDiaEst
'        lErro = Rotina_CMP_MontaComandosSQL(iMes, sComandoSQL())
'        If lErro <> SUCESSO Then gError 25238
'
'        'Inicializa os 6 comandos
'        lErro = Rotina_CMP_InicializaComandos(sComandoSQL(), alComando(), tProduto, dCPAtual, tTipoMovEst, tMovEstoque, tMovEstoque2, tSldMesEst, tSldDiaEst, tSldDiaEstAlm, iAno, iMes, iFilialEmpresa, tSldMesEst2)
'        If lErro <> SUCESSO Then gError 25239
'
'        tSldMesEst.iAno = iAno
'        tSldMesEst.iFilialEmpresa = iFilialEmpresa
'        tSldMesEst2.iAno = iAno
'        tSldMesEst2.iFilialEmpresa = iFilialEmpresa
'        tSldMesEst1.iAno = iAno
'        tSldMesEst1.iFilialEmpresa = iFilialEmpresa
'
'        Do While True 'trecho que é repetido POR PRODUTO
'
'            tSldMesEst1.sProduto = tMovEstoque.sProduto
'
'            'Atualiza custos de PRODUCAO dos movimentos (ENTRADAS) e calcula CustoMedioProducaoAtual (CMPAtual) do Produto
'            lErro = Rotina_CMP_AtualizaCustosProducao(sComandoSQL(), alComando(), tProduto, dCPAtual, tTipoMovEst, tMovEstoque, tMovEstoque2, tSldMesEst, tSldDiaEst, iAno, iMes, iFilialEmpresa, dCMPAtual, colAlmoxInfo, tSldMesEst1)
'            If lErro <> SUCESSO And lErro <> 25258 Then gError 25240
'
'            'Se não tem mais Movimentos para atualizar custos
'            If lErro = 25258 Then
'                'Atualiza tela de acompanhamento do Batch
'                lErro = Rotina_CMP_AtualizaTelaBatch()
'                If lErro <> SUCESSO Then gError 25780
'                Exit Do
'            End If
'
'            'Atualiza custos dos movimentos , exceto PRODUCAO, baseado em CMPAtual e os custos dos Escaninhos
'            lErro = Rotina_CMP_AtualizaCustos(sComandoSQL(), alComando(), tProduto, tTipoMovEst, tMovEstoque, tMovEstoque2, tSldMesEst, tSldDiaEst, iAno, iMes, iFilialEmpresa, dCMPAtual, colAlmoxInfo, tSldMesEst2, tSldMesEst1)
'            If lErro <> SUCESSO And lErro <> 69708 And lErro <> 69720 Then gError 25292
'
'            'Se não tem mais Movimentos para atualizar custos
'            If lErro <> SUCESSO Then
'                'Atualiza tela de acompanhamento do Batch
'                lErro = Rotina_CMP_AtualizaTelaBatch()
'                If lErro <> SUCESSO Then gError 25782
'                Exit Do
'            End If
'
'            'Atualiza tela de acompanhamento do Batch
'            lErro = Rotina_CMP_AtualizaTelaBatch()
'            If lErro <> SUCESSO Then gError 25781
'
'        Loop
'
'    End If
'
'    'Altera campo CustoProdApurado na tabela EstoqueMes. Comando de lock já foi feito.
'    lErro = Rotina_CMP_EstoqueMes_Atualiza(alComando(8), iFilialEmpresa, iAno, iMes)
'    If lErro <> SUCESSO Then gError 25324
'
''    'Coloca o Valor dos custos dos movimentos com estorno
''    lErro = Rotina_Atualiza_Custo_Movimento_Estorno(iMes, iAno, iFilialEmpresa)
''    If lErro <> SUCESSO Then Error 78034
'
'    If iMes = 12 Then
'
'        'Atualiza os valores iniciais para o mês de Desembro
'        lErro = Transfere_ValoresIniciais(iAno, iFilialEmpresa)
'        If lErro <> SUCESSO Then gError 69763
'
'    End If
'
'    'Confirma a transação
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then gError 25241
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Rotina_CustoMedioProducao_Int = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CustoMedioProducao_Int:
'
'    Rotina_CustoMedioProducao_Int = gErr
'
'    Select Case gErr
'
'       Case 25234
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 25235
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 25236, 25237, 25238, 25239, 25240, 25292, 25324, 25780, 25781, 25782, 69021, 69763, 78034 'Tratados na rotina chamada
'
'        Case 25241
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case 83756
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE", gErr, iFilialEmpresa, iAno, iMes)
'
'        Case 83757
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_ABERTO", gErr, iFilialEmpresa, iAno, iMes)
'
'        Case 83758
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_CMP_APURADO", gErr, iFilialEmpresa, iAno, iMes)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159492)
'
'    End Select
'
'    'Rollback
'    Call Transacao_Rollback
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Function Rotina_CustoMedioProducao_Reproc(ByVal iFilialEmpresa As Integer, iAno As Integer, iMes As Integer) As Long
''calcula o custo médio de produção para mes/ano passados e valora os movimentos de estoque
'
'Dim alComando(1 To 27) As Long
'Dim sComandoSQL(1 To 9) As String
'Dim iIndice As Integer
'Dim lErro As Long
'Dim lTotalProdutos As Long 'nº de produtos que participam processo
'Dim tMovEstoque As typeItemMovEstoque
'Dim tMovEstoque2 As typeItemMovEstoque
'Dim tSldMesEst As typeSldMesEst
'Dim tSldDiaEst As typeSldDiaEst
'Dim tSldDiaEstAlm As typeSldDiaEstAlm
'Dim dCPAtual As Double 'Custo Producao do mês atual (iMes)
'Dim tProduto As typeProduto
'Dim dCMPAtual As Double 'Custo Medio Producao do mês atual (iMes)
'Dim colAlmoxInfo As Collection
'Dim tTipoMovEst As typeTipoMovEst
'Dim lTotalProdutos2 As Long
'Dim tSldMesEst2 As typeSldMesEst2
'Dim tSldMesEst1 As typeSldMesEst1
'
'On Error GoTo Erro_Rotina_CustoMedioProducao_Reproc
'
'    giReprocessamento = REPROCESSAMENTO_REFAZ
'
'    'Abre comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 25234
'    Next
'
'    'Critica se EstoqueMes ñ tem CMP apurado e está fechado. Faz lock exclusive.
'    lErro = CF("Rotina_CMP_EstoqueMes_CriticaLock",alComando(8), iFilialEmpresa, iAno, iMes)
'    If lErro <> SUCESSO And lErro <> 25287 And lErro <> 25289 And lErro <> 25290 And lErro <> 25290 Then gError 25236
'
'    'so apura o custo de producao se o custo tiver anteriormente sido apurado
'    If lErro = 25290 Then
'
'        'Monta 6 comandos SQL -> leitura MovEstoque, atualizacao MovEstoque, atualizacao SaldoMesEst, atualizacao SaldoDiaEst
'        lErro = Rotina_CMP_MontaComandosSQL(iMes, sComandoSQL())
'        If lErro <> SUCESSO Then gError 25238
'
'        'Inicializa os 6 comandos
'        lErro = Rotina_CMP_InicializaComandos(sComandoSQL(), alComando(), tProduto, dCPAtual, tTipoMovEst, tMovEstoque, tMovEstoque2, tSldMesEst, tSldDiaEst, tSldDiaEstAlm, iAno, iMes, iFilialEmpresa, tSldMesEst2)
'        If lErro <> SUCESSO Then gError 25239
'
'        tSldMesEst.iAno = iAno
'        tSldMesEst.iFilialEmpresa = iFilialEmpresa
'        tSldMesEst2.iAno = iAno
'        tSldMesEst2.iFilialEmpresa = iFilialEmpresa
'        tSldMesEst1.iAno = iAno
'        tSldMesEst1.iFilialEmpresa = iFilialEmpresa
'
'        Do While True 'trecho que é repetido POR PRODUTO
'
'            tSldMesEst1.sProduto = tMovEstoque.sProduto
'
'            'Atualiza custos de PRODUCAO dos movimentos (ENTRADAS) e calcula CustoMedioProducaoAtual (CMPAtual) do Produto
'            lErro = Rotina_CMP_AtualizaCustosProducao(sComandoSQL(), alComando(), tProduto, dCPAtual, tTipoMovEst, tMovEstoque, tMovEstoque2, tSldMesEst, tSldDiaEst, iAno, iMes, iFilialEmpresa, dCMPAtual, colAlmoxInfo, tSldMesEst1)
'            If lErro <> SUCESSO And lErro <> 25258 Then gError 25240
'
'            'Se não tem mais Movimentos para atualizar custos
'            If lErro = 25258 Then
'                Exit Do
'            End If
'
'            'Atualiza custos dos movimentos , exceto PRODUCAO, baseado em CMPAtual e os custos dos Escaninhos
'            lErro = Rotina_CMP_AtualizaCustos(sComandoSQL(), alComando(), tProduto, tTipoMovEst, tMovEstoque, tMovEstoque2, tSldMesEst, tSldDiaEst, iAno, iMes, iFilialEmpresa, dCMPAtual, colAlmoxInfo, tSldMesEst2, tSldMesEst1)
'            If lErro <> SUCESSO And lErro <> 69708 And lErro <> 69720 Then gError 25292
'
'            'Se não tem mais Movimentos para atualizar custos
'            If lErro <> SUCESSO Then
'                Exit Do
'            End If
'
'        Loop
'
''        'Coloca o Valor dos custos dos movimentos com estorno
''        lErro = Rotina_Atualiza_Custo_Movimento_Estorno(iMes, iAno, iFilialEmpresa)
''        If lErro <> SUCESSO Then Error 78034
'
'        If iMes = 12 Then
'
'            'Atualiza os valores iniciais para o mês de Desembro
'            lErro = Transfere_ValoresIniciais(iAno, iFilialEmpresa)
'            If lErro <> SUCESSO Then gError 69763
'
'        End If
'
'    End If
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Rotina_CustoMedioProducao_Reproc = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CustoMedioProducao_Reproc:
'
'    Rotina_CustoMedioProducao_Reproc = gErr
'
'    Select Case gErr
'
'       Case 25234
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 25235
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 25236, 25237, 25238, 25239, 25240, 25292, 25324, 25780, 25781, 25782, 69021, 69763, 78034 'Tratados na rotina chamada
'
'        Case 25241
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159493)
'
'    End Select
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'
'Function Transfere_ValoresIniciais(iAno As Integer, iFilialEmpresa As Integer) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Transfere_ValoresIniciais
'
'    'Transfere para o proximo ano o Valor Inicial e o CustoMedioProducaoInicial de SaldoMesEst
'    lErro = Rotina_CMP_Transfere_SldMesEst_ValorInicial(iAno, iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 64486
'
'    'Transfere para o proximo Ano os Valor Inicial e o CustoMedioProducaoInicial
'    lErro = Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial(iAno, iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 69034
'
'    'Transfere para o proximo ano o Valor Inicial e o CustoMedioProducaoInicial de SaldoMesEst
'    lErro = Rotina_CMP_Transfere_SldMesEst1_ValorInicial(iAno, iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 89842
'
'    'Transfere para o proximo ano o Valor Inicial e o CustoMedioProducaoInicial de SaldoMesEst
'    lErro = Rotina_CMP_Transfere_SldMesEst2_ValorInicial(iAno, iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 69757
'
'    'Transfere para o proximo Ano os Valor Inicial e o CustoMedioProducaoInicial
'    lErro = Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial(iAno, iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 89844
'
'    'Transfere para o proximo Ano os Valor Inicial e o CustoMedioProducaoInicial
'    lErro = Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial(iAno, iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 69758
'
'    Transfere_ValoresIniciais = SUCESSO
'
'    Exit Function
'
'Erro_Transfere_ValoresIniciais:
'
'    Transfere_ValoresIniciais = Err
'
'    Select Case Err
'
'        Case 64486, 69034, 69757, 69758, 89842, 89844
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159494)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaTelaBatch()
''Atualiza tela de acompanhamento do Batch
'
'On Error GoTo Erro_Rotina_CMP_AtualizaTelaBatch
'
'    If giReprocessamento = 0 Then
'
'        'Atualiza tela de acompanhamento do Batch
'        TelaAcompanhaBatchEST.dValorAtual = TelaAcompanhaBatchEST.dValorAtual + 1
'        TelaAcompanhaBatchEST.TotReg.Caption = CStr(TelaAcompanhaBatchEST.dValorAtual)
'        TelaAcompanhaBatchEST.ProgressBar1.Value = CInt((TelaAcompanhaBatchEST.dValorAtual / TelaAcompanhaBatchEST.dValorTotal) * 100)
'
'    End If
'
'    Rotina_CMP_AtualizaTelaBatch = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaTelaBatch:
'
'    Rotina_CMP_AtualizaTelaBatch = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159495)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_EstoqueMes_Atualiza(lComando As Long, iFilialEmpresa As Integer, iAno As Integer, iMes As Integer) As Long
''Atualiza campo CustoProdApurado para CUSTO_APURADO em EstoqueMes
''Chamada EM TRANSAÇÃO.
'
'Dim lErro As Long
'Dim lComando2 As Long  'Comando do UPDATE - fica local
'
'On Error GoTo Erro_Rotina_CMP_EstoqueMes_Atualiza
'
'    'Abre comando
'    lComando2 = Comando_Abrir()
'    If lComando2 = 0 Then gError 25322
'
'    'Atualiza tabela SaldoMesEst, campos ValEnt e ValSai de iMes
'    lErro = Comando_ExecutarPos(lComando2, "UPDATE EstoqueMes SET CustoProdApurado = ?", lComando, CUSTO_APURADO)
'    If lErro <> AD_SQL_SUCESSO Then gError 25323
'
'    'Fecha o comando de UPDATE
'    Call Comando_Fechar(lComando2)
'
'    Rotina_CMP_EstoqueMes_Atualiza = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_EstoqueMes_Atualiza:
'
'    Rotina_CMP_EstoqueMes_Atualiza = gErr
'
'    Select Case gErr
'
'        Case 25322
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 25323
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ESTOQUEMES", gErr, iFilialEmpresa, iAno, iMes)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159496)
'
'    End Select
'
'    'Fecha o comando de UPDATE
'    Call Comando_Fechar(lComando2)
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_MontaComandosSQL(iMes As Integer, sComandoSQL() As String) As Long
''retorna os comandos SQL que participam do processamento de custo médio de produção
'
'Dim lErro As Long
'Dim sFiltroMovEstoque As String
'Dim sFiltroMovEstoque2 As String
'Dim sOrdemMovEstoque As String  'por Produto, Apropr, EntradaouSaida, Data, Almoxarifado
'Dim iIndice As Integer
'Dim sOrdemMovEstoque2 As String  'por Produto, Apropr, Data, Almoxarifado
'
'On Error GoTo Erro_Rotina_CMP_MontaComandosSQL
'
'    sFiltroMovEstoque = "MovimentoEstoque.FilialEmpresa = ? AND Data >= ? AND Data <= ? AND (MovimentoEstoque.Apropriacao = ? OR MovimentoEstoque.Apropriacao = ?) "
'    sFiltroMovEstoque2 = "MovimentoEstoqueES.FilialEmpresa = ? AND Data >= ? AND Data <= ? AND (MovimentoEstoqueES.Apropriacao = ? OR MovimentoEstoqueES.Apropriacao = ?) "
'    sOrdemMovEstoque = "ORDER BY MovimentoEstoque.Produto, MovimentoEstoque.Apropriacao"
'    sOrdemMovEstoque2 = "ORDER BY MovimentoEstoqueES.Produto, MovimentoEstoqueES.Apropriacao"
'
'    If APROPR_CUSTO_REAL > APROPR_CUSTO_MEDIO_PRODUCAO Then
'        sOrdemMovEstoque = sOrdemMovEstoque & " DESC"
'        sOrdemMovEstoque2 = sOrdemMovEstoque2 & " DESC"
'    End If
'
'    sOrdemMovEstoque = sOrdemMovEstoque + ", TiposMovimentoEstoque.EntradaSaidaCMP DESC"
'    sOrdemMovEstoque = sOrdemMovEstoque & ", MovimentoEstoque.Data, MovimentoEstoque.Almoxarifado, MovimentoEstoque.NumIntDoc"
'    sOrdemMovEstoque2 = sOrdemMovEstoque2 + ", MovimentoEstoqueES.EntradaSaidaCMP DESC"
'    sOrdemMovEstoque2 = sOrdemMovEstoque2 & ", MovimentoEstoqueES.Data, MovimentoEstoqueES.Almoxarifado, MovimentoEstoqueES.NumIntDoc"
'
'    'Comando para ler dados de MovimetosEstoque, TipoMovimento
'    '(para saber se é Entrada ou Saída)
'    '-------------------------------------------------------
'
'    'Campos selecionados de TipoMovEstoque e de MovEstoque
'    sComandoSQL(1) = "SELECT ClasseUM, SiglaUMEstoque, EntradaSaidaCMP, AtualizaConsumo, AtualizaVenda, MovimentoEstoque.NumIntDoc, MovimentoEstoque.Produto, Quantidade, SiglaUM, MovimentoEstoque.Apropriacao, Data, Almoxarifado, TiposMovimentoEstoque.AtualizaConsig, TiposMovimentoEstoque.AtualizaDemo, TiposMovimentoEstoque.AtualizaConserto, TiposMovimentoEstoque.ProdutodeTerc,TiposMovimentoEstoque.AtualizaOutras, TiposMovimentoEstoque.AtualizaBenef, TiposMovimentoEstoque.CustoMedio, TiposMovimentoEstoque.CodigoOrig, TiposMovimentoEstoque.Codigo "
'    'Tabelas
'    sComandoSQL(1) = sComandoSQL(1) & "FROM Produtos, TiposMovimentoEstoque, MovimentoEstoque "
'    'Links
'    sComandoSQL(1) = sComandoSQL(1) & "WHERE TiposMovimentoEstoque.Codigo = MovimentoEstoque.TipoMov AND MovimentoEstoque.Produto=Produtos.Codigo "
'    'Filtro MovEstoque
'    sComandoSQL(1) = sComandoSQL(1) & "AND " & sFiltroMovEstoque
'    'Ordem
'    sComandoSQL(1) = sComandoSQL(1) & sOrdemMovEstoque
'
'    'Comando para atualizar Custos de Movimentos de Estoque
'    '-------------------------------------------------------
'
'    sComandoSQL(2) = "SELECT NumIntDoc FROM MovimentoEstoqueES WHERE " & sFiltroMovEstoque2 & sOrdemMovEstoque2
'
'    'Comando para atualizar ValorEnt(i) e ValorSai(i) e ValorCons(i) na tabela SaldoMesEst,
'    'onde i é o Mes cujo CMP está sendo calculado
'    '--------------------------------------------------------------------------------------
'
'    'CustoMedioProducaoInicial e CP Mes Atual
'    sComandoSQL(3) = "SELECT CustoMedioProducaoInicial, CustoProducao" & CStr(iMes) & ", CustoMedio" & CStr(iMes) & ", CustoStandard" & CStr(iMes) & ", "
'    'Quantidade e valor inicial
'    sComandoSQL(3) = sComandoSQL(3) & "QuantInicial, ValorInicial, "
'    sComandoSQL(3) = sComandoSQL(3) & "QuantInicialCusto, ValorInicialCusto, "
'    'Quantidades e valores de entrada e de saida mensais
'    For iIndice = 1 To 12
'        sComandoSQL(3) = sComandoSQL(3) & "QuantEnt" & CStr(iIndice) & ", " & "QuantSai" & CStr(iIndice) & ", " & "ValorEnt" & CStr(iIndice) & ", " & "ValorSai" & CStr(iIndice) & ", " & "SaldoQuantCusto" & CStr(iIndice) & ", " & "SaldoValorCusto" & CStr(iIndice) & ", "
'    Next
'    'Produto
'    sComandoSQL(3) = sComandoSQL(3) & "Produto "
'    'Tabela, Filtro, Ordem
'    sComandoSQL(3) = sComandoSQL(3) & "FROM SldMesEst WHERE Ano = ? AND FilialEmpresa = ? ORDER BY Produto"
'
'    'Comando para atualizar ValorEntrada, ValorSaida, ValorCons e ValorVend na tabela SaldoDiaEst
'    '----------------------------------------------------------------------------------
'
''    'Produto e Data
''    sComandoSQL(4) = "SELECT Produto, Data "
''    'Tabela, Filtro, Ordem
''    sComandoSQL(4) = sComandoSQL(4) & "FROM SldDiaEst WHERE Data >= ? AND Data <= ? AND FilialEmpresa = ? ORDER BY Produto, Data"
'
'    'Comando para atualizar ValorEntrada, ValorSaida, ValorCons e VaorVend na tabela SaldoDiaEstAlm
'    '--------------------------------------------------------------------------------------
'
''    'Produto e Data e Almoxarifado
''    sComandoSQL(5) = "SELECT Produto, Data, Almoxarifado "
''    'Tabela, Filtro, Ordem
''    sComandoSQL(5) = sComandoSQL(5) & "FROM SldDiaEstAlm WHERE Data >= ? AND Data <= ? ORDER BY Produto, Data, Almoxarifado"
'
'    '----------------------------------------------------------------
'    'Select para o cálculo do Custo e Atualização do Saldo dos escaninhos
'    sComandoSQL(6) = sComandoSQL(6) & "SELECT QuantInicialConsig, ValorInicialConsig, QuantInicialDemo, ValorInicialDemo, QuantInicialConserto, ValorInicialConserto, QuantInicialOutros, ValorInicialOutros, QuantInicialBenef, ValorInicialBenef, "
'    'Quantidades e valores
'    For iIndice = 1 To 12
'        sComandoSQL(6) = sComandoSQL(6) & "SaldoQuantConsig" & CStr(iIndice) & ", " & "SaldoValorConsig" & CStr(iIndice) & "," & "SaldoQuantDemo" & CStr(iIndice) & ", " & "SaldoValorDemo" & CStr(iIndice) & "," & "SaldoQuantConserto" & CStr(iIndice) & ", " & "SaldoValorConserto" & CStr(iIndice) & "," & "SaldoQuantOutros" & CStr(iIndice) & ", " & "SaldoValorOutros" & CStr(iIndice) & "," & "SaldoQuantBenef" & CStr(iIndice) & ", " & "SaldoValorBenef" & CStr(iIndice) & ","
'    Next
'    'Produto
'    sComandoSQL(6) = sComandoSQL(6) & " Produto "
'    'Tabela, Filtro, Ordem
'    sComandoSQL(6) = sComandoSQL(6) & "FROM SldMesEst2 WHERE Ano = ? AND FilialEmpresa = ? ORDER BY Produto"
'
'
'    'Select para o cálculo do Custo e Atualização do Saldo dos escaninhos de Terceiros
'    sComandoSQL(9) = sComandoSQL(9) & "SELECT QuantInicialConsig3, ValorInicialConsig3, QuantInicialDemo3, ValorInicialDemo3, QuantInicialConserto3, ValorInicialConserto3, QuantInicialOutros3, ValorInicialOutros3, QuantInicialBenef3, ValorInicialBenef3, "
'    'Quantidades e valores
'    For iIndice = 1 To 12
'        sComandoSQL(9) = sComandoSQL(9) & "SaldoQuantConsig3" & CStr(iIndice) & ", " & "SaldoValorConsig3" & CStr(iIndice) & "," & "SaldoQuantDemo3" & CStr(iIndice) & ", " & "SaldoValorDemo3" & CStr(iIndice) & "," & "SaldoQuantConserto3" & CStr(iIndice) & ", " & "SaldoValorConserto3" & CStr(iIndice) & "," & "SaldoQuantOutros3" & CStr(iIndice) & ", " & "SaldoValorOutros3" & CStr(iIndice) & "," & "SaldoQuantBenef3" & CStr(iIndice) & ", " & "SaldoValorBenef3" & CStr(iIndice) & ","
'    Next
'    'Produto
'    sComandoSQL(9) = sComandoSQL(9) & " Produto "
'    'Tabela, Filtro, Ordem
'    sComandoSQL(9) = sComandoSQL(9) & "FROM SldMesEst1 WHERE Ano = ? AND FilialEmpresa = ? ORDER BY Produto"
'
'    Rotina_CMP_MontaComandosSQL = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_MontaComandosSQL:
'
'    Rotina_CMP_MontaComandosSQL = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159497)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_InicializaComandos(sComandoSQL() As String, alComando() As Long, tProduto As typeProduto, dCPAtual As Double, tTipoMovEst As typeTipoMovEst, tMovEstoque As typeItemMovEstoque, tMovEstoque2 As typeItemMovEstoque, tSldMesEst As typeSldMesEst, tSldDiaEst As typeSldDiaEst, tSldDiaEstAlm As typeSldDiaEstAlm, iAno As Integer, iMes As Integer, iFilialEmpresa As Integer, tSldMesEst2 As typeSldMesEst2) As Long
''Chamada EM TRANSAÇÃO
''Inicializa os 5 comandos da Rotina_CustoMedioProducao
'
'Dim iDiasMes As Integer
'Dim dtDataInicial As Date
'Dim dtDataFinal As Date
'Dim lErro As Long
'Dim dQuantInicialProxAno As Double
'Dim objTipoMovEstoque As New ClassTipoMovEst
'
'On Error GoTo Erro_Rotina_CMP_InicializaComandos
'
'    'Preparação das Strings para Execução de comandos SQL
'    With tMovEstoque
'        .sProduto = String(STRING_PRODUTO, 0)
'        .sSiglaUM = String(STRING_UM_SIGLA, 0)
'    End With
'
'    tTipoMovEst.sEntradaOuSaida = String(STRING_ENTRADAOUSAIDA, 0)
'    tProduto.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
'
'    tSldMesEst.sProduto = String(STRING_PRODUTO, 0)
'    tSldDiaEst.sProduto = String(STRING_PRODUTO, 0)
'    tSldDiaEstAlm.sProduto = String(STRING_PRODUTO, 0)
'
'    'Determinação de faixa de datas
'    dtDataInicial = CDate("1/" & CStr(iMes) & "/" & CStr(iAno))
'    iDiasMes = Dias_Mes(iMes, iAno)
'    dtDataFinal = CDate(CStr(iDiasMes) & "/" & CStr(iMes) & "/" & CStr(iAno))
'
'    'Busca primeiro registro de MovimentoEstoque vinculado a TipoMovimento e Produto
'    '-------------------------------------------------------------------------------
'    With tMovEstoque
'        lErro = Comando_Executar(alComando(1), sComandoSQL(1), tProduto.iClasseUM, tProduto.sSiglaUMEstoque, tTipoMovEst.sEntradaOuSaida, tTipoMovEst.iAtualizaConsumo, tTipoMovEst.iAtualizaVenda, .lNumIntDoc, .sProduto, .dQuantidade, .sSiglaUM, .iApropriacao, .dtData, .iAlmoxarifado, tTipoMovEst.iAtualizaConsig, tTipoMovEst.iAtualizaDemo, tTipoMovEst.iAtualizaConserto, tTipoMovEst.iProdutoDeTerc, tTipoMovEst.iAtualizaOutras, tTipoMovEst.iAtualizaBenef, tTipoMovEst.iCustoMedio, tTipoMovEst.iCodigoOrig, tTipoMovEst.iCodigo, iFilialEmpresa, dtDataInicial, dtDataFinal, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO)
'    End With
'    If lErro <> AD_SQL_SUCESSO Then gError 25242
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25243
'    If lErro = AD_SQL_SEM_DADOS Then gError 25244
'
'    If tTipoMovEst.iCodigoOrig <> 0 Then
'
'        objTipoMovEstoque.iCodigo = tTipoMovEst.iCodigo
'
'        'ler os dados referentes ao tipo de movimento
'        lErro = CF("TiposMovEst_Le1",alComando(21), objTipoMovEstoque)
'        If lErro <> SUCESSO Then gError 89157
'
'        tTipoMovEst.sEntradaOuSaida = objTipoMovEstoque.sEntradaOuSaida
'        tTipoMovEst.iAtualizaConsumo = objTipoMovEstoque.iAtualizaConsumo
'        tTipoMovEst.iAtualizaVenda = objTipoMovEstoque.iAtualizaVenda
'        tTipoMovEst.iAtualizaConsig = objTipoMovEstoque.iAtualizaConsig
'        tTipoMovEst.iAtualizaDemo = objTipoMovEstoque.iAtualizaDemo
'        tTipoMovEst.iAtualizaConserto = objTipoMovEstoque.iAtualizaConserto
'        tTipoMovEst.iAtualizaOutras = objTipoMovEstoque.iAtualizaOutras
'        tTipoMovEst.iAtualizaBenef = objTipoMovEstoque.iAtualizaBenef
'        tTipoMovEst.iCustoMedio = objTipoMovEstoque.iCustoMedio
'
'        tMovEstoque.dQuantidade = -tMovEstoque.dQuantidade
'
'    End If
'
'    'Busca primeiro registro de MovimentoEstoque para atualizar custo
'    '----------------------------------------------------------------
'    lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, tMovEstoque2.lNumIntDoc, iFilialEmpresa, dtDataInicial, dtDataFinal, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO)
'    If lErro <> AD_SQL_SUCESSO Then gError 25245
'
'    lErro = Comando_BuscarPrimeiro(alComando(2))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25246
'    If lErro = AD_SQL_SEM_DADOS Then gError 25247
'
'    'Se os NumInt de MovimentoEstoque dos 2 comandos não baterem ERRO
'    If tMovEstoque.lNumIntDoc <> tMovEstoque2.lNumIntDoc Then gError 25248
'
'    'Busca primeiro registro de SaldoMesEst para atualizar Valores de Entrada, Saída, Consumo, Venda
'    '----------------------------------------------------------------------------------------
'    With tSldMesEst
'        lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), 0, .dCustoMedioProducaoInicial, dCPAtual, .adCustoMedio(iMes), .adCustoStandard(iMes), .dQuantInicial, .dValorInicial, .dQuantInicialCusto, .dValorInicialCusto, .adQuantEnt(1), .adQuantSai(1), .adValorEnt(1), .adValorSai(1), .adSaldoQuantCusto(1), .adSaldoValorCusto(1), .adQuantEnt(2), .adQuantSai(2), .adValorEnt(2), .adValorSai(2), .adSaldoQuantCusto(2), .adSaldoValorCusto(2), .adQuantEnt(3), .adQuantSai(3), .adValorEnt(3), .adValorSai(3), .adSaldoQuantCusto(3), .adSaldoValorCusto(3), .adQuantEnt(4), .adQuantSai(4), .adValorEnt(4), .adValorSai(4), .adSaldoQuantCusto(4), .adSaldoValorCusto(4), .adQuantEnt(5), .adQuantSai(5), .adValorEnt(5), .adValorSai(5), .adSaldoQuantCusto(5), .adSaldoValorCusto(5), .adQuantEnt(6), .adQuantSai(6), .adValorEnt(6), .adValorSai(6), .adSaldoQuantCusto(6), .adSaldoValorCusto(6), _
'                .adQuantEnt(7), .adQuantSai(7), .adValorEnt(7), .adValorSai(7), .adSaldoQuantCusto(7), .adSaldoValorCusto(7), .adQuantEnt(8), .adQuantSai(8), .adValorEnt(8), .adValorSai(8), .adSaldoQuantCusto(8), .adSaldoValorCusto(8), .adQuantEnt(9), .adQuantSai(9), .adValorEnt(9), .adValorSai(9), .adSaldoQuantCusto(9), .adSaldoValorCusto(9), .adQuantEnt(10), .adQuantSai(10), .adValorEnt(10), .adValorSai(10), .adSaldoQuantCusto(10), .adSaldoValorCusto(10), .adQuantEnt(11), .adQuantSai(11), .adValorEnt(11), .adValorSai(11), .adSaldoQuantCusto(11), .adSaldoValorCusto(11), .adQuantEnt(12), .adQuantSai(12), .adValorEnt(12), .adValorSai(12), .adSaldoQuantCusto(12), .adSaldoValorCusto(12), .sProduto, iAno, iFilialEmpresa)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 25249
'
'    lErro = Comando_BuscarPrimeiro(alComando(3))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25250
'    If lErro = AD_SQL_SEM_DADOS Then gError 25251
'
'    'Sincronismo deste comando com o anterior a nível de Produto
'
'    'Enquanto código do Produto de SaldoMesEst for inferior segue adiante
'    Do While tSldMesEst.sProduto < tMovEstoque.sProduto
'        lErro = Comando_BuscarProximo(alComando(3))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25252
'        If lErro = AD_SQL_SEM_DADOS Then gError 25253
'    Loop
'
'    'Se o código de Produto ultrapassou o de tMovEstoque.sProduto ERRO
'    If tSldMesEst.sProduto > tMovEstoque.sProduto Then gError 25254
'
''    'Busca primeiro registro de SaldoDiaEst para atualizar valores de Entrada, Saída, Consumo, Venda
''    '----------------------------------------------------------------------------------------
''
''    'Habilita cursor a movimentar para trás
''    Call Comando_DefScroll(alComando(4), True)
''
''    With tSldDiaEst
''        lErro = Comando_ExecutarPos(alComando(4), sComandoSQL(4), 0, .sProduto, .dtData, dtDataInicial, dtDataFinal, iFilialEmpresa)
''    End With
''
''    If lErro <> AD_SQL_SUCESSO Then gError 25293
''
''    lErro = Comando_BuscarPrimeiro(alComando(4))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25294
''    If lErro = AD_SQL_SEM_DADOS Then gError 25295
''
''    'Sincronismo deste comando com comando de MovEstoque a nível de Produto e Data
''
''    'Enquanto código do Produto e Data de SaldoDiaEst for inferior segue adiante
''    Do While tSldDiaEst.sProduto < tMovEstoque.sProduto Or (tSldDiaEst.sProduto = tMovEstoque.sProduto And tSldDiaEst.dtData < tMovEstoque.dtData)
''        lErro = Comando_BuscarProximo(alComando(4))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25296
''        If lErro = AD_SQL_SEM_DADOS Then gError 25297
''    Loop
''
''    'Se o código de Produto ou a Data ultrapassaram os de tMovEstoque ERRO
''    If tSldDiaEst.sProduto > tMovEstoque.sProduto Or tSldDiaEst.dtData > tMovEstoque.dtData Then gError 25298
'
''    'Busca primeiro registro de SaldoDiaEstAlm para atualizar valores de Entrada, Saída, Consumo, Venda
''    '----------------------------------------------------------------------------------------
''
''    'Habilita cursor a movimentar para trás
''    Call Comando_DefScroll(alComando(5), True)
''
''    With tSldDiaEstAlm
''        lErro = Comando_ExecutarPos(alComando(5), sComandoSQL(5), 0, .sProduto, .dtData, .iAlmoxarifado, dtDataInicial, dtDataFinal)
''    End With
''
''    If lErro <> AD_SQL_SUCESSO Then gError 25882
''
''    lErro = Comando_BuscarPrimeiro(alComando(5))
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25883
''    If lErro = AD_SQL_SEM_DADOS Then gError 25884
''
''    'Sincronismo deste comando com comando de MovEstoque a nível de Produto, Data, Almoxarifado
''
''    'Enquanto código do Produto, Data, Almoxarifado de SaldoDiaEstAlm for inferior segue adiante
''    Do While tSldDiaEstAlm.sProduto < tMovEstoque.sProduto Or (tSldDiaEstAlm.sProduto = tMovEstoque.sProduto And tSldDiaEstAlm.dtData < tMovEstoque.dtData) Or (tSldDiaEstAlm.sProduto = tMovEstoque.sProduto And tSldDiaEstAlm.dtData = tMovEstoque.dtData And tSldDiaEstAlm.iAlmoxarifado < tMovEstoque.iAlmoxarifado)
''        lErro = Comando_BuscarProximo(alComando(5))
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25885
''        If lErro = AD_SQL_SEM_DADOS Then gError 25886
''    Loop
''
''    'Se o código de Produto ou a Data ultrapassaram os de tMovEstoque ERRO
''    If tSldDiaEstAlm.sProduto > tMovEstoque.sProduto Or tSldDiaEstAlm.dtData > tMovEstoque.dtData Or tSldDiaEstAlm.iAlmoxarifado > tMovEstoque.iAlmoxarifado Then gError 25887
'
'    'Busca primeiro registro de SaldoMesEst2 para atualizar os Escaninhos
'    '----------------------------------------------------------------------------------------
'    Call Comando_DefScroll(alComando(6), True)
'
'    With tSldMesEst2
'        .sProduto = String(STRING_PRODUTO, 0)
'        lErro = Comando_ExecutarPos(alComando(17), sComandoSQL(6), 0, .dQuantInicialConsig, .dValorInicialConsig, .dQuantInicialDemo, .dValorInicialDemo, .dQuantInicialConserto, .dValorInicialConserto, .dQuantInicialOutros, .dValorInicialOutros, .dQuantInicialBenef, .dValorInicialBenef, .adSaldoQuantConsig(1), .adSaldoValorConsig(1), .adSaldoQuantDemo(1), .adSaldoValorDemo(1), .adSaldoQuantConserto(1), .adSaldoValorConserto(1), .adSaldoQuantOutros(1), .adSaldoValorOutros(1), .adSaldoQuantBenef(1), .adSaldoValorBenef(1), .adSaldoQuantConsig(2), .adSaldoValorConsig(2), .adSaldoQuantDemo(2), .adSaldoValorDemo(2), .adSaldoQuantConserto(2), .adSaldoValorConserto(2), .adSaldoQuantOutros(2), .adSaldoValorOutros(2), .adSaldoQuantBenef(2), .adSaldoValorBenef(2) _
'        , .adSaldoQuantConsig(3), .adSaldoValorConsig(3), .adSaldoQuantDemo(3), .adSaldoValorDemo(3), .adSaldoQuantConserto(3), .adSaldoValorConserto(3), .adSaldoQuantOutros(3), .adSaldoValorOutros(3), .adSaldoQuantBenef(3), .adSaldoValorBenef(3), .adSaldoQuantConsig(4), .adSaldoValorConsig(4), .adSaldoQuantDemo(4), .adSaldoValorDemo(4), .adSaldoQuantConserto(4), .adSaldoValorConserto(4), .adSaldoQuantOutros(4), .adSaldoValorOutros(4), .adSaldoQuantBenef(4), .adSaldoValorBenef(4), .adSaldoQuantConsig(5), .adSaldoValorConsig(5), .adSaldoQuantDemo(5), .adSaldoValorDemo(5), .adSaldoQuantConserto(5), .adSaldoValorConserto(5), .adSaldoQuantOutros(5), .adSaldoValorOutros(5), .adSaldoQuantBenef(5), .adSaldoValorBenef(5), .adSaldoQuantConsig(6), .adSaldoValorConsig(6), .adSaldoQuantDemo(6), .adSaldoValorDemo(6), .adSaldoQuantConserto(6), .adSaldoValorConserto(6), .adSaldoQuantOutros(6), .adSaldoValorOutros(6), .adSaldoQuantBenef(6), .adSaldoValorBenef(6) _
'        , .adSaldoQuantConsig(7), .adSaldoValorConsig(7), .adSaldoQuantDemo(7), .adSaldoValorDemo(7), .adSaldoQuantConserto(7), .adSaldoValorConserto(7), .adSaldoQuantOutros(7), .adSaldoValorOutros(7), .adSaldoQuantBenef(7), .adSaldoValorBenef(7), .adSaldoQuantConsig(8), .adSaldoValorConsig(8), .adSaldoQuantDemo(8), .adSaldoValorDemo(8), .adSaldoQuantConserto(8), .adSaldoValorConserto(8), .adSaldoQuantOutros(8), .adSaldoValorOutros(8), .adSaldoQuantBenef(8), .adSaldoValorBenef(8), .adSaldoQuantConsig(9), .adSaldoValorConsig(9), .adSaldoQuantDemo(9), .adSaldoValorDemo(9), .adSaldoQuantConserto(9), .adSaldoValorConserto(9), .adSaldoQuantOutros(9), .adSaldoValorOutros(9), .adSaldoQuantBenef(9), .adSaldoValorBenef(9), .adSaldoQuantConsig(10), .adSaldoValorConsig(10), .adSaldoQuantDemo(10), .adSaldoValorDemo(10), .adSaldoQuantConserto(10), .adSaldoValorConserto(10), .adSaldoQuantOutros(10), .adSaldoValorOutros(10), .adSaldoQuantBenef(10), .adSaldoValorBenef(10) _
'        , .adSaldoQuantConsig(11), .adSaldoValorConsig(11), .adSaldoQuantDemo(11), .adSaldoValorDemo(11), .adSaldoQuantConserto(11), .adSaldoValorConserto(11), .adSaldoQuantOutros(11), .adSaldoValorOutros(11), .adSaldoQuantBenef(11), .adSaldoValorBenef(11), .adSaldoQuantConsig(12), .adSaldoValorConsig(12), .adSaldoQuantDemo(12), .adSaldoValorDemo(12), .adSaldoQuantConserto(12), .adSaldoValorConserto(12), .adSaldoQuantOutros(12), .adSaldoValorOutros(12), .adSaldoQuantBenef(12), .adSaldoValorBenef(12), .sProduto, iAno, iFilialEmpresa)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 69664
'
'    lErro = Comando_BuscarPrimeiro(alComando(17))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69665
'    If lErro = AD_SQL_SEM_DADOS Then gError 69666
'
'    'Sincronismo deste comando com o anterior a nível de Produto
'
'    'Enquanto código do Produto de SaldoMesEst2 for inferior segue adiante
'    Do While tSldMesEst2.sProduto < tMovEstoque.sProduto
'        lErro = Comando_BuscarProximo(alComando(17))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69667
'        If lErro = AD_SQL_SEM_DADOS Then gError 69668
'    Loop
'
'    'Se o código de Produto ultrapassou o de tMovEstoque.sProduto ERRO
'    If tSldMesEst2.sProduto > tMovEstoque.sProduto Then gError 69669
'
'    Rotina_CMP_InicializaComandos = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_InicializaComandos:
'
'    Rotina_CMP_InicializaComandos = gErr
'
'    Select Case gErr
'
'        Case 25242
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL(1))
'
'        Case 25243, 25246
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, iFilialEmpresa)
'
'        Case 25244
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_MOVTOS_PRODUTOS_PRODUZIDOS", gErr, iFilialEmpresa, iMes, iAno)
'
'        Case 25245
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL(2))
'
'        Case 25247, 25248
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_ITEMMOVEST_MOVTOESTOQUE", gErr, tMovEstoque.lNumIntDoc)
'
'        Case 25249
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL(3))
'
'        Case 25250, 25252, 69664, 69665, 69667
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, iFilialEmpresa, iAno)
'
'        Case 25251, 25253, 25254, 69666, 69668, 69669
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SALDOMESEST", gErr, iFilialEmpresa, iAno, tMovEstoque.sProduto)
'
'        Case 25293
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL(4))
'
'        Case 25294, 25296
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST2", gErr, iFilialEmpresa)
'
'        Case 25295, 25297, 25298
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SLDDIAEST", gErr, iFilialEmpresa, tMovEstoque.sProduto, tMovEstoque.dtData)
'
'        Case 25882
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL(6))
'
'        Case 25883, 25885
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM2", gErr)
'
'        Case 25884, 25886, 25887
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SLDDIAESTALM", gErr, tMovEstoque.sProduto, tMovEstoque.dtData, tMovEstoque.iAlmoxarifado)
'
'        Case 89157
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159498)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaCustosProducao(sComandoSQL() As String, alComando() As Long, tProduto As typeProduto, dCPAtual As Double, tTipoMovEst As typeTipoMovEst, tMovEstoque As typeItemMovEstoque, tMovEstoque2 As typeItemMovEstoque, tSldMesEst As typeSldMesEst, tSldDiaEst As typeSldDiaEst, iAno As Integer, iMes As Integer, iFilialEmpresa As Integer, dCMPAtual As Double, colAlmoxInfo As Collection, tSldMesEst1 As typeSldMesEst1) As Long
''Chamada EM TRANSAÇÃO
''Atualiza custos dos movtos de produção de materiais produzidos
''Entradas de Producao pelo CustoProducaoReal
''Devolve Custo Medio Producao atual em dCMPAtual
'
'Dim lErro As Long
'Dim dQuantEntApropCustoProd As Double
'Dim iApropriacao As Integer
'Dim objAlmoxInfo As ClassAlmoxInfo
'Dim tSldDiaEstAlm As typeSldDiaEstAlm
'Dim dSaldoValorCustoInformado As Double
'Dim tTipoMovEst1 As typeTipoMovEst
'
'On Error GoTo Erro_Rotina_CMP_AtualizaCustosProducao
'
'    'Zera as variáveis de memória com valores de entrada, saída, consumo
'    tSldMesEst.adValorEnt(iMes) = 0
'    tSldMesEst.adValorSai(iMes) = 0
'    tSldMesEst.adValorCons(iMes) = 0
'
'    dSaldoValorCustoInformado = tSldMesEst.adSaldoValorCusto(iMes)
'    tSldMesEst.adSaldoValorCusto(iMes) = 0
'
'    tSldMesEst1.adSaldoValorConsig3(iMes) = 0
'    tSldMesEst1.adSaldoValorDemo3(iMes) = 0
'    tSldMesEst1.adSaldoValorConserto3(iMes) = 0
'    tSldMesEst1.adSaldoValorOutros3(iMes) = 0
'    tSldMesEst1.adSaldoValorBenef3(iMes) = 0
'    tSldMesEst1.adSaldoQuantBenef3(iMes) = 0
'
'    'Reinicializa coleção de Alomoxarifados para o Produto
'    Set colAlmoxInfo = New Collection
'
'    iApropriacao = APROPR_CUSTO_REAL
'
'    'Se houve movto de Producao, testa se CustoProducao de iMes foi informado
'    If tMovEstoque.sProduto = tSldMesEst.sProduto And tMovEstoque.iApropriacao = APROPR_CUSTO_REAL And dCPAtual = 0 Then gError 25255
'
'    'Movimentos com APROPRIACAO=CustoRealProducao (Producao)
'    'Loop POR DIA
'    Do While tMovEstoque.sProduto = tSldMesEst.sProduto And tMovEstoque.iApropriacao = APROPR_CUSTO_REAL
'
'        tSldDiaEst.dValorEntrada = 0
'        tSldDiaEst.dValorSaida = 0
'        tSldDiaEst.dValorCons = 0
'        tSldDiaEst.dValorEntCusto = 0
'        tSldDiaEst.dValorSaiCusto = 0
'        tSldDiaEst.dValorEntConsig = 0
'        tSldDiaEst.dValorEntConsig3 = 0
'        tSldDiaEst.dValorEntDemo = 0
'        tSldDiaEst.dValorEntDemo3 = 0
'        tSldDiaEst.dValorEntConserto = 0
'        tSldDiaEst.dValorEntConserto3 = 0
'        tSldDiaEst.dValorEntOutros = 0
'        tSldDiaEst.dValorEntOutros3 = 0
'        tSldDiaEst.dValorEntBenef = 0
'        tSldDiaEst.dValorEntBenef3 = 0
'        tSldDiaEst.dValorSaiConsig = 0
'        tSldDiaEst.dValorSaiConsig3 = 0
'        tSldDiaEst.dValorSaiDemo = 0
'        tSldDiaEst.dValorSaiDemo3 = 0
'        tSldDiaEst.dValorSaiConserto = 0
'        tSldDiaEst.dValorSaiConserto3 = 0
'        tSldDiaEst.dValorSaiOutros = 0
'        tSldDiaEst.dValorSaiOutros3 = 0
'        tSldDiaEst.dValorSaiBenef = 0
'        tSldDiaEst.dValorSaiBenef3 = 0
'
'        tSldDiaEstAlm.sProduto = tMovEstoque.sProduto
'        tSldDiaEstAlm.dtData = tMovEstoque.dtData
'        tSldDiaEst.sProduto = tMovEstoque.sProduto
'        tSldDiaEst.dtData = tMovEstoque.dtData
'
'        'Loop por ALMOXARIFADO
'        Do While tMovEstoque.sProduto = tSldMesEst.sProduto And tMovEstoque.iApropriacao = APROPR_CUSTO_REAL And tMovEstoque.dtData = tSldDiaEst.dtData
'
'            tSldDiaEstAlm.dValorEntrada = 0
'            tSldDiaEstAlm.dValorEntCusto = 0
'            tSldDiaEstAlm.dValorSaida = 0
'            tSldDiaEstAlm.dValorSaiCusto = 0
'            tSldDiaEstAlm.dValorCons = 0
'            tSldDiaEstAlm.dValorEntConsig = 0
'            tSldDiaEstAlm.dValorEntConsig3 = 0
'            tSldDiaEstAlm.dValorEntDemo = 0
'            tSldDiaEstAlm.dValorEntDemo3 = 0
'            tSldDiaEstAlm.dValorEntConserto = 0
'            tSldDiaEstAlm.dValorEntConserto3 = 0
'            tSldDiaEstAlm.dValorEntOutros = 0
'            tSldDiaEstAlm.dValorEntOutros3 = 0
'            tSldDiaEstAlm.dValorEntBenef = 0
'            tSldDiaEstAlm.dValorEntBenef3 = 0
'            tSldDiaEstAlm.dValorSaiConsig = 0
'            tSldDiaEstAlm.dValorSaiConsig3 = 0
'            tSldDiaEstAlm.dValorSaiDemo = 0
'            tSldDiaEstAlm.dValorSaiDemo3 = 0
'            tSldDiaEstAlm.dValorSaiConserto = 0
'            tSldDiaEstAlm.dValorSaiConserto3 = 0
'            tSldDiaEstAlm.dValorSaiOutros = 0
'            tSldDiaEstAlm.dValorSaiOutros3 = 0
'            tSldDiaEstAlm.dValorSaiBenef = 0
'            tSldDiaEstAlm.dValorSaiBenef3 = 0
'
'            'Testa existencia de Almoxarifado do Movimento na coleção, devolve em objAlmoxInfo
'            lErro = Rotina_CMP_TestaAlmoxarifado(tMovEstoque, colAlmoxInfo, objAlmoxInfo)
'            If lErro <> SUCESSO Then gError 25912
'
'            tSldDiaEstAlm.iAlmoxarifado = tMovEstoque.iAlmoxarifado
'
'            'Loop enquanto fica o mesmo ALMOXARIFADO
'            Do While tMovEstoque.sProduto = tSldMesEst.sProduto And tMovEstoque.iApropriacao = APROPR_CUSTO_REAL And tMovEstoque.dtData = tSldDiaEst.dtData And tMovEstoque.iAlmoxarifado = tSldDiaEstAlm.iAlmoxarifado
'
'                'Atualiza CustoRealProducao nos movimentos
'                lErro = Rotina_CMP_AtualizaCP(sComandoSQL(), alComando(), dQuantEntApropCustoProd, tProduto, dCPAtual, tTipoMovEst, tMovEstoque, tMovEstoque2, tSldMesEst, tSldDiaEst, tSldDiaEstAlm, objAlmoxInfo, iMes, iFilialEmpresa, tSldMesEst1)
'                If lErro <> SUCESSO And lErro <> 25274 Then gError 25256
'
'                If lErro = 25274 Then 'Não tem + movimentos
'
'                    'Atualiza Valor de Entrada em SaldoDiaEstAlm
'                    lErro = Rotina_CMP_AtualizaSldDiaEstAlm(alComando(), tSldDiaEstAlm)
'                    If lErro <> SUCESSO Then gError 25890
'
'                    'Atualiza valor de Entrada em SaldoDiaEst
'                    lErro = Rotina_CMP_AtualizaSldDiaEst(alComando(), tSldDiaEst, iFilialEmpresa)
'                    If lErro <> SUCESSO Then gError 25299
'
'                    'Cálculo de Custo Médio de Produção Atual
'                    lErro = Rotina_CMP_CMPAtualCalcula(alComando(), tTipoMovEst, tSldMesEst, tSldMesEst1, dQuantEntApropCustoProd, iMes, iAno, dCMPAtual, dSaldoValorCustoInformado, iFilialEmpresa)
'                    If lErro <> SUCESSO Then gError 78026
'
'                    'Atualiza Custo (Médio de Produção) em SldMesEst
'                    lErro = Rotina_CMP_AlteraCMP(tSldMesEst, iMes, dCMPAtual)
'                    If lErro <> SUCESSO Then gError 78027
'
'                    'Atualiza Valores de Entrada de iMes em SaldoMesEstAlm
'                    lErro = Rotina_CMP_AtualizaSldMesEstAlm(alComando(), tSldMesEst.sProduto, colAlmoxInfo, iMes, iAno)
'                    If lErro <> SUCESSO Then gError 25891
'
'                    'Atualiza os Valores dos escaninhos de terceiros em SaldoMesEstAlm1
'                    lErro = Rotina_CMP_AtualizaSldMesEstAlm1(alComando(), tSldMesEst1.sProduto, colAlmoxInfo, iMes, iAno)
'                    If lErro <> SUCESSO Then gError 89830
'
'                    'Atualiza Valor de Entrada de iMes em SaldoMesEst
'                    lErro = Rotina_CMP_AtualizaSldMesEst(alComando(), tSldMesEst, iMes, iAno, iFilialEmpresa)
'                    If lErro <> SUCESSO Then gError 25257
'
'                    'Atualiza Valores dos escaninhos de terceiros em SaldoMesEst1
'                    lErro = Rotina_CMP_AtualizaSldMesEst1(alComando(), tSldMesEst1, iMes, iAno, iFilialEmpresa, dCMPAtual, tTipoMovEst)
'                    If lErro <> SUCESSO Then gError 89831
'
'                    gError 25258  'Não tem + movimentos. Sai da função.
'
'                End If
'
'            Loop
'
'            'Atualiza Valor de Entrada em SaldoDiaEstAlm
'            lErro = Rotina_CMP_AtualizaSldDiaEstAlm(alComando(), tSldDiaEstAlm)
'            If lErro <> SUCESSO Then gError 25888
'
'        Loop
'
'        'Atualiza Valor de Entrada em SaldoDiaEst
'        lErro = Rotina_CMP_AtualizaSldDiaEst(alComando(), tSldDiaEst, iFilialEmpresa)
'        If lErro <> SUCESSO Then gError 25300
'
'    Loop
'
'    'Cálculo de Custo Médio de Produção Atual
'    lErro = Rotina_CMP_CMPAtualCalcula(alComando(), tTipoMovEst, tSldMesEst, tSldMesEst1, dQuantEntApropCustoProd, iMes, iAno, dCMPAtual, dSaldoValorCustoInformado, iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 25259
'
'    'Atualiza Custo (Médio de Produção) em SldMesEst
'    lErro = Rotina_CMP_AlteraCMP(tSldMesEst, iMes, dCMPAtual)
'    If lErro <> SUCESSO Then gError 62557
'
'    Rotina_CMP_AtualizaCustosProducao = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaCustosProducao:
'
'    Rotina_CMP_AtualizaCustosProducao = gErr
'
'    Select Case gErr
'
'        Case 25255
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTO_PRODUCAO_NAO_INFORMADO", gErr, tMovEstoque.sProduto, iMes, iAno)
'
'        Case 25256, 25299, 25257, 25300, 25301, 25259, 25888, 25889, 25890, 25891, 25912, 78026, 78027, 89830, 89831  'Tratado na rotina chamada
'
'        Case 25258 'Não tem mais registros de movimentos
'                   'Tratado na rotina chamadora
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159499)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaCustos(sComandoSQL() As String, alComando() As Long, tProduto As typeProduto, tTipoMovEst As typeTipoMovEst, tMovEstoque As typeItemMovEstoque, tMovEstoque2 As typeItemMovEstoque, tSldMesEst As typeSldMesEst, tSldDiaEst As typeSldDiaEst, iAno As Integer, iMes As Integer, iFilialEmpresa As Integer, dCMPAtual As Double, colAlmoxInfo As Collection, tSldMesEst2 As typeSldMesEst2, tSldMesEst1 As typeSldMesEst1) As Long
''Chamada EM TRANSAÇÃO
''Atualiza custos dos movtos de materiais produzidos com EXCECAO da PRODUCAO
''Entradas e Saidas pelo CustoMédioProdução passado em dCMPAtual
'
'Dim lErro As Long
'Dim iApropriacao As Integer
'Dim objAlmoxInfo As ClassAlmoxInfo
'Dim dCMPConsigAtual As Double, dCMPDemoAtual As Double, dCMPConsertoAtual As Double, dCMPOutrasAtual As Double, dCMPBenefAtual As Double
'Dim tSldDiaEstAlm As typeSldDiaEstAlm
'
'On Error GoTo Erro_Rotina_CMP_AtualizaCustos
'
'    'Zera os Saldos para o Mês que está sendo apurado porque só me interessa as Quantidades e Valores de Saídas
'    tSldMesEst2.adSaldoValorConsig(iMes) = 0
'    tSldMesEst2.adSaldoValorDemo(iMes) = 0
'    tSldMesEst2.adSaldoValorConserto(iMes) = 0
'    tSldMesEst2.adSaldoValorOutros(iMes) = 0
'    tSldMesEst2.adSaldoValorBenef(iMes) = 0
'    tSldMesEst2.adSaldoQuantConsig(iMes) = 0
'    tSldMesEst2.adSaldoQuantDemo(iMes) = 0
'    tSldMesEst2.adSaldoQuantConserto(iMes) = 0
'    tSldMesEst2.adSaldoQuantOutros(iMes) = 0
'    tSldMesEst2.adSaldoQuantBenef(iMes) = 0
'
'    iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO
'
'    'Movimentos com APROPRIACAO_CUSTO_MEDIO_PRODUCAO
'    'Loop por Entrada e Saída
'    Do While tMovEstoque.sProduto = tSldMesEst.sProduto And tMovEstoque.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO
'
'        'Loop por Produto/Apropriacao
'        'Acumula os valores de Saídas dos escaninhos
'        Do While tMovEstoque.sProduto = tSldMesEst.sProduto And tMovEstoque.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And tTipoMovEst.sEntradaOuSaida = TIPOMOV_EST_SAIDA
'
'            tSldDiaEstAlm.sProduto = tMovEstoque.sProduto
'            tSldDiaEstAlm.dtData = tMovEstoque.dtData
'
'            tSldDiaEst.sProduto = tMovEstoque.sProduto
'            tSldDiaEst.dtData = tMovEstoque.dtData
'
'            tSldDiaEst.dValorEntrada = 0
'            tSldDiaEst.dValorSaida = 0
'            tSldDiaEst.dValorCons = 0
'            tSldDiaEst.dValorVend = 0
'            tSldDiaEst.dValorEntCusto = 0
'            tSldDiaEst.dValorSaiCusto = 0
'            tSldDiaEst.dValorEntConsig = 0
'            tSldDiaEst.dValorEntConsig3 = 0
'            tSldDiaEst.dValorEntDemo = 0
'            tSldDiaEst.dValorEntDemo3 = 0
'            tSldDiaEst.dValorEntConserto = 0
'            tSldDiaEst.dValorEntConserto3 = 0
'            tSldDiaEst.dValorEntOutros = 0
'            tSldDiaEst.dValorEntOutros3 = 0
'            tSldDiaEst.dValorEntBenef = 0
'            tSldDiaEst.dValorEntBenef3 = 0
'            tSldDiaEst.dValorSaiConsig = 0
'            tSldDiaEst.dValorSaiConsig3 = 0
'            tSldDiaEst.dValorSaiDemo = 0
'            tSldDiaEst.dValorSaiDemo3 = 0
'            tSldDiaEst.dValorSaiConserto = 0
'            tSldDiaEst.dValorSaiConserto3 = 0
'            tSldDiaEst.dValorSaiOutros = 0
'            tSldDiaEst.dValorSaiOutros3 = 0
'            tSldDiaEst.dValorSaiBenef = 0
'            tSldDiaEst.dValorSaiBenef3 = 0
'
'            'Loop por dia
'            Do While tMovEstoque.sProduto = tSldMesEst.sProduto And tMovEstoque.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And tTipoMovEst.sEntradaOuSaida = TIPOMOV_EST_SAIDA And tMovEstoque.dtData = tSldDiaEst.dtData
'
'                tSldDiaEstAlm.dValorEntrada = 0
'                tSldDiaEstAlm.dValorEntCusto = 0
'                tSldDiaEstAlm.dValorSaida = 0
'                tSldDiaEstAlm.dValorSaiCusto = 0
'                tSldDiaEstAlm.dValorCons = 0
'                tSldDiaEstAlm.dValorVend = 0
'                tSldDiaEstAlm.dValorEntCusto = 0
'                tSldDiaEstAlm.dValorSaiCusto = 0
'                tSldDiaEstAlm.dValorEntConsig = 0
'                tSldDiaEstAlm.dValorEntConsig3 = 0
'                tSldDiaEstAlm.dValorEntDemo = 0
'                tSldDiaEstAlm.dValorEntDemo3 = 0
'                tSldDiaEstAlm.dValorEntConserto = 0
'                tSldDiaEstAlm.dValorEntConserto3 = 0
'                tSldDiaEstAlm.dValorEntOutros = 0
'                tSldDiaEstAlm.dValorEntOutros3 = 0
'                tSldDiaEstAlm.dValorEntBenef = 0
'                tSldDiaEstAlm.dValorEntBenef3 = 0
'                tSldDiaEstAlm.dValorSaiConsig = 0
'                tSldDiaEstAlm.dValorSaiConsig3 = 0
'                tSldDiaEstAlm.dValorSaiDemo = 0
'                tSldDiaEstAlm.dValorSaiDemo3 = 0
'                tSldDiaEstAlm.dValorSaiConserto = 0
'                tSldDiaEstAlm.dValorSaiConserto3 = 0
'                tSldDiaEstAlm.dValorSaiOutros = 0
'                tSldDiaEstAlm.dValorSaiOutros3 = 0
'                tSldDiaEstAlm.dValorSaiBenef = 0
'                tSldDiaEstAlm.dValorSaiBenef3 = 0
'
'                'Testa existencia de Almoxarifado do Movimento na coleção, devolve em objAlmoxInfo
'                lErro = Rotina_CMP_TestaAlmoxarifado(tMovEstoque, colAlmoxInfo, objAlmoxInfo)
'                If lErro <> SUCESSO Then gError 69700
'
'                tSldDiaEstAlm.iAlmoxarifado = tMovEstoque.iAlmoxarifado
'
'                'Loop enquanto fica o mesmo ALMOXARIFADO
'                Do While tMovEstoque.sProduto = tSldMesEst.sProduto And tMovEstoque.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And tTipoMovEst.sEntradaOuSaida = TIPOMOV_EST_SAIDA And tMovEstoque.dtData = tSldDiaEst.dtData And tMovEstoque.iAlmoxarifado = tSldDiaEstAlm.iAlmoxarifado
'
'                    'Acumula as Saídas para cada escaninhos
'                    lErro = Rotina_CMP_AtualizaSaidasCMP(sComandoSQL(), alComando(), tProduto, dCMPAtual, tTipoMovEst, tMovEstoque, tMovEstoque2, tSldMesEst, tSldDiaEst, tSldDiaEstAlm, objAlmoxInfo, iMes, iFilialEmpresa, tSldMesEst2, tSldMesEst1)
'                    If lErro <> SUCESSO And lErro <> 25270 Then gError 69701
'
''                    lErro = Rotina_CMP_AtualizaCMP(sComandoSQL(), alComando(), tProduto, dCMPAtual, tTipoMovEst, tMovEstoque, tMovEstoque2, tSldMesEst, tSldDiaEst, tSldDiaEstAlm, objAlmoxInfo, iMes, iFilialEmpresa)
''                    If lErro <> SUCESSO And lErro <> 25270 Then gError 25260
'
'                    If lErro = 25270 Then  'Não tem + registros de movimento
'
'                        'Atualiza Valores de Entrada, Saida, Consumo em SaldoDiaEstAlm
'                        lErro = Rotina_CMP_AtualizaSldDiaEstAlm(alComando(), tSldDiaEstAlm)
'                        If lErro <> SUCESSO Then gError 69702
'
'                        'Atualiza valores de Entrada, Saida, Consumo em SaldoDiaEst
'                        lErro = Rotina_CMP_AtualizaSldDiaEst(alComando(), tSldDiaEst, iFilialEmpresa)
'                        If lErro <> SUCESSO Then gError 69703
'
'                        'Atualiza Valores de Entrada, Saida, Consumo de iMes em SaldoMesEstAlm
'                        lErro = Rotina_CMP_AtualizaSldMesEstAlm(alComando(), tSldMesEst.sProduto, colAlmoxInfo, iMes, iAno)
'                        If lErro <> SUCESSO Then gError 69704
'
'                        'Atualiza os Valores dos escaninhos de terceiros em SaldoMesEstAlm1
'                        lErro = Rotina_CMP_AtualizaSldMesEstAlm1(alComando(), tSldMesEst1.sProduto, colAlmoxInfo, iMes, iAno)
'                        If lErro <> SUCESSO Then gError 89820
'
'                        'Atualiza Valores dos produtos nossos em terceiros SaldoMesEstAlm2
'                        lErro = Rotina_CMP_AtualizaSldMesEstAlm2(alComando(), tSldMesEst2.sProduto, colAlmoxInfo, iMes, iAno)
'                        If lErro <> SUCESSO Then gError 69705
'
'                        'Atualiza Valores de Entrada, Saida, Consumo de iMes em SaldoMesEst
'                        lErro = Rotina_CMP_AtualizaSldMesEst(alComando(), tSldMesEst, iMes, iAno, iFilialEmpresa)
'                        If lErro <> SUCESSO Then gError 69706
'
'                        'Atualiza Valores dos escaninhos de terceiros em SaldoMesEst1
'                        lErro = Rotina_CMP_AtualizaSldMesEst1(alComando(), tSldMesEst1, iMes, iAno, iFilialEmpresa, dCMPAtual, tTipoMovEst)
'                        If lErro <> SUCESSO Then gError 89826
'
'                        'Atualiza Valores dos escaninhos nossos em poder de terceiros SaldoMesEst2
'                        lErro = Rotina_CMP_AtualizaSldMesEst2(alComando(), tSldMesEst2, iMes, iAno, iFilialEmpresa)
'                        If lErro <> SUCESSO Then gError 69707
'
'                        gError 69708 'Não tem + registros de movimento. Sai da função.
'
'                    End If
'
'                Loop
'
'                'Atualiza Valores de Entrada, Saida, Consumo em SaldoDiaEstAlm
'                lErro = Rotina_CMP_AtualizaSldDiaEstAlm(alComando(), tSldDiaEstAlm)
'                If lErro <> SUCESSO Then gError 69709
'
'            Loop
'
'            'Atualiza Valores de Entrada, Saida, Consumo em SaldoDiaEst
'            lErro = Rotina_CMP_AtualizaSldDiaEst(alComando(), tSldDiaEst, iFilialEmpresa)
'            If lErro <> SUCESSO Then gError 69761
'
'        Loop
'
'        'Calcula CustoMedioProducao dos escaninhos
'        lErro = Rotina_CMP_CMPTerceirosCalcula(tSldMesEst2, iMes, dCMPConsigAtual, dCMPDemoAtual, dCMPConsertoAtual, dCMPOutrasAtual, dCMPBenefAtual)
'        If lErro <> SUCESSO Then gError 69711
'
'        'Loop por dia
'        'Depois de Calculado o Custo por escaninho acumula os valores de Entrada utilizando o custo dos escaninhos
'        Do While tMovEstoque.sProduto = tSldMesEst.sProduto And tMovEstoque.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And tTipoMovEst.sEntradaOuSaida = TIPOMOV_EST_ENTRADA
'
'            tSldDiaEstAlm.sProduto = tMovEstoque.sProduto
'            tSldDiaEstAlm.dtData = tMovEstoque.dtData
'
'            tSldDiaEst.sProduto = tMovEstoque.sProduto
'            tSldDiaEst.dtData = tMovEstoque.dtData
'
'            'Loop por ALMOXARIFADO
'            Do While tMovEstoque.sProduto = tSldMesEst.sProduto And tMovEstoque.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And tTipoMovEst.sEntradaOuSaida = TIPOMOV_EST_ENTRADA And tMovEstoque.dtData = tSldDiaEst.dtData
'
'                tSldDiaEstAlm.dValorEntrada = 0
'                tSldDiaEstAlm.dValorEntCusto = 0
'                tSldDiaEstAlm.dValorSaida = 0
'                tSldDiaEstAlm.dValorSaiCusto = 0
'                tSldDiaEstAlm.dValorCons = 0
'                tSldDiaEstAlm.dValorVend = 0
'                tSldDiaEstAlm.dValorEntConsig = 0
'                tSldDiaEstAlm.dValorEntConsig3 = 0
'                tSldDiaEstAlm.dValorEntDemo = 0
'                tSldDiaEstAlm.dValorEntDemo3 = 0
'                tSldDiaEstAlm.dValorEntConserto = 0
'                tSldDiaEstAlm.dValorEntConserto3 = 0
'                tSldDiaEstAlm.dValorEntOutros = 0
'                tSldDiaEstAlm.dValorEntOutros3 = 0
'                tSldDiaEstAlm.dValorEntBenef = 0
'                tSldDiaEstAlm.dValorEntBenef3 = 0
'                tSldDiaEstAlm.dValorSaiConsig = 0
'                tSldDiaEstAlm.dValorSaiConsig3 = 0
'                tSldDiaEstAlm.dValorSaiDemo = 0
'                tSldDiaEstAlm.dValorSaiDemo3 = 0
'                tSldDiaEstAlm.dValorSaiConserto = 0
'                tSldDiaEstAlm.dValorSaiConserto3 = 0
'                tSldDiaEstAlm.dValorSaiOutros = 0
'                tSldDiaEstAlm.dValorSaiOutros3 = 0
'                tSldDiaEstAlm.dValorSaiBenef = 0
'                tSldDiaEstAlm.dValorSaiBenef3 = 0
'
'                'Testa existencia de Almoxarifado do Movimento na coleção, devolve em objAlmoxInfo
'                lErro = Rotina_CMP_TestaAlmoxarifado(tMovEstoque, colAlmoxInfo, objAlmoxInfo)
'                If lErro <> SUCESSO Then gError 69712
'
'                tSldDiaEstAlm.iAlmoxarifado = tMovEstoque.iAlmoxarifado
'
'                'Loop enquanto fica o mesmo ALMOXARIFADO
'                Do While tMovEstoque.sProduto = tSldMesEst.sProduto And tMovEstoque.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And tTipoMovEst.sEntradaOuSaida = TIPOMOV_EST_ENTRADA And tMovEstoque.dtData = tSldDiaEst.dtData And tMovEstoque.iAlmoxarifado = tSldDiaEstAlm.iAlmoxarifado
'
'                    'Atualiza as Entradas e Atualiza CustoMedioProducao nos Movimentos
'                    lErro = Rotina_CMP_AtualizaEntradasCMP(sComandoSQL(), alComando(), tProduto, dCMPAtual, tTipoMovEst, tMovEstoque, tMovEstoque2, tSldMesEst, tSldDiaEst, tSldDiaEstAlm, objAlmoxInfo, iMes, iFilialEmpresa, tSldMesEst2, dCMPConsigAtual, dCMPDemoAtual, dCMPConsertoAtual, dCMPOutrasAtual, dCMPBenefAtual, tSldMesEst1)
'                    If lErro <> SUCESSO And lErro <> 69767 Then gError 69713
'
'                    If lErro = 69767 Then  'Não tem + registros de movimento
'
'                        'Atualiza Valores de Entrada, Saida, Consumo, Venda em SaldoDiaEstAlm
'                        lErro = Rotina_CMP_AtualizaSldDiaEstAlm(alComando(), tSldDiaEstAlm)
'                        If lErro <> SUCESSO Then gError 69714
'
'                        'Atualiza valores de Entrada, Saida, Consumo, Venda em SaldoDiaEst
'                        lErro = Rotina_CMP_AtualizaSldDiaEst(alComando(), tSldDiaEst, iFilialEmpresa)
'                        If lErro <> SUCESSO Then gError 69715
'
'                        'Atualiza Valores de Entrada, Saida, Consumo, Venda de iMes em SaldoMesEstAlm
'                        lErro = Rotina_CMP_AtualizaSldMesEstAlm(alComando(), tSldMesEst.sProduto, colAlmoxInfo, iMes, iAno)
'                        If lErro <> SUCESSO Then gError 69716
'
'                        'Atualiza os Valores dos escaninhos de terceiros em SaldoMesEstAlm1
'                        lErro = Rotina_CMP_AtualizaSldMesEstAlm1(alComando(), tSldMesEst2.sProduto, colAlmoxInfo, iMes, iAno)
'                        If lErro <> SUCESSO Then gError 89821
'
'                        'Atualiza Valores de Entrada, Saida, Consumo, Venda de iMes em SaldoMesEstAlm
'                        lErro = Rotina_CMP_AtualizaSldMesEstAlm2(alComando(), tSldMesEst2.sProduto, colAlmoxInfo, iMes, iAno)
'                        If lErro <> SUCESSO Then gError 69717
'
'                        'Atualiza Valores de Entrada, Saida, Consumo, Venda de iMes em SaldoMesEst
'                        lErro = Rotina_CMP_AtualizaSldMesEst(alComando(), tSldMesEst, iMes, iAno, iFilialEmpresa)
'                        If lErro <> SUCESSO Then gError 69718
'
'                        'Atualiza Valores dos escaninhos de terceiros em SaldoMesEst1
'                        lErro = Rotina_CMP_AtualizaSldMesEst1(alComando(), tSldMesEst1, iMes, iAno, iFilialEmpresa, dCMPAtual, tTipoMovEst)
'                        If lErro <> SUCESSO Then gError 89827
'
'                        'Atualiza Valores de Entrada, Saida, Consumo, Venda de iMes em SaldoMesEst
'                        lErro = Rotina_CMP_AtualizaSldMesEst2(alComando(), tSldMesEst2, iMes, iAno, iFilialEmpresa)
'                        If lErro <> SUCESSO Then gError 69719
'
'                        gError 69720 'Não tem + registros de movimento. Sai da função.
'                    End If
'
'                Loop
'
'                'Atualiza Valores de Entrada, Saida, Consumo em SaldoDiaEstAlm
'                lErro = Rotina_CMP_AtualizaSldDiaEstAlm(alComando(), tSldDiaEstAlm)
'                If lErro <> SUCESSO Then gError 69721
'
'            Loop
'
'            'Atualiza Valores de Entrada, Saida, Consumo, Venda em SaldoDiaEst
'            lErro = Rotina_CMP_AtualizaSldDiaEst(alComando(), tSldDiaEst, iFilialEmpresa)
'            If lErro <> SUCESSO Then gError 69723
'
'        Loop
'
'    Loop
'
'    'Atualiza Valores de Entrada, Saida, Consumo, Venda de iMes em SaldoMesEstAlm
'    lErro = Rotina_CMP_AtualizaSldMesEstAlm(alComando(), tSldMesEst.sProduto, colAlmoxInfo, iMes, iAno)
'    If lErro <> SUCESSO Then gError 69725
'
'    'Atualiza os Valores dos escaninhos de terceiros em SaldoMesEstAlm1
'    lErro = Rotina_CMP_AtualizaSldMesEstAlm1(alComando(), tSldMesEst2.sProduto, colAlmoxInfo, iMes, iAno)
'    If lErro <> SUCESSO Then gError 89822
'
'    'Atualiza Valores dos escaninhos
'    lErro = Rotina_CMP_AtualizaSldMesEstAlm2(alComando(), tSldMesEst2.sProduto, colAlmoxInfo, iMes, iAno)
'    If lErro <> SUCESSO Then gError 69726
'
'    'Atualiza Valores de Entrada e Saida de iMes em SaldoMesEst
'    lErro = Rotina_CMP_AtualizaSldMesEst(alComando(), tSldMesEst, iMes, iAno, iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 69727
'
'    'Atualiza Valores dos escaninhos de terceiros em SaldoMesEst1
'    lErro = Rotina_CMP_AtualizaSldMesEst1(alComando(), tSldMesEst1, iMes, iAno, iFilialEmpresa, dCMPAtual, tTipoMovEst)
'    If lErro <> SUCESSO Then gError 89828
'
'    'Atualiza Valores dos escaninhos
'    lErro = Rotina_CMP_AtualizaSldMesEst2(alComando(), tSldMesEst2, iMes, iAno, iFilialEmpresa)
'    If lErro <> SUCESSO Then gError 69728
'
'    'Busca próximo registro de SldMesEst
'    lErro = Rotina_CMP_ProximoSldMesEst(alComando(), tSldMesEst, tMovEstoque, iFilialEmpresa, iAno)
'    If lErro <> SUCESSO Then gError 69729
'
'    'Busca próximo registro de SldMesEst
'    lErro = Rotina_CMP_ProximoSldMesEst2(alComando(), tSldMesEst2, tMovEstoque, iFilialEmpresa, iAno)
'    If lErro <> SUCESSO Then gError 69730
'
'    Rotina_CMP_AtualizaCustos = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaCustos:
'
'    Rotina_CMP_AtualizaCustos = gErr
'
'    Select Case gErr
'
'        Case 69700 To 69707, 69709 To 69719, 69721 To 69730, 69761, 69762
'
'        Case 69708, 69720, 89820, 89821, 89822, 89826, 89827, 89828
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159500)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_CMPTerceirosCalcula(tSldMesEst2 As typeSldMesEst2, iMes As Integer, dCMPConsigAtual As Double, dCMPDemoAtual As Double, dCMPConsertoAtual As Double, dCMPOutrasAtual As Double, dCMPBenefAtual As Double) As Long
''Calcula Custo Médio de Produção Atual para os Escaninhos de terceiros
'
'Dim dQuantAcumuladaConsig As Double, dQuantAcumuladaDemo As Double, dQuantAcumuladaConserto As Double, dQuantAcumuladaOutras As Double, dQuantAcumuladaBenef As Double
'Dim dValorAcumuladaConsig As Double, dValorAcumuladaDemo As Double, dValorAcumuladaConserto As Double, dValorAcumuladaOutras As Double, dValorAcumuladaBenef As Double
'Dim iMesFinal As Integer
'Dim iIndice As Integer
'
'On Error GoTo Erro_Rotina_CMP_CMPTerceirosCalcula
'
'    'Acumula quantidade inicial
'    dQuantAcumuladaBenef = tSldMesEst2.dQuantInicialBenef
'    dQuantAcumuladaConserto = tSldMesEst2.dQuantInicialConserto
'    dQuantAcumuladaConsig = tSldMesEst2.dQuantInicialConsig
'    dQuantAcumuladaDemo = tSldMesEst2.dQuantInicialDemo
'    dQuantAcumuladaOutras = tSldMesEst2.dQuantInicialOutros
'
'    'Adiciona saldos dos meses anteriores
'    For iIndice = 1 To iMes
'        dQuantAcumuladaBenef = dQuantAcumuladaBenef + tSldMesEst2.adSaldoQuantBenef(iIndice)
'        dQuantAcumuladaConserto = dQuantAcumuladaConserto + tSldMesEst2.adSaldoQuantConserto(iIndice)
'        dQuantAcumuladaConsig = dQuantAcumuladaConsig + tSldMesEst2.adSaldoQuantConsig(iIndice)
'        dQuantAcumuladaDemo = dQuantAcumuladaDemo + tSldMesEst2.adSaldoQuantDemo(iIndice)
'        dQuantAcumuladaOutras = dQuantAcumuladaOutras + tSldMesEst2.adSaldoQuantOutros(iIndice)
'    Next
'
'    'Acumula valor inicial
'    dValorAcumuladaBenef = tSldMesEst2.dValorInicialBenef
'    dValorAcumuladaConserto = tSldMesEst2.dValorInicialConserto
'    dValorAcumuladaConsig = tSldMesEst2.dValorInicialConsig
'    dValorAcumuladaDemo = tSldMesEst2.dValorInicialDemo
'    dValorAcumuladaOutras = tSldMesEst2.dValorInicialOutros
'
'    'Adiciona saldos nos Meses
'    For iIndice = 1 To iMes
'        dValorAcumuladaBenef = dValorAcumuladaBenef + tSldMesEst2.adSaldoValorBenef(iIndice)
'        dValorAcumuladaConserto = dValorAcumuladaConserto + tSldMesEst2.adSaldoValorConserto(iIndice)
'        dValorAcumuladaConsig = dValorAcumuladaConsig + tSldMesEst2.adSaldoValorConsig(iIndice)
'        dValorAcumuladaDemo = dValorAcumuladaDemo + tSldMesEst2.adSaldoValorDemo(iIndice)
'        dValorAcumuladaOutras = dValorAcumuladaOutras + tSldMesEst2.adSaldoValorOutros(iIndice)
'    Next
'
'    'Calcula CustoMedioProducaoAtual
'    If dQuantAcumuladaBenef > 0 Then dCMPBenefAtual = dValorAcumuladaBenef / dQuantAcumuladaBenef
'    If dQuantAcumuladaConserto > 0 Then dCMPConsertoAtual = dValorAcumuladaConserto / dQuantAcumuladaConserto
'    If dQuantAcumuladaConsig > 0 Then dCMPConsigAtual = dValorAcumuladaConsig / dQuantAcumuladaConsig
'    If dQuantAcumuladaDemo > 0 Then dCMPDemoAtual = dValorAcumuladaDemo / dQuantAcumuladaDemo
'    If dQuantAcumuladaOutras > 0 Then dCMPOutrasAtual = dValorAcumuladaOutras / dQuantAcumuladaOutras
'
'    Rotina_CMP_CMPTerceirosCalcula = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_CMPTerceirosCalcula:
'
'    Rotina_CMP_CMPTerceirosCalcula = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159501)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaSaidasCMP(sComandoSQL() As String, alComando() As Long, tProduto As typeProduto, dCMPAtual As Double, tTipoMovEst As typeTipoMovEst, tMovEstoque As typeItemMovEstoque, tMovEstoque2 As typeItemMovEstoque, tSldMesEst As typeSldMesEst, tSldDiaEst As typeSldDiaEst, tSldDiaEstAlm As typeSldDiaEstAlm, objAlmoxInfo As ClassAlmoxInfo, iMes As Integer, iFilialEmpresa As Integer, tSldMesEst2 As typeSldMesEst2, tSldMesEst1 As typeSldMesEst1) As Long
''Atualiza o Custo do Movimento e Acumula os valores de Saídas para os Escaninhos
'
'Dim dFator As Double
'Dim lErro As Long
'
'On Error GoTo Erro_Rotina_CMP_AtualizaSaidasCMP
'
'    'Fator de conv. de UMs
'    lErro = CF("UM_Conversao",tProduto.iClasseUM, tMovEstoque.sSiglaUM, tProduto.sSiglaUMEstoque, dFator)
'    If lErro <> SUCESSO Then gError 25268
'
'    'Calcula custo do Movimento
'    tMovEstoque.dCusto = dCMPAtual * tMovEstoque.dQuantidade * dFator
'
'    'se não for um estorno
'    If tTipoMovEst.iCodigoOrig = 0 Then
'
'        'Atualiza custo do Movimento
'        lErro = Comando_ExecutarPos(alComando(12), "UPDATE MovimentoEstoqueES SET Custo = ?", alComando(2), tMovEstoque.dCusto)
'        If lErro <> AD_SQL_SUCESSO Then gError 25269
'
'    Else
'
'        'Atualiza custo do Movimento. Torna o sinal do custo positivo já que a quantidade (tMovEstoque.dQuantidade) foi colocada com valor negativo.
'        lErro = Comando_ExecutarPos(alComando(12), "UPDATE MovimentoEstoqueES SET Custo = ?", alComando(2), -tMovEstoque.dCusto)
'        If lErro <> AD_SQL_SUCESSO Then gError 25269
'
'    End If
'
'    'Acumula
'    tSldDiaEstAlm.dValorSaida = tSldDiaEstAlm.dValorSaida + tMovEstoque.dCusto
'    tSldDiaEstAlm.dValorSaiCusto = tSldDiaEstAlm.dValorSaiCusto + tMovEstoque.dCusto
'    tSldDiaEst.dValorSaida = tSldDiaEst.dValorSaida + tMovEstoque.dCusto
'    tSldDiaEst.dValorSaiCusto = tSldDiaEst.dValorSaiCusto + tMovEstoque.dCusto
'    objAlmoxInfo.dValorSaida = objAlmoxInfo.dValorSaida + tMovEstoque.dCusto
'    tSldMesEst.adValorSai(iMes) = tSldMesEst.adValorSai(iMes) + tMovEstoque.dCusto
'    tSldMesEst.adSaldoValorCusto(iMes) = tSldMesEst.adSaldoValorCusto(iMes) - tMovEstoque.dCusto
'    objAlmoxInfo.dSaldoValorCusto = objAlmoxInfo.dSaldoValorCusto - tMovEstoque.dCusto
'
'    'os produtos nossos em poder de terceiros acumulam a quantidade nas saidas do nosso estoque para ficar em poder de terceiros
'    'pois quando for processada a saida do escaninho de mat. nosso em poder de terceiros e entrada no nosso estoque o custo da saida
'    'será feito pelo custo médio dos escaninhos.
'
'    'Acumula os escaninhos
'    If tTipoMovEst.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'
'        'Consignação
'        If tTipoMovEst.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
'            tSldMesEst2.adSaldoValorConsig(iMes) = tSldMesEst2.adSaldoValorConsig(iMes) + tMovEstoque.dCusto
'            tSldMesEst2.adSaldoQuantConsig(iMes) = tSldMesEst2.adSaldoQuantConsig(iMes) + tMovEstoque.dQuantidade * dFator
'            objAlmoxInfo.dSaldoValorConsig = objAlmoxInfo.dSaldoValorConsig + tMovEstoque.dCusto
'            tSldDiaEst.dValorEntConsig = tSldDiaEst.dValorEntConsig + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorEntConsig = tSldDiaEstAlm.dValorEntConsig + tMovEstoque.dCusto
'
'        'Demostração
'        ElseIf tTipoMovEst.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
'            tSldMesEst2.adSaldoValorDemo(iMes) = tSldMesEst2.adSaldoValorDemo(iMes) + tMovEstoque.dCusto
'            tSldMesEst2.adSaldoQuantDemo(iMes) = tSldMesEst2.adSaldoQuantDemo(iMes) + tMovEstoque.dQuantidade * dFator
'            objAlmoxInfo.dSaldoValorDemo = objAlmoxInfo.dSaldoValorDemo + tMovEstoque.dCusto
'            tSldDiaEst.dValorEntDemo = tSldDiaEst.dValorEntDemo + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorEntDemo = tSldDiaEstAlm.dValorEntDemo + tMovEstoque.dCusto
'
'        'Conserto
'        ElseIf tTipoMovEst.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
'            tSldMesEst2.adSaldoValorConserto(iMes) = tSldMesEst2.adSaldoValorConserto(iMes) + tMovEstoque.dCusto
'            tSldMesEst2.adSaldoQuantConserto(iMes) = tSldMesEst2.adSaldoQuantConserto(iMes) + tMovEstoque.dQuantidade * dFator
'            objAlmoxInfo.dSaldoValorConserto = objAlmoxInfo.dSaldoValorConserto + tMovEstoque.dCusto
'            tSldDiaEst.dValorEntConserto = tSldDiaEst.dValorEntConserto + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorEntConserto = tSldDiaEstAlm.dValorEntConserto + tMovEstoque.dCusto
'
'        'Outras
'        ElseIf tTipoMovEst.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
'            tSldMesEst2.adSaldoValorOutros(iMes) = tSldMesEst2.adSaldoValorOutros(iMes) + tMovEstoque.dCusto
'            tSldMesEst2.adSaldoQuantOutros(iMes) = tSldMesEst2.adSaldoQuantOutros(iMes) + tMovEstoque.dQuantidade * dFator
'            objAlmoxInfo.dSaldoValorOutros = objAlmoxInfo.dSaldoValorOutros + tMovEstoque.dCusto
'            tSldDiaEst.dValorEntOutros = tSldDiaEst.dValorEntOutros + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorEntOutros = tSldDiaEstAlm.dValorEntOutros + tMovEstoque.dCusto
'
'        'Beneficiamento
'        ElseIf tTipoMovEst.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
'            tSldMesEst2.adSaldoValorBenef(iMes) = tSldMesEst2.adSaldoValorBenef(iMes) + tMovEstoque.dCusto
'            tSldMesEst2.adSaldoQuantBenef(iMes) = tSldMesEst2.adSaldoQuantBenef(iMes) + tMovEstoque.dQuantidade * dFator
'            objAlmoxInfo.dSaldoValorBenef = objAlmoxInfo.dSaldoValorBenef + tMovEstoque.dCusto
'            tSldDiaEst.dValorEntBenef = tSldDiaEst.dValorEntBenef + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorEntBenef = tSldDiaEstAlm.dValorEntBenef + tMovEstoque.dCusto
'        End If
'
'    'em produtos de terceiros só está sendo tratado as subtrações pois somente as saidas de materiais
'    'de terceiros serão tratadas neste momento, tais como: Notas Fiscais de Saida de Material de Terceiros Beneficiado, Remessa de Material de Terceiros e Devolução de Material de Terceiros.
'    ElseIf tTipoMovEst.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
'
'        'Consignação
'        If tTipoMovEst.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
'            tSldMesEst1.adSaldoValorConsig3(iMes) = tSldMesEst1.adSaldoValorConsig3(iMes) - tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorConsig3 = objAlmoxInfo.dSaldoValorConsig3 - tMovEstoque.dCusto
'            tSldDiaEst.dValorSaiConsig3 = tSldDiaEst.dValorSaiConsig3 + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorSaiConsig3 = tSldDiaEstAlm.dValorSaiConsig3 + tMovEstoque.dCusto
'
'        'Demostração
'        ElseIf tTipoMovEst.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
'            tSldMesEst1.adSaldoValorDemo3(iMes) = tSldMesEst1.adSaldoValorDemo3(iMes) - tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorDemo3 = objAlmoxInfo.dSaldoValorDemo3 - tMovEstoque.dCusto
'            tSldDiaEst.dValorSaiDemo3 = tSldDiaEst.dValorSaiDemo3 + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorSaiDemo3 = tSldDiaEstAlm.dValorSaiDemo3 + tMovEstoque.dCusto
'
'        'Conserto
'        ElseIf tTipoMovEst.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
'            tSldMesEst1.adSaldoValorConserto3(iMes) = tSldMesEst1.adSaldoValorConserto3(iMes) - tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorConserto3 = objAlmoxInfo.dSaldoValorConserto3 - tMovEstoque.dCusto
'            tSldDiaEst.dValorSaiConserto3 = tSldDiaEst.dValorSaiConserto3 + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorSaiConserto3 = tSldDiaEstAlm.dValorSaiConserto3 + tMovEstoque.dCusto
'
'        'Outras
'        ElseIf tTipoMovEst.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
'            tSldMesEst1.adSaldoValorOutros3(iMes) = tSldMesEst1.adSaldoValorOutros3(iMes) - tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorOutros3 = objAlmoxInfo.dSaldoValorOutros3 - tMovEstoque.dCusto
'            tSldDiaEst.dValorSaiOutros3 = tSldDiaEst.dValorSaiOutros3 + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorSaiOutros3 = tSldDiaEstAlm.dValorSaiOutros3 + tMovEstoque.dCusto
'
'        'Beneficiamento
'        ElseIf tTipoMovEst.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
'            tSldMesEst1.adSaldoValorBenef3(iMes) = tSldMesEst1.adSaldoValorBenef3(iMes) - tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorBenef3 = objAlmoxInfo.dSaldoValorBenef3 - tMovEstoque.dCusto
'            tSldDiaEst.dValorSaiBenef3 = tSldDiaEst.dValorSaiBenef3 + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorSaiBenef3 = tSldDiaEstAlm.dValorSaiBenef3 + tMovEstoque.dCusto
'
'        End If
'
'    End If
'
'    'Acumula valores de consumo
'    If tTipoMovEst.iAtualizaConsumo = TIPOMOV_EST_ADICIONACONSUMO Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorCons = tSldDiaEstAlm.dValorCons + tMovEstoque.dCusto
'        tSldDiaEst.dValorCons = tSldDiaEst.dValorCons + tMovEstoque.dCusto
'        objAlmoxInfo.dValorCons = objAlmoxInfo.dValorCons + tMovEstoque.dCusto
'        tSldMesEst.adValorCons(iMes) = tSldMesEst.adValorCons(iMes) + tMovEstoque.dCusto
'
'    ElseIf tTipoMovEst.iAtualizaConsumo = TIPOMOV_EST_SUBTRAICONSUMO Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorCons = tSldDiaEstAlm.dValorCons - tMovEstoque.dCusto
'        tSldDiaEst.dValorCons = tSldDiaEst.dValorCons - tMovEstoque.dCusto
'        objAlmoxInfo.dValorCons = objAlmoxInfo.dValorCons - tMovEstoque.dCusto
'        tSldMesEst.adValorCons(iMes) = tSldMesEst.adValorCons(iMes) - tMovEstoque.dCusto
'
'    End If
'
'    'Acumula valores de Venda
'    If tTipoMovEst.iAtualizaVenda = TIPOMOV_EST_ADICIONAVENDA Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorVend = tSldDiaEstAlm.dValorVend + tMovEstoque.dCusto
'        tSldDiaEst.dValorVend = tSldDiaEst.dValorVend + tMovEstoque.dCusto
'        objAlmoxInfo.dValorVenda = objAlmoxInfo.dValorVenda + tMovEstoque.dCusto
'        tSldMesEst.adValorVend(iMes) = tSldMesEst.adValorVend(iMes) + tMovEstoque.dCusto
'
'    ElseIf tTipoMovEst.iAtualizaVenda = TIPOMOV_EST_SUBTRAIVENDA Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorVend = tSldDiaEstAlm.dValorVend - tMovEstoque.dCusto
'        tSldDiaEst.dValorVend = tSldDiaEst.dValorVend - tMovEstoque.dCusto
'        objAlmoxInfo.dValorVenda = objAlmoxInfo.dValorVenda - tMovEstoque.dCusto
'        tSldMesEst.adValorVend(iMes) = tSldMesEst.adValorVend(iMes) - tMovEstoque.dCusto
'
'    End If
'
'    'Busca próximo movimento
'    lErro = Rotina_CMP_ProximoMovimento(sComandoSQL(), alComando(), tProduto, tTipoMovEst, tMovEstoque, tMovEstoque2, iFilialEmpresa)
'    If lErro <> SUCESSO And lErro <> 25277 Then gError 25291
'
'    If lErro = 25277 Then gError 25270 'Não tem + movimentos
'
'    Rotina_CMP_AtualizaSaidasCMP = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaSaidasCMP:
'
'    Rotina_CMP_AtualizaSaidasCMP = gErr
'
'    Select Case gErr
'
'        Case 25268, 25291
'
'        Case 25269
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MOVESTOQUE", gErr, tMovEstoque.lNumIntDoc)
'
'        Case 25270 'Acabou os movimentos
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159502)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaEntradasCMP(sComandoSQL() As String, alComando() As Long, tProduto As typeProduto, dCMPAtual As Double, tTipoMovEst As typeTipoMovEst, tMovEstoque As typeItemMovEstoque, tMovEstoque2 As typeItemMovEstoque, tSldMesEst As typeSldMesEst, tSldDiaEst As typeSldDiaEst, tSldDiaEstAlm As typeSldDiaEstAlm, objAlmoxInfo As ClassAlmoxInfo, iMes As Integer, iFilialEmpresa As Integer, tSldMesEst2 As typeSldMesEst2, dCMPConsigAtual As Double, dCMPDemoAtual As Double, dCMPConsertoAtual As Double, dCMPOutrasAtual As Double, dCMPBenefAtual As Double, tSldMesEst1 As typeSldMesEst1) As Long
''Atualiza o Custo do Movimento e Acumula os valores de Entradas no Estoque (Saida dos Escaninhos)
'
'Dim dFator As Double
'Dim lErro As Long
'
'On Error GoTo Erro_Rotina_CMP_AtualizaEntradasCMP
'
'    'Fator de conv. de UMs
'    lErro = CF("UM_Conversao",tProduto.iClasseUM, tMovEstoque.sSiglaUM, tProduto.sSiglaUMEstoque, dFator)
'    If lErro <> SUCESSO Then gError 69764
'
'    If tTipoMovEst.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'
'        'Acumula os escaninhos
'        'Consignação
'        If tTipoMovEst.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
'            'Calcula custo do Movimento
'            tMovEstoque.dCusto = dCMPConsigAtual * tMovEstoque.dQuantidade * dFator
'            tSldMesEst2.adSaldoValorConsig(iMes) = tSldMesEst2.adSaldoValorConsig(iMes) - tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorConsig = objAlmoxInfo.dSaldoValorConsig - tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorSaiConsig = tSldDiaEstAlm.dValorSaiConsig + tMovEstoque.dCusto
'            tSldDiaEst.dValorSaiConsig = tSldDiaEst.dValorSaiConsig + tMovEstoque.dCusto
'
'        'Demonstração
'        ElseIf tTipoMovEst.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
'            tMovEstoque.dCusto = dCMPDemoAtual * tMovEstoque.dQuantidade * dFator
'            tSldMesEst2.adSaldoValorDemo(iMes) = tSldMesEst2.adSaldoValorDemo(iMes) - tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorDemo = objAlmoxInfo.dSaldoValorDemo - tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorSaiDemo = tSldDiaEstAlm.dValorSaiDemo + tMovEstoque.dCusto
'            tSldDiaEst.dValorSaiDemo = tSldDiaEst.dValorSaiDemo + tMovEstoque.dCusto
'
'        'Conserto
'        ElseIf tTipoMovEst.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
'            tMovEstoque.dCusto = dCMPConsertoAtual * tMovEstoque.dQuantidade * dFator
'            tSldMesEst2.adSaldoValorConserto(iMes) = tSldMesEst2.adSaldoValorConserto(iMes) - tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorConserto = objAlmoxInfo.dSaldoValorConserto - tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorSaiConserto = tSldDiaEstAlm.dValorSaiConserto + tMovEstoque.dCusto
'            tSldDiaEst.dValorSaiConserto = tSldDiaEst.dValorSaiConserto + tMovEstoque.dCusto
'
'        'Outras
'        ElseIf tTipoMovEst.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
'            tMovEstoque.dCusto = dCMPOutrasAtual * tMovEstoque.dQuantidade * dFator
'            tSldMesEst2.adSaldoValorOutros(iMes) = tSldMesEst2.adSaldoValorOutros(iMes) - tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorOutros = objAlmoxInfo.dSaldoValorOutros - tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorSaiOutros = tSldDiaEstAlm.dValorSaiOutros + tMovEstoque.dCusto
'            tSldDiaEst.dValorSaiOutros = tSldDiaEst.dValorSaiOutros + tMovEstoque.dCusto
'
'        'Beneficiamento
'        ElseIf tTipoMovEst.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
'            tMovEstoque.dCusto = dCMPBenefAtual * tMovEstoque.dQuantidade * dFator
'            tSldMesEst2.adSaldoValorBenef(iMes) = tSldMesEst2.adSaldoValorBenef(iMes) - tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorBenef = objAlmoxInfo.dSaldoValorBenef - tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorSaiBenef = tSldDiaEstAlm.dValorSaiBenef + tMovEstoque.dCusto
'            tSldDiaEst.dValorSaiBenef = tSldDiaEst.dValorSaiBenef + tMovEstoque.dCusto
'
'        'Calcula o Custo caso seja do tipo: Disponivel Origem Nossa Consignação
'        ElseIf tTipoMovEst.iCustoMedio = TIPOMOV_EST_CUSTOMEDIO_CONSIG Then
'            tMovEstoque.dCusto = dCMPConsigAtual * tMovEstoque.dQuantidade * dFator
'
'        'Calcula custo do Movimento caso vá para o escaninho de disponivel
'        Else
'            tMovEstoque.dCusto = dCMPAtual * tMovEstoque.dQuantidade * dFator
'        End If
'
'    ElseIf tTipoMovEst.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
'
'        'por enquanto está usando o custo médio de produção pois a alternativa seria a criação de um custo real para cada escaninho de material de terceiros (ou pelo menos para o beneficiamento)
'        tMovEstoque.dCusto = dCMPAtual * tMovEstoque.dQuantidade * dFator
'
'        'Consignação
'        If tTipoMovEst.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
'            tSldMesEst1.adSaldoValorConsig3(iMes) = tSldMesEst1.adSaldoValorConsig3(iMes) + tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorConsig3 = objAlmoxInfo.dSaldoValorConsig3 + tMovEstoque.dCusto
'            tSldDiaEst.dValorEntConsig3 = tSldDiaEst.dValorEntConsig3 + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorEntConsig3 = tSldDiaEstAlm.dValorEntConsig3 + tMovEstoque.dCusto
'
'        'Demostração
'        ElseIf tTipoMovEst.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
'            tSldMesEst1.adSaldoValorDemo3(iMes) = tSldMesEst1.adSaldoValorDemo3(iMes) + tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorDemo3 = objAlmoxInfo.dSaldoValorDemo3 + tMovEstoque.dCusto
'            tSldDiaEst.dValorEntDemo3 = tSldDiaEst.dValorEntDemo3 + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorEntDemo3 = tSldDiaEstAlm.dValorEntDemo3 + tMovEstoque.dCusto
'
'        'Conserto
'        ElseIf tTipoMovEst.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
'            tSldMesEst1.adSaldoValorConserto3(iMes) = tSldMesEst1.adSaldoValorConserto3(iMes) + tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorConserto3 = objAlmoxInfo.dSaldoValorConserto3 + tMovEstoque.dCusto
'            tSldDiaEst.dValorEntConserto3 = tSldDiaEst.dValorEntConserto3 + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorEntConserto3 = tSldDiaEstAlm.dValorEntConserto3 + tMovEstoque.dCusto
'
'        'Outras
'        ElseIf tTipoMovEst.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
'            tSldMesEst1.adSaldoValorOutros3(iMes) = tSldMesEst1.adSaldoValorOutros3(iMes) + tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorOutros3 = objAlmoxInfo.dSaldoValorOutros3 + tMovEstoque.dCusto
'            tSldDiaEst.dValorEntOutros3 = tSldDiaEst.dValorEntOutros3 + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorEntOutros3 = tSldDiaEstAlm.dValorEntOutros3 + tMovEstoque.dCusto
'
'        'Beneficiamento
'        ElseIf tTipoMovEst.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
'            tSldMesEst1.adSaldoValorBenef3(iMes) = tSldMesEst1.adSaldoValorBenef3(iMes) + tMovEstoque.dCusto
'            objAlmoxInfo.dSaldoValorBenef3 = objAlmoxInfo.dSaldoValorBenef3 + tMovEstoque.dCusto
'            tSldDiaEst.dValorEntBenef3 = tSldDiaEst.dValorEntBenef3 + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorEntBenef3 = tSldDiaEstAlm.dValorEntBenef3 + tMovEstoque.dCusto
'
'        End If
'
'    End If
'
'    If tTipoMovEst.iCustoMedio <> TIPOMOV_EST_CUSTOMEDIO_CONSIG Then
'        'Acumula independente do escaninho
'        tSldDiaEstAlm.dValorEntrada = tSldDiaEstAlm.dValorEntrada + tMovEstoque.dCusto
'        tSldDiaEstAlm.dValorEntCusto = tSldDiaEstAlm.dValorEntCusto + tMovEstoque.dCusto
'        tSldDiaEst.dValorEntrada = tSldDiaEst.dValorEntrada + tMovEstoque.dCusto
'        tSldDiaEst.dValorEntCusto = tSldDiaEst.dValorEntCusto + tMovEstoque.dCusto
'        objAlmoxInfo.dValorEntrada = objAlmoxInfo.dValorEntrada + tMovEstoque.dCusto
'        tSldMesEst.adValorEnt(iMes) = tSldMesEst.adValorEnt(iMes) + tMovEstoque.dCusto
'        tSldMesEst.adSaldoValorCusto(iMes) = tSldMesEst.adSaldoValorCusto(iMes) + tMovEstoque.dCusto
'        objAlmoxInfo.dSaldoValorCusto = objAlmoxInfo.dSaldoValorCusto + tMovEstoque.dCusto
'    End If
'
'    'se não for um estorno
'    If tTipoMovEst.iCodigoOrig = 0 Then
'
'        'Atualiza custo do Movimento
'        lErro = Comando_ExecutarPos(alComando(12), "UPDATE MovimentoEstoqueES SET Custo = ?", alComando(2), tMovEstoque.dCusto)
'        If lErro <> AD_SQL_SUCESSO Then gError 69765
'
'    Else
'
'        'Atualiza custo do Movimento. Torna o sinal do custo positivo já que a quantidade (tMovEstoque.dQuantidade) foi colocada com valor negativo.
'        lErro = Comando_ExecutarPos(alComando(12), "UPDATE MovimentoEstoqueES SET Custo = ?", alComando(2), -tMovEstoque.dCusto)
'        If lErro <> AD_SQL_SUCESSO Then gError 69765
'
'    End If
'
'    'Acumula valores de consumo
'    If tTipoMovEst.iAtualizaConsumo = TIPOMOV_EST_ADICIONACONSUMO Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorCons = tSldDiaEstAlm.dValorCons + tMovEstoque.dCusto
'        tSldDiaEst.dValorCons = tSldDiaEst.dValorCons + tMovEstoque.dCusto
'        objAlmoxInfo.dValorCons = objAlmoxInfo.dValorCons + tMovEstoque.dCusto
'        tSldMesEst.adValorCons(iMes) = tSldMesEst.adValorCons(iMes) + tMovEstoque.dCusto
'
'    ElseIf tTipoMovEst.iAtualizaConsumo = TIPOMOV_EST_SUBTRAICONSUMO Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorCons = tSldDiaEstAlm.dValorCons - tMovEstoque.dCusto
'        tSldDiaEst.dValorCons = tSldDiaEst.dValorCons - tMovEstoque.dCusto
'        objAlmoxInfo.dValorCons = objAlmoxInfo.dValorCons - tMovEstoque.dCusto
'        tSldMesEst.adValorCons(iMes) = tSldMesEst.adValorCons(iMes) - tMovEstoque.dCusto
'
'    End If
'
'    'Acumula valores de Venda
'    If tTipoMovEst.iAtualizaVenda = TIPOMOV_EST_ADICIONAVENDA Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorVend = tSldDiaEstAlm.dValorVend + tMovEstoque.dCusto
'        tSldDiaEst.dValorVend = tSldDiaEst.dValorVend + tMovEstoque.dCusto
'        objAlmoxInfo.dValorVenda = objAlmoxInfo.dValorVenda + tMovEstoque.dCusto
'        tSldMesEst.adValorVend(iMes) = tSldMesEst.adValorVend(iMes) + tMovEstoque.dCusto
'
'    ElseIf tTipoMovEst.iAtualizaVenda = TIPOMOV_EST_SUBTRAIVENDA Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorVend = tSldDiaEstAlm.dValorVend - tMovEstoque.dCusto
'        tSldDiaEst.dValorVend = tSldDiaEst.dValorVend - tMovEstoque.dCusto
'        objAlmoxInfo.dValorVenda = objAlmoxInfo.dValorVenda - tMovEstoque.dCusto
'        tSldMesEst.adValorVend(iMes) = tSldMesEst.adValorVend(iMes) - tMovEstoque.dCusto
'
'    End If
'
'    'Busca próximo movimento
'    lErro = Rotina_CMP_ProximoMovimento(sComandoSQL(), alComando(), tProduto, tTipoMovEst, tMovEstoque, tMovEstoque2, iFilialEmpresa)
'    If lErro <> SUCESSO And lErro <> 25277 Then gError 69766
'
'    If lErro = 25277 Then gError 69767 'Não tem + movimentos
'
'    Rotina_CMP_AtualizaEntradasCMP = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaEntradasCMP:
'
'    Rotina_CMP_AtualizaEntradasCMP = gErr
'
'    Select Case gErr
'
'        Case 69764, 69766
'
'        Case 69765
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MOVESTOQUE", gErr, tMovEstoque.lNumIntDoc)
'
'        Case 69767 'Acabou os movimentos
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159503)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Private Function Rotina_CMP_ProximoSldDiaEst(alComando() As Long, iApropriacao As Integer, tMovEstoque As typeItemMovEstoque, tSldDiaEst As typeSldDiaEst, iFilialEmpresa As Integer) As Long
'''Chamada EM TRANSAÇÃO
'''Busca próximo registro de SldDiaEst
''
''Dim lErro As Long
''
''On Error GoTo Erro_Rotina_CMP_ProximoSldDiaEst
''
''    'Mudou apenas a Data do MovEstoque
''    If tSldDiaEst.sProduto = tMovEstoque.sProduto And tMovEstoque.iApropriacao = iApropriacao Then
''
''        'Iguala Produto e Data em SaldoDiaEst ao de MovEstoque
''        Do While tSldDiaEst.sProduto = tMovEstoque.sProduto And tSldDiaEst.dtData < tMovEstoque.dtData
''            lErro = Comando_BuscarProximo(alComando(4))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25306
''            If lErro = AD_SQL_SEM_DADOS Then gError 25307
''        Loop
''
''        'Se Produto ficou diferente ou Data ultrapassou a de Movimento, erro
''        If tSldDiaEst.sProduto <> tMovEstoque.sProduto Or tSldDiaEst.dtData > tMovEstoque.dtData Then gError 25308
''
''    'Mudou a apropriação do MovEstoque mas não o Produto
''    ElseIf tSldDiaEst.sProduto = tMovEstoque.sProduto And tMovEstoque.iApropriacao <> iApropriacao Then
''
''        'Iguala Produto e Data em SaldoDiaEst ao de MovEstoque
''        If tSldDiaEst.dtData > tMovEstoque.dtData Then
''
''            Do While tSldDiaEst.sProduto = tMovEstoque.sProduto And tSldDiaEst.dtData > tMovEstoque.dtData
''                lErro = Comando_BuscarAnterior(alComando(4))
''                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25309
''                If lErro = AD_SQL_SEM_DADOS Then gError 25311
''            Loop
''
''            'Se Produto ficou diferente ou Data ficou menor que a de Movimento, erro
''            If tSldDiaEst.sProduto <> tMovEstoque.sProduto Or tSldDiaEst.dtData < tMovEstoque.dtData Then gError 25313
''
''        ElseIf tSldDiaEst.dtData < tMovEstoque.dtData Then
''
''            Do While tSldDiaEst.sProduto = tMovEstoque.sProduto And tSldDiaEst.dtData < tMovEstoque.dtData
''                lErro = Comando_BuscarProximo(alComando(4))
''                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25310
''                If lErro = AD_SQL_SEM_DADOS Then gError 25312
''            Loop
''
''            'Se Produto ficou diferente ou Data ficou menor que a de Movimento, erro
''            If tSldDiaEst.sProduto <> tMovEstoque.sProduto Or tSldDiaEst.dtData > tMovEstoque.dtData Then gError 25314
''
''        End If
''
''    'Mudou o Produto do MovEstoque (necessariamente MAIOR)
''    ElseIf tSldDiaEst.sProduto <> tMovEstoque.sProduto Then
''
''        'Iguala Produto e Data em SaldoDiaEst ao de MovEstoque
''        Do While tSldDiaEst.sProduto < tMovEstoque.sProduto Or (tSldDiaEst.sProduto = tMovEstoque.sProduto And tSldDiaEst.dtData < tMovEstoque.dtData)
''            lErro = Comando_BuscarProximo(alComando(4))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25315
''            If lErro = AD_SQL_SEM_DADOS Then gError 25316
''        Loop
''
''        'Se Produto ficou diferente ou Data ultrapassou a de Movimento, erro
''        If tSldDiaEst.sProduto <> tMovEstoque.sProduto Or tSldDiaEst.dtData > tMovEstoque.dtData Then gError 25317
''
''    End If
''
''    Rotina_CMP_ProximoSldDiaEst = SUCESSO
''
''    Exit Function
''
''Erro_Rotina_CMP_ProximoSldDiaEst:
''
''    Rotina_CMP_ProximoSldDiaEst = gErr
''
''    Select Case gErr
''
''        Case 25306, 25309, 25310, 25315
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST2", gErr, iFilialEmpresa)
''
''        Case 25307, 25308, 25311, 25312, 25313, 25314, 25316, 25317
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SLDDIAEST", gErr, iFilialEmpresa, tMovEstoque.sProduto, tMovEstoque.dtData)
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159504)
''
''    End Select
''
''    Exit Function
''
''End Function
'
''Private Function Rotina_CMP_ProximoSldDiaEstAlm(alComando() As Long, iApropriacao As Integer, tMovEstoque As typeItemMovEstoque, tSldDiaEstAlm As typeSldDiaEstAlm) As Long
'''Chamada EM TRANSAÇÃO
'''Busca próximo registro de SldDiaEstAlm
''
''Dim lErro As Long
''
''On Error GoTo Erro_Rotina_CMP_ProximoSldDiaEstAlm
''
''    'Mudou apenas o Almoxarifado do MovEstoque
''    If tMovEstoque.sProduto = tSldDiaEstAlm.sProduto And tMovEstoque.iApropriacao = iApropriacao And tMovEstoque.dtData = tSldDiaEstAlm.dtData Then
''
''        'Iguala Produto, Data, Almoxarifado em SaldoDiaEstAlm ao de MovEstoque
''        Do While tSldDiaEstAlm.sProduto = tMovEstoque.sProduto And tSldDiaEstAlm.dtData = tMovEstoque.dtData And tSldDiaEstAlm.iAlmoxarifado < tMovEstoque.iAlmoxarifado
''            lErro = Comando_BuscarProximo(alComando(5))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25896
''            If lErro = AD_SQL_SEM_DADOS Then gError 25898
''        Loop
''
''        'Se Produto ou Data ou Almoxarifado ficaram diferentes dos de Movimento, erro
''        If tSldDiaEstAlm.sProduto <> tMovEstoque.sProduto Or tSldDiaEstAlm.dtData <> tMovEstoque.dtData Or tSldDiaEstAlm.iAlmoxarifado <> tMovEstoque.iAlmoxarifado Then gError 25900
''
''    'Mudou a Data de MovEstoque mas não o Produto nem a Apropriação
''    ElseIf tMovEstoque.sProduto = tSldDiaEstAlm.sProduto And tMovEstoque.iApropriacao = iApropriacao Then
''
''        'Iguala Produto, Data, Almoxarifado em SaldoDiaEstAlm ao de MovEstoque
''        Do While tSldDiaEstAlm.sProduto = tMovEstoque.sProduto And (tSldDiaEstAlm.dtData < tMovEstoque.dtData Or (tSldDiaEstAlm.dtData = tMovEstoque.dtData And tSldDiaEstAlm.iAlmoxarifado < tMovEstoque.iAlmoxarifado))
''            lErro = Comando_BuscarProximo(alComando(5))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25897
''            If lErro = AD_SQL_SEM_DADOS Then gError 25899
''        Loop
''
''        'Se Produto, Data ou Almoxarifado ficaram diferentes dos de Movimento, erro
''        If tSldDiaEstAlm.sProduto <> tMovEstoque.sProduto Or tSldDiaEstAlm.dtData <> tMovEstoque.dtData Or tSldDiaEstAlm.iAlmoxarifado <> tMovEstoque.iAlmoxarifado Then gError 25901
''
''    'Mudou a apropriação do MovEstoque mas não o Produto
''    ElseIf tMovEstoque.sProduto = tSldDiaEstAlm.sProduto Then
''
''        'Iguala Produto, Data e Almoxarifado em SaldoDiaEstAlm ao de MovEstoque
''        If tSldDiaEstAlm.dtData > tMovEstoque.dtData Or (tSldDiaEstAlm.dtData = tMovEstoque.dtData And tSldDiaEstAlm.iAlmoxarifado > tMovEstoque.iAlmoxarifado) Then
''
''            Do While tSldDiaEstAlm.sProduto = tMovEstoque.sProduto And (tSldDiaEstAlm.dtData > tMovEstoque.dtData Or (tSldDiaEstAlm.dtData = tMovEstoque.dtData And tSldDiaEstAlm.iAlmoxarifado > tMovEstoque.iAlmoxarifado))
''                lErro = Comando_BuscarAnterior(alComando(5))
''                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25902
''                If lErro = AD_SQL_SEM_DADOS Then gError 25903
''            Loop
''
''            'Se Produto ou Data ou Almoxarifado ficaram diferentes dos de Movimento, erro
''            If tSldDiaEstAlm.sProduto <> tMovEstoque.sProduto Or tSldDiaEstAlm.dtData <> tMovEstoque.dtData Or tSldDiaEstAlm.iAlmoxarifado <> tMovEstoque.iAlmoxarifado Then gError 25904
''
''        ElseIf tSldDiaEstAlm.dtData < tMovEstoque.dtData Or (tSldDiaEstAlm.dtData = tMovEstoque.dtData And tSldDiaEstAlm.iAlmoxarifado < tMovEstoque.iAlmoxarifado) Then
''
''            Do While tSldDiaEstAlm.sProduto = tMovEstoque.sProduto And (tSldDiaEstAlm.dtData < tMovEstoque.dtData Or (tSldDiaEstAlm.dtData = tMovEstoque.dtData And tSldDiaEstAlm.iAlmoxarifado < tMovEstoque.iAlmoxarifado))
''                lErro = Comando_BuscarProximo(alComando(5))
''                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25905
''                If lErro = AD_SQL_SEM_DADOS Then gError 25906
''            Loop
''
''            'Se Produto ou Data ou Almoxarifado ficaram diferentes dos de Movimento, erro
''            If tSldDiaEstAlm.sProduto <> tMovEstoque.sProduto Or tSldDiaEstAlm.dtData <> tMovEstoque.dtData Or tSldDiaEstAlm.iAlmoxarifado <> tMovEstoque.iAlmoxarifado Then gError 25907
''
''        End If
''
''    'Mudou o Produto do MovEstoque (necessariamente MAIOR)
''    ElseIf tSldDiaEstAlm.sProduto <> tMovEstoque.sProduto Then
''
''        'Iguala Produto e Data em SaldoDiaEstAlm ao de MovEstoque
''        Do While tSldDiaEstAlm.sProduto < tMovEstoque.sProduto Or (tSldDiaEstAlm.sProduto = tMovEstoque.sProduto And tSldDiaEstAlm.dtData < tMovEstoque.dtData) Or (tSldDiaEstAlm.sProduto = tMovEstoque.sProduto And tSldDiaEstAlm.dtData = tMovEstoque.dtData And tSldDiaEstAlm.iAlmoxarifado < tMovEstoque.iAlmoxarifado)
''            lErro = Comando_BuscarProximo(alComando(5))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25908
''            If lErro = AD_SQL_SEM_DADOS Then gError 25909
''        Loop
''
''        'Se Produto ou Data ou Almoxarifado ficaram diferentes dos de Movimento, erro
''        If tSldDiaEstAlm.sProduto <> tMovEstoque.sProduto Or tSldDiaEstAlm.dtData <> tMovEstoque.dtData Or tSldDiaEstAlm.iAlmoxarifado <> tMovEstoque.iAlmoxarifado Then gError 25910
''
''    End If
''
''    Rotina_CMP_ProximoSldDiaEstAlm = SUCESSO
''
''    Exit Function
''
''Erro_Rotina_CMP_ProximoSldDiaEstAlm:
''
''    Rotina_CMP_ProximoSldDiaEstAlm = gErr
''
''    Select Case gErr
''
''        Case 25896, 25897, 25902, 25905, 25908
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM2", gErr)
''
''        Case 25898, 25899, 25900, 25901, 25903, 25906, 25904, 25907, 25909, 25910
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SLDDIAESTALM", gErr, tMovEstoque.sProduto, tMovEstoque.dtData, tMovEstoque.iAlmoxarifado)
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159505)
''
''    End Select
''
''    Exit Function
''
''End Function
'
'Private Function Rotina_CMP_ProximoSldMesEst(alComando() As Long, tSldMesEst As typeSldMesEst, tMovEstoque As typeItemMovEstoque, iFilialEmpresa As Integer, iAno As Integer) As Long
''Chamada EM TRANSAÇÃO
''Busca próximo registro de SldMesEst
'
'Dim lErro As Long
'
'On Error GoTo Erro_Rotina_CMP_ProximoSldMesEst
'
'    'Iguala produto em SaldoMesEst ao de MovEstoque
'    Do While tSldMesEst.sProduto < tMovEstoque.sProduto
'        lErro = Comando_BuscarProximo(alComando(3))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25318
'        If lErro = AD_SQL_SEM_DADOS Then gError 25319
'    Loop
'
'    If tSldMesEst.sProduto > tMovEstoque.sProduto Then gError 25320
'
'    Rotina_CMP_ProximoSldMesEst = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_ProximoSldMesEst:
'
'    Rotina_CMP_ProximoSldMesEst = gErr
'
'    Select Case gErr
'
'        Case 25318
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, iFilialEmpresa, iAno)
'
'        Case 25319, 25320
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SALDOMESEST", gErr, iFilialEmpresa, iAno, tMovEstoque.sProduto)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159506)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_ProximoSldMesEst2(alComando() As Long, tSldMesEst2 As typeSldMesEst2, tMovEstoque As typeItemMovEstoque, iFilialEmpresa As Integer, iAno As Integer) As Long
''Chamada EM TRANSAÇÃO
''Busca próximo registro de SldMesEst
'
'Dim lErro As Long
'
'On Error GoTo Erro_Rotina_CMP_ProximoSldMesEst2
'
'    'Iguala produto em SaldoMesEst ao de MovEstoque
'    Do While tSldMesEst2.sProduto < tMovEstoque.sProduto
'        lErro = Comando_BuscarProximo(alComando(17))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69731
'        If lErro = AD_SQL_SEM_DADOS Then gError 69732
'    Loop
'
'    If tSldMesEst2.sProduto > tMovEstoque.sProduto Then gError 69733
'
'    Rotina_CMP_ProximoSldMesEst2 = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_ProximoSldMesEst2:
'
'    Rotina_CMP_ProximoSldMesEst2 = gErr
'
'    Select Case gErr
'
'        Case 69731
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST2", gErr, iFilialEmpresa, iAno)
'
'        Case 69732, 69733
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SALDOMESEST", gErr, iFilialEmpresa, iAno, tMovEstoque.sProduto)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159507)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaSldDiaEst(alComando() As Long, tSldDiaEst As typeSldDiaEst, iFilialEmpresa As Integer) As Long
''Atualiza Valores Entrada e Saida na tabela SaldoDiaEst
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim sComandoSQL(1 To 2) As String
'Dim sCodProduto As String
'
'On Error GoTo Erro_Rotina_CMP_AtualizaSldDiaEst
'
'    'SELECT em SldDiaEstAlm filtrando pela chave primária
'    sComandoSQL(1) = "SELECT Produto FROM SldDiaEst WHERE FilialEmpresa = ? AND Produto = ? AND Data = ?"
'
'    'Monta comando SQL para UPDATE de SldDiaEst
'    sComandoSQL(2) = "UPDATE SldDiaEst SET ValorEntrada = ValorEntrada + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaida = ValorSaida + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorVend = ValorVend + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorCons = ValorCons + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntCusto = ValorEntCusto + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiCusto = ValorSaiCusto + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntConsig = ValorEntConsig + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiConsig = ValorSaiConsig + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntDemo = ValorEntDemo + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiDemo = ValorSaiDemo + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntConserto = ValorEntConserto + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiConserto = ValorSaiConserto + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntOutros = ValorEntOutros + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiOutros = ValorSaiOutros + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntBenef = ValorEntBenef + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiBenef = ValorSaiBenef + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntConsig3 = ValorEntConsig3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiConsig3 = ValorSaiConsig3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntDemo3 = ValorEntDemo3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiDemo3 = ValorSaiDemo3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntConserto3 = ValorEntConserto3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiConserto3 = ValorSaiConserto3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntOutros3 = ValorEntOutros3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiOutros3 = ValorSaiOutros3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntBenef3 = ValorEntBenef3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiBenef3 = ValorSaiBenef3 + ?"
'
'    'Inicializa comando da tabela SaldoDiaEstAlm para atualizar Valores de Entrada, Saída, Consumo
'    sCodProduto = String(STRING_PRODUTO, 0)
'
'    lErro = Comando_ExecutarPos(alComando(4), sComandoSQL(1), 0, sCodProduto, iFilialEmpresa, tSldDiaEst.sProduto, tSldDiaEst.dtData)
'    If lErro <> AD_SQL_SUCESSO Then gError 71954
'
'    lErro = Comando_BuscarPrimeiro(alComando(4))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71955
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 71956
'
'    'Atualiza tabela SaldoDiaEst, campos ValEntrada e ValSaida
'    lErro = Comando_ExecutarPos(alComando(7), sComandoSQL(2), alComando(4), tSldDiaEst.dValorEntrada, tSldDiaEst.dValorSaida, tSldDiaEst.dValorVend, tSldDiaEst.dValorCons, tSldDiaEst.dValorEntCusto, tSldDiaEst.dValorSaiCusto, tSldDiaEst.dValorEntConsig, tSldDiaEst.dValorSaiConsig, tSldDiaEst.dValorEntDemo, tSldDiaEst.dValorSaiDemo, tSldDiaEst.dValorEntConserto, tSldDiaEst.dValorSaiConserto, tSldDiaEst.dValorEntOutros, tSldDiaEst.dValorSaiOutros, tSldDiaEst.dValorEntBenef, tSldDiaEst.dValorSaiBenef, tSldDiaEst.dValorEntConsig3, tSldDiaEst.dValorSaiConsig3, tSldDiaEst.dValorEntDemo3, tSldDiaEst.dValorSaiDemo3, tSldDiaEst.dValorEntConserto3, tSldDiaEst.dValorSaiConserto3, tSldDiaEst.dValorEntOutros3, tSldDiaEst.dValorSaiOutros3, tSldDiaEst.dValorEntBenef3, tSldDiaEst.dValorSaiBenef3)
'    If lErro <> AD_SQL_SUCESSO Then gError 25321
'
'    Rotina_CMP_AtualizaSldDiaEst = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaSldDiaEst:
'
'    Rotina_CMP_AtualizaSldDiaEst = gErr
'
'    Select Case gErr
'
'         Case 25321
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDDIAEST", gErr, iFilialEmpresa, tSldDiaEst.sProduto, tSldDiaEst.dtData)
'
'        Case 71954, 71955
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST", gErr, iFilialEmpresa, tSldDiaEst.sProduto, tSldDiaEst.dtData)
'
'        Case 71956
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SLDDIAEST", gErr, iFilialEmpresa, tSldDiaEst.sProduto, tSldDiaEst.dtData)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159508)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Private Function Rotina_CMP_AtualizaSldDiaEstAlm(alComando() As Long, tSldDiaEstAlm As typeSldDiaEstAlm) As Long
'''Atualiza Valores Entrada e Saida na tabela SaldoDiaEst
'''Chamada EM TRANSAÇÃO
''
''Dim lErro As Long
''Dim sComandoSQL As String
''
''On Error GoTo Erro_Rotina_CMP_AtualizaSldDiaEstAlm
''
''    'Monta comando SQL para UPDATE de SldDiaEstAlm
''    sComandoSQL = "UPDATE SldDiaEstAlm SET ValorEntrada = ValorEntrada + ?, "
''    sComandoSQL = sComandoSQL & "ValorSaida = ValorSaida + ?, "
''    sComandoSQL = sComandoSQL & "ValorVend = ValorVend + ?, "
''    sComandoSQL = sComandoSQL & "ValorCons = ValorCons + ?"
''
''    'Atualiza tabela SaldoDiaEstAlm, campos ValorEntrada, ValorSaida, ValorCons
''    lErro = Comando_ExecutarPos(alComando(11), sComandoSQL, alComando(5), tSldDiaEstAlm.dValorEntrada, tSldDiaEstAlm.dValorSaida, tSldDiaEstAlm.dValorVend, tSldDiaEstAlm.dValorCons)
''    If lErro <> AD_SQL_SUCESSO Then gError 25892
''
''    Rotina_CMP_AtualizaSldDiaEstAlm = SUCESSO
''
''    Exit Function
''
''Erro_Rotina_CMP_AtualizaSldDiaEstAlm:
''
''    Rotina_CMP_AtualizaSldDiaEstAlm = gErr
''
''    Select Case gErr
''
''         Case 25892
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDDIAESTALM", gErr, tSldDiaEstAlm.iAlmoxarifado, tSldDiaEstAlm.sProduto, tSldDiaEstAlm.dtData)
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159509)
''
''    End Select
''
''    Exit Function
''
''End Function
'
'Private Function Rotina_CMP_AtualizaSldDiaEstAlm(alComando() As Long, tSldDiaEstAlm As typeSldDiaEstAlm) As Long
''Atualiza Valores Entrada e Saida na tabela SaldoDiaEst
''Chamada EM TRANSAÇÃO
'
''Parei aqui.
'
'Dim lErro As Long
'Dim sComandoSQL(1 To 2) As String
'Dim sCodProduto As String
'
'On Error GoTo Erro_Rotina_CMP_AtualizaSldDiaEstAlm
'
'    'SELECT em SldDiaEstAlm filtrando pela chave primária
'    sComandoSQL(1) = "SELECT Produto FROM SldDiaEstAlm WHERE Almoxarifado = ? AND Produto = ? AND Data = ?"
'
'    'Monta comando SQL para UPDATE de SldDiaEstAlm
'    sComandoSQL(2) = "UPDATE SldDiaEstAlm SET ValorEntrada = ValorEntrada + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaida = ValorSaida + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorVend = ValorVend + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorCons = ValorCons + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntCusto = ValorEntCusto + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiCusto = ValorSaiCusto + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntConsig = ValorEntConsig + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiConsig = ValorSaiConsig + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntDemo = ValorEntDemo + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiDemo = ValorSaiDemo + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntConserto = ValorEntConserto + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiConserto = ValorSaiConserto + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntOutros = ValorEntOutros + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiOutros = ValorSaiOutros + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntBenef = ValorEntBenef + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiBenef = ValorSaiBenef + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntConsig3 = ValorEntConsig3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiConsig3 = ValorSaiConsig3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntDemo3 = ValorEntDemo3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiDemo3 = ValorSaiDemo3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntConserto3 = ValorEntConserto3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiConserto3 = ValorSaiConserto3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntOutros3 = ValorEntOutros3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiOutros3 = ValorSaiOutros3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorEntBenef3 = ValorEntBenef3 + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSaiBenef3 = ValorSaiBenef3 + ?"
'
'    'Inicializa comando da tabela SaldoDiaEstAlm para atualizar Valores de Entrada, Saída, Consumo
'    sCodProduto = String(STRING_PRODUTO, 0)
'
'    lErro = Comando_ExecutarPos(alComando(5), sComandoSQL(1), 0, sCodProduto, tSldDiaEstAlm.iAlmoxarifado, tSldDiaEstAlm.sProduto, tSldDiaEstAlm.dtData)
'    If lErro <> AD_SQL_SUCESSO Then gError 71950
'
'    lErro = Comando_BuscarPrimeiro(alComando(5))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71951
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 71952
'
'    'Atualiza tabela SaldoDiaEstAlm, campos ValorEntrada, ValorSaida, ValorVend, ValorCons
'    lErro = Comando_ExecutarPos(alComando(11), sComandoSQL(2), alComando(5), tSldDiaEstAlm.dValorEntrada, tSldDiaEstAlm.dValorSaida, tSldDiaEstAlm.dValorVend, tSldDiaEstAlm.dValorCons, tSldDiaEstAlm.dValorEntCusto, tSldDiaEstAlm.dValorSaiCusto, tSldDiaEstAlm.dValorEntConsig, tSldDiaEstAlm.dValorSaiConsig, tSldDiaEstAlm.dValorEntDemo, tSldDiaEstAlm.dValorSaiDemo, tSldDiaEstAlm.dValorEntConserto, tSldDiaEstAlm.dValorSaiConserto, tSldDiaEstAlm.dValorEntOutros, tSldDiaEstAlm.dValorSaiOutros, tSldDiaEstAlm.dValorEntBenef, tSldDiaEstAlm.dValorSaiBenef, tSldDiaEstAlm.dValorEntConsig3, tSldDiaEstAlm.dValorSaiConsig3, tSldDiaEstAlm.dValorEntDemo3, tSldDiaEstAlm.dValorSaiDemo3, tSldDiaEstAlm.dValorEntConserto3, tSldDiaEstAlm.dValorSaiConserto3, tSldDiaEstAlm.dValorEntOutros3, tSldDiaEstAlm.dValorSaiOutros3, tSldDiaEstAlm.dValorEntBenef3, tSldDiaEstAlm.dValorSaiBenef3)
'    If lErro <> AD_SQL_SUCESSO Then gError 71953
'
'    Rotina_CMP_AtualizaSldDiaEstAlm = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaSldDiaEstAlm:
'
'    Rotina_CMP_AtualizaSldDiaEstAlm = gErr
'
'    Select Case gErr
'
'        Case 71950, 71951
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM", gErr, tSldDiaEstAlm.iAlmoxarifado, tSldDiaEstAlm.sProduto, tSldDiaEstAlm.dtData)
'
'        Case 71952
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SLDDIAESTALM", gErr, tSldDiaEstAlm.iAlmoxarifado, tSldDiaEstAlm.sProduto, tSldDiaEstAlm.dtData)
'
'        Case 71953
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDDIAESTALM", gErr, tSldDiaEstAlm.iAlmoxarifado, tSldDiaEstAlm.sProduto, tSldDiaEstAlm.dtData)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159510)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_CMPAtualCalcula(alComando() As Long, tTipoMovEst As typeTipoMovEst, tSldMesEst As typeSldMesEst, tSldMesEst1 As typeSldMesEst1, dQuantEntApropCustoProd As Double, iMes As Integer, iAno As Integer, dCMPAtual As Double, dSaldoValorCustoInformado As Double, iFilialEmpresa As Integer) As Long
''Calcula Custo Médio de Produção Atual
'' *** ALTERADA POR LUIZ G.F.NOGUEIRA EM 18/07/2001 ***
'' A alteração foi feita para que, quando o movimento for de material de terceiros
'' essa função chame outra função que irá calcular o custo médio de produção do produto
'
'Dim dQuantAcumulada As Double
'Dim dValorAcumulado As Double
'Dim iMesFinal As Integer
'Dim iIndice As Integer
'Dim dQuantInformadaMesAtual As Double
'Dim lErro As Long
'
'On Error GoTo Erro_Rotina_CMP_CMPAtualCalcula
'
'    ' *** ALTERADA POR LUIZ G.F.NOGUEIRA EM 18/07/2001 ***
'    'Se o movimento é referente a produto de terceiros
'    If tTipoMovEst.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
'
'        'Calcula o Custo Médio para o Escaninho e atualiza a tabela SldMesEst1
'        lErro = Rotina_CMP_AtualizaSldMesEst1(alComando(), tSldMesEst1, iMes, iAno, iFilialEmpresa, dCMPAtual, tTipoMovEst)
'        If lErro <> SUCESSO Then gError 90565
'
'    'Senão
'    Else
'
'        'Acumula quantidade inicial
'        dQuantAcumulada = tSldMesEst.dQuantInicialCusto
'
'        'Adiciona saldos dos meses anteriores
'        If iMes > 1 Then
'            For iIndice = 1 To iMes - 1
'                dQuantAcumulada = dQuantAcumulada + tSldMesEst.adSaldoQuantCusto(iIndice)
'            Next
'        End If
'
'        If tSldMesEst.adCustoMedio(iMes) <> 0 Then
'            dQuantInformadaMesAtual = dSaldoValorCustoInformado / tSldMesEst.adCustoMedio(iMes)
'        End If
'
'        'Adiciona QuantEntrada de Produção do mês atual
'        dQuantAcumulada = dQuantAcumulada + dQuantEntApropCustoProd + dQuantInformadaMesAtual
'
'        'Acumula valor inicial
'        dValorAcumulado = tSldMesEst.dValorInicialCusto + dSaldoValorCustoInformado
'
'        'Adiciona saldos nos Meses
'        For iIndice = 1 To iMes
'            dValorAcumulado = dValorAcumulado + tSldMesEst.adSaldoValorCusto(iIndice)
'        Next
'
'        If dQuantAcumulada <> 0 And dValorAcumulado <> 0 Then
'
'            'Calcula CustoMedioProducaoAtual
'            dCMPAtual = dValorAcumulado / dQuantAcumulada
'
'        ElseIf tSldMesEst.adCustoStandard(iMes) <> 0 Then
'
'            dCMPAtual = tSldMesEst.adCustoStandard(iMes)
'
'        Else
'            gError 83008
'        End If
'
'    End If
'
'    Rotina_CMP_CMPAtualCalcula = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_CMPAtualCalcula:
'
'    Rotina_CMP_CMPAtualCalcula = gErr
'
'    Select Case gErr
'
'        Case 90565
'
'        Case 83008
'            Call Rotina_Erro(vbOKOnly, "ERRO_CMP_E_CST_ZERADOS", gErr, tSldMesEst.sProduto, iMes)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159511)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaSldMesEst(alComando() As Long, tSldMesEst As typeSldMesEst, iMes As Integer, iAno As Integer, iFilialEmpresa As Integer) As Long
''Atualiza ValorEnt, ValorSai, ValorCons do iMes na tabela SaldoMesEst
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim sComandoSQL As String
'
'On Error GoTo Erro_Rotina_CMP_AtualizaSldMesEst
''Acrescenta no sComandoSQL código de atualizacao de dSaldoValorCusto
'
'    'Monta comando SQL para UPDATE de SldMesEst
'    sComandoSQL = "UPDATE SldMesEst SET ValorEnt" & CStr(iMes) & " = ValorEnt" & CStr(iMes) & " + ?, "
'    sComandoSQL = sComandoSQL & "ValorSai" & CStr(iMes) & " = ValorSai" & CStr(iMes) & " +  ?, "
'    sComandoSQL = sComandoSQL & "ValorCons" & CStr(iMes) & " = ValorCons" & CStr(iMes) & " + ?, "
'    sComandoSQL = sComandoSQL & "ValorVend" & CStr(iMes) & " = ValorVend" & CStr(iMes) & " + ?, "
'    sComandoSQL = sComandoSQL & "SaldoValorCusto" & CStr(iMes) & " = SaldoValorCusto" & CStr(iMes) & " + ?"
'
'    'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
'    lErro = Comando_ExecutarPos(alComando(9), sComandoSQL, alComando(3), tSldMesEst.adValorEnt(iMes), tSldMesEst.adValorSai(iMes), tSldMesEst.adValorCons(iMes), tSldMesEst.adValorVend(iMes), tSldMesEst.adSaldoValorCusto(iMes))
'    If lErro <> AD_SQL_SUCESSO Then gError 25267
'
'    Rotina_CMP_AtualizaSldMesEst = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaSldMesEst:
'
'    Rotina_CMP_AtualizaSldMesEst = gErr
'
'    Select Case gErr
'
'         Case 25267
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESEST", gErr, iAno, iFilialEmpresa, tSldMesEst.sProduto)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159512)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaSldMesEst2(alComando() As Long, tSldMesEst2 As typeSldMesEst2, iMes As Integer, iAno As Integer, iFilialEmpresa As Integer) As Long
''Atualiza os saldos e quantidades dos escaninhos de SldMesEst2
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim sComandoSQL As String
'Dim dCustoMedioConsig As Double
'Dim dCustoMedioDemo As Double
'Dim dCustoMedioConserto As Double
'Dim dCustoMedioOutros As Double
'Dim dCustoMedioBenef As Double
'Dim iIndice As Integer
'
'On Error GoTo Erro_Rotina_CMP_AtualizaSldMesEst2
'
'    lErro = Comando_ExecutarPos(alComando(26), "SELECT QuantInicialConsig, ValorInicialConsig, SaldoQuantConsig1, SaldoValorConsig1, SaldoQuantConsig2, SaldoValorConsig2, SaldoQuantConsig3, SaldoValorConsig3, SaldoQuantConsig4, SaldoValorConsig4, SaldoQuantConsig5, SaldoValorConsig5, SaldoQuantConsig6, SaldoValorConsig6, SaldoQuantConsig7, SaldoValorConsig7, SaldoQuantConsig8, SaldoValorConsig8, SaldoQuantConsig9, SaldoValorConsig9, SaldoQuantConsig10, SaldoValorConsig10, SaldoQuantConsig11, SaldoValorConsig11, SaldoQuantConsig12, SaldoValorConsig12, " & _
'                                                "QuantInicialDemo, ValorInicialDemo, SaldoQuantDemo1, SaldoValorDemo1, SaldoQuantDemo2, SaldoValorDemo2, SaldoQuantDemo3, SaldoValorDemo3, SaldoQuantDemo4, SaldoValorDemo4, SaldoQuantDemo5, SaldoValorDemo5, SaldoQuantDemo6, SaldoValorDemo6, SaldoQuantDemo7, SaldoValorDemo7, SaldoQuantDemo8, SaldoValorDemo8, SaldoQuantDemo9, SaldoValorDemo9, SaldoQuantDemo10, SaldoValorDemo10, SaldoQuantDemo11, SaldoValorDemo11, SaldoQuantDemo12, SaldoValorDemo12, " & _
'                                                "QuantInicialConserto, ValorInicialConserto, SaldoQuantConserto1, SaldoValorConserto1, SaldoQuantConserto2, SaldoValorConserto2, SaldoQuantConserto3, SaldoValorConserto3, SaldoQuantConserto4, SaldoValorConserto4, SaldoQuantConserto5, SaldoValorConserto5, SaldoQuantConserto6, SaldoValorConserto6, SaldoQuantConserto7, SaldoValorConserto7, SaldoQuantConserto8, SaldoValorConserto8, SaldoQuantConserto9, SaldoValorConserto9, SaldoQuantConserto10, SaldoValorConserto10, SaldoQuantConserto11, SaldoValorConserto11, SaldoQuantConserto12, SaldoValorConserto12, " & _
'                                                "QuantInicialOutros, ValorInicialOutros, SaldoQuantOutros1, SaldoValorOutros1, SaldoQuantOutros2, SaldoValorOutros2, SaldoQuantOutros3, SaldoValorOutros3, SaldoQuantOutros4, SaldoValorOutros4, SaldoQuantOutros5, SaldoValorOutros5, SaldoQuantOutros6, SaldoValorOutros6, SaldoQuantOutros7, SaldoValorOutros7, SaldoQuantOutros8, SaldoValorOutros8, SaldoQuantOutros9, SaldoValorOutros9, SaldoQuantOutros10, SaldoValorOutros10, SaldoQuantOutros11, SaldoValorOutros11, SaldoQuantOutros12, SaldoValorOutros12, " & _
'                                                "QuantInicialBenef, ValorInicialBenef, SaldoQuantBenef1, SaldoValorBenef1, SaldoQuantBenef2, SaldoValorBenef2, SaldoQuantBenef3, SaldoValorBenef3, SaldoQuantBenef4, SaldoValorBenef4, SaldoQuantBenef5, SaldoValorBenef5, SaldoQuantBenef6, SaldoValorBenef6, SaldoQuantBenef7, SaldoValorBenef7, SaldoQuantBenef8, SaldoValorBenef8, SaldoQuantBenef9, SaldoValorBenef9, SaldoQuantBenef10, SaldoValorBenef10, SaldoQuantBenef11, SaldoValorBenef11, SaldoQuantBenef12, SaldoValorBenef12 " & _
'                                                "FROM SldMesEst2 WHERE FilialEmpresa = ? AND Produto = ? AND Ano = ?", 0, _
'                                                tSldMesEst2.dQuantInicialConsig, tSldMesEst2.dValorInicialConsig, tSldMesEst2.adSaldoQuantConsig(1), tSldMesEst2.adSaldoValorConsig(1), tSldMesEst2.adSaldoQuantConsig(2), tSldMesEst2.adSaldoValorConsig(2), tSldMesEst2.adSaldoQuantConsig(3), tSldMesEst2.adSaldoValorConsig(3), tSldMesEst2.adSaldoQuantConsig(4), tSldMesEst2.adSaldoValorConsig(4), tSldMesEst2.adSaldoQuantConsig(5), tSldMesEst2.adSaldoValorConsig(5), tSldMesEst2.adSaldoQuantConsig(6), tSldMesEst2.adSaldoValorConsig(6), tSldMesEst2.adSaldoQuantConsig(7), tSldMesEst2.adSaldoValorConsig(7), tSldMesEst2.adSaldoQuantConsig(8), tSldMesEst2.adSaldoValorConsig(8), tSldMesEst2.adSaldoQuantConsig(9), tSldMesEst2.adSaldoValorConsig(9), _
'                                                tSldMesEst2.adSaldoQuantConsig(10), tSldMesEst2.adSaldoValorConsig(10), tSldMesEst2.adSaldoQuantConsig(11), tSldMesEst2.adSaldoValorConsig(11), tSldMesEst2.adSaldoQuantConsig(12), tSldMesEst2.adSaldoValorConsig(12), _
'                                                tSldMesEst2.dQuantInicialDemo, tSldMesEst2.dValorInicialDemo, tSldMesEst2.adSaldoQuantDemo(1), tSldMesEst2.adSaldoValorDemo(1), tSldMesEst2.adSaldoQuantDemo(2), tSldMesEst2.adSaldoValorDemo(2), tSldMesEst2.adSaldoQuantDemo(3), tSldMesEst2.adSaldoValorDemo(3), tSldMesEst2.adSaldoQuantDemo(4), tSldMesEst2.adSaldoValorDemo(4), tSldMesEst2.adSaldoQuantDemo(5), tSldMesEst2.adSaldoValorDemo(5), tSldMesEst2.adSaldoQuantDemo(6), tSldMesEst2.adSaldoValorDemo(6), tSldMesEst2.adSaldoQuantDemo(7), tSldMesEst2.adSaldoValorDemo(7), tSldMesEst2.adSaldoQuantDemo(8), tSldMesEst2.adSaldoValorDemo(8), tSldMesEst2.adSaldoQuantDemo(9), tSldMesEst2.adSaldoValorDemo(9), _
'                                                tSldMesEst2.adSaldoQuantDemo(10), tSldMesEst2.adSaldoValorDemo(10), tSldMesEst2.adSaldoQuantDemo(11), tSldMesEst2.adSaldoValorDemo(11), tSldMesEst2.adSaldoQuantDemo(12), tSldMesEst2.adSaldoValorDemo(12), _
'                                                tSldMesEst2.dQuantInicialConserto, tSldMesEst2.dValorInicialConserto, tSldMesEst2.adSaldoQuantConserto(1), tSldMesEst2.adSaldoValorConserto(1), tSldMesEst2.adSaldoQuantConserto(2), tSldMesEst2.adSaldoValorConserto(2), tSldMesEst2.adSaldoQuantConserto(3), tSldMesEst2.adSaldoValorConserto(3), tSldMesEst2.adSaldoQuantConserto(4), tSldMesEst2.adSaldoValorConserto(4), tSldMesEst2.adSaldoQuantConserto(5), tSldMesEst2.adSaldoValorConserto(5), tSldMesEst2.adSaldoQuantConserto(6), tSldMesEst2.adSaldoValorConserto(6), tSldMesEst2.adSaldoQuantConserto(7), tSldMesEst2.adSaldoValorConserto(7), tSldMesEst2.adSaldoQuantConserto(8), tSldMesEst2.adSaldoValorConserto(8), tSldMesEst2.adSaldoQuantConserto(9), tSldMesEst2.adSaldoValorConserto(9), _
'                                                tSldMesEst2.adSaldoQuantConserto(10), tSldMesEst2.adSaldoValorConserto(10), tSldMesEst2.adSaldoQuantConserto(11), tSldMesEst2.adSaldoValorConserto(11), tSldMesEst2.adSaldoQuantConserto(12), tSldMesEst2.adSaldoValorConserto(12), _
'                                                tSldMesEst2.dQuantInicialOutros, tSldMesEst2.dValorInicialOutros, tSldMesEst2.adSaldoQuantOutros(1), tSldMesEst2.adSaldoValorOutros(1), tSldMesEst2.adSaldoQuantOutros(2), tSldMesEst2.adSaldoValorOutros(2), tSldMesEst2.adSaldoQuantOutros(3), tSldMesEst2.adSaldoValorOutros(3), tSldMesEst2.adSaldoQuantOutros(4), tSldMesEst2.adSaldoValorOutros(4), tSldMesEst2.adSaldoQuantOutros(5), tSldMesEst2.adSaldoValorOutros(5), tSldMesEst2.adSaldoQuantOutros(6), tSldMesEst2.adSaldoValorOutros(6), tSldMesEst2.adSaldoQuantOutros(7), tSldMesEst2.adSaldoValorOutros(7), tSldMesEst2.adSaldoQuantOutros(8), tSldMesEst2.adSaldoValorOutros(8), tSldMesEst2.adSaldoQuantOutros(9), tSldMesEst2.adSaldoValorOutros(9), _
'                                                tSldMesEst2.adSaldoQuantOutros(10), tSldMesEst2.adSaldoValorOutros(10), tSldMesEst2.adSaldoQuantOutros(11), tSldMesEst2.adSaldoValorOutros(11), tSldMesEst2.adSaldoQuantOutros(12), tSldMesEst2.adSaldoValorOutros(12), _
'                                                tSldMesEst2.dQuantInicialBenef, tSldMesEst2.dValorInicialBenef, tSldMesEst2.adSaldoQuantBenef(1), tSldMesEst2.adSaldoValorBenef(1), tSldMesEst2.adSaldoQuantBenef(2), tSldMesEst2.adSaldoValorBenef(2), tSldMesEst2.adSaldoQuantBenef(3), tSldMesEst2.adSaldoValorBenef(3), tSldMesEst2.adSaldoQuantBenef(4), tSldMesEst2.adSaldoValorBenef(4), tSldMesEst2.adSaldoQuantBenef(5), tSldMesEst2.adSaldoValorBenef(5), tSldMesEst2.adSaldoQuantBenef(6), tSldMesEst2.adSaldoValorBenef(6), tSldMesEst2.adSaldoQuantBenef(7), tSldMesEst2.adSaldoValorBenef(7), tSldMesEst2.adSaldoQuantBenef(8), tSldMesEst2.adSaldoValorBenef(8), tSldMesEst2.adSaldoQuantBenef(9), tSldMesEst2.adSaldoValorBenef(9), _
'                                                tSldMesEst2.adSaldoQuantBenef(10), tSldMesEst2.adSaldoValorBenef(10), tSldMesEst2.adSaldoQuantBenef(11), tSldMesEst2.adSaldoValorBenef(11), tSldMesEst2.adSaldoQuantBenef(12), tSldMesEst2.adSaldoValorBenef(12), _
'                                                tSldMesEst2.iFilialEmpresa, tSldMesEst2.sProduto, tSldMesEst2.iAno)
'    If lErro <> AD_SQL_SUCESSO Then gError 89832
'
'    lErro = Comando_BuscarPrimeiro(alComando(26))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89833
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 89834
'
'    For iIndice = 1 To iMes
'
'        tSldMesEst2.dQuantInicialConsig = tSldMesEst2.dQuantInicialConsig + tSldMesEst2.adSaldoQuantConsig(iIndice)
'        tSldMesEst2.dValorInicialConsig = tSldMesEst2.dValorInicialConsig + tSldMesEst2.adSaldoValorConsig(iIndice)
'
'        tSldMesEst2.dQuantInicialDemo = tSldMesEst2.dQuantInicialDemo + tSldMesEst2.adSaldoQuantDemo(iIndice)
'        tSldMesEst2.dValorInicialDemo = tSldMesEst2.dValorInicialDemo + tSldMesEst2.adSaldoValorDemo(iIndice)
'
'        tSldMesEst2.dQuantInicialConserto = tSldMesEst2.dQuantInicialConserto + tSldMesEst2.adSaldoQuantConserto(iIndice)
'        tSldMesEst2.dValorInicialConserto = tSldMesEst2.dValorInicialConserto + tSldMesEst2.adSaldoValorConserto(iIndice)
'
'        tSldMesEst2.dQuantInicialOutros = tSldMesEst2.dQuantInicialOutros + tSldMesEst2.adSaldoQuantOutros(iIndice)
'        tSldMesEst2.dValorInicialOutros = tSldMesEst2.dValorInicialOutros + tSldMesEst2.adSaldoValorOutros(iIndice)
'
'        tSldMesEst2.dQuantInicialBenef = tSldMesEst2.dQuantInicialBenef + tSldMesEst2.adSaldoQuantBenef(iIndice)
'        tSldMesEst2.dValorInicialBenef = tSldMesEst2.dValorInicialBenef + tSldMesEst2.adSaldoValorBenef(iIndice)
'
'    Next
'
'    If tSldMesEst2.dQuantInicialConsig > 0 Then
'        dCustoMedioConsig = tSldMesEst2.dValorInicialConsig / tSldMesEst2.dQuantInicialConsig
'    Else
'        dCustoMedioConsig = 0
'    End If
'
'    If tSldMesEst2.dQuantInicialDemo > 0 Then
'        dCustoMedioDemo = tSldMesEst2.dValorInicialDemo / tSldMesEst2.dQuantInicialDemo
'    Else
'        dCustoMedioDemo = 0
'    End If
'
'    If tSldMesEst2.dQuantInicialConserto > 0 Then
'        dCustoMedioConserto = tSldMesEst2.dValorInicialDemo / tSldMesEst2.dQuantInicialConserto
'    Else
'        dCustoMedioConserto = 0
'    End If
'
'    If tSldMesEst2.dQuantInicialOutros > 0 Then
'        dCustoMedioOutros = tSldMesEst2.dValorInicialOutros / tSldMesEst2.dQuantInicialOutros
'    Else
'        dCustoMedioOutros = 0
'    End If
'
'    If tSldMesEst2.dQuantInicialBenef > 0 Then
'        dCustoMedioBenef = tSldMesEst2.dValorInicialBenef / tSldMesEst2.dQuantInicialBenef
'    Else
'        dCustoMedioBenef = 0
'    End If
'
'    'Monta comando SQL para UPDATE de SldMesEst
'    sComandoSQL = "UPDATE SldMesEst2 SET SaldoValorConsig" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "SaldoValorDemo" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "SaldoValorConserto" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "SaldoValorOutros" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "SaldoValorBenef" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "CustoMedioConsig" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "CustoMedioDemo" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "CustoMedioConserto" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "CustoMedioOutros" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "CustoMedioBenef" & CStr(iMes) & " = ?"
'
'    'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
'    lErro = Comando_ExecutarPos(alComando(27), sComandoSQL, alComando(26), tSldMesEst2.adSaldoValorConsig(iMes), tSldMesEst2.adSaldoValorDemo(iMes), tSldMesEst2.adSaldoValorConserto(iMes), tSldMesEst2.adSaldoValorOutros(iMes), tSldMesEst2.adSaldoValorBenef(iMes), dCustoMedioConsig, dCustoMedioDemo, dCustoMedioConserto, dCustoMedioOutros, dCustoMedioBenef)
'    If lErro <> AD_SQL_SUCESSO Then gError 89835
'
'    Rotina_CMP_AtualizaSldMesEst2 = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaSldMesEst2:
'
'    Rotina_CMP_AtualizaSldMesEst2 = gErr
'
'    Select Case gErr
'
'        Case 89832, 89833
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, tSldMesEst2.iAno, tSldMesEst2.iFilialEmpresa, tSldMesEst2.sProduto)
'
'        Case 89834
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESEST2", gErr, tSldMesEst2.iAno, tSldMesEst2.iFilialEmpresa, tSldMesEst2.sProduto)
'
'        Case 89835
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST2_NAO_CADASTRADO", gErr, tSldMesEst2.iAno, tSldMesEst2.iFilialEmpresa, tSldMesEst2.sProduto)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159513)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaSldMesEst1(alComando() As Long, tSldMesEst1 As typeSldMesEst1, iMes As Integer, iAno As Integer, iFilialEmpresa As Integer, dCMPAtual As Double, tTipoMovEst As typeTipoMovEst) As Long
''Atualiza os saldos e quantidades dos escaninhos de SldMesEst1
''Chamada EM TRANSAÇÃO
'
'' *** ALTERADA POR LUIZ G.F. NOGUEIRA EM 18/07/2001 ***
'' A alteração foi feita para que a função, além de atualizar SldMesEst1, disponibilize
'' o custo médio do escaninho de terceiros ao qual o movimento se refere
'' ******************************************************************************************
'
'Dim lErro As Long
'Dim sComandoSQL As String
'Dim dCustoMedioConsig3 As Double
'Dim dCustoMedioDemo3 As Double
'Dim dCustoMedioConserto3 As Double
'Dim dCustoMedioOutros3 As Double
'Dim dCustoMedioBenef3 As Double
'Dim iIndice As Integer
'Dim tSldMesEst1Aux As typeSldMesEst1
'
'On Error GoTo Erro_Rotina_CMP_AtualizaSldMesEst1
'
'    ' *** ALTERADO POR LUIZ G.F. NOGUEIRA EM 17/07/2001 ***
'    'Guarda os valores que foram acumulados para as colunas SaldoValor e para a coluna SaldoQuantBenef3
'    'Isso evita que a leitura de BD sobreponha os valores que foram calculados
'    tSldMesEst1Aux.adSaldoValorBenef3(iMes) = tSldMesEst1.adSaldoValorBenef3(iMes)
'    tSldMesEst1Aux.adSaldoQuantBenef3(iMes) = tSldMesEst1.adSaldoQuantBenef3(iMes)
'    tSldMesEst1Aux.adSaldoValorConserto3(iMes) = tSldMesEst1.adSaldoValorConserto3(iMes)
'    tSldMesEst1Aux.adSaldoValorConsig3(iMes) = tSldMesEst1.adSaldoValorConsig3(iMes)
'    tSldMesEst1Aux.adSaldoValorDemo3(iMes) = tSldMesEst1.adSaldoValorDemo3(iMes)
'    tSldMesEst1Aux.adSaldoValorOutros3(iMes) = tSldMesEst1.adSaldoValorOutros3(iMes)
'    ' *******************************************************
'
'    lErro = Comando_ExecutarPos(alComando(24), "SELECT QuantInicialConsig3, ValorInicialConsig3, SaldoQuantConsig31, SaldoValorConsig31, SaldoQuantConsig32, SaldoValorConsig32, SaldoQuantConsig33, SaldoValorConsig33, SaldoQuantConsig34, SaldoValorConsig34, SaldoQuantConsig35, SaldoValorConsig35, SaldoQuantConsig36, SaldoValorConsig36, SaldoQuantConsig37, SaldoValorConsig37, SaldoQuantConsig38, SaldoValorConsig38, SaldoQuantConsig39, SaldoValorConsig39, SaldoQuantConsig310, SaldoValorConsig310, SaldoQuantConsig311, SaldoValorConsig311, SaldoQuantConsig312, SaldoValorConsig312, " & _
'                                                "QuantInicialDemo3, ValorInicialDemo3, SaldoQuantDemo31, SaldoValorDemo31, SaldoQuantDemo32, SaldoValorDemo32, SaldoQuantDemo33, SaldoValorDemo33, SaldoQuantDemo34, SaldoValorDemo34, SaldoQuantDemo35, SaldoValorDemo35, SaldoQuantDemo36, SaldoValorDemo36, SaldoQuantDemo37, SaldoValorDemo37, SaldoQuantDemo38, SaldoValorDemo38, SaldoQuantDemo39, SaldoValorDemo39, SaldoQuantDemo310, SaldoValorDemo310, SaldoQuantDemo311, SaldoValorDemo311, SaldoQuantDemo312, SaldoValorDemo312, " & _
'                                                "QuantInicialConserto3, ValorInicialConserto3, SaldoQuantConserto31, SaldoValorConserto31, SaldoQuantConserto32, SaldoValorConserto32, SaldoQuantConserto33, SaldoValorConserto33, SaldoQuantConserto34, SaldoValorConserto34, SaldoQuantConserto35, SaldoValorConserto35, SaldoQuantConserto36, SaldoValorConserto36, SaldoQuantConserto37, SaldoValorConserto37, SaldoQuantConserto38, SaldoValorConserto38, SaldoQuantConserto39, SaldoValorConserto39, SaldoQuantConserto310, SaldoValorConserto310, SaldoQuantConserto311, SaldoValorConserto311, SaldoQuantConserto312, SaldoValorConserto312, " & _
'                                                "QuantInicialOutros3, ValorInicialOutros3, SaldoQuantOutros31, SaldoValorOutros31, SaldoQuantOutros32, SaldoValorOutros32, SaldoQuantOutros33, SaldoValorOutros33, SaldoQuantOutros34, SaldoValorOutros34, SaldoQuantOutros35, SaldoValorOutros35, SaldoQuantOutros36, SaldoValorOutros36, SaldoQuantOutros37, SaldoValorOutros37, SaldoQuantOutros38, SaldoValorOutros38, SaldoQuantOutros39, SaldoValorOutros39, SaldoQuantOutros310, SaldoValorOutros310, SaldoQuantOutros311, SaldoValorOutros311, SaldoQuantOutros312, SaldoValorOutros312, " & _
'                                                "QuantInicialBenef3, ValorInicialBenef3, SaldoQuantBenef31, SaldoValorBenef31, SaldoQuantBenef32, SaldoValorBenef32, SaldoQuantBenef33, SaldoValorBenef33, SaldoQuantBenef34, SaldoValorBenef34, SaldoQuantBenef35, SaldoValorBenef35, SaldoQuantBenef36, SaldoValorBenef36, SaldoQuantBenef37, SaldoValorBenef37, SaldoQuantBenef38, SaldoValorBenef38, SaldoQuantBenef39, SaldoValorBenef39, SaldoQuantBenef310, SaldoValorBenef310, SaldoQuantBenef311, SaldoValorBenef311, SaldoQuantBenef312, SaldoValorBenef312 " & _
'                                                "FROM SldMesEst1 WHERE FilialEmpresa = ? AND Produto = ? AND Ano = ?", 0, _
'                                                tSldMesEst1.dQuantInicialConsig3, tSldMesEst1.dValorInicialConsig3, tSldMesEst1.adSaldoQuantConsig3(1), tSldMesEst1.adSaldoValorConsig3(1), tSldMesEst1.adSaldoQuantConsig3(2), tSldMesEst1.adSaldoValorConsig3(2), tSldMesEst1.adSaldoQuantConsig3(3), tSldMesEst1.adSaldoValorConsig3(3), tSldMesEst1.adSaldoQuantConsig3(4), tSldMesEst1.adSaldoValorConsig3(4), tSldMesEst1.adSaldoQuantConsig3(5), tSldMesEst1.adSaldoValorConsig3(5), tSldMesEst1.adSaldoQuantConsig3(6), tSldMesEst1.adSaldoValorConsig3(6), tSldMesEst1.adSaldoQuantConsig3(7), tSldMesEst1.adSaldoValorConsig3(7), tSldMesEst1.adSaldoQuantConsig3(8), tSldMesEst1.adSaldoValorConsig3(8), tSldMesEst1.adSaldoQuantConsig3(9), tSldMesEst1.adSaldoValorConsig3(9), _
'                                                tSldMesEst1.adSaldoQuantConsig3(10), tSldMesEst1.adSaldoValorConsig3(10), tSldMesEst1.adSaldoQuantConsig3(11), tSldMesEst1.adSaldoValorConsig3(11), tSldMesEst1.adSaldoQuantConsig3(12), tSldMesEst1.adSaldoValorConsig3(12), _
'                                                tSldMesEst1.dQuantInicialDemo3, tSldMesEst1.dValorInicialDemo3, tSldMesEst1.adSaldoQuantDemo3(1), tSldMesEst1.adSaldoValorDemo3(1), tSldMesEst1.adSaldoQuantDemo3(2), tSldMesEst1.adSaldoValorDemo3(2), tSldMesEst1.adSaldoQuantDemo3(3), tSldMesEst1.adSaldoValorDemo3(3), tSldMesEst1.adSaldoQuantDemo3(4), tSldMesEst1.adSaldoValorDemo3(4), tSldMesEst1.adSaldoQuantDemo3(5), tSldMesEst1.adSaldoValorDemo3(5), tSldMesEst1.adSaldoQuantDemo3(6), tSldMesEst1.adSaldoValorDemo3(6), tSldMesEst1.adSaldoQuantDemo3(7), tSldMesEst1.adSaldoValorDemo3(7), tSldMesEst1.adSaldoQuantDemo3(8), tSldMesEst1.adSaldoValorDemo3(8), tSldMesEst1.adSaldoQuantDemo3(9), tSldMesEst1.adSaldoValorDemo3(9), _
'                                                tSldMesEst1.adSaldoQuantDemo3(10), tSldMesEst1.adSaldoValorDemo3(10), tSldMesEst1.adSaldoQuantDemo3(11), tSldMesEst1.adSaldoValorDemo3(11), tSldMesEst1.adSaldoQuantDemo3(12), tSldMesEst1.adSaldoValorDemo3(12), _
'                                                tSldMesEst1.dQuantInicialConserto3, tSldMesEst1.dValorInicialConserto3, tSldMesEst1.adSaldoQuantConserto3(1), tSldMesEst1.adSaldoValorConserto3(1), tSldMesEst1.adSaldoQuantConserto3(2), tSldMesEst1.adSaldoValorConserto3(2), tSldMesEst1.adSaldoQuantConserto3(3), tSldMesEst1.adSaldoValorConserto3(3), tSldMesEst1.adSaldoQuantConserto3(4), tSldMesEst1.adSaldoValorConserto3(4), tSldMesEst1.adSaldoQuantConserto3(5), tSldMesEst1.adSaldoValorConserto3(5), tSldMesEst1.adSaldoQuantConserto3(6), tSldMesEst1.adSaldoValorConserto3(6), tSldMesEst1.adSaldoQuantConserto3(7), tSldMesEst1.adSaldoValorConserto3(7), tSldMesEst1.adSaldoQuantConserto3(8), tSldMesEst1.adSaldoValorConserto3(8), tSldMesEst1.adSaldoQuantConserto3(9), tSldMesEst1.adSaldoValorConserto3(9), _
'                                                tSldMesEst1.adSaldoQuantConserto3(10), tSldMesEst1.adSaldoValorConserto3(10), tSldMesEst1.adSaldoQuantConserto3(11), tSldMesEst1.adSaldoValorConserto3(11), tSldMesEst1.adSaldoQuantConserto3(12), tSldMesEst1.adSaldoValorConserto3(12), _
'                                                tSldMesEst1.dQuantInicialOutros3, tSldMesEst1.dValorInicialOutros3, tSldMesEst1.adSaldoQuantOutros3(1), tSldMesEst1.adSaldoValorOutros3(1), tSldMesEst1.adSaldoQuantOutros3(2), tSldMesEst1.adSaldoValorOutros3(2), tSldMesEst1.adSaldoQuantOutros3(3), tSldMesEst1.adSaldoValorOutros3(3), tSldMesEst1.adSaldoQuantOutros3(4), tSldMesEst1.adSaldoValorOutros3(4), tSldMesEst1.adSaldoQuantOutros3(5), tSldMesEst1.adSaldoValorOutros3(5), tSldMesEst1.adSaldoQuantOutros3(6), tSldMesEst1.adSaldoValorOutros3(6), tSldMesEst1.adSaldoQuantOutros3(7), tSldMesEst1.adSaldoValorOutros3(7), tSldMesEst1.adSaldoQuantOutros3(8), tSldMesEst1.adSaldoValorOutros3(8), tSldMesEst1.adSaldoQuantOutros3(9), tSldMesEst1.adSaldoValorOutros3(9), _
'                                                tSldMesEst1.adSaldoQuantOutros3(10), tSldMesEst1.adSaldoValorOutros3(10), tSldMesEst1.adSaldoQuantOutros3(11), tSldMesEst1.adSaldoValorOutros3(11), tSldMesEst1.adSaldoQuantOutros3(12), tSldMesEst1.adSaldoValorOutros3(12), _
'                                                tSldMesEst1.dQuantInicialBenef3, tSldMesEst1.dValorInicialBenef3, tSldMesEst1.adSaldoQuantBenef3(1), tSldMesEst1.adSaldoValorBenef3(1), tSldMesEst1.adSaldoQuantBenef3(2), tSldMesEst1.adSaldoValorBenef3(2), tSldMesEst1.adSaldoQuantBenef3(3), tSldMesEst1.adSaldoValorBenef3(3), tSldMesEst1.adSaldoQuantBenef3(4), tSldMesEst1.adSaldoValorBenef3(4), tSldMesEst1.adSaldoQuantBenef3(5), tSldMesEst1.adSaldoValorBenef3(5), tSldMesEst1.adSaldoQuantBenef3(6), tSldMesEst1.adSaldoValorBenef3(6), tSldMesEst1.adSaldoQuantBenef3(7), tSldMesEst1.adSaldoValorBenef3(7), tSldMesEst1.adSaldoQuantBenef3(8), tSldMesEst1.adSaldoValorBenef3(8), tSldMesEst1.adSaldoQuantBenef3(9), tSldMesEst1.adSaldoValorBenef3(9), _
'                                                tSldMesEst1.adSaldoQuantBenef3(10), tSldMesEst1.adSaldoValorBenef3(10), tSldMesEst1.adSaldoQuantBenef3(11), tSldMesEst1.adSaldoValorBenef3(11), tSldMesEst1.adSaldoQuantBenef3(12), tSldMesEst1.adSaldoValorBenef3(12), _
'                                                tSldMesEst1.iFilialEmpresa, tSldMesEst1.sProduto, tSldMesEst1.iAno)
'    If lErro <> AD_SQL_SUCESSO Then gError 89822
'
'    lErro = Comando_BuscarPrimeiro(alComando(24))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89823
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 89824
'
'    ' *** ALTERADO POR LUIZ G.F. NOGUEIRA EM 17/07/2001 ***
'    'Devolve os valores acumulados que foram guardados
'    tSldMesEst1.adSaldoValorBenef3(iMes) = tSldMesEst1Aux.adSaldoValorBenef3(iMes)
'    tSldMesEst1.adSaldoQuantBenef3(iMes) = tSldMesEst1Aux.adSaldoQuantBenef3(iMes)
'    tSldMesEst1.adSaldoValorConserto3(iMes) = tSldMesEst1Aux.adSaldoValorConserto3(iMes)
'    tSldMesEst1.adSaldoValorConsig3(iMes) = tSldMesEst1Aux.adSaldoValorConsig3(iMes)
'    tSldMesEst1.adSaldoValorDemo3(iMes) = tSldMesEst1Aux.adSaldoValorDemo3(iMes)
'    tSldMesEst1.adSaldoValorOutros3(iMes) = tSldMesEst1Aux.adSaldoValorOutros3(iMes)
'    ' *******************************************************
'
'    For iIndice = 1 To iMes
'
'        tSldMesEst1.dQuantInicialConsig3 = tSldMesEst1.dQuantInicialConsig3 + tSldMesEst1.adSaldoQuantConsig3(iIndice)
'        tSldMesEst1.dValorInicialConsig3 = tSldMesEst1.dValorInicialConsig3 + tSldMesEst1.adSaldoValorConsig3(iIndice)
'
'        tSldMesEst1.dQuantInicialDemo3 = tSldMesEst1.dQuantInicialDemo3 + tSldMesEst1.adSaldoQuantDemo3(iIndice)
'        tSldMesEst1.dValorInicialDemo3 = tSldMesEst1.dValorInicialDemo3 + tSldMesEst1.adSaldoValorDemo3(iIndice)
'
'        tSldMesEst1.dQuantInicialConserto3 = tSldMesEst1.dQuantInicialConserto3 + tSldMesEst1.adSaldoQuantConserto3(iIndice)
'        tSldMesEst1.dValorInicialConserto3 = tSldMesEst1.dValorInicialConserto3 + tSldMesEst1.adSaldoValorConserto3(iIndice)
'
'        tSldMesEst1.dQuantInicialOutros3 = tSldMesEst1.dQuantInicialOutros3 + tSldMesEst1.adSaldoQuantOutros3(iIndice)
'        tSldMesEst1.dValorInicialOutros3 = tSldMesEst1.dValorInicialOutros3 + tSldMesEst1.adSaldoValorOutros3(iIndice)
'
'        tSldMesEst1.dQuantInicialBenef3 = tSldMesEst1.dQuantInicialBenef3 + tSldMesEst1.adSaldoQuantBenef3(iIndice)
'        tSldMesEst1.dValorInicialBenef3 = tSldMesEst1.dValorInicialBenef3 + tSldMesEst1.adSaldoValorBenef3(iIndice)
'
'    Next
'
'    If tSldMesEst1.dQuantInicialConsig3 > 0 Then
'        dCustoMedioConsig3 = tSldMesEst1.dValorInicialConsig3 / tSldMesEst1.dQuantInicialConsig3
'    Else
'        dCustoMedioConsig3 = 0
'    End If
'
'    If tSldMesEst1.dQuantInicialDemo3 > 0 Then
'        dCustoMedioDemo3 = tSldMesEst1.dValorInicialDemo3 / tSldMesEst1.dQuantInicialDemo3
'    Else
'        dCustoMedioDemo3 = 0
'    End If
'
'    If tSldMesEst1.dQuantInicialConserto3 > 0 Then
'        dCustoMedioConserto3 = tSldMesEst1.dValorInicialConserto3 / tSldMesEst1.dQuantInicialConserto3
'    Else
'        dCustoMedioConserto3 = 0
'    End If
'
'    If tSldMesEst1.dQuantInicialOutros3 > 0 Then
'        dCustoMedioOutros3 = tSldMesEst1.dValorInicialOutros3 / tSldMesEst1.dQuantInicialOutros3
'    Else
'        dCustoMedioOutros3 = 0
'    End If
'
'    If tSldMesEst1.dQuantInicialBenef3 > 0 Then
'        dCustoMedioBenef3 = tSldMesEst1.dValorInicialBenef3 / tSldMesEst1.dQuantInicialBenef3
'    Else
'        dCustoMedioBenef3 = 0
'    End If
'
'    ' *** INCLUÍDO POR LUIZ G.F. NOGUEIRA EM 18/07/2001 ***
'    'Seta a variável iEscaninho de acordo com o escaninho que está sendo movimentado
'    If tTipoMovEst.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Or tTipoMovEst.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then dCMPAtual = dCustoMedioBenef3
'    If tTipoMovEst.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Or tTipoMovEst.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then dCMPAtual = dCustoMedioConserto3
'    If tTipoMovEst.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Or tTipoMovEst.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then dCMPAtual = dCustoMedioConsig3
'    If tTipoMovEst.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Or tTipoMovEst.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then dCMPAtual = dCustoMedioDemo3
'    If tTipoMovEst.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Or tTipoMovEst.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then dCMPAtual = dCustoMedioOutros3
'    '*************************************************************
'
'    'Monta comando SQL para UPDATE de SldMesEst
'    sComandoSQL = "UPDATE SldMesEst1 SET SaldoValorConsig3" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "SaldoValorDemo3" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "SaldoValorConserto3" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "SaldoValorOutros3" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "SaldoValorBenef3" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "CustoMedioConsig3" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "CustoMedioDemo3" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "CustoMedioConserto3" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "CustoMedioOutros3" & CStr(iMes) & " = ?, "
'    sComandoSQL = sComandoSQL & "CustoMedioBenef3" & CStr(iMes) & " = ?"
'
'    'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
'    lErro = Comando_ExecutarPos(alComando(25), sComandoSQL, alComando(24), tSldMesEst1.adSaldoValorConsig3(iMes), tSldMesEst1.adSaldoValorDemo3(iMes), tSldMesEst1.adSaldoValorConserto3(iMes), tSldMesEst1.adSaldoValorOutros3(iMes), tSldMesEst1.adSaldoValorBenef3(iMes), dCustoMedioConsig3, dCustoMedioDemo3, dCustoMedioConserto3, dCustoMedioOutros3, dCustoMedioBenef3)
'    If lErro <> AD_SQL_SUCESSO Then gError 89825
'
'    Rotina_CMP_AtualizaSldMesEst1 = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaSldMesEst1:
'
'    Rotina_CMP_AtualizaSldMesEst1 = gErr
'
'    Select Case gErr
'
'        Case 89822, 89823
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, tSldMesEst1.iAno, tSldMesEst1.iFilialEmpresa, tSldMesEst1.sProduto)
'
'        Case 89824
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESEST1", gErr, tSldMesEst1.iAno, tSldMesEst1.iFilialEmpresa, tSldMesEst1.sProduto)
'
'        Case 89825
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST1_NAO_CADASTRADO", gErr, tSldMesEst1.iAno, tSldMesEst1.iFilialEmpresa, tSldMesEst1.sProduto)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159514)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaSldMesEstAlm(alComando() As Long, sProduto As String, colAlmoxInfo As Collection, iMes As Integer, iAno As Integer) As Long
''Atualiza para um dado Produto em vários Almoxarifados, ValorEnt, ValorSai, ValorCons do iMes na tabela SaldoMesEstAlm
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim sComandoSQL(1 To 2) As String
'Dim objAlmoxInfo As ClassAlmoxInfo
'Dim sCodProduto As String
'
'On Error GoTo Erro_Rotina_CMP_AtualizaSldMesEstAlm
'
'    'SELECT em SldMesEstAlm filtrando pela chave primária
'    sComandoSQL(1) = "SELECT Produto FROM SldMesEstAlm WHERE Almoxarifado = ? AND Produto = ? AND Ano = ?"
'
'    'Comando SQL para UPDATE de SldMesEstAlm
'    sComandoSQL(2) = "UPDATE SldMesEstAlm SET ValorEnt" & CStr(iMes) & " = ValorEnt" & CStr(iMes) & " + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorSai" & CStr(iMes) & " = ValorSai" & CStr(iMes) & " + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorVend" & CStr(iMes) & " = ValorVend" & CStr(iMes) & " + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "ValorCons" & CStr(iMes) & " = ValorCons" & CStr(iMes) & " + ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "SaldoValorCusto" & CStr(iMes) & " = SaldoValorCusto" & CStr(iMes) & " + ?"
'
'    For Each objAlmoxInfo In colAlmoxInfo
'
'        'Inicializa comando da tabela SaldoMesEstAlm para atualizar Valores de Entrada, Saída, Consumo
'        sCodProduto = String(STRING_PRODUTO, 0)
'
'        lErro = Comando_ExecutarPos(alComando(6), sComandoSQL(1), 0, sCodProduto, objAlmoxInfo.iAlmoxarifado, sProduto, iAno)
'        If lErro <> AD_SQL_SUCESSO Then gError 25881
'
'        lErro = Comando_BuscarProximo(alComando(6))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25895
'        If lErro = AD_SQL_SEM_DADOS Then gError 25894
'
'        'Atualiza tabela SaldoMesEstAlm, campos ValorEnt, ValorSai, ValorCons de iMes
'        lErro = Comando_ExecutarPos(alComando(14), sComandoSQL(2), alComando(6), objAlmoxInfo.dValorEntrada, objAlmoxInfo.dValorSaida, objAlmoxInfo.dValorVenda, objAlmoxInfo.dValorCons, objAlmoxInfo.dSaldoValorCusto)
'        If lErro <> AD_SQL_SUCESSO Then gError 25893
'
'    Next
'
'    Rotina_CMP_AtualizaSldMesEstAlm = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaSldMesEstAlm:
'
'    Rotina_CMP_AtualizaSldMesEstAlm = gErr
'
'    Select Case gErr
'
'        Case 25881
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL(1))
'
'        Case 25893
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESESTALM", gErr, iAno, objAlmoxInfo.iAlmoxarifado, sProduto)
'
'        Case 25894
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SLDMESESTALM", gErr, iAno, sProduto, objAlmoxInfo.iAlmoxarifado)
'
'        Case 25895
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM", gErr, iAno, objAlmoxInfo.iAlmoxarifado, sProduto)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159515)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaSldMesEstAlm2(alComando() As Long, sProduto As String, colAlmoxInfo As Collection, iMes As Integer, iAno As Integer) As Long
''Atualiza para um dado Produto em vários Almoxarifados o escaninho do iMes na tabela SaldoMesEstAlm2
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim sComandoSQL(1 To 2) As String
'Dim objAlmoxInfo As ClassAlmoxInfo
'Dim sCodProduto As String
'
'On Error GoTo Erro_Rotina_CMP_AtualizaSldMesEstAlm2
'
'    'SELECT em SldMesEstAlm filtrando pela chave primária
'    sComandoSQL(1) = "SELECT Produto FROM SldMesEstAlm2 WHERE Almoxarifado = ? AND Produto = ? AND Ano = ?"
'
'    'Comando SQL para UPDATE de SldMesEstAlm
'    sComandoSQL(2) = "UPDATE SldMesEstAlm2 SET SaldoValorConsig" & CStr(iMes) & " = ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "SaldoValorDemo" & CStr(iMes) & " = ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "SaldoValorConserto" & CStr(iMes) & " = ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "SaldoValorOutros" & CStr(iMes) & " = ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "SaldoValorBenef" & CStr(iMes) & " = ?"
'
'    For Each objAlmoxInfo In colAlmoxInfo
'
'        'Inicializa comando da tabela SaldoMesEstAlm para atualizar Valores de Entrada, Saída, Consumo
'        sCodProduto = String(STRING_PRODUTO, 0)
'
'        lErro = Comando_ExecutarPos(alComando(19), sComandoSQL(1), 0, sCodProduto, objAlmoxInfo.iAlmoxarifado, sProduto, iAno)
'        If lErro <> AD_SQL_SUCESSO Then gError 69735
'
'        lErro = Comando_BuscarProximo(alComando(19))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69736
'        If lErro = AD_SQL_SEM_DADOS Then gError 69737
'
'        'Atualiza tabela SaldoMesEstAlm, campos ValorEnt, ValorSai, ValorCons de iMes
'        lErro = Comando_ExecutarPos(alComando(20), sComandoSQL(2), alComando(19), objAlmoxInfo.dSaldoValorConsig, objAlmoxInfo.dSaldoValorDemo, objAlmoxInfo.dSaldoValorConserto, objAlmoxInfo.dSaldoValorOutros, objAlmoxInfo.dSaldoValorBenef)
'        If lErro <> AD_SQL_SUCESSO Then gError 69738
'
'    Next
'
'    Rotina_CMP_AtualizaSldMesEstAlm2 = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaSldMesEstAlm2:
'
'    Rotina_CMP_AtualizaSldMesEstAlm2 = gErr
'
'    Select Case gErr
'
'        Case 69735, 69736
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM21", gErr, iAno, objAlmoxInfo.iAlmoxarifado, sProduto)
'
'        Case 69737
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESESTALM2", gErr, iAno, objAlmoxInfo.iAlmoxarifado, sProduto)
'
'        Case 69738
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESESTALM2_NAO_CADASTRADO", gErr, iAno, sProduto, objAlmoxInfo.iAlmoxarifado)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159516)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaSldMesEstAlm1(alComando() As Long, sProduto As String, colAlmoxInfo As Collection, iMes As Integer, iAno As Integer) As Long
''Atualiza para um dado Produto em vários Almoxarifados o escaninho do iMes na tabela SaldoMesEstAlm2
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim sComandoSQL(1 To 2) As String
'Dim objAlmoxInfo As ClassAlmoxInfo
'Dim sCodProduto As String
'
'On Error GoTo Erro_Rotina_CMP_AtualizaSldMesEstAlm1
'
'    'SELECT em SldMesEstAlm filtrando pela chave primária
'    sComandoSQL(1) = "SELECT Produto FROM SldMesEstAlm1 WHERE Almoxarifado = ? AND Produto = ? AND Ano = ?"
'
'    'Comando SQL para UPDATE de SldMesEstAlm
'    sComandoSQL(2) = "UPDATE SldMesEstAlm1 SET SaldoValorConsig3" & CStr(iMes) & " = ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "SaldoValorDemo3" & CStr(iMes) & " = ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "SaldoValorConserto3" & CStr(iMes) & " = ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "SaldoValorOutros3" & CStr(iMes) & " = ?, "
'    sComandoSQL(2) = sComandoSQL(2) & "SaldoValorBenef3" & CStr(iMes) & " = ?"
'
'    For Each objAlmoxInfo In colAlmoxInfo
'
'        'Inicializa comando da tabela SaldoMesEstAlm para atualizar Valores de Entrada, Saída, Consumo
'        sCodProduto = String(STRING_PRODUTO, 0)
'
'        lErro = Comando_ExecutarPos(alComando(22), sComandoSQL(1), 0, sCodProduto, objAlmoxInfo.iAlmoxarifado, sProduto, iAno)
'        If lErro <> AD_SQL_SUCESSO Then gError 89817
'
'        lErro = Comando_BuscarProximo(alComando(22))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89818
'        If lErro = AD_SQL_SEM_DADOS Then gError 89819
'
'        'Atualiza tabela SaldoMesEstAlm, campos ValorEnt, ValorSai, ValorCons de iMes
'        lErro = Comando_ExecutarPos(alComando(23), sComandoSQL(2), alComando(22), objAlmoxInfo.dSaldoValorConsig3, objAlmoxInfo.dSaldoValorDemo3, objAlmoxInfo.dSaldoValorConserto3, objAlmoxInfo.dSaldoValorOutros3, objAlmoxInfo.dSaldoValorBenef3)
'        If lErro <> AD_SQL_SUCESSO Then gError 89820
'
'    Next
'
'    Rotina_CMP_AtualizaSldMesEstAlm1 = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaSldMesEstAlm1:
'
'    Rotina_CMP_AtualizaSldMesEstAlm1 = gErr
'
'    Select Case gErr
'
'        Case 89817, 89818
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM11", gErr, iAno, objAlmoxInfo.iAlmoxarifado, sProduto)
'
'        Case 89819
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESESTALM1", gErr, iAno, objAlmoxInfo.iAlmoxarifado, sProduto)
'
'        Case 89820
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESESTALM1_NAO_CADASTRADO", gErr, iAno, sProduto, objAlmoxInfo.iAlmoxarifado)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159517)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaCMP(sComandoSQL() As String, alComando() As Long, tProduto As typeProduto, dCMPAtual As Double, tTipoMovEst As typeTipoMovEst, tMovEstoque As typeItemMovEstoque, tMovEstoque2 As typeItemMovEstoque, tSldMesEst As typeSldMesEst, tSldDiaEst As typeSldDiaEst, tSldDiaEstAlm As typeSldDiaEstAlm, objAlmoxInfo As ClassAlmoxInfo, iMes As Integer, iFilialEmpresa As Integer) As Long
''Atualiza CMP nos movimentos. Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim dFator As Double
'
'On Error GoTo Erro_Rotina_CMP_AtualizaCMP
'
'    'Fator de conv. de UMs
'    lErro = CF("UM_Conversao",tProduto.iClasseUM, tMovEstoque.sSiglaUM, tProduto.sSiglaUMEstoque, dFator)
'    If lErro <> SUCESSO Then gError 25268
'
'    'Calcula custo do Movimento
'    tMovEstoque.dCusto = dCMPAtual * tMovEstoque.dQuantidade * dFator
'
'    'se não for um estorno
'    If tTipoMovEst.iCodigoOrig = 0 Then
'
'        'Atualiza custo do Movimento
'        lErro = Comando_ExecutarPos(alComando(12), "UPDATE MovimentoEstoqueES SET Custo = ?", alComando(2), tMovEstoque.dCusto)
'        If lErro <> AD_SQL_SUCESSO Then gError 25269
'
'    Else
'
'        'Atualiza custo do Movimento. Torna o sinal do custo positivo já que a quantidade (tMovEstoque.dQuantidade) foi colocada com valor negativo.
'        lErro = Comando_ExecutarPos(alComando(12), "UPDATE MovimentoEstoqueES SET Custo = ?", alComando(2), -tMovEstoque.dCusto)
'        If lErro <> AD_SQL_SUCESSO Then gError 25269
'
'    End If
'
'    'Acumula valores de entrada e saída
'    If tTipoMovEst.sEntradaOuSaida = TIPOMOV_EST_ENTRADA Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorEntrada = tSldDiaEstAlm.dValorEntrada + tMovEstoque.dCusto
'        tSldDiaEstAlm.dValorEntCusto = tSldDiaEstAlm.dValorEntCusto + tMovEstoque.dCusto
'        tSldDiaEst.dValorEntrada = tSldDiaEst.dValorEntrada + tMovEstoque.dCusto
'        tSldDiaEst.dValorEntCusto = tSldDiaEst.dValorEntCusto + tMovEstoque.dCusto
'        objAlmoxInfo.dValorEntrada = objAlmoxInfo.dValorEntrada + tMovEstoque.dCusto
'        tSldMesEst.adValorEnt(iMes) = tSldMesEst.adValorEnt(iMes) + tMovEstoque.dCusto
'        tSldMesEst.adSaldoValorCusto(iMes) = tSldMesEst.adSaldoValorCusto(iMes) + tMovEstoque.dCusto
'
'    ElseIf tTipoMovEst.sEntradaOuSaida = TIPOMOV_EST_SAIDA Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorSaida = tSldDiaEstAlm.dValorSaida + tMovEstoque.dCusto
'        tSldDiaEstAlm.dValorSaiCusto = tSldDiaEstAlm.dValorSaiCusto + tMovEstoque.dCusto
'        tSldDiaEst.dValorSaida = tSldDiaEst.dValorSaida + tMovEstoque.dCusto
'        tSldDiaEst.dValorSaiCusto = tSldDiaEst.dValorSaiCusto + tMovEstoque.dCusto
'        objAlmoxInfo.dValorSaida = objAlmoxInfo.dValorSaida + tMovEstoque.dCusto
'        tSldMesEst.adValorSai(iMes) = tSldMesEst.adValorSai(iMes) + tMovEstoque.dCusto
'        tSldMesEst.adSaldoValorCusto(iMes) = tSldMesEst.adSaldoValorCusto(iMes) - tMovEstoque.dCusto
'
'    End If
'
'    'Acumula valores de consumo
'    If tTipoMovEst.iAtualizaConsumo = TIPOMOV_EST_ADICIONACONSUMO Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorCons = tSldDiaEstAlm.dValorCons + tMovEstoque.dCusto
'        tSldDiaEst.dValorCons = tSldDiaEst.dValorCons + tMovEstoque.dCusto
'        objAlmoxInfo.dValorCons = objAlmoxInfo.dValorCons + tMovEstoque.dCusto
'        tSldMesEst.adValorCons(iMes) = tSldMesEst.adValorCons(iMes) + tMovEstoque.dCusto
'
'    ElseIf tTipoMovEst.iAtualizaConsumo = TIPOMOV_EST_SUBTRAICONSUMO Then
'
'        'Acumula
'        tSldDiaEstAlm.dValorCons = tSldDiaEstAlm.dValorCons - tMovEstoque.dCusto
'        tSldDiaEst.dValorCons = tSldDiaEst.dValorCons - tMovEstoque.dCusto
'        objAlmoxInfo.dValorCons = objAlmoxInfo.dValorCons - tMovEstoque.dCusto
'        tSldMesEst.adValorCons(iMes) = tSldMesEst.adValorCons(iMes) - tMovEstoque.dCusto
'
'    End If
'
'    'Busca próximo movimento
'    lErro = Rotina_CMP_ProximoMovimento(sComandoSQL(), alComando(), tProduto, tTipoMovEst, tMovEstoque, tMovEstoque2, iFilialEmpresa)
'    If lErro = 25277 Then gError 25270 'Não tem + movimentos
'    If lErro <> SUCESSO And lErro <> 25277 Then gError 25291
'
'    Rotina_CMP_AtualizaCMP = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaCMP:
'
'    Rotina_CMP_AtualizaCMP = gErr
'
'    Select Case gErr
'
'       Case 25268, 25291  'Tratado na rotina chamada
'
'       Case 25269
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MOVESTOQUE", gErr, tMovEstoque.lNumIntDoc)
'
'       Case 25270    'Não tem mais registros de movimentos
'                     'Tratado na rotina chamadora
'       Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159518)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AtualizaCP(sComandoSQL() As String, alComando() As Long, dQuantEntApropCustoProd As Double, tProduto As typeProduto, dCPAtual As Double, tTipoMovEst As typeTipoMovEst, tMovEstoque As typeItemMovEstoque, tMovEstoque2 As typeItemMovEstoque, tSldMesEst As typeSldMesEst, tSldDiaEst As typeSldDiaEst, tSldDiaEstAlm As typeSldDiaEstAlm, objAlmoxInfo As ClassAlmoxInfo, iMes As Integer, iFilialEmpresa As Integer, tSldMesEst1 As typeSldMesEst1) As Long
''Atualiza CustoRealProdução nos movimento. Chamada em TRANSAÇÃO
'
'Dim lErro As Long
'Dim dFator As Double
'
'On Error GoTo Erro_Rotina_CMP_AtualizaCP
'
'    'Fator de conv. de UMs
'    lErro = CF("UM_Conversao",tProduto.iClasseUM, tMovEstoque.sSiglaUM, tProduto.sSiglaUMEstoque, dFator)
'    If lErro <> SUCESSO Then gError 25271
'
'    'Calcula custo do Movimento
'    tMovEstoque.dCusto = dCPAtual * tMovEstoque.dQuantidade * dFator
'
'    'se não for um estorno
'    If tTipoMovEst.iCodigoOrig = 0 Then
'
'        'Atualiza custo do Movimento
'        lErro = Comando_ExecutarPos(alComando(13), "UPDATE MovimentoEstoqueES SET Custo = ?", alComando(2), tMovEstoque.dCusto)
'        If lErro <> AD_SQL_SUCESSO Then gError 25272
'
'    Else
'
'        'Atualiza custo do Movimento
'        lErro = Comando_ExecutarPos(alComando(13), "UPDATE MovimentoEstoqueES SET Custo = ?", alComando(2), -tMovEstoque.dCusto)
'        If lErro <> AD_SQL_SUCESSO Then gError 25272
'
'    End If
'
'    'Não pode ser uma SAÍDA com apropr. CUSTO_REAL_PRODUCAO
'    If tTipoMovEst.sEntradaOuSaida = TIPOMOV_EST_SAIDA Then gError 25273
'
'    'Acumula valores de ENTRADA
'    tSldDiaEstAlm.dValorEntrada = tSldDiaEstAlm.dValorEntrada + tMovEstoque.dCusto
'    tSldDiaEstAlm.dValorEntCusto = tSldDiaEstAlm.dValorEntCusto + tMovEstoque.dCusto
'    tSldDiaEst.dValorEntrada = tSldDiaEst.dValorEntrada + tMovEstoque.dCusto
'    objAlmoxInfo.dValorEntrada = objAlmoxInfo.dValorEntrada + tMovEstoque.dCusto
'    tSldMesEst.adValorEnt(iMes) = tSldMesEst.adValorEnt(iMes) + tMovEstoque.dCusto
'    tSldMesEst.adSaldoValorCusto(iMes) = tSldMesEst.adSaldoValorCusto(iMes) + tMovEstoque.dCusto
'    dQuantEntApropCustoProd = dQuantEntApropCustoProd + tMovEstoque.dQuantidade * dFator
'    tSldDiaEst.dValorEntCusto = tSldDiaEst.dValorEntCusto + tMovEstoque.dCusto
'    objAlmoxInfo.dSaldoValorCusto = objAlmoxInfo.dSaldoValorCusto + tMovEstoque.dCusto
'
'    If tTipoMovEst.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
'
'        'trata produção de material para terceiros (Beneficiamento)
'        If tTipoMovEst.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
'            tSldMesEst1.adSaldoValorBenef3(iMes) = tSldMesEst1.adSaldoValorBenef3(iMes) + tMovEstoque.dCusto
'            tSldMesEst1.adSaldoQuantBenef3(iMes) = tSldMesEst1.adSaldoQuantBenef3(iMes) + tMovEstoque.dQuantidade
'            objAlmoxInfo.dSaldoValorBenef3 = objAlmoxInfo.dSaldoValorBenef3 + tMovEstoque.dCusto
'            tSldDiaEst.dValorEntBenef3 = tSldDiaEst.dValorEntBenef3 + tMovEstoque.dCusto
'            tSldDiaEstAlm.dValorEntBenef3 = tSldDiaEstAlm.dValorEntBenef3 + tMovEstoque.dCusto
'        End If
'
'    End If
'
'
'    'Busca próximo movimento
'    lErro = Rotina_CMP_ProximoMovimento(sComandoSQL(), alComando(), tProduto, tTipoMovEst, tMovEstoque, tMovEstoque2, iFilialEmpresa)
'    If lErro = 25277 Then gError 25274 'Não tem + movimentos
'    If lErro <> SUCESSO And lErro <> 25277 Then gError 25275
'
'    Rotina_CMP_AtualizaCP = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AtualizaCP:
'
'    Rotina_CMP_AtualizaCP = gErr
'
'    Select Case gErr
'
'       Case 25271, 25275  'Tratado na rotina chamada
'
'       Case 25273
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_SAIDA_APROPR_CRP", gErr, tMovEstoque.lNumIntDoc)
'
'       Case 25272
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MOVESTOQUE", gErr, tMovEstoque.lNumIntDoc)
'
'       Case 25274    'Não tem mais registros de movimentos
'                     'Tratado na rotina chamadora
'       Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159519)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Function Rotina_CMP_ProximoMovimento(sComandoSQL() As String, alComando() As Long, tProduto As typeProduto, tTipoMovEst As typeTipoMovEst, tMovEstoque As typeItemMovEstoque, tMovEstoque2 As typeItemMovEstoque, iFilialEmpresa As Integer) As Long
''Chamada EM TRANSAÇÃO
''Busca próximo registro nos 2 comandos de Movimentos da Rotina_CustoMedioProducao
'
'Dim lErro As Long
'Dim objTipoMovEstoque As New ClassTipoMovEst
'
'On Error GoTo Erro_Rotina_CMP_ProximoMovimento
'
'    'Busca próximo registro de MovimentoEstoque vinculado a TipoMovimento e Produto
'    '------------------------------------------------------------------------------
'    lErro = Comando_BuscarProximo(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25276
'    If lErro = AD_SQL_SEM_DADOS Then gError 25277
'
'    If tTipoMovEst.iCodigoOrig <> 0 Then
'
'        objTipoMovEstoque.iCodigo = tTipoMovEst.iCodigo
'
'        'ler os dados referentes ao tipo de movimento
'        lErro = CF("TiposMovEst_Le1",alComando(21), objTipoMovEstoque)
'        If lErro <> SUCESSO Then gError 89158
'
'        tTipoMovEst.sEntradaOuSaida = objTipoMovEstoque.sEntradaOuSaida
'        tTipoMovEst.iAtualizaConsumo = objTipoMovEstoque.iAtualizaConsumo
'        tTipoMovEst.iAtualizaVenda = objTipoMovEstoque.iAtualizaVenda
'        tTipoMovEst.iAtualizaConsig = objTipoMovEstoque.iAtualizaConsig
'        tTipoMovEst.iAtualizaDemo = objTipoMovEstoque.iAtualizaDemo
'        tTipoMovEst.iAtualizaConserto = objTipoMovEstoque.iAtualizaConserto
'        tTipoMovEst.iAtualizaOutras = objTipoMovEstoque.iAtualizaOutras
'        tTipoMovEst.iAtualizaBenef = objTipoMovEstoque.iAtualizaBenef
'        tTipoMovEst.iCustoMedio = objTipoMovEstoque.iCustoMedio
'
'        tMovEstoque.dQuantidade = -tMovEstoque.dQuantidade
'
'    End If
'
'    'Busca próximo registro de MovimentoEstoque para atualizar custo
'    '------------------------------------------------------------------------------
'    lErro = Comando_BuscarProximo(alComando(2))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25278
'    If lErro = AD_SQL_SEM_DADOS Then gError 25279
'    If tMovEstoque.lNumIntDoc <> tMovEstoque2.lNumIntDoc Then gError 25280
'
'    Rotina_CMP_ProximoMovimento = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_ProximoMovimento:
'
'    Rotina_CMP_ProximoMovimento = gErr
'
'    Select Case gErr
'
'        Case 25276, 25278
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, iFilialEmpresa)
'
'        Case 25277 'Não tem + registro. Será tratado na rotina chamadora.
'
'        Case 25279, 25280
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_ITEMMOVEST_MOVTOESTOQUE", gErr, tMovEstoque.lNumIntDoc)
'
'        Case 89158
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159520)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_TotalMovEstoque(iFilialEmpresa As Integer, iAno As Integer, iMes As Integer, lTotalProdutos As Long) As Long
''Chamada EM TRANSAÇÃO
''retorna o número de produtos que participam do processamento de custo médio de produção
'
'Dim lComando As Long
'Dim lErro As Long
'Dim dtDataInicial As Date
'Dim dtDataFinal As Date
'Dim iDiasMes As Integer
'Dim sComandoSQL As String
'Dim sProduto As String
'
'On Error GoTo Erro_Rotina_CMP_TotalMovEstoque
'
'    'Abre comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 25281
'
'    'Determinação de faixa de datas
'    dtDataInicial = CDate("1/" & CStr(iMes) & "/" & CStr(iAno))
'    iDiasMes = Dias_Mes(iMes, iAno)
'    dtDataFinal = CDate(CStr(iDiasMes) & "/" & CStr(iMes) & "/" & CStr(iAno))
'
'    sComandoSQL = "SELECT DISTINCT Produto FROM MovimentoEstoque WHERE MovimentoEstoque.FilialEmpresa = ? AND Data >= ? AND Data <= ? AND (Apropriacao = ? OR Apropriacao = ?)"
'
'    sProduto = String(STRING_PRODUTO, 0)
'    lErro = Comando_Executar(lComando, sComandoSQL, sProduto, iFilialEmpresa, dtDataInicial, dtDataFinal, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO)
'    If lErro <> AD_SQL_SUCESSO Then gError 25282
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25283
'
'    lTotalProdutos = 0
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        lTotalProdutos = lTotalProdutos + 1
'
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25779
'
'    Loop
'
'    'Fechamento comando
'    Call Comando_Fechar(lComando)
'
'    Rotina_CMP_TotalMovEstoque = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_TotalMovEstoque:
'
'    Rotina_CMP_TotalMovEstoque = gErr
'
'    Select Case gErr
'
'        Case 25281
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 25282
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL)
'
'        Case 25283, 25779
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, iFilialEmpresa)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159521)
'
'    End Select
'
'   'Fechamento comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_EstoqueMes_CriticaLock(lComando As Long, iFilialEmpresa As Integer, iAno As Integer, iMes As Integer) As Long
''Verifica se EstoqueMes tem Custo Medio Produção não apurado e
''está fechado. Faz lock exclusive. Chamada EM TRANSAÇÃO.
'
'Dim lErro As Long
'Dim iFechamento As Integer
'Dim iCustoProdApurado As Integer
'
'On Error GoTo Erro_Rotina_CMP_EstoqueMes_CriticaLock
'
'    'Seleciona EstoqueMes
'    lErro = Comando_ExecutarPos(lComando, "SELECT CustoProdApurado, Fechamento FROM EstoqueMes WHERE FilialEmpresa=? AND Ano=? AND Mes=?", 0, iCustoProdApurado, iFechamento, iFilialEmpresa, iAno, iMes)
'    If lErro <> AD_SQL_SUCESSO Then gError 25285
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 25286
'
'    If lErro = AD_SQL_SEM_DADOS Then gError 25287 'não encontrou EstoqueMes
'
'    'Faz lock em EstoqueMes
'    lErro = Comando_LockExclusive(lComando)
'    If lErro <> AD_SQL_SUCESSO Then gError 25288
'
'    'Verifica se EstoqueMes está fechado
'    If iFechamento = ESTOQUEMES_FECHAMENTO_ABERTO Then gError 25289
'
'    'Verifica se EstoqueMes teve CMP apurado
'    If iCustoProdApurado = CUSTO_APURADO Then gError 25290
'
'    Rotina_CMP_EstoqueMes_CriticaLock = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_EstoqueMes_CriticaLock:
'
'    Rotina_CMP_EstoqueMes_CriticaLock = gErr
'
'    Select Case gErr
'
'        Case 25285
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, "SELECT Fechamento FROM EstoqueMes WHERE FilialEmpresa=?, Ano=?, Mes=?")
'
'        Case 25286
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES", gErr, iFilialEmpresa, iAno, iMes)
'
'        Case 25287, 25289, 25290
'
'        Case 25288
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ESTOQUEMES", gErr, iFilialEmpresa, iAno, iMes)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159522)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_TestaAlmoxarifado(tMovEstoque As typeItemMovEstoque, colAlmoxInfo As Collection, objAlmoxInfo As ClassAlmoxInfo) As Long
'
'Dim bEncontrou As Boolean
'
'On Error GoTo Erro_Rotina_CMP_TestaAlmoxarifado
'
'    bEncontrou = False
'
'    For Each objAlmoxInfo In colAlmoxInfo
'        If objAlmoxInfo.iAlmoxarifado = tMovEstoque.iAlmoxarifado Then
'            bEncontrou = True
'            Exit For
'        End If
'    Next
'
'    If Not bEncontrou Then
'        Set objAlmoxInfo = New ClassAlmoxInfo
'        objAlmoxInfo.iAlmoxarifado = tMovEstoque.iAlmoxarifado
'        colAlmoxInfo.Add objAlmoxInfo
'    End If
'
'    Rotina_CMP_TestaAlmoxarifado = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_TestaAlmoxarifado:
'
'    Rotina_CMP_TestaAlmoxarifado = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159523)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_AlteraCMP(tSldMesEst As typeSldMesEst, iMes As Integer, dCMPAtual As Double) As Long
''Atualiza o campo ("CustoMedio" & iMes)  para o produto e o imes passado
'
'Dim alComando(1 To 2) As Long
'Dim iIndice As Integer
'Dim dCusto As Double
'Dim lErro As Long
'Dim sComandoSQL As String
'
'On Error GoTo Erro_Rotina_CMP_AlteraCMP
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 62552
'    Next
'
'    sComandoSQL = "SELECT CustoMedio" & iMes & " FROM SldMesEst WHERE FilialEmpresa =? AND Ano =? AND Produto =? "
'
'    lErro = Comando_ExecutarPos(alComando(1), sComandoSQL, 0, dCusto, tSldMesEst.iFilialEmpresa, tSldMesEst.iAno, tSldMesEst.sProduto)
'    If lErro <> AD_SQL_SUCESSO Then gError 62553
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62554
'    If lErro = AD_SQL_SEM_DADOS Then gError 62555
'
'    sComandoSQL = "UPDATE SldMesEst SET CustoMedio" & iMes & " = ?"
'
'    lErro = Comando_ExecutarPos(alComando(2), sComandoSQL, alComando(1), dCMPAtual)
'    If lErro <> AD_SQL_SUCESSO Then gError 62556
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Rotina_CMP_AlteraCMP = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_AlteraCMP:
'
'    Rotina_CMP_AlteraCMP = gErr
'
'    Select Case gErr
'
'        Case 62552
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 62553, 62554
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, tSldMesEst.iAno, tSldMesEst.iFilialEmpresa, tSldMesEst.sProduto)
'
'        Case 62555
'            Call Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_INEXISTENTE", gErr, tSldMesEst.iAno, tSldMesEst.iFilialEmpresa, tSldMesEst.sProduto)
'
'        Case 62556
'            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESEST", gErr, tSldMesEst.iAno, tSldMesEst.iFilialEmpresa, tSldMesEst.sProduto)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159524)
'
'    End Select
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_Transfere_SldMesEst_ValorInicial(iAno As Integer, iFilialEmpresa As Integer) As Long
''Atualiza Valor Inicial e o CustoMedioProducaoInicial para o Ano sequinte na tabela SaldoMesEst
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim dValorInicialProxAno As Double
'Dim dQuantInicialProxAno As Double
'Dim dCustoMedioProducaoInicial As Double
'Dim iMesFinal As Integer
'Dim iIndice As Integer
'Dim sComandoSQL(1 To 3) As String
'Dim alComando(1 To 3) As Long
'Dim tSldMesEst As typeSldMesEst
'
'On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEst_ValorInicial
'
'    'Abre comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 69005
'    Next
'
'    'Le os Valores de Entrada e de Saida para todos os meses
'    sComandoSQL(1) = "SELECT CustoMedioProducaoInicial, "
'
'    'Quantidade e valor inicial
'    sComandoSQL(1) = sComandoSQL(1) & "QuantInicial, ValorInicial, QuantInicialCusto, ValorInicialCusto, "
'
'    'Quantidades e valores de entrada e de saida mensais
'    For iIndice = 1 To 12
'        sComandoSQL(1) = sComandoSQL(1) & "QuantEnt" & CStr(iIndice) & ", " & "QuantSai" & CStr(iIndice) & ", " & "ValorEnt" & CStr(iIndice) & ", " & "ValorSai" & CStr(iIndice) & ", "
'    Next
'
'    For iIndice = 1 To 12
'        sComandoSQL(1) = sComandoSQL(1) & "SaldoQuantCusto" & CStr(iIndice) & ", " & "SaldoValorCusto" & CStr(iIndice) & ", "
'    Next
'
'    sComandoSQL(1) = sComandoSQL(1) & "Produto "
'    'Tabela, Filtro, Ordem
'    sComandoSQL(1) = sComandoSQL(1) & "FROM Produtos, SldMesEst WHERE Produtos.Codigo = SldMesEst.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? AND FilialEmpresa = ? ORDER BY Produtos.Codigo"
'
'
'    With tSldMesEst
'
'        .sProduto = String(STRING_PRODUTO, 0)
'
'        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .dCustoMedioProducaoInicial, .dQuantInicial, .dValorInicial, .dQuantInicialCusto, .dValorInicialCusto, .adQuantEnt(1), .adQuantSai(1), .adValorEnt(1), .adValorSai(1), .adQuantEnt(2), .adQuantSai(2), .adValorEnt(2), .adValorSai(2), .adQuantEnt(3), .adQuantSai(3), .adValorEnt(3), .adValorSai(3), .adQuantEnt(4), .adQuantSai(4), .adValorEnt(4), .adValorSai(4), .adQuantEnt(5), .adQuantSai(5), .adValorEnt(5), .adValorSai(5), .adQuantEnt(6), .adQuantSai(6), .adValorEnt(6), .adValorSai(6), _
'                .adQuantEnt(7), .adQuantSai(7), .adValorEnt(7), .adValorSai(7), .adQuantEnt(8), .adQuantSai(8), .adValorEnt(8), .adValorSai(8), .adQuantEnt(9), .adQuantSai(9), .adValorEnt(9), .adValorSai(9), .adQuantEnt(10), .adQuantSai(10), .adValorEnt(10), .adValorSai(10), .adQuantEnt(11), .adQuantSai(11), .adValorEnt(11), .adValorSai(11), .adQuantEnt(12), .adQuantSai(12), .adValorEnt(12), .adValorSai(12), _
'                .adSaldoQuantCusto(1), .adSaldoValorCusto(1), .adSaldoQuantCusto(2), .adSaldoValorCusto(2), .adSaldoQuantCusto(3), .adSaldoValorCusto(3), .adSaldoQuantCusto(4), .adSaldoValorCusto(4), .adSaldoQuantCusto(5), .adSaldoValorCusto(5), .adSaldoQuantCusto(6), .adSaldoValorCusto(6), .adSaldoQuantCusto(7), .adSaldoValorCusto(7), .adSaldoQuantCusto(8), .adSaldoValorCusto(8), .adSaldoQuantCusto(9), .adSaldoValorCusto(9), .adSaldoQuantCusto(10), .adSaldoValorCusto(10), .adSaldoQuantCusto(11), .adSaldoValorCusto(11), .adSaldoQuantCusto(12), .adSaldoValorCusto(12), _
'                .sProduto, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno, iFilialEmpresa)
'
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 69006
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69007
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        'Quantitade Inicial
'        sComandoSQL(2) = "SELECT QuantInicial "
'        'Tabela, Filtro, Ordem
'        sComandoSQL(2) = sComandoSQL(2) & "FROM SldMesEst WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? ORDER BY Produto"
'
'        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
'        '----------------------------------------------------------------------------------------
'        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, dQuantInicialProxAno, iAno + 1, iFilialEmpresa, tSldMesEst.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 69008
'
'        lErro = Comando_BuscarPrimeiro(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69009
'
'        'Se não encontrou
'        If lErro = AD_SQL_SEM_DADOS Then gError 69010
'
'        'Cálculo de saldo de quantidade
'        '----------------------------------------------------------------
'
'        'Quantidade no início do ano
'        dQuantInicialProxAno = tSldMesEst.dQuantInicial
'
'        'Acumula entrada e saída dos meses
'        For iIndice = 1 To 12
'            dQuantInicialProxAno = dQuantInicialProxAno + tSldMesEst.adQuantEnt(iIndice) - tSldMesEst.adQuantSai(iIndice)
'        Next
'
'        'Cálculo de saldo de valor
'        '-------------------------
'        'Valor inicial
'        dValorInicialProxAno = tSldMesEst.dValorInicial
'
'        'Meses
'        For iIndice = 1 To 12
'            tSldMesEst.dValorInicialCusto = tSldMesEst.dValorInicialCusto + tSldMesEst.adSaldoValorCusto(iIndice)
'        Next
'
'        'Meses
'        For iIndice = 1 To 12
'            dValorInicialProxAno = dValorInicialProxAno + tSldMesEst.adValorEnt(iIndice) - tSldMesEst.adValorSai(iIndice)
'        Next
'
'        If dQuantInicialProxAno > 0 Then
'
'            'Calcula CustoMedioProducaoAtual
'            '-------------------------------
'            dCustoMedioProducaoInicial = dValorInicialProxAno / dQuantInicialProxAno
'
'        ElseIf dQuantInicialProxAno = 0 Then
'
'            'Procura o último mês em que houve saída (apropr=CMP)
'            For iMesFinal = 12 To 1 Step -1
'
'                If tSldMesEst.adQuantSai(iMesFinal) > 0 Then Exit For
'
'            Next
'
'            If tSldMesEst.adQuantSai(iMesFinal) > 0 Then
'
'                'Custo Médio é o valor da última saída mensal dividido pela quantidade
'                dCustoMedioProducaoInicial = tSldMesEst.adValorSai(iMesFinal) / tSldMesEst.adQuantSai(iMesFinal)
'
'            Else 'Todas as quantidades de saída e de entrada do Produto estão zeradas nesse ano
'
'                dCustoMedioProducaoInicial = tSldMesEst.dCustoMedioProducaoInicial
'
'            End If
'
'        End If
'
'        '------------Atualiza SldMesEst para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
'
'        'Monta comando SQL para UPDATE de SldMesEst
'        sComandoSQL(3) = "UPDATE SldMesEst SET ValorInicial = ?, ValorInicialCusto = ?, CustoMedioProducaoInicial  = ?"
'
'        'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
'        lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorInicialProxAno, tSldMesEst.dValorInicialCusto, dCustoMedioProducaoInicial)
'        If lErro <> AD_SQL_SUCESSO Then gError 69011
'
'        'Atualiza tela de acompanhamento do Batch
'        lErro = Rotina_CMP_AtualizaTelaBatch()
'        If lErro <> SUCESSO Then gError 69017
'
'        '-----------------Busca o Proximo -------------------
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69012
'
'    Loop
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Rotina_CMP_Transfere_SldMesEst_ValorInicial = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_Transfere_SldMesEst_ValorInicial:
'
'    Rotina_CMP_Transfere_SldMesEst_ValorInicial = gErr
'
'    Select Case gErr
'
'        Case 69005
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 69006, 69007, 69012
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, iFilialEmpresa, iAno)
'
'        Case 69008, 69009
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, iFilialEmpresa, iAno + 1)
'
'        Case 69010
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SALDOMESEST", gErr, iFilialEmpresa, iAno + 1, tSldMesEst.sProduto)
'
'        Case 69011
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESEST", gErr, iAno, iFilialEmpresa, tSldMesEst.sProduto)
'
'        Case 69017
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159525)
'
'    End Select
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_Transfere_SldMesEst1_ValorInicial(iAno As Integer, iFilialEmpresa As Integer) As Long
''Atualiza Valor Inicial para o Ano sequinte na tabela SaldoMesEst1
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim iMesFinal As Integer
'Dim iIndice As Integer
'Dim sComandoSQL(1 To 3) As String
'Dim alComando(1 To 3) As Long
'Dim tSldMesEst1 As typeSldMesEst1
'Dim dValorAcumuladaConsig3 As Double, dValorAcumuladaDemo3 As Double, dValorAcumuladaConserto3 As Double, dValorAcumuladaOutras3 As Double, dValorAcumuladaBenef3 As Double
'
'On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEst1_ValorInicial
'
'    'Abre comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 89833
'    Next
'
'    'Le os Valores de Entrada e de Saida para todos os meses
'    'Quantidade e valor inicial
'    sComandoSQL(1) = "SELECT ValorInicialConsig3, ValorInicialDemo3, ValorInicialConserto3, ValorInicialOutros3, ValorInicialBenef3, "
'    'Quantidades e valores de entrada e de saida mensais
'    For iIndice = 1 To 12
'        sComandoSQL(1) = sComandoSQL(1) & "SaldoValorConsig3" & CStr(iIndice) & ", " & "SaldoValorDemo3" & CStr(iIndice) & ", " & "SaldoValorConserto3" & CStr(iIndice) & ", " & "SaldoValorOutros3" & CStr(iIndice) & ", " & "SaldoValorBenef3" & CStr(iIndice) & ", "
'    Next
'
'    sComandoSQL(1) = sComandoSQL(1) & "Produto "
'
'    'Tabela, Filtro, Ordem
'    sComandoSQL(1) = sComandoSQL(1) & "FROM Produtos, SldMesEst1 WHERE Produtos.Codigo = SldMesEst1.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? AND FilialEmpresa = ? ORDER BY Produtos.Codigo"
'
'
'    With tSldMesEst1
'        .sProduto = String(STRING_PRODUTO, 0)
'
'        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .dValorInicialConsig3, .dValorInicialDemo3, .dValorInicialConserto3, .dValorInicialOutros3, .dValorInicialBenef3, .adSaldoValorConsig3(1), .adSaldoValorDemo3(1), .adSaldoValorConserto3(1), .adSaldoValorOutros3(1), .adSaldoValorBenef3(1), .adSaldoValorConsig3(2), .adSaldoValorDemo3(2), .adSaldoValorConserto3(2), .adSaldoValorOutros3(2), .adSaldoValorBenef3(2), .adSaldoValorConsig3(3), .adSaldoValorDemo3(3), .adSaldoValorConserto3(3), .adSaldoValorOutros3(3), .adSaldoValorBenef3(3), .adSaldoValorConsig3(4), .adSaldoValorDemo3(4), .adSaldoValorConserto3(4), .adSaldoValorOutros3(4), .adSaldoValorBenef3(4), .adSaldoValorConsig3(5), .adSaldoValorDemo3(5), .adSaldoValorConserto3(5), .adSaldoValorOutros3(5), .adSaldoValorBenef3(5), .adSaldoValorConsig3(6), .adSaldoValorDemo3(6), .adSaldoValorConserto3(6), .adSaldoValorOutros3(6), .adSaldoValorBenef3(6), _
'        .adSaldoValorConsig3(7), .adSaldoValorDemo3(7), .adSaldoValorConserto3(7), .adSaldoValorOutros3(7), .adSaldoValorBenef3(7), .adSaldoValorConsig3(8), .adSaldoValorDemo3(8), .adSaldoValorConserto3(8), .adSaldoValorOutros3(8), .adSaldoValorBenef3(8), .adSaldoValorConsig3(9), .adSaldoValorDemo3(9), .adSaldoValorConserto3(9), .adSaldoValorOutros3(9), .adSaldoValorBenef3(9), .adSaldoValorConsig3(10), .adSaldoValorDemo3(10), .adSaldoValorConserto3(10), .adSaldoValorOutros3(10), .adSaldoValorBenef3(10), .adSaldoValorConsig3(11), .adSaldoValorDemo3(11), .adSaldoValorConserto3(11), .adSaldoValorOutros3(11), .adSaldoValorBenef3(11), .adSaldoValorConsig3(12), .adSaldoValorDemo3(12), .adSaldoValorConserto3(12), .adSaldoValorOutros3(12), .adSaldoValorBenef3(12), .sProduto, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno, iFilialEmpresa)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 89834
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89835
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        'Quantitade Inicial
'        sComandoSQL(2) = "SELECT FilialEmpresa "
'        'Tabela, Filtro, Ordem
'        sComandoSQL(2) = sComandoSQL(2) & "FROM SldMesEst1 WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? ORDER BY Produto"
'
'        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
'        '----------------------------------------------------------------------------------------
'        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, iFilialEmpresa, iAno + 1, iFilialEmpresa, tSldMesEst1.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 89836
'
'        lErro = Comando_BuscarPrimeiro(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89837
'
'        'Se não encontrou
'        If lErro = AD_SQL_SEM_DADOS Then gError 89838
'
'        'Meses
'        For iIndice = 1 To 12
'            dValorAcumuladaBenef3 = dValorAcumuladaBenef3 + tSldMesEst1.adSaldoValorBenef3(iIndice)
'            dValorAcumuladaConserto3 = dValorAcumuladaConserto3 + tSldMesEst1.adSaldoValorConserto3(iIndice)
'            dValorAcumuladaConsig3 = dValorAcumuladaConsig3 + tSldMesEst1.adSaldoValorConsig3(iIndice)
'            dValorAcumuladaDemo3 = dValorAcumuladaDemo3 + tSldMesEst1.adSaldoValorDemo3(iIndice)
'            dValorAcumuladaOutras3 = dValorAcumuladaOutras3 + tSldMesEst1.adSaldoValorOutros3(iIndice)
'        Next
'
'        '------------Atualiza SldMesEst para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
'
'        'Monta comando SQL para UPDATE de SldMesEst
'        sComandoSQL(3) = "UPDATE SldMesEst1 SET ValorInicialConsig3 = ?, ValorInicialDemo3 = ?, ValorInicialConserto3 = ?, ValorInicialOutros3 = ?, ValorInicialBenef3 = ?"
'
'        'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
'        lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorAcumuladaConsig3, dValorAcumuladaDemo3, dValorAcumuladaConserto3, dValorAcumuladaOutras3, dValorAcumuladaBenef3)
'        If lErro <> AD_SQL_SUCESSO Then gError 89839
'
'        'Atualiza tela de acompanhamento do Batch
'        lErro = Rotina_CMP_AtualizaTelaBatch()
'        If lErro <> SUCESSO Then gError 89840
'
'        '-----------------Busca o Proximo -------------------
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89841
'
'    Loop
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Rotina_CMP_Transfere_SldMesEst1_ValorInicial = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_Transfere_SldMesEst1_ValorInicial:
'
'    Rotina_CMP_Transfere_SldMesEst1_ValorInicial = gErr
'
'    Select Case gErr
'
'        Case 89833
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 89834, 89835, 89841
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST1_2", gErr, iAno, iFilialEmpresa)
'
'        Case 89836, 89837
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, iAno + 1, iFilialEmpresa, tSldMesEst1.sProduto)
'
'        Case 89838
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDOMESEST1_NAO_CADASTRADO", gErr, iFilialEmpresa, iAno + 1, tSldMesEst1.sProduto)
'
'        Case 89839
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESEST1", gErr, iAno + 1, iFilialEmpresa, tSldMesEst1.sProduto)
'
'        Case 89840
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159526)
'
'    End Select
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_Transfere_SldMesEst2_ValorInicial(iAno As Integer, iFilialEmpresa As Integer) As Long
''Atualiza Valor Inicial para o Ano sequinte na tabela SaldoMesEst2
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim iMesFinal As Integer
'Dim iIndice As Integer
'Dim sComandoSQL(1 To 3) As String
'Dim alComando(1 To 3) As Long
'Dim tSldMesEst2 As typeSldMesEst2
'Dim dValorAcumuladaConsig As Double, dValorAcumuladaDemo As Double, dValorAcumuladaConserto As Double, dValorAcumuladaOutras As Double, dValorAcumuladaBenef As Double
'
'On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEst2_ValorInicial
'
'    'Abre comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 69740
'    Next
'
'    'Le os Valores de Entrada e de Saida para todos os meses
'    'Quantidade e valor inicial
'    sComandoSQL(1) = "SELECT ValorInicialConsig, ValorInicialDemo, ValorInicialConserto, ValorInicialOutros, ValorInicialBenef, "
'    'Quantidades e valores de entrada e de saida mensais
'    For iIndice = 1 To 12
'        sComandoSQL(1) = sComandoSQL(1) & "SaldoValorConsig" & CStr(iIndice) & ", " & "SaldoValorDemo" & CStr(iIndice) & ", " & "SaldoValorConserto" & CStr(iIndice) & ", " & "SaldoValorOutros" & CStr(iIndice) & ", " & "SaldoValorBenef" & CStr(iIndice) & ", "
'    Next
'    sComandoSQL(1) = sComandoSQL(1) & "Produto "
'    'Tabela, Filtro, Ordem
'    sComandoSQL(1) = sComandoSQL(1) & "FROM Produtos, SldMesEst2 WHERE Produtos.Codigo = SldMesEst2.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? AND FilialEmpresa = ? ORDER BY Produtos.Codigo"
'
'
'    With tSldMesEst2
'        .sProduto = String(STRING_PRODUTO, 0)
'
'        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .dValorInicialConsig, .dValorInicialDemo, .dValorInicialConserto, .dValorInicialOutros, .dValorInicialBenef, .adSaldoValorConsig(1), .adSaldoValorDemo(1), .adSaldoValorConserto(1), .adSaldoValorOutros(1), .adSaldoValorBenef(1), .adSaldoValorConsig(2), .adSaldoValorDemo(2), .adSaldoValorConserto(2), .adSaldoValorOutros(2), .adSaldoValorBenef(2), .adSaldoValorConsig(3), .adSaldoValorDemo(3), .adSaldoValorConserto(3), .adSaldoValorOutros(3), .adSaldoValorBenef(3), .adSaldoValorConsig(4), .adSaldoValorDemo(4), .adSaldoValorConserto(4), .adSaldoValorOutros(4), .adSaldoValorBenef(4), .adSaldoValorConsig(5), .adSaldoValorDemo(5), .adSaldoValorConserto(5), .adSaldoValorOutros(5), .adSaldoValorBenef(5), .adSaldoValorConsig(6), .adSaldoValorDemo(6), .adSaldoValorConserto(6), .adSaldoValorOutros(6), .adSaldoValorBenef(6), _
'        .adSaldoValorConsig(7), .adSaldoValorDemo(7), .adSaldoValorConserto(7), .adSaldoValorOutros(7), .adSaldoValorBenef(7), .adSaldoValorConsig(8), .adSaldoValorDemo(8), .adSaldoValorConserto(8), .adSaldoValorOutros(8), .adSaldoValorBenef(8), .adSaldoValorConsig(9), .adSaldoValorDemo(9), .adSaldoValorConserto(9), .adSaldoValorOutros(9), .adSaldoValorBenef(9), .adSaldoValorConsig(10), .adSaldoValorDemo(10), .adSaldoValorConserto(10), .adSaldoValorOutros(10), .adSaldoValorBenef(10), .adSaldoValorConsig(11), .adSaldoValorDemo(11), .adSaldoValorConserto(11), .adSaldoValorOutros(11), .adSaldoValorBenef(11), .adSaldoValorConsig(12), .adSaldoValorDemo(12), .adSaldoValorConserto(12), .adSaldoValorOutros(12), .adSaldoValorBenef(12), .sProduto, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno, iFilialEmpresa)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 69741
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69742
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        'Quantitade Inicial
'        sComandoSQL(2) = "SELECT FilialEmpresa "
'        'Tabela, Filtro, Ordem
'        sComandoSQL(2) = sComandoSQL(2) & "FROM SldMesEst2 WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? ORDER BY Produto"
'
'        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
'        '----------------------------------------------------------------------------------------
'        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, iFilialEmpresa, iAno + 1, iFilialEmpresa, tSldMesEst2.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 69743
'
'        lErro = Comando_BuscarPrimeiro(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69744
'
'        'Se não encontrou
'        If lErro = AD_SQL_SEM_DADOS Then gError 69745
'
'        'Meses
'        For iIndice = 1 To 12
'            dValorAcumuladaBenef = dValorAcumuladaBenef + tSldMesEst2.adSaldoValorBenef(iIndice)
'            dValorAcumuladaConserto = dValorAcumuladaConserto + tSldMesEst2.adSaldoValorConserto(iIndice)
'            dValorAcumuladaConsig = dValorAcumuladaConsig + tSldMesEst2.adSaldoValorConsig(iIndice)
'            dValorAcumuladaDemo = dValorAcumuladaDemo + tSldMesEst2.adSaldoValorDemo(iIndice)
'            dValorAcumuladaOutras = dValorAcumuladaOutras + tSldMesEst2.adSaldoValorOutros(iIndice)
'        Next
'
'        '------------Atualiza SldMesEst para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
'
'        'Monta comando SQL para UPDATE de SldMesEst
'        sComandoSQL(3) = "UPDATE SldMesEst2 SET ValorInicialConsig = ?, ValorInicialDemo = ?, ValorInicialConserto = ?, ValorInicialOutros = ?, ValorInicialBenef = ?"
'
'        'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
'        lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorAcumuladaConsig, dValorAcumuladaDemo, dValorAcumuladaConserto, dValorAcumuladaOutras, dValorAcumuladaBenef)
'        If lErro <> AD_SQL_SUCESSO Then gError 69746
'
'        'Atualiza tela de acompanhamento do Batch
'        lErro = Rotina_CMP_AtualizaTelaBatch()
'        If lErro <> SUCESSO Then gError 69747
'
'        '-----------------Busca o Proximo -------------------
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69748
'
'    Loop
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Rotina_CMP_Transfere_SldMesEst2_ValorInicial = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_Transfere_SldMesEst2_ValorInicial:
'
'    Rotina_CMP_Transfere_SldMesEst2_ValorInicial = gErr
'
'    Select Case gErr
'
'        Case 69740
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 69741, 69742
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST2_2", gErr, iAno, iFilialEmpresa)
'
'        Case 69743, 69744, 69748
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno + 1, iFilialEmpresa, tSldMesEst2.sProduto)
'
'        Case 69745
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDOMESEST2_NAO_CADASTRADO", gErr, iFilialEmpresa, iAno + 1, tSldMesEst2.sProduto)
'
'        Case 69746
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESEST1", gErr, iAno + 1, iFilialEmpresa, tSldMesEst2.sProduto)
'
'        Case 69747
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159527)
'
'    End Select
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial(iAno As Integer, iFilialEmpresa As Integer) As Long
''Atualiza Valor Inicial para o Ano sequinte na tabela SaldoMesEstAlm2
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim iMesFinal As Integer
'Dim iIndice As Integer
'Dim sComandoSQL(1 To 3) As String
'Dim alComando(1 To 3) As Long
'Dim tSldMesEstAlm1 As typeSldMesEstAlm1
'Dim iAlmoxarifado As Integer
'Dim dValorAcumuladaConsig3 As Double, dValorAcumuladaDemo3 As Double, dValorAcumuladaConserto3 As Double, dValorAcumuladaOutras3 As Double, dValorAcumuladaBenef3 As Double
'
'On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial
'
'    'Abre comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 89843
'    Next
'
'    'Le os Valores de Entrada e de Saida para todos os meses
'    'Quantidade e valor inicial
'    sComandoSQL(1) = "SELECT ValorInicialConsig3, ValorInicialDemo3, ValorInicialConserto3, ValorInicialOutros3, ValorInicialBenef3, "
'    'Quantidades e valores de entrada e de saida mensais
'    For iIndice = 1 To 12
'        sComandoSQL(1) = sComandoSQL(1) & "SaldoValorConsig3" & CStr(iIndice) & ", " & "SaldoValorDemo3" & CStr(iIndice) & ", " & "SaldoValorConserto3" & CStr(iIndice) & ", " & "SaldoValorOutros3" & CStr(iIndice) & ", " & "SaldoValorBenef3" & CStr(iIndice) & ", "
'    Next
'    sComandoSQL(1) = sComandoSQL(1) & " Produto, "
'    sComandoSQL(1) = sComandoSQL(1) & " Almoxarifado "
'    'Tabela, Filtro, Ordem
'    sComandoSQL(1) = sComandoSQL(1) & " FROM Produtos, SldMesEstAlm1, Almoxarifado WHERE Produtos.Codigo = SldMesEstAlm1.Produto AND Almoxarifado.Codigo = SldMesEstAlm1.Almoxarifado AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? AND Almoxarifado.FilialEmpresa = ? ORDER BY Produtos.Codigo"
'
'    With tSldMesEstAlm1
'        .sProduto = String(STRING_PRODUTO, 0)
'
'        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .dValorInicialConsig3, .dValorInicialDemo3, .dValorInicialConserto3, .dValorInicialOutros3, .dValorInicialBenef3, .adSaldoValorConsig3(1), .adSaldoValorDemo3(1), .adSaldoValorConserto3(1), .adSaldoValorOutros3(1), .adSaldoValorBenef3(1), .adSaldoValorConsig3(2), .adSaldoValorDemo3(2), .adSaldoValorConserto3(2), .adSaldoValorOutros3(2), .adSaldoValorBenef3(2), .adSaldoValorConsig3(3), .adSaldoValorDemo3(3), .adSaldoValorConserto3(3), .adSaldoValorOutros3(3), .adSaldoValorBenef3(3), .adSaldoValorConsig3(4), .adSaldoValorDemo3(4), .adSaldoValorConserto3(4), .adSaldoValorOutros3(4), .adSaldoValorBenef3(4), .adSaldoValorConsig3(5), .adSaldoValorDemo3(5), .adSaldoValorConserto3(5), .adSaldoValorOutros3(5), .adSaldoValorBenef3(5), .adSaldoValorConsig3(6), .adSaldoValorDemo3(6), .adSaldoValorConserto3(6), .adSaldoValorOutros3(6), .adSaldoValorBenef3(6), _
'        .adSaldoValorConsig3(7), .adSaldoValorDemo3(7), .adSaldoValorConserto3(7), .adSaldoValorOutros3(7), .adSaldoValorBenef3(7), .adSaldoValorConsig3(8), .adSaldoValorDemo3(8), .adSaldoValorConserto3(8), .adSaldoValorOutros3(8), .adSaldoValorBenef3(8), .adSaldoValorConsig3(9), .adSaldoValorDemo3(9), .adSaldoValorConserto3(9), .adSaldoValorOutros3(9), .adSaldoValorBenef3(9), .adSaldoValorConsig3(10), .adSaldoValorDemo3(10), .adSaldoValorConserto3(10), .adSaldoValorOutros3(10), .adSaldoValorBenef3(10), .adSaldoValorConsig3(11), .adSaldoValorDemo3(11), .adSaldoValorConserto3(11), .adSaldoValorOutros3(11), .adSaldoValorBenef3(11), .adSaldoValorConsig3(12), .adSaldoValorDemo3(12), .adSaldoValorConserto3(12), .adSaldoValorOutros3(12), .adSaldoValorBenef3(12), .sProduto, .iAlmoxarifado, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno, iFilialEmpresa)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 89844
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89845
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        'Quantitade Inicial
'        sComandoSQL(2) = "SELECT Almoxarifado "
'        'Tabela, Filtro, Ordem
'        sComandoSQL(2) = sComandoSQL(2) & "FROM SldMesEstAlm1 WHERE Ano = ? AND Produto = ? AND Almoxarifado = ? ORDER BY Produto"
'
'        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
'        '----------------------------------------------------------------------------------------
'        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, iAlmoxarifado, iAno + 1, tSldMesEstAlm1.sProduto, tSldMesEstAlm1.iAlmoxarifado)
'        If lErro <> AD_SQL_SUCESSO Then gError 89846
'
'        lErro = Comando_BuscarPrimeiro(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89847
'
'        'Se não encontrou
'        If lErro = AD_SQL_SEM_DADOS Then gError 89848
'
'        'Meses
'        For iIndice = 1 To 12
'            dValorAcumuladaBenef3 = dValorAcumuladaBenef3 + tSldMesEstAlm1.adSaldoValorBenef3(iIndice)
'            dValorAcumuladaConserto3 = dValorAcumuladaConserto3 + tSldMesEstAlm1.adSaldoValorConserto3(iIndice)
'            dValorAcumuladaConsig3 = dValorAcumuladaConsig3 + tSldMesEstAlm1.adSaldoValorConsig3(iIndice)
'            dValorAcumuladaDemo3 = dValorAcumuladaDemo3 + tSldMesEstAlm1.adSaldoValorDemo3(iIndice)
'            dValorAcumuladaOutras3 = dValorAcumuladaOutras3 + tSldMesEstAlm1.adSaldoValorOutros3(iIndice)
'        Next
'
'        '------------Atualiza SldMesEst para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
'
'        'Monta comando SQL para UPDATE de SldMesEst
'        sComandoSQL(3) = "UPDATE SldMesEstAlm1 SET ValorInicialConsig3 = ?, ValorInicialDemo3 = ?, ValorInicialConserto3 = ?, ValorInicialOutros3 = ?, ValorInicialBenef3 = ?"
'
'        'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
'        lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorAcumuladaConsig3, dValorAcumuladaDemo3, dValorAcumuladaConserto3, dValorAcumuladaOutras3, dValorAcumuladaBenef3)
'        If lErro <> AD_SQL_SUCESSO Then gError 89849
'
'        'Atualiza tela de acompanhamento do Batch
'        lErro = Rotina_CMP_AtualizaTelaBatch()
'        If lErro <> SUCESSO Then gError 89850
'
'        '-----------------Busca o Proximo -------------------
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89851
'
'    Loop
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial:
'
'    Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial = gErr
'
'    Select Case gErr
'
'        Case 89843
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 89844, 89845
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM12", gErr, iAno)
'
'        Case 89846, 89847, 89851
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM11", gErr, iAno + 1, tSldMesEstAlm1.iAlmoxarifado, tSldMesEstAlm1.sProduto)
'
'        Case 89848
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESESTALM1_NAO_CADASTRADO", gErr, iAno + 1, tSldMesEstAlm1.sProduto, tSldMesEstAlm1.iAlmoxarifado)
'
'        Case 89849
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESESTALM1", gErr, iAno + 1, tSldMesEstAlm1.iAlmoxarifado, tSldMesEstAlm1.sProduto)
'
'        Case 89850
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159528)
'
'    End Select
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial(iAno As Integer, iFilialEmpresa As Integer) As Long
''Atualiza Valor Inicial para o Ano sequinte na tabela SaldoMesEstAlm2
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim iMesFinal As Integer
'Dim iIndice As Integer
'Dim sComandoSQL(1 To 3) As String
'Dim alComando(1 To 3) As Long
'Dim tSldMesEstAlm2 As typeSldMesEstAlm2
'Dim iAlmoxarifado As Integer
'Dim dValorAcumuladaConsig As Double, dValorAcumuladaDemo As Double, dValorAcumuladaConserto As Double, dValorAcumuladaOutras As Double, dValorAcumuladaBenef As Double
'
'On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial
'
'    'Abre comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 69748
'    Next
'
'    'Le os Valores de Entrada e de Saida para todos os meses
'    'Quantidade e valor inicial
'    sComandoSQL(1) = "SELECT ValorInicialConsig, ValorInicialDemo, ValorInicialConserto, ValorInicialOutros, ValorInicialBenef, "
'    'Quantidades e valores de entrada e de saida mensais
'    For iIndice = 1 To 12
'        sComandoSQL(1) = sComandoSQL(1) & "SaldoValorConsig" & CStr(iIndice) & ", " & "SaldoValorDemo" & CStr(iIndice) & ", " & "SaldoValorConserto" & CStr(iIndice) & ", " & "SaldoValorOutros" & CStr(iIndice) & ", " & "SaldoValorBenef" & CStr(iIndice) & ", "
'    Next
'    sComandoSQL(1) = sComandoSQL(1) & " Produto, "
'    sComandoSQL(1) = sComandoSQL(1) & " Almoxarifado "
'    'Tabela, Filtro, Ordem
'    sComandoSQL(1) = sComandoSQL(1) & " FROM Produtos, SldMesEstAlm2, Almoxarifado WHERE Produtos.Codigo = SldMesEstAlm2.Produto AND Almoxarifado.Codigo = SldMesEstAlm2.Almoxarifado AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? AND Almoxarifado.FilialEmpresa = ? ORDER BY Produtos.Codigo"
'
'    With tSldMesEstAlm2
'        .sProduto = String(STRING_PRODUTO, 0)
'
'        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .dValorInicialConsig, .dValorInicialDemo, .dValorInicialConserto, .dValorInicialOutros, .dValorInicialBenef, .adSaldoValorConsig(1), .adSaldoValorDemo(1), .adSaldoValorConserto(1), .adSaldoValorOutros(1), .adSaldoValorBenef(1), .adSaldoValorConsig(2), .adSaldoValorDemo(2), .adSaldoValorConserto(2), .adSaldoValorOutros(2), .adSaldoValorBenef(2), .adSaldoValorConsig(3), .adSaldoValorDemo(3), .adSaldoValorConserto(3), .adSaldoValorOutros(3), .adSaldoValorBenef(3), .adSaldoValorConsig(4), .adSaldoValorDemo(4), .adSaldoValorConserto(4), .adSaldoValorOutros(4), .adSaldoValorBenef(4), .adSaldoValorConsig(5), .adSaldoValorDemo(5), .adSaldoValorConserto(5), .adSaldoValorOutros(5), .adSaldoValorBenef(5), .adSaldoValorConsig(6), .adSaldoValorDemo(6), .adSaldoValorConserto(6), .adSaldoValorOutros(6), .adSaldoValorBenef(6), _
'        .adSaldoValorConsig(7), .adSaldoValorDemo(7), .adSaldoValorConserto(7), .adSaldoValorOutros(7), .adSaldoValorBenef(7), .adSaldoValorConsig(8), .adSaldoValorDemo(8), .adSaldoValorConserto(8), .adSaldoValorOutros(8), .adSaldoValorBenef(8), .adSaldoValorConsig(9), .adSaldoValorDemo(9), .adSaldoValorConserto(9), .adSaldoValorOutros(9), .adSaldoValorBenef(9), .adSaldoValorConsig(10), .adSaldoValorDemo(10), .adSaldoValorConserto(10), .adSaldoValorOutros(10), .adSaldoValorBenef(10), .adSaldoValorConsig(11), .adSaldoValorDemo(11), .adSaldoValorConserto(11), .adSaldoValorOutros(11), .adSaldoValorBenef(11), .adSaldoValorConsig(12), .adSaldoValorDemo(12), .adSaldoValorConserto(12), .adSaldoValorOutros(12), .adSaldoValorBenef(12), .sProduto, .iAlmoxarifado, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno, iFilialEmpresa)
'    End With
'
'    If lErro <> AD_SQL_SUCESSO Then gError 69749
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69750
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        'Quantitade Inicial
'        sComandoSQL(2) = "SELECT Almoxarifado "
'        'Tabela, Filtro, Ordem
'        sComandoSQL(2) = sComandoSQL(2) & "FROM SldMesEstAlm2 WHERE Ano = ? AND Produto = ? AND Almoxarifado = ? ORDER BY Produto"
'
'        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
'        '----------------------------------------------------------------------------------------
'        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, iAlmoxarifado, iAno + 1, tSldMesEstAlm2.sProduto, tSldMesEstAlm2.iAlmoxarifado)
'        If lErro <> AD_SQL_SUCESSO Then gError 69751
'
'        lErro = Comando_BuscarPrimeiro(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69752
'
'        'Se não encontrou
'        If lErro = AD_SQL_SEM_DADOS Then gError 69753
'
'        'Meses
'        For iIndice = 1 To 12
'            dValorAcumuladaBenef = dValorAcumuladaBenef + tSldMesEstAlm2.adSaldoValorBenef(iIndice)
'            dValorAcumuladaConserto = dValorAcumuladaConserto + tSldMesEstAlm2.adSaldoValorConserto(iIndice)
'            dValorAcumuladaConsig = dValorAcumuladaConsig + tSldMesEstAlm2.adSaldoValorConsig(iIndice)
'            dValorAcumuladaDemo = dValorAcumuladaDemo + tSldMesEstAlm2.adSaldoValorDemo(iIndice)
'            dValorAcumuladaOutras = dValorAcumuladaOutras + tSldMesEstAlm2.adSaldoValorOutros(iIndice)
'        Next
'
'        '------------Atualiza SldMesEst para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
'
'        'Monta comando SQL para UPDATE de SldMesEst
'        sComandoSQL(3) = "UPDATE SldMesEstAlm2 SET ValorInicialConsig = ?, ValorInicialDemo = ?, ValorInicialConserto = ?, ValorInicialOutros = ?, ValorInicialBenef = ?"
'
'        'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
'        lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorAcumuladaConsig, dValorAcumuladaDemo, dValorAcumuladaConserto, dValorAcumuladaOutras, dValorAcumuladaBenef)
'        If lErro <> AD_SQL_SUCESSO Then gError 69754
'
'        'Atualiza tela de acompanhamento do Batch
'        lErro = Rotina_CMP_AtualizaTelaBatch()
'        If lErro <> SUCESSO Then gError 69755
'
'        '-----------------Busca o Proximo -------------------
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69756
'
'    Loop
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial:
'
'    Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial = gErr
'
'    Select Case gErr
'
'        Case 69748
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 69749, 69750
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM22", gErr, iAno)
'
'        Case 69751, 69752, 69756
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM21", gErr, iAno + 1, tSldMesEstAlm2.iAlmoxarifado, tSldMesEstAlm2.sProduto)
'
'        Case 69753
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESESTALM2_NAO_CADASTRADO", gErr, iAno + 1, tSldMesEstAlm2.sProduto, tSldMesEstAlm2.iAlmoxarifado)
'
'        Case 69754
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESESTALM2", gErr, iAno + 1, tSldMesEstAlm2.iAlmoxarifado, tSldMesEstAlm2.sProduto)
'
'        Case 69755
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159529)
'
'    End Select
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Function Rotina_CMP_TotalTransfereValorInicial(iFilialEmpresa As Integer, iAno As Integer, iMes As Integer, lTotalProdutos As Long) As Long
''Chamada EM TRANSAÇÃO
''Retorna o Número Total de Produtos que terao valores inicial tranferidos
'
'Dim lComando As Long
'Dim lErro As Long
'Dim sProduto As String
'Dim sComandoSQL(1 To 2) As String
'Dim lSubTotal1 As Long, lSubTotal2 As Long
'
'On Error GoTo Erro_Rotina_CMP_TotalTransfereValorInicial
'
'    'Abre comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 69013
'
'    'Soma a Quantidade de Produtos de SldMesEst que terão os valores iniciais transferidos
'    sComandoSQL(1) = "SELECT COUNT(*) FROM Produtos, SldMesEst WHERE Produtos.Codigo = SldMesEst.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND SldMesEst.FilialEmpresa = ? AND SldMesEst.Ano = ?"
'
'    lErro = Comando_Executar(lComando, sComandoSQL(1), lSubTotal1, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iFilialEmpresa, iAno)
'    If lErro <> AD_SQL_SUCESSO Then gError 69014
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69015
'
'    'Soma a Quantidade de Produtos de SldMesEstAlm que teram os valores iniciais transferidos
'    sComandoSQL(2) = "SELECT COUNT(*) FROM Produtos, SldMesEstAlm, Almoxarifado WHERE Produtos.Codigo = SldMesEstAlm.Produto AND SldMesEstAlm.Almoxarifado = Almoxarifado.Codigo AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Almoxarifado.FilialEmpresa = ? AND SldMesEstAlm.Ano = ?"
'
'    lErro = Comando_Executar(lComando, sComandoSQL(2), lSubTotal2, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iFilialEmpresa, iAno)
'    If lErro <> AD_SQL_SUCESSO Then gError 69022
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69023
'
'    lTotalProdutos = lSubTotal1 + lSubTotal2
'
'    'Fechamento comando
'    Call Comando_Fechar(lComando)
'
'    Rotina_CMP_TotalTransfereValorInicial = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_TotalTransfereValorInicial:
'
'    Rotina_CMP_TotalTransfereValorInicial = gErr
'
'    Select Case gErr
'
'        Case 69013
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 69014, 69015, 69016
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL(1))
'
'        Case 69022, 69023, 69024
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL(2))
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159530)
'
'    End Select
'
'   'Fechamento comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial(iAno As Integer, iFilialEmpresa As Integer) As Long
''Atualiza Valor Inicial e o CustoMedioProducaoInicial para o Ano sequinte na tabela SaldoMesEstAlm
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim dValorInicialProxAno As Double
'Dim iMesFinal As Integer
'Dim iIndice As Integer
'Dim sComandoSQL(1 To 3) As String
'Dim alComando(1 To 3) As Long
'Dim tSldMesEstAlm As typeSldMesEstAlm
'
'On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial
'
'    'Abre comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 69026
'    Next
'
'    'Le os Valores de Entrada e de Saida para todos os meses
'    sComandoSQL(1) = "SELECT "
'
'    'valor inicial
'    sComandoSQL(1) = sComandoSQL(1) & "ValorInicial, ValorInicialCusto, "
'
'    'valores de entrada e de saida mensais
'    For iIndice = 1 To 12
'        sComandoSQL(1) = sComandoSQL(1) & "ValorEnt" & CStr(iIndice) & ", " & "ValorSai" & CStr(iIndice) & ", "
'    Next
'
'    For iIndice = 1 To 12
'        sComandoSQL(1) = sComandoSQL(1) & "SaldoValorCusto" & CStr(iIndice) & ", "
'    Next
'
'    sComandoSQL(1) = sComandoSQL(1) & "Produto "
'    sComandoSQL(1) = sComandoSQL(1) & ",Almoxarifado "
'
'    'Tabela, Filtro, Ordem
'    sComandoSQL(1) = sComandoSQL(1) & "FROM Produtos, SldMesEstAlm, Almoxarifado WHERE Produtos.Codigo = SldMesEstAlm.Produto AND  Almoxarifado.Codigo = SldMesEstAlm.Almoxarifado AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? AND Almoxarifado.FilialEmpresa = ? ORDER BY Produtos.Codigo"
'
'    With tSldMesEstAlm
'
'        .sProduto = String(STRING_PRODUTO, 0)
'        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .dValorInicial, .dValorInicialCusto, .adValorEnt(1), .adValorSai(1), .adValorEnt(2), .adValorSai(2), .adValorEnt(3), .adValorSai(3), .adValorEnt(4), .adValorSai(4), .adValorEnt(5), .adValorSai(5), .adValorEnt(6), .adValorSai(6), .adValorEnt(7), .adValorSai(7), .adValorEnt(8), .adValorSai(8), .adValorEnt(9), .adValorSai(9), .adValorEnt(10), .adValorSai(10), .adValorEnt(11), .adValorSai(11), .adValorEnt(12), .adValorSai(12), _
'        .adSaldoValorCusto(1), .adSaldoValorCusto(2), .adSaldoValorCusto(3), .adSaldoValorCusto(4), .adSaldoValorCusto(5), .adSaldoValorCusto(6), .adSaldoValorCusto(7), .adSaldoValorCusto(8), .adSaldoValorCusto(9), .adSaldoValorCusto(10), .adSaldoValorCusto(11), .adSaldoValorCusto(12), _
'        .sProduto, .iAlmoxarifado, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno, iFilialEmpresa)
'        If lErro <> AD_SQL_SUCESSO Then gError 69027
'
'    End With
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69028
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        'Valor Inicial
'        sComandoSQL(2) = "SELECT ValorInicial "
'        'Tabela, Filtro, Ordem
'        sComandoSQL(2) = sComandoSQL(2) & "FROM SldMesEstAlm WHERE Ano = ? AND Almoxarifado = ? AND Produto = ? ORDER BY Produto"
'
'        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
'        '----------------------------------------------------------------------------------------
'        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, dValorInicialProxAno, iAno + 1, tSldMesEstAlm.iAlmoxarifado, tSldMesEstAlm.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 69029
'
'        lErro = Comando_BuscarPrimeiro(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69030
'
'        'Se não encontrou
'        If lErro = AD_SQL_SEM_DADOS Then gError 69031
'
'        'Cálculo de saldo de valor
'        '-------------------------
'
'        dValorInicialProxAno = tSldMesEstAlm.dValorInicial
'
'        'Meses
'        For iIndice = 1 To 12
'            dValorInicialProxAno = dValorInicialProxAno + tSldMesEstAlm.adValorEnt(iIndice) - tSldMesEstAlm.adValorSai(iIndice)
'        Next
'
'        'Meses
'        For iIndice = 1 To 12
'            tSldMesEstAlm.dValorInicialCusto = tSldMesEstAlm.dValorInicialCusto + tSldMesEstAlm.adSaldoValorCusto(iIndice)
'        Next
'
'        '------------Atualiza SldMesEstAlm para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
'
'        'Monta comando SQL para UPDATE de SldMesEstAlm
'        sComandoSQL(3) = "UPDATE SldMesEstAlm SET ValorInicial = ?, ValorInicialCusto = ?"
'
'        'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
'        lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorInicialProxAno, tSldMesEstAlm.dValorInicialCusto)
'        If lErro <> AD_SQL_SUCESSO Then gError 69032
'
'        'Atualiza tela de acompanhamento do Batch
'        lErro = Rotina_CMP_AtualizaTelaBatch()
'        If lErro <> SUCESSO Then gError 69033
'
'        '-----------------Busca o Proximo -------------------
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69034
'
'    Loop
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial:
'
'    Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial = gErr
'
'    Select Case gErr
'
'        Case 69026
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 69027, 69028, 69034
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM2", gErr, iAno)
'
'        Case 69029, 69030
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM2", gErr, iAno + 1)
'
'        Case 69031
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SLDMESESTALM", gErr, iAno + 1, tSldMesEstAlm.sProduto, tSldMesEstAlm.iAlmoxarifado)
'
'        Case 69032
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESESTALM", gErr, iAno, tSldMesEstAlm.iAlmoxarifado, tSldMesEstAlm.sProduto)
'
'        Case 69033
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159531)
'
'    End Select
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Private Function Rotina_Atualiza_Custo_Movimento_Estorno(iMes As Integer, iAno As Integer, iFilialEmpresa As Integer) As Long
''Atualiza os movimentos deste mês que foram estornados
''Chamada EM TRANSAÇÃO
'
'Dim lErro As Long
'Dim alComando(1 To 3) As Long
'Dim dCustoExtorno As Double
'Dim dCustoMovimento As Double
'Dim lNumIntDoc As Long
'Dim lCodigo As Long
'Dim iIndice As Integer
'
'On Error GoTo Erro_Rotina_Atualiza_Custo_Movimento_Estorno
'
'    'Abre comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 78027
'    Next
'
'    lErro = Comando_Executar(alComando(1), "SELECT NumIntDocEst, Custo FROM MovimentoEstoque WHERE NumIntDocEst > 0 AND MONTH(Data) = ? AND YEAR(Data) = ? AND FilialEmpresa = ?", lNumIntDoc, dCustoMovimento, iMes, iAno, iFilialEmpresa)
'    If lErro <> AD_SQL_SUCESSO Then gError 78028
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78029
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        lErro = Comando_ExecutarPos(alComando(2), "SELECT Custo, Codigo FROM MovimentoEstoque WHERE NumIntDoc = ? AND MONTH(Data) = ? AND YEAR(Data) = ? AND FilialEmpresa = ?", 0, dCustoExtorno, lCodigo, lNumIntDoc, iMes, iAno, iFilialEmpresa)
'        If lErro <> AD_SQL_SUCESSO Then gError 78030
'
'        lErro = Comando_BuscarPrimeiro(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78031
'
'        If lErro = AD_SQL_SUCESSO Then
'
'            'Atualiza o Custo do Movimento
'            lErro = Comando_ExecutarPos(alComando(3), "UPDATE MovimentoEstoque SET Custo = ?", alComando(2), dCustoMovimento)
'            If lErro <> AD_SQL_SUCESSO Then gError 78033
'
'        End If
'
'        '-----------------Busca o Proximo -------------------
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78032
'
'    Loop
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Rotina_Atualiza_Custo_Movimento_Estorno = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_Atualiza_Custo_Movimento_Estorno:
'
'    Rotina_Atualiza_Custo_Movimento_Estorno = gErr
'
'    Select Case gErr
'
'        Case 78027
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 78028, 78029, 78030, 78031, 78032
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
'
'        Case 78033
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MOVIMENTOESTOQUE", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159532)
'
'    End Select
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'
'
