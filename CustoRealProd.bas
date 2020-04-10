Attribute VB_Name = "CustoRealProd"
Option Explicit

'**** Quando iMes for janeiro e houver Dezembro do ano anterior,
'**** transportar só valor (saldo do ano) para ValorInicial

Const HORA_NAO_ANTECIPAR = 6.94444444444444E-04 'cdbl(cDate("00:01:00"))

'???? Tranferir para ErrosMAT
Const ERRO_CMP_E_CST_ZERADOS = 0 'Parametro: sProduto, iMes
'Atenção, o produto %s está com o custo médio de produção e o custo standard do mes %i zerados. Favor colocar o custo standard.

Dim giReprocessamento As Integer 'indica se esta executando a rotina de reprocessamento dos movimentos de estoque

Function Rotina_CustoMedioProducao_Int(iAno As Integer, iMes As Integer) As Long
 'calcula o custo médio de produção para mes/ano passados e valora os movimentos de estoque

Dim alComando1(1 To NUM_MAX_LCOMANDO_MOVESTOQUE) As Long
Dim lTransacao As Long
Dim alComando(1 To 7) As Long
Dim iIndice As Integer
Dim lErro As Long, iOrdem As Integer
Dim tMovEstoque As typeItemMovEstoque
Dim lTotalProdutos As Long
Dim colEstoqueMes As New Collection
Dim objEstoqueMes As ClassEstoqueMes
Dim colApropriacaoInsumo As Collection
Dim lNumMovs As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim objItemMovEst As New ClassItemMovEstoque
Dim iDiasMes As Integer
Dim iAchou As Integer
Dim objEstoqueMes1 As New ClassEstoqueMes
Dim colEstoqueMesProduto As New Collection
Dim objTipoMovEstoque As New ClassTipoMovEst
Dim colMovsAntecipados As New Collection, bAntecipado As Boolean, dtDataMovsAntecipados As Date
Dim iFilialEmpresa As Integer
Dim colExercicio As New Collection, bComMovtoCustoZero As Boolean
Dim objFiliais As AdmFiliais, lNumIntDoc As Long

On Error GoTo Erro_Rotina_CustoMedioProducao_Int

    bComMovtoCustoZero = False
    giReprocessamento = 0 'indica que nao se trata de um reprocessamento
    dtDataMovsAntecipados = DATA_NULA

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92536
    Next

    'libera comandos
    For iIndice = LBound(alComando1) To UBound(alComando1)
        alComando1(iIndice) = Comando_Abrir()
        If alComando1(iIndice) = 0 Then gError 92537
    Next

    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92546

    'desativa os locks dos comandos a seguir
    lErro = Conexao_DesativarLocks(DESATIVAR_LOCKS)
    If lErro <> SUCESSO Then gError 94580

    'faz lock em estoqueMes, le os Gastos Diretos e Indiretos de todas as filiais e coloca estes valores em colEstoqueMes.
    lErro = CF("Rotina_CMP_EstoqueMes_CriticaLock", alComando(1), iAno, iMes, colEstoqueMes)
    If lErro <> SUCESSO Then gError 92547
    
    'preenche uma colecao com os produtos que tiveram gastos informados e que portanto não terão seu calculo feito com os demais produtos
    lErro = CF("EstoqueMesProduto_Le", iAno, iMes, colEstoqueMesProduto)
    If lErro <> SUCESSO Then gError 92891
    
    'Determinação de faixa de datas
    dtDataInicial = CDate("1/" & CStr(iMes) & "/" & CStr(iAno))
    dtDataFinal = DateAdd("m", 1, dtDataInicial) - 1
    
    'elimina apropriacoes automaticas criadas em execucao anterior da rotina de custo
    lErro = CustoProd_LimpaApropriacoesAutomaticas(dtDataInicial, dtDataFinal, alComando)
    If lErro <> SUCESSO Then gError 106529
    
    'Apura o total de horas maquina e custo das matérias primas para cada filial e coloca em colEstoqueMes
    lErro = CF("MovEstoque_Le_HorasMaq_CustoMPrim", colEstoqueMes, iMes, iAno, colEstoqueMesProduto)
    If lErro <> SUCESSO Then gError 92548
    
    'coloca todos os objetos na coleção como apurado para quando passar pela rotina Estoque_AtualizaItemMov poder calcular o custo de produção
    For Each objEstoqueMes In colEstoqueMes
        
        'se está processando o mes 12 de algum ano ==> é necessário que o mes 1 do proximo ano esteja aberto para receber os saldos iniciais
        If iMes = 12 Then
            objEstoqueMes1.iFilialEmpresa = objEstoqueMes.iFilialEmpresa
            objEstoqueMes1.iAno = iAno + 1
            objEstoqueMes1.iMes = 1
        
            lErro = CF("EstoqueMes_Le", objEstoqueMes1)
            If lErro <> SUCESSO And lErro <> 36513 Then gError 92856
            
            'se o ano ainda não foi aberto para a filial em questão ==> erro
            If lErro = 36513 Then gError 92857
        
        End If
        
        objEstoqueMes.iCustoProdApurado = CUSTO_APURADO
        
    Next
    
    'Grava o total de horas maquina e custo das matérias primas para cada filial
    lErro = CF("EstoqueMes_Grava", colEstoqueMes)
    If lErro <> SUCESSO Then gError 92549
    
    'Grava a quantidade total produzida no mes de cada produto que teve o gasto especificado
    lErro = CF("EstoqueMesProduto_Grava", colEstoqueMesProduto)
    If lErro <> SUCESSO Then gError 92896
    
    'Retorna o número de Movimentos de Produtos Produzidos no Mes/Ano em questão
    lErro = Rotina_CMP_TotalMovEstoque(iAno, iMes, lNumMovs)
    If lErro <> SUCESSO Then gError 92550
            
    If iMes = 12 Then
    
        'Retorna o número total de Produtos a serem transferidos para o proximo Ano
        lErro = Rotina_CMP_TotalTransfereValorInicial(iAno, iMes, lTotalProdutos)
        If lErro <> SUCESSO Then gError 92551
    
    End If
    
    'Tela acompanhamento Batch inicializa dValorTotal
    TelaAcompanhaBatchEST.dValorTotal = lNumMovs + (lTotalProdutos * 3)
        
    'zera os totalizadores de estoque SldMesEst, SldMesEst1, SldMesEst2, SldMesEstAlm, SldMesEstAlm1, SldMesEstAlm2
    lErro = CF("Zera_Totalizadores_Estoque", iAno, iMes)
    If lErro <> SUCESSO Then gError 92635
    
    tMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tMovEstoque.sProduto = String(STRING_PRODUTO, 0)
    tMovEstoque.sOPCodigo = String(STRING_OPCODIGO, 0)
    tMovEstoque.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    'le o primeiro movimento de estoque
    With tMovEstoque
        lErro = Comando_Executar(alComando(2), "SELECT Ordem, FilialEmpresa, MovimentoEstoque.Codigo, NumIntDoc, NumIntDocEst, NumIntDocOrigem, Produto, CodigoOP, Quantidade, SiglaUM, Almoxarifado, TipoMov, MovimentoEstoque.Apropriacao, Data, Hora, MovimentoEstoque.HorasMaquina, Custo, Fornecedor, MovimentoEstoque.TipoNumIntDocOrigem, DocOrigem FROM MovimentoEstoque, TiposMovimentoEstoque, TiposOrdemCusto, Produtos WHERE Produtos.Codigo = MovimentoEstoque.Produto AND MovimentoEstoque.TipoMov = TiposMovimentoEstoque.Codigo AND TiposMovimentoEstoque.OrdemCusto = TiposOrdemCusto.Codigo AND Data >= ? AND Data <= ? AND Produtos.Apropriacao = ? ORDER BY Data, TiposOrdemCusto.Ordem, NumIntDoc", _
            iOrdem, .iFilialEmpresa, .lCodigo, .lNumIntDoc, .lNumIntDocEst, .lNumIntDocOrigem, .sProduto, .sOPCodigo, .dQuantidade, .sSiglaUM, .iAlmoxarifado, .iTipoMov, .iApropriacao, .dtData, .dHora, .lHorasMaquina, .dCusto, .lFornecedor, .iTipoNumIntDocOrigem, .sDocOrigem, dtDataInicial, dtDataFinal, APROPR_CUSTO_REAL)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 92552

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92553
    
    Do While lErro = AD_SQL_SUCESSO
        
        objTipoMovEstoque.iCodigo = tMovEstoque.iTipoMov
        
        'ler os dados referentes ao tipo de movimento
        lErro = CF("TiposMovEst_Le1", alComando(3), objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 81873
    
        If objTipoMovEstoque.iAtualizaMovEstoque <> TIPOMOV_EST_ESTORNOMOV Then
        
'            '??? apenas p/debug, retirar
'            If tMovEstoque.lCodigo = 424 Then
'
'                MsgBox ("ok")
'
'            End If
'            If tMovEstoque.sOPCodigo = "8207" Or tMovEstoque.sOPCodigo = "8216" Then
'
'                MsgBox ("ok")
'
'            End If
            
            'identifica se o movto foi processado antecipadamento
            lErro = MovtoProcAntecipado_Verifica(tMovEstoque, colMovsAntecipados, bAntecipado)
            If lErro <> SUCESSO Then gError 81904
            
            'se o processamento do movto nao foi antecipado
            If bAntecipado = False Then
            
                'processa-lo e em seguida o seu estorno, se houver
                
                objItemMovEst.iFilialEmpresa = tMovEstoque.iFilialEmpresa
                objItemMovEst.lCodigo = tMovEstoque.lCodigo
                objItemMovEst.lNumIntDoc = tMovEstoque.lNumIntDoc
                objItemMovEst.lNumIntDocEst = tMovEstoque.lNumIntDocEst
                objItemMovEst.lNumIntDocOrigem = tMovEstoque.lNumIntDocOrigem
                objItemMovEst.sProduto = tMovEstoque.sProduto
                objItemMovEst.sOPCodigo = tMovEstoque.sOPCodigo
                objItemMovEst.dQuantidade = tMovEstoque.dQuantidade
                objItemMovEst.sSiglaUM = tMovEstoque.sSiglaUM
                objItemMovEst.iAlmoxarifado = tMovEstoque.iAlmoxarifado
                objItemMovEst.iTipoMov = tMovEstoque.iTipoMov
                objItemMovEst.iApropriacao = tMovEstoque.iApropriacao
                objItemMovEst.dtData = tMovEstoque.dtData
                objItemMovEst.dtHora = tMovEstoque.dHora
                objItemMovEst.lHorasMaquina = tMovEstoque.lHorasMaquina
                objItemMovEst.dCusto = tMovEstoque.dCusto
                objItemMovEst.lFornecedor = tMovEstoque.lFornecedor
                objItemMovEst.iTipoNumIntDocOrigem = tMovEstoque.iTipoNumIntDocOrigem
                objItemMovEst.sDocOrigem = tMovEstoque.sDocOrigem
                
                Set colApropriacaoInsumo = New Collection
                
                If objItemMovEst.iTipoMov = MOV_EST_PRODUCAO Or objItemMovEst.iTipoMov = MOV_EST_PRODUCAO_BENEF3 Then
                
                    'Le as Apriações do Item
                    lErro = CF("ApropriacaoInsumo_Le_NumIntDocOrigem", tMovEstoque.lNumIntDoc, colApropriacaoInsumo)
                    If lErro <> SUCESSO Then gError 92554
                                
                End If
                                
                Set objItemMovEst.colApropriacaoInsumo = colApropriacaoInsumo
                
                iAchou = 0
                
                For Each objEstoqueMes In colEstoqueMes
                    If objEstoqueMes.iFilialEmpresa = objItemMovEst.iFilialEmpresa Then
                        iAchou = 1
                        Exit For
                    End If
                Next
                
                If iAchou = 0 Then gError 92555
                
                'processa movtos de producao entrada dos insumos para a data sendo processada
                lErro = ProdEntradaInsumos_Processa(objItemMovEst, colMovsAntecipados, alComando1, colEstoqueMes, colEstoqueMesProduto, dtDataMovsAntecipados, iOrdem, objItemMovEst.lNumIntDoc)
                If lErro <> SUCESSO Then gError 106510
                
                'processa movimentos de insumos do produto produzido
                lErro = Insumos_Processa(objItemMovEst, colMovsAntecipados, alComando1, objEstoqueMes, colEstoqueMesProduto, dtDataMovsAntecipados)
                If lErro <> SUCESSO Then gError 81905
                
                'se é uma entrada por transferencia entre filiais da mesma empresa vou precisar garantir que a saida já tenha sido processada
                If gobjCRFAT.lFornEmp <> 0 And objItemMovEst.lFornecedor = gobjCRFAT.lFornEmp And objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_ENTRADA And _
                    objTipoMovEstoque.iCodigoOrig = 0 And objTipoMovEstoque.iAtualizaSoLote = 0 And objTipoMovEstoque.iAtualizaRecebIndisp = 0 And objTipoMovEstoque.iInventario = 0 Then
                    
                    'antecipa tratamento de saida por transferencia entre filiais da mesma empresa
                    lErro = TransfFilialSaida_Processa(objItemMovEst, colMovsAntecipados, alComando1, objEstoqueMes, colEstoqueMesProduto, dtDataMovsAntecipados)
                    If lErro <> SUCESSO And lErro <> 106506 Then gError 106510
                
                End If
                
                'guarda o custo anterior para poder guardar os valores em estoqueproduto pela diferença. Permite reexecução da apuração.
                objItemMovEst.dCustoAnt = objItemMovEst.dCusto
                
                lErro = CF("Estoque_ApuraCustoProducao", alComando1, objItemMovEst, objEstoqueMes, colEstoqueMesProduto)
                If lErro <> SUCESSO Then gError 92556
                
                'se o movto possui um estorno
                If objItemMovEst.lNumIntDocEst <> 0 Then
                
                    tMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
                    tMovEstoque.sProduto = String(STRING_PRODUTO, 0)
                    tMovEstoque.sOPCodigo = String(STRING_OPCODIGO, 0)
                    
                    'le o movimento de estoque de estorno
                    With tMovEstoque
                        lErro = Comando_Executar(alComando(4), "SELECT FilialEmpresa, Codigo, NumIntDoc, NumIntDocEst, NumIntDocOrigem, Produto, CodigoOP, Quantidade, SiglaUM, Almoxarifado, TipoMov, Apropriacao, Data, Hora, HorasMaquina, Custo, TipoNumIntDocOrigem FROM MovimentoEstoque WHERE NumIntDoc = ?", .iFilialEmpresa, .lCodigo, .lNumIntDoc, .lNumIntDocEst, .lNumIntDocOrigem, .sProduto, .sOPCodigo, .dQuantidade, .sSiglaUM, .iAlmoxarifado, .iTipoMov, .iApropriacao, .dtData, .dHora, .lHorasMaquina, .dCusto, .iTipoNumIntDocOrigem, objItemMovEst.lNumIntDocEst)
                    End With
                    If lErro <> AD_SQL_SUCESSO Then gError 92552
                
                    lErro = Comando_BuscarPrimeiro(alComando(4))
                    If lErro <> AD_SQL_SUCESSO Then gError 92553
        
                    objItemMovEst.iFilialEmpresa = tMovEstoque.iFilialEmpresa
                    objItemMovEst.lCodigo = tMovEstoque.lCodigo
                    objItemMovEst.lNumIntDoc = tMovEstoque.lNumIntDoc
                    objItemMovEst.lNumIntDocEst = tMovEstoque.lNumIntDocEst
                    objItemMovEst.iTipoNumIntDocOrigem = tMovEstoque.iTipoNumIntDocOrigem
                    objItemMovEst.lNumIntDocOrigem = tMovEstoque.lNumIntDocOrigem
                    objItemMovEst.sProduto = tMovEstoque.sProduto
                    objItemMovEst.sOPCodigo = tMovEstoque.sOPCodigo
                    objItemMovEst.dQuantidade = tMovEstoque.dQuantidade
                    objItemMovEst.sSiglaUM = tMovEstoque.sSiglaUM
                    objItemMovEst.iAlmoxarifado = tMovEstoque.iAlmoxarifado
                    objItemMovEst.iTipoMov = tMovEstoque.iTipoMov
                    objItemMovEst.iApropriacao = tMovEstoque.iApropriacao
                    objItemMovEst.dtData = tMovEstoque.dtData
                    objItemMovEst.dtHora = tMovEstoque.dHora
                    objItemMovEst.lHorasMaquina = tMovEstoque.lHorasMaquina
                    objItemMovEst.dCusto = tMovEstoque.dCusto
                
                    Set objItemMovEst.colApropriacaoInsumo = New Collection
                
                    'guarda o custo anterior para poder guardar os valores em estoqueproduto pela diferença. Permite reexecução da apuração.
                    objItemMovEst.dCustoAnt = objItemMovEst.dCusto
                    
                    lErro = CF("Estoque_ApuraCustoProducao", alComando1, objItemMovEst, objEstoqueMes, colEstoqueMesProduto)
                    If lErro <> SUCESSO Then gError 92556
                    
                End If
            
            End If
            
        End If
        
        'Atualiza tela de acompanhamento do Batch
        lErro = Rotina_CMP_AtualizaTelaBatch()
        If lErro <> SUCESSO Then gError 92557
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92558
        
    Loop
    
    If iMes = 12 Then
    
        'Atualiza os valores iniciais para o mês de Desembro
        lErro = Transfere_ValoresIniciais(iAno)
        If lErro <> SUCESSO Then gError 69763
    
    End If
    
'Início do Trecho de código alterado por Daniel e Mário em 02/10/2002
    For Each objFiliais In gcolFiliais
        
        iFilialEmpresa = objFiliais.iCodFilial
        
        If iFilialEmpresa <> EMPRESA_TODA Then
            
            'Ajustar a contabilidade do mes estoque em questão
            lErro = CF("Rotina_Reprocessamento_CProd", iFilialEmpresa, iMes, iAno, colExercicio)
            If lErro <> SUCESSO Then gError 106524
                    
        End If
        
    Next
'Fim do Trecho de código alterado por Daniel e Mário em 02/10/2002
    
    'verificar se ficou algum movto sem ser custeado
    lErro = Comando_Executar(alComando(7), "SELECT NumIntDoc FROM MovimentoEstoque, TiposMovimentoEstoque WHERE MovimentoEstoque.TipoMov = TiposMovimentoEstoque.Codigo AND Custo = 0 AND Data BETWEEN ? AND ? AND NumIntDocEst = 0 AND AtualizaSoLote = 0", lNumIntDoc, dtDataInicial, dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 124003
    
    lErro = Comando_BuscarProximo(alComando(7))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124004
            
    'apenas registrar que vai ter ser analisado
    If lErro <> AD_SQL_SEM_DADOS Then bComMovtoCustoZero = True
    
    'reativa os locks
    lErro = Conexao_DesativarLocks(REATIVAR_LOCKS)
    If lErro <> SUCESSO Then gError 94581
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 25241

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
   'Fechamento comando
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    
    If bComMovtoCustoZero Then 'gError 124005
        Call Rotina_Aviso(vbOKOnly, "AVISO_MOVTOS_SEM_CUSTO", gErr)
    End If
    
    Rotina_CustoMedioProducao_Int = SUCESSO

    Exit Function

Erro_Rotina_CustoMedioProducao_Int:

    Rotina_CustoMedioProducao_Int = gErr

    Select Case gErr

       Case 92536, 92537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 92546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
          
        Case 92547, 92548, 92549, 92550, 92551, 92554, 92556, 92557, 92635, 92856, 92891, 92896, 81873, 81904, 81905, 106510, 106524, 106529
          
        Case 92552, 92553, 92558, 124003, 124004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
          
        Case 92555
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE4", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
          
        Case 92857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_AINDA_NAO_ABERTO", gErr, objEstoqueMes1.iFilialEmpresa, objEstoqueMes1.iAno, objEstoqueMes1.iMes)
          
        Case 94580
            Call Rotina_Erro(vbOKOnly, "ERRO_DESATIVACAO_LOCKS", gErr)
             
        Case 94581
            Call Rotina_Erro(vbOKOnly, "ERRO_REATIVACAO_LOCKS", gErr)
        
        Case 124005
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVTOS_SEM_CUSTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158691)

    End Select

    'reativa os locks
    Call Conexao_DesativarLocks(REATIVAR_LOCKS)
    
    'Rollback
    Call Transacao_Rollback

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
   'Fechamento comando
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    
    Exit Function

End Function

Function Rotina_CustoMedioProducao_Reproc(ByVal iFilialEmpresa As Integer, iAno As Integer, iMes As Integer) As Long
'calcula o custo médio de produção para mes/ano passados e valora os movimentos de estoque

Dim alComando(1 To 27) As Long
Dim sComandoSQL(1 To 9) As String
Dim iIndice As Integer
Dim lErro As Long
Dim lTotalProdutos As Long 'nº de produtos que participam processo
Dim tMovEstoque As typeItemMovEstoque
Dim tMovEstoque2 As typeItemMovEstoque
Dim tSldMesEst As typeSldMesEst
Dim tSldDiaEst As typeSldDiaEst
Dim tSldDiaEstAlm As typeSldDiaEstAlm
Dim dCPAtual As Double 'Custo Producao do mês atual (iMes)
Dim tProduto As typeProduto
Dim dCMPAtual As Double 'Custo Medio Producao do mês atual (iMes)
Dim colAlmoxInfo As Collection
Dim tTipoMovEst As typeTipoMovEst
Dim lTotalProdutos2 As Long
Dim tSldMesEst2 As typeSldMesEst2
Dim tSldMesEst1 As typeSldMesEst1

On Error GoTo Erro_Rotina_CustoMedioProducao_Reproc

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
    Rotina_CustoMedioProducao_Reproc = SUCESSO

    Exit Function

Erro_Rotina_CustoMedioProducao_Reproc:

    Rotina_CustoMedioProducao_Reproc = gErr

    Select Case gErr

       Case 25234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 25235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 25236, 25237, 25238, 25239, 25240, 25292, 25324, 25780, 25781, 25782, 69021, 69763, 78034 'Tratados na rotina chamada

        Case 25241
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158692)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Transfere_ValoresIniciais(iAno As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Transfere_ValoresIniciais

    'Transfere para o proximo ano o Valor Inicial e o CustoMedioProducaoInicial de SaldoMesEst
    lErro = Rotina_CMP_Transfere_SldMesEst_ValorInicial(iAno)
    If lErro <> SUCESSO Then gError 64486

    'Transfere para o proximo Ano os Valor Inicial e o CustoMedioProducaoInicial
    lErro = Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial(iAno)
    If lErro <> SUCESSO Then gError 69034
    
    'Transfere para o proximo ano o Valor Inicial e o CustoMedioProducaoInicial de SaldoMesEst
    lErro = Rotina_CMP_Transfere_SldMesEst1_ValorInicial(iAno)
    If lErro <> SUCESSO Then gError 89842
    
    'Transfere para o proximo ano o Valor Inicial e o CustoMedioProducaoInicial de SaldoMesEst
    lErro = Rotina_CMP_Transfere_SldMesEst2_ValorInicial(iAno)
    If lErro <> SUCESSO Then gError 69757

    'Transfere para o proximo Ano os Valor Inicial e o CustoMedioProducaoInicial
    lErro = Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial(iAno)
    If lErro <> SUCESSO Then gError 89844

    'Transfere para o proximo Ano os Valor Inicial e o CustoMedioProducaoInicial
    lErro = Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial(iAno)
    If lErro <> SUCESSO Then gError 69758
        
    Transfere_ValoresIniciais = SUCESSO
    
    Exit Function
    
Erro_Transfere_ValoresIniciais:

    Transfere_ValoresIniciais = Err
    
    Select Case Err
            
        Case 64486, 69034, 69757, 69758, 89842, 89844
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158693)

    End Select
    
    Exit Function
    
End Function

Private Function Rotina_CMP_AtualizaTelaBatch()
'Atualiza tela de acompanhamento do Batch

On Error GoTo Erro_Rotina_CMP_AtualizaTelaBatch

    If giReprocessamento = 0 Then

        DoEvents

        'Atualiza tela de acompanhamento do Batch
        TelaAcompanhaBatchEST.dValorAtual = TelaAcompanhaBatchEST.dValorAtual + 1
        TelaAcompanhaBatchEST.TotReg.Caption = CStr(TelaAcompanhaBatchEST.dValorAtual)
        TelaAcompanhaBatchEST.ProgressBar1.Value = 0.5 'StrParaInt((TelaAcompanhaBatchEST.dValorAtual / TelaAcompanhaBatchEST.dValorTotal) * 100)
    
    End If
    
    Rotina_CMP_AtualizaTelaBatch = SUCESSO
    
    Exit Function

Erro_Rotina_CMP_AtualizaTelaBatch:

    Rotina_CMP_AtualizaTelaBatch = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158694)

    End Select
       
    Exit Function

End Function

'Private Function Rotina_CMP_Transfere_SldMesEst_ValorInicial(iAno As Integer) As Long
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
'Dim alComando(1 To 4) As Long
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
'    sComandoSQL(1) = "SELECT FilialEmpresa, CustoMedioProducaoInicial, "
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
'    sComandoSQL(1) = sComandoSQL(1) & "FROM Produtos, SldMesEst WHERE Produtos.Codigo = SldMesEst.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? ORDER BY Produtos.Codigo"
'
'
'    With tSldMesEst
'
'        .sProduto = String(STRING_PRODUTO, 0)
'
'        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .iFilialEmpresa, .dCustoMedioProducaoInicial, .dQuantInicial, .dValorInicial, .dQuantInicialCusto, .dValorInicialCusto, .adQuantEnt(1), .adQuantSai(1), .adValorEnt(1), .adValorSai(1), .adQuantEnt(2), .adQuantSai(2), .adValorEnt(2), .adValorSai(2), .adQuantEnt(3), .adQuantSai(3), .adValorEnt(3), .adValorSai(3), .adQuantEnt(4), .adQuantSai(4), .adValorEnt(4), .adValorSai(4), .adQuantEnt(5), .adQuantSai(5), .adValorEnt(5), .adValorSai(5), .adQuantEnt(6), .adQuantSai(6), .adValorEnt(6), .adValorSai(6), _
'                .adQuantEnt(7), .adQuantSai(7), .adValorEnt(7), .adValorSai(7), .adQuantEnt(8), .adQuantSai(8), .adValorEnt(8), .adValorSai(8), .adQuantEnt(9), .adQuantSai(9), .adValorEnt(9), .adValorSai(9), .adQuantEnt(10), .adQuantSai(10), .adValorEnt(10), .adValorSai(10), .adQuantEnt(11), .adQuantSai(11), .adValorEnt(11), .adValorSai(11), .adQuantEnt(12), .adQuantSai(12), .adValorEnt(12), .adValorSai(12), _
'                .adSaldoQuantCusto(1), .adSaldoValorCusto(1), .adSaldoQuantCusto(2), .adSaldoValorCusto(2), .adSaldoQuantCusto(3), .adSaldoValorCusto(3), .adSaldoQuantCusto(4), .adSaldoValorCusto(4), .adSaldoQuantCusto(5), .adSaldoValorCusto(5), .adSaldoQuantCusto(6), .adSaldoValorCusto(6), .adSaldoQuantCusto(7), .adSaldoValorCusto(7), .adSaldoQuantCusto(8), .adSaldoValorCusto(8), .adSaldoQuantCusto(9), .adSaldoValorCusto(9), .adSaldoQuantCusto(10), .adSaldoValorCusto(10), .adSaldoQuantCusto(11), .adSaldoValorCusto(11), .adSaldoQuantCusto(12), .adSaldoValorCusto(12), _
'                .sProduto, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno)
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
'        'Tabela, Filtro, Ordem
'        sComandoSQL(2) = "SELECT QuantInicial FROM SldMesEst WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? ORDER BY Produto"
'
'        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
'        '----------------------------------------------------------------------------------------
'        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, dQuantInicialProxAno, iAno + 1, tSldMesEst.iFilialEmpresa, tSldMesEst.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 69008
'
'        lErro = Comando_BuscarPrimeiro(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69009
'
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
'            If iMesFinal > 0 Then
'
'                If tSldMesEst.adQuantSai(iMesFinal) > 0 Then
'
'                    'Custo Médio é o valor da última saída mensal dividido pela quantidade
'                    dCustoMedioProducaoInicial = tSldMesEst.adValorSai(iMesFinal) / tSldMesEst.adQuantSai(iMesFinal)
'
'                Else 'Todas as quantidades de saída e de entrada do Produto estão zeradas nesse ano
'
'                    dCustoMedioProducaoInicial = tSldMesEst.dCustoMedioProducaoInicial
'
'                End If
'            Else 'Todas as quantidades de saída e de entrada do Produto estão zeradas nesse ano
'
'                dCustoMedioProducaoInicial = tSldMesEst.dCustoMedioProducaoInicial
'            End If
'
'        End If
'
'        '------------Atualiza SldMesEst para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
'        If lErro = AD_SQL_SUCESSO Then
'
'
'            'Monta comando SQL para UPDATE de SldMesEst
'            sComandoSQL(3) = "UPDATE SldMesEst SET ValorInicial = ?, ValorInicialCusto = ?, CustoMedioProducaoInicial  = ?"
'
'            'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
'            lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorInicialProxAno, tSldMesEst.dValorInicialCusto, dCustoMedioProducaoInicial)
'            If lErro <> AD_SQL_SUCESSO Then gError 69011
'
'        Else
'
'            'Insere os dados no BD
'            lErro = Comando_Executar(alComando(4), "INSERT INTO SldMesEst (Ano, FilialEmpresa, Produto, ValorInicial, ValorInicialCusto, CustoMedioProducaoInicial) VALUES (?,?,?,?,?,?)", iAno + 1, tSldMesEst.iFilialEmpresa, tSldMesEst.sProduto, dValorInicialProxAno, tSldMesEst.dValorInicialCusto, dCustoMedioProducaoInicial)
'            If lErro <> AD_SQL_SUCESSO Then gError 40691
'
'        End If
'
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
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, tSldMesEst.iFilialEmpresa, iAno)
'
'        Case 69008, 69009
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, tSldMesEst.iFilialEmpresa, iAno + 1)
'
'        Case 69010
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SALDOMESEST", gErr, tSldMesEst.iFilialEmpresa, iAno + 1, tSldMesEst.sProduto)
'
'        Case 69011
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESEST", gErr, iAno, tSldMesEst.iFilialEmpresa, tSldMesEst.sProduto)
'
'        Case 69017
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158695)
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

Private Function Rotina_CMP_Transfere_SldMesEst_ValorInicial(iAno As Integer) As Long
'Atualiza Valor Inicial e o CustoMedioProducaoInicial para o Ano sequinte na tabela SaldoMesEst
'Chamada EM TRANSAÇÃO

'Refeita em 27/10/2015 por Wagner para considerar o Valor Inicial e Quantidade Inicial de Custo e
'antes estava considerando ValorInicial e QuantInicial que tinham saldo de terceiros conosco

Dim lErro As Long
Dim dValorInicialProxAno As Double
Dim dQuantInicialProxAno As Double
Dim dValorInicialCustoProxAno As Double
Dim dQuantInicialCustoProxAno As Double
Dim dCustoMedioProducaoInicial As Double
Dim iMesFinal As Integer
Dim iIndice As Integer
Dim sComandoSQL(1 To 3) As String
Dim alComando(1 To 4) As Long
Dim tSldMesEst As typeSldMesEst

On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEst_ValorInicial

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69005
    Next

    'Le os Valores de Entrada e de Saida para todos os meses
    sComandoSQL(1) = "SELECT FilialEmpresa, CustoMedioProducaoInicial, "
    
    'Quantidade e valor inicial
    sComandoSQL(1) = sComandoSQL(1) & "QuantInicial, ValorInicial, QuantInicialCusto, ValorInicialCusto, "
    
    'Quantidades e valores de entrada e de saida mensais
    For iIndice = 1 To 12
        sComandoSQL(1) = sComandoSQL(1) & "QuantEnt" & CStr(iIndice) & ", " & "QuantSai" & CStr(iIndice) & ", " & "ValorEnt" & CStr(iIndice) & ", " & "ValorSai" & CStr(iIndice) & ", "
    Next
    
    For iIndice = 1 To 12
        sComandoSQL(1) = sComandoSQL(1) & "SaldoQuantCusto" & CStr(iIndice) & ", " & "SaldoValorCusto" & CStr(iIndice) & ", "
    Next
    
    sComandoSQL(1) = sComandoSQL(1) & "Produto "
    'Tabela, Filtro, Ordem
    sComandoSQL(1) = sComandoSQL(1) & "FROM Produtos, SldMesEst WHERE Produtos.Codigo = SldMesEst.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? ORDER BY Produtos.Codigo"
    
    
    With tSldMesEst
        
        .sProduto = String(STRING_PRODUTO, 0)
        
        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .iFilialEmpresa, .dCustoMedioProducaoInicial, .dQuantInicial, .dValorInicial, .dQuantInicialCusto, .dValorInicialCusto, .adQuantEnt(1), .adQuantSai(1), .adValorEnt(1), .adValorSai(1), .adQuantEnt(2), .adQuantSai(2), .adValorEnt(2), .adValorSai(2), .adQuantEnt(3), .adQuantSai(3), .adValorEnt(3), .adValorSai(3), .adQuantEnt(4), .adQuantSai(4), .adValorEnt(4), .adValorSai(4), .adQuantEnt(5), .adQuantSai(5), .adValorEnt(5), .adValorSai(5), .adQuantEnt(6), .adQuantSai(6), .adValorEnt(6), .adValorSai(6), _
                .adQuantEnt(7), .adQuantSai(7), .adValorEnt(7), .adValorSai(7), .adQuantEnt(8), .adQuantSai(8), .adValorEnt(8), .adValorSai(8), .adQuantEnt(9), .adQuantSai(9), .adValorEnt(9), .adValorSai(9), .adQuantEnt(10), .adQuantSai(10), .adValorEnt(10), .adValorSai(10), .adQuantEnt(11), .adQuantSai(11), .adValorEnt(11), .adValorSai(11), .adQuantEnt(12), .adQuantSai(12), .adValorEnt(12), .adValorSai(12), _
                .adSaldoQuantCusto(1), .adSaldoValorCusto(1), .adSaldoQuantCusto(2), .adSaldoValorCusto(2), .adSaldoQuantCusto(3), .adSaldoValorCusto(3), .adSaldoQuantCusto(4), .adSaldoValorCusto(4), .adSaldoQuantCusto(5), .adSaldoValorCusto(5), .adSaldoQuantCusto(6), .adSaldoValorCusto(6), .adSaldoQuantCusto(7), .adSaldoValorCusto(7), .adSaldoQuantCusto(8), .adSaldoValorCusto(8), .adSaldoQuantCusto(9), .adSaldoValorCusto(9), .adSaldoQuantCusto(10), .adSaldoValorCusto(10), .adSaldoQuantCusto(11), .adSaldoValorCusto(11), .adSaldoQuantCusto(12), .adSaldoValorCusto(12), _
                .sProduto, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno)
    
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 69006
   
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69007
    
    Do While lErro <> AD_SQL_SEM_DADOS
            
        'Tabela, Filtro, Ordem
        sComandoSQL(2) = "SELECT QuantInicial FROM SldMesEst WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? ORDER BY Produto"

        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
        '----------------------------------------------------------------------------------------
        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, dQuantInicialProxAno, iAno + 1, tSldMesEst.iFilialEmpresa, tSldMesEst.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69008
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69009
        
    
        'Cálculo de saldo de quantidade
        '----------------------------------------------------------------
        
        'Quantidade no início do ano
        dQuantInicialProxAno = tSldMesEst.dQuantInicial
        dQuantInicialCustoProxAno = tSldMesEst.dQuantInicialCusto
        
        'Acumula entrada e saída dos meses
        For iIndice = 1 To 12
            dQuantInicialProxAno = dQuantInicialProxAno + tSldMesEst.adQuantEnt(iIndice) - tSldMesEst.adQuantSai(iIndice)
            dQuantInicialCustoProxAno = dQuantInicialCustoProxAno + tSldMesEst.adSaldoQuantCusto(iIndice)
        Next
            
        'Cálculo de saldo de valor
        '-------------------------
        'Valor inicial
        dValorInicialProxAno = tSldMesEst.dValorInicial
        dValorInicialCustoProxAno = tSldMesEst.dValorInicialCusto
                
        'Meses
        For iIndice = 1 To 12
            dValorInicialProxAno = dValorInicialProxAno + tSldMesEst.adValorEnt(iIndice) - tSldMesEst.adValorSai(iIndice)
            dValorInicialCustoProxAno = dValorInicialCustoProxAno + tSldMesEst.adSaldoValorCusto(iIndice)
        Next
        
        If dQuantInicialCustoProxAno > 0 Then
            
            'Calcula CustoMedioProducaoAtual
            '-------------------------------
            dCustoMedioProducaoInicial = dValorInicialCustoProxAno / dQuantInicialCustoProxAno
        
        ElseIf dQuantInicialCustoProxAno = 0 Then
        
            'Procura o último mês em que houve saída (apropr=CMP)
            For iMesFinal = 12 To 1 Step -1
            
                If tSldMesEst.adQuantSai(iMesFinal) > 0 Then Exit For
            
            Next
        
            If iMesFinal > 0 Then
        
                If tSldMesEst.adSaldoQuantCusto(iMesFinal) > 0 Then
            
                    'Custo Médio é o valor da última saída mensal dividido pela quantidade
                    dCustoMedioProducaoInicial = tSldMesEst.adSaldoValorCusto(iMesFinal) / tSldMesEst.adSaldoQuantCusto(iMesFinal)
                
                Else 'Todas as quantidades de saída e de entrada do Produto estão zeradas nesse ano
            
                    dCustoMedioProducaoInicial = tSldMesEst.dCustoMedioProducaoInicial
                            
                End If
            Else 'Todas as quantidades de saída e de entrada do Produto estão zeradas nesse ano
            
                dCustoMedioProducaoInicial = tSldMesEst.dCustoMedioProducaoInicial
            End If
        
        End If
            
        '------------Atualiza SldMesEst para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
        If lErro = AD_SQL_SUCESSO Then
        
        
            'Monta comando SQL para UPDATE de SldMesEst
            sComandoSQL(3) = "UPDATE SldMesEst SET ValorInicial = ?, ValorInicialCusto = ?, CustoMedioProducaoInicial  = ?"
            
            'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
            lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorInicialProxAno, dValorInicialCustoProxAno, dCustoMedioProducaoInicial)
            If lErro <> AD_SQL_SUCESSO Then gError 69011
        
        Else
        
            'Insere os dados no BD
            lErro = Comando_Executar(alComando(4), "INSERT INTO SldMesEst (Ano, FilialEmpresa, Produto, ValorInicial, ValorInicialCusto, CustoMedioProducaoInicial) VALUES (?,?,?,?,?,?)", iAno + 1, tSldMesEst.iFilialEmpresa, tSldMesEst.sProduto, dValorInicialProxAno, dValorInicialCustoProxAno, dCustoMedioProducaoInicial)
            If lErro <> AD_SQL_SUCESSO Then gError 40691
        
        End If
        
        
        'Atualiza tela de acompanhamento do Batch
        lErro = Rotina_CMP_AtualizaTelaBatch()
        If lErro <> SUCESSO Then gError 69017
        
        '-----------------Busca o Proximo -------------------
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69012
    
    Loop
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Rotina_CMP_Transfere_SldMesEst_ValorInicial = SUCESSO

    Exit Function

Erro_Rotina_CMP_Transfere_SldMesEst_ValorInicial:

    Rotina_CMP_Transfere_SldMesEst_ValorInicial = gErr

    Select Case gErr
                
        Case 69005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69006, 69007, 69012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, tSldMesEst.iFilialEmpresa, iAno)
        
        Case 69008, 69009
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, tSldMesEst.iFilialEmpresa, iAno + 1)
        
        Case 69010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SALDOMESEST", gErr, tSldMesEst.iFilialEmpresa, iAno + 1, tSldMesEst.sProduto)
         
        Case 69011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESEST", gErr, iAno, tSldMesEst.iFilialEmpresa, tSldMesEst.sProduto)
        
        Case 69017
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158695)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Rotina_CMP_Transfere_SldMesEst1_ValorInicial(iAno As Integer) As Long
'Atualiza Valor Inicial para o Ano sequinte na tabela SaldoMesEst1
'Chamada EM TRANSAÇÃO

Dim lErro As Long
Dim iMesFinal As Integer
Dim iIndice As Integer
Dim sComandoSQL(1 To 3) As String
Dim alComando(1 To 4) As Long
Dim tSldMesEst1 As typeSldMesEst1
Dim dValorAcumuladaConsig3 As Double, dValorAcumuladaDemo3 As Double, dValorAcumuladaConserto3 As Double, dValorAcumuladaOutras3 As Double, dValorAcumuladaBenef3 As Double
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEst1_ValorInicial

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89833
    Next

    'Le os Valores de Entrada e de Saida para todos os meses
    'Quantidade e valor inicial
    sComandoSQL(1) = "SELECT FilialEmpresa, ValorInicialConsig3, ValorInicialDemo3, ValorInicialConserto3, ValorInicialOutros3, ValorInicialBenef3, "
    'Quantidades e valores de entrada e de saida mensais
    For iIndice = 1 To 12
        sComandoSQL(1) = sComandoSQL(1) & "SaldoValorConsig3" & CStr(iIndice) & ", " & "SaldoValorDemo3" & CStr(iIndice) & ", " & "SaldoValorConserto3" & CStr(iIndice) & ", " & "SaldoValorOutros3" & CStr(iIndice) & ", " & "SaldoValorBenef3" & CStr(iIndice) & ", "
    Next
    
    sComandoSQL(1) = sComandoSQL(1) & "Produto "
    
    'Tabela, Filtro, Ordem
    sComandoSQL(1) = sComandoSQL(1) & "FROM Produtos, SldMesEst1 WHERE Produtos.Codigo = SldMesEst1.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? ORDER BY Produtos.Codigo"
    
    
    With tSldMesEst1
        .sProduto = String(STRING_PRODUTO, 0)
        
        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .iFilialEmpresa, .dValorInicialConsig3, .dValorInicialDemo3, .dValorInicialConserto3, .dValorInicialOutros3, .dValorInicialBenef3, .adSaldoValorConsig3(1), .adSaldoValorDemo3(1), .adSaldoValorConserto3(1), .adSaldoValorOutros3(1), .adSaldoValorBenef3(1), .adSaldoValorConsig3(2), .adSaldoValorDemo3(2), .adSaldoValorConserto3(2), .adSaldoValorOutros3(2), .adSaldoValorBenef3(2), .adSaldoValorConsig3(3), .adSaldoValorDemo3(3), .adSaldoValorConserto3(3), .adSaldoValorOutros3(3), .adSaldoValorBenef3(3), .adSaldoValorConsig3(4), .adSaldoValorDemo3(4), .adSaldoValorConserto3(4), .adSaldoValorOutros3(4), .adSaldoValorBenef3(4), .adSaldoValorConsig3(5), .adSaldoValorDemo3(5), .adSaldoValorConserto3(5), .adSaldoValorOutros3(5), .adSaldoValorBenef3(5), .adSaldoValorConsig3(6), .adSaldoValorDemo3(6), .adSaldoValorConserto3(6), .adSaldoValorOutros3(6), .adSaldoValorBenef3(6), _
        .adSaldoValorConsig3(7), .adSaldoValorDemo3(7), .adSaldoValorConserto3(7), .adSaldoValorOutros3(7), .adSaldoValorBenef3(7), .adSaldoValorConsig3(8), .adSaldoValorDemo3(8), .adSaldoValorConserto3(8), .adSaldoValorOutros3(8), .adSaldoValorBenef3(8), .adSaldoValorConsig3(9), .adSaldoValorDemo3(9), .adSaldoValorConserto3(9), .adSaldoValorOutros3(9), .adSaldoValorBenef3(9), .adSaldoValorConsig3(10), .adSaldoValorDemo3(10), .adSaldoValorConserto3(10), .adSaldoValorOutros3(10), .adSaldoValorBenef3(10), .adSaldoValorConsig3(11), .adSaldoValorDemo3(11), .adSaldoValorConserto3(11), .adSaldoValorOutros3(11), .adSaldoValorBenef3(11), .adSaldoValorConsig3(12), .adSaldoValorDemo3(12), .adSaldoValorConserto3(12), .adSaldoValorOutros3(12), .adSaldoValorBenef3(12), .sProduto, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno)
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 89834
   
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89835
    
    Do While lErro <> AD_SQL_SEM_DADOS
            
        'Quantitade Inicial
        sComandoSQL(2) = "SELECT FilialEmpresa "
        'Tabela, Filtro, Ordem
        sComandoSQL(2) = sComandoSQL(2) & "FROM SldMesEst1 WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? ORDER BY Produto"

        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
        '----------------------------------------------------------------------------------------
        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, iFilialEmpresa, iAno + 1, tSldMesEst1.iFilialEmpresa, tSldMesEst1.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 89836
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89837
        
        With tSldMesEst1
            dValorAcumuladaBenef3 = .dValorInicialBenef3
            dValorAcumuladaConserto3 = .dValorInicialConserto3
            dValorAcumuladaConsig3 = .dValorInicialConsig3
            dValorAcumuladaDemo3 = .dValorInicialDemo3
            dValorAcumuladaOutras3 = .dValorInicialOutros3
        End With
    
        'Meses
        For iIndice = 1 To 12
            dValorAcumuladaBenef3 = dValorAcumuladaBenef3 + tSldMesEst1.adSaldoValorBenef3(iIndice)
            dValorAcumuladaConserto3 = dValorAcumuladaConserto3 + tSldMesEst1.adSaldoValorConserto3(iIndice)
            dValorAcumuladaConsig3 = dValorAcumuladaConsig3 + tSldMesEst1.adSaldoValorConsig3(iIndice)
            dValorAcumuladaDemo3 = dValorAcumuladaDemo3 + tSldMesEst1.adSaldoValorDemo3(iIndice)
            dValorAcumuladaOutras3 = dValorAcumuladaOutras3 + tSldMesEst1.adSaldoValorOutros3(iIndice)
        Next
        
        '------------Atualiza SldMesEst para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
            
        If lErro = AD_SQL_SUCESSO Then
            
            'Monta comando SQL para UPDATE de SldMesEst
            sComandoSQL(3) = "UPDATE SldMesEst1 SET ValorInicialConsig3 = ?, ValorInicialDemo3 = ?, ValorInicialConserto3 = ?, ValorInicialOutros3 = ?, ValorInicialBenef3 = ?"
            
            'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
            lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorAcumuladaConsig3, dValorAcumuladaDemo3, dValorAcumuladaConserto3, dValorAcumuladaOutras3, dValorAcumuladaBenef3)
            If lErro <> AD_SQL_SUCESSO Then gError 89839
        
        Else
        
            'Insere os dados no BD
            lErro = Comando_Executar(alComando(4), "INSERT INTO SldMesEst1 (Ano, FilialEmpresa, Produto, ValorInicialConsig3, ValorInicialDemo3, ValorInicialConserto3, ValorInicialOutros3,ValorInicialBenef3) VALUES (?,?,?,?,?,?,?,?)", iAno + 1, tSldMesEst1.iFilialEmpresa, tSldMesEst1.sProduto, dValorAcumuladaConsig3, dValorAcumuladaDemo3, dValorAcumuladaConserto3, dValorAcumuladaOutras3, dValorAcumuladaBenef3)
            If lErro <> AD_SQL_SUCESSO Then gError 40691
                
        End If
        
        'Atualiza tela de acompanhamento do Batch
        lErro = Rotina_CMP_AtualizaTelaBatch()
        If lErro <> SUCESSO Then gError 89840
        
        '-----------------Busca o Proximo -------------------
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89841
    
    Loop
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Rotina_CMP_Transfere_SldMesEst1_ValorInicial = SUCESSO

    Exit Function

Erro_Rotina_CMP_Transfere_SldMesEst1_ValorInicial:

    Rotina_CMP_Transfere_SldMesEst1_ValorInicial = gErr

    Select Case gErr
                
        Case 89833
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89834, 89835, 89841
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST1_2", gErr, iAno, tSldMesEst1.iFilialEmpresa)
        
        Case 89836, 89837
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, iAno + 1, tSldMesEst1.iFilialEmpresa, tSldMesEst1.sProduto)
        
        Case 89838
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDOMESEST1_NAO_CADASTRADO", gErr, tSldMesEst1.iFilialEmpresa, iAno + 1, tSldMesEst1.sProduto)
         
        Case 89839
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESEST1", gErr, iAno + 1, tSldMesEst1.iFilialEmpresa, tSldMesEst1.sProduto)
        
        Case 89840
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158696)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Rotina_CMP_Transfere_SldMesEst2_ValorInicial(iAno As Integer) As Long
'Atualiza Valor Inicial para o Ano sequinte na tabela SaldoMesEst2
'Chamada EM TRANSAÇÃO

Dim lErro As Long
Dim iMesFinal As Integer
Dim iIndice As Integer
Dim sComandoSQL(1 To 3) As String
Dim alComando(1 To 4) As Long
Dim tSldMesEst2 As typeSldMesEst2
Dim dValorAcumuladaConsig As Double, dValorAcumuladaDemo As Double, dValorAcumuladaConserto As Double, dValorAcumuladaOutras As Double, dValorAcumuladaBenef As Double
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEst2_ValorInicial

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69740
    Next

    'Le os Valores de Entrada e de Saida para todos os meses
    'Quantidade e valor inicial
    sComandoSQL(1) = "SELECT FilialEmpresa, ValorInicialConsig, ValorInicialDemo, ValorInicialConserto, ValorInicialOutros, ValorInicialBenef, "
    'Quantidades e valores de entrada e de saida mensais
    For iIndice = 1 To 12
        sComandoSQL(1) = sComandoSQL(1) & "SaldoValorConsig" & CStr(iIndice) & ", " & "SaldoValorDemo" & CStr(iIndice) & ", " & "SaldoValorConserto" & CStr(iIndice) & ", " & "SaldoValorOutros" & CStr(iIndice) & ", " & "SaldoValorBenef" & CStr(iIndice) & ", "
    Next
    sComandoSQL(1) = sComandoSQL(1) & "Produto "
    'Tabela, Filtro, Ordem
    sComandoSQL(1) = sComandoSQL(1) & "FROM Produtos, SldMesEst2 WHERE Produtos.Codigo = SldMesEst2.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? ORDER BY Produtos.Codigo"
    
    
    With tSldMesEst2
        .sProduto = String(STRING_PRODUTO, 0)
        
        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .iFilialEmpresa, .dValorInicialConsig, .dValorInicialDemo, .dValorInicialConserto, .dValorInicialOutros, .dValorInicialBenef, .adSaldoValorConsig(1), .adSaldoValorDemo(1), .adSaldoValorConserto(1), .adSaldoValorOutros(1), .adSaldoValorBenef(1), .adSaldoValorConsig(2), .adSaldoValorDemo(2), .adSaldoValorConserto(2), .adSaldoValorOutros(2), .adSaldoValorBenef(2), .adSaldoValorConsig(3), .adSaldoValorDemo(3), .adSaldoValorConserto(3), .adSaldoValorOutros(3), .adSaldoValorBenef(3), .adSaldoValorConsig(4), .adSaldoValorDemo(4), .adSaldoValorConserto(4), .adSaldoValorOutros(4), .adSaldoValorBenef(4), .adSaldoValorConsig(5), .adSaldoValorDemo(5), .adSaldoValorConserto(5), .adSaldoValorOutros(5), .adSaldoValorBenef(5), .adSaldoValorConsig(6), .adSaldoValorDemo(6), .adSaldoValorConserto(6), .adSaldoValorOutros(6), .adSaldoValorBenef(6), _
        .adSaldoValorConsig(7), .adSaldoValorDemo(7), .adSaldoValorConserto(7), .adSaldoValorOutros(7), .adSaldoValorBenef(7), .adSaldoValorConsig(8), .adSaldoValorDemo(8), .adSaldoValorConserto(8), .adSaldoValorOutros(8), .adSaldoValorBenef(8), .adSaldoValorConsig(9), .adSaldoValorDemo(9), .adSaldoValorConserto(9), .adSaldoValorOutros(9), .adSaldoValorBenef(9), .adSaldoValorConsig(10), .adSaldoValorDemo(10), .adSaldoValorConserto(10), .adSaldoValorOutros(10), .adSaldoValorBenef(10), .adSaldoValorConsig(11), .adSaldoValorDemo(11), .adSaldoValorConserto(11), .adSaldoValorOutros(11), .adSaldoValorBenef(11), .adSaldoValorConsig(12), .adSaldoValorDemo(12), .adSaldoValorConserto(12), .adSaldoValorOutros(12), .adSaldoValorBenef(12), .sProduto, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno)
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 69741
   
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69742
    
    Do While lErro <> AD_SQL_SEM_DADOS
            
        'Quantitade Inicial
        sComandoSQL(2) = "SELECT FilialEmpresa "
        'Tabela, Filtro, Ordem
        sComandoSQL(2) = sComandoSQL(2) & "FROM SldMesEst2 WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ? ORDER BY Produto"

        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
        '----------------------------------------------------------------------------------------
        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, iFilialEmpresa, iAno + 1, tSldMesEst2.iFilialEmpresa, tSldMesEst2.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69743
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69744
        
        With tSldMesEst2
            dValorAcumuladaBenef = .dValorInicialBenef
            dValorAcumuladaConserto = .dValorInicialConserto
            dValorAcumuladaConsig = .dValorInicialConsig
            dValorAcumuladaDemo = .dValorInicialDemo
            dValorAcumuladaOutras = .dValorInicialOutros
        End With
        
        'Meses
        For iIndice = 1 To 12
            dValorAcumuladaBenef = dValorAcumuladaBenef + tSldMesEst2.adSaldoValorBenef(iIndice)
            dValorAcumuladaConserto = dValorAcumuladaConserto + tSldMesEst2.adSaldoValorConserto(iIndice)
            dValorAcumuladaConsig = dValorAcumuladaConsig + tSldMesEst2.adSaldoValorConsig(iIndice)
            dValorAcumuladaDemo = dValorAcumuladaDemo + tSldMesEst2.adSaldoValorDemo(iIndice)
            dValorAcumuladaOutras = dValorAcumuladaOutras + tSldMesEst2.adSaldoValorOutros(iIndice)
        Next
        
        '------------Atualiza SldMesEst para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
            
        If lErro = AD_SQL_SUCESSO Then
            
            'Monta comando SQL para UPDATE de SldMesEst
            sComandoSQL(3) = "UPDATE SldMesEst2 SET ValorInicialConsig = ?, ValorInicialDemo = ?, ValorInicialConserto = ?, ValorInicialOutros = ?, ValorInicialBenef = ?"
            
            'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
            lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorAcumuladaConsig, dValorAcumuladaDemo, dValorAcumuladaConserto, dValorAcumuladaOutras, dValorAcumuladaBenef)
            If lErro <> AD_SQL_SUCESSO Then gError 69746
        
        Else
        
            'Insere os dados no BD
            lErro = Comando_Executar(alComando(4), "INSERT INTO SldMesEst2 (Ano, FilialEmpresa, Produto, ValorInicialConsig, ValorInicialDemo, ValorInicialConserto, ValorInicialOutros,ValorInicialBenef) VALUES (?,?,?,?,?,?,?,?)", iAno + 1, tSldMesEst2.iFilialEmpresa, tSldMesEst2.sProduto, dValorAcumuladaConsig, dValorAcumuladaDemo, dValorAcumuladaConserto, dValorAcumuladaOutras, dValorAcumuladaBenef)
            If lErro <> AD_SQL_SUCESSO Then gError 40691
                
        End If
        
        
        'Atualiza tela de acompanhamento do Batch
        lErro = Rotina_CMP_AtualizaTelaBatch()
        If lErro <> SUCESSO Then gError 69747
        
        '-----------------Busca o Proximo -------------------
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69748
    
    Loop
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Rotina_CMP_Transfere_SldMesEst2_ValorInicial = SUCESSO

    Exit Function

Erro_Rotina_CMP_Transfere_SldMesEst2_ValorInicial:

    Rotina_CMP_Transfere_SldMesEst2_ValorInicial = gErr

    Select Case gErr
                
        Case 69740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69741, 69742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST2_2", gErr, iAno, tSldMesEst2.iFilialEmpresa)
        
        Case 69743, 69744, 69748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno + 1, tSldMesEst2.iFilialEmpresa, tSldMesEst2.sProduto)
        
        Case 69745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDOMESEST2_NAO_CADASTRADO", gErr, tSldMesEst2.iFilialEmpresa, iAno + 1, tSldMesEst2.sProduto)
         
        Case 69746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESEST1", gErr, iAno + 1, tSldMesEst2.iFilialEmpresa, tSldMesEst2.sProduto)
        
        Case 69747
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158697)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial(iAno As Integer) As Long
'Atualiza Valor Inicial para o Ano sequinte na tabela SaldoMesEstAlm2
'Chamada EM TRANSAÇÃO

Dim lErro As Long
Dim iMesFinal As Integer
Dim iIndice As Integer
Dim sComandoSQL(1 To 3) As String
Dim alComando(1 To 4) As Long
Dim tSldMesEstAlm1 As typeSldMesEstAlm1
Dim iAlmoxarifado As Integer
Dim dValorAcumuladaConsig3 As Double, dValorAcumuladaDemo3 As Double, dValorAcumuladaConserto3 As Double, dValorAcumuladaOutras3 As Double, dValorAcumuladaBenef3 As Double

On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89843
    Next

    'Le os Valores de Entrada e de Saida para todos os meses
    'Quantidade e valor inicial
    sComandoSQL(1) = "SELECT ValorInicialConsig3, ValorInicialDemo3, ValorInicialConserto3, ValorInicialOutros3, ValorInicialBenef3, "
    'Quantidades e valores de entrada e de saida mensais
    For iIndice = 1 To 12
        sComandoSQL(1) = sComandoSQL(1) & "SaldoValorConsig3" & CStr(iIndice) & ", " & "SaldoValorDemo3" & CStr(iIndice) & ", " & "SaldoValorConserto3" & CStr(iIndice) & ", " & "SaldoValorOutros3" & CStr(iIndice) & ", " & "SaldoValorBenef3" & CStr(iIndice) & ", "
    Next
    sComandoSQL(1) = sComandoSQL(1) & " Produto, "
    sComandoSQL(1) = sComandoSQL(1) & " Almoxarifado "
    'Tabela, Filtro, Ordem
    sComandoSQL(1) = sComandoSQL(1) & " FROM Produtos, SldMesEstAlm1 WHERE Produtos.Codigo = SldMesEstAlm1.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? ORDER BY Produtos.Codigo"
    
    With tSldMesEstAlm1
        .sProduto = String(STRING_PRODUTO, 0)
        
        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .dValorInicialConsig3, .dValorInicialDemo3, .dValorInicialConserto3, .dValorInicialOutros3, .dValorInicialBenef3, .adSaldoValorConsig3(1), .adSaldoValorDemo3(1), .adSaldoValorConserto3(1), .adSaldoValorOutros3(1), .adSaldoValorBenef3(1), .adSaldoValorConsig3(2), .adSaldoValorDemo3(2), .adSaldoValorConserto3(2), .adSaldoValorOutros3(2), .adSaldoValorBenef3(2), .adSaldoValorConsig3(3), .adSaldoValorDemo3(3), .adSaldoValorConserto3(3), .adSaldoValorOutros3(3), .adSaldoValorBenef3(3), .adSaldoValorConsig3(4), .adSaldoValorDemo3(4), .adSaldoValorConserto3(4), .adSaldoValorOutros3(4), .adSaldoValorBenef3(4), .adSaldoValorConsig3(5), .adSaldoValorDemo3(5), .adSaldoValorConserto3(5), .adSaldoValorOutros3(5), .adSaldoValorBenef3(5), .adSaldoValorConsig3(6), .adSaldoValorDemo3(6), .adSaldoValorConserto3(6), .adSaldoValorOutros3(6), .adSaldoValorBenef3(6), _
        .adSaldoValorConsig3(7), .adSaldoValorDemo3(7), .adSaldoValorConserto3(7), .adSaldoValorOutros3(7), .adSaldoValorBenef3(7), .adSaldoValorConsig3(8), .adSaldoValorDemo3(8), .adSaldoValorConserto3(8), .adSaldoValorOutros3(8), .adSaldoValorBenef3(8), .adSaldoValorConsig3(9), .adSaldoValorDemo3(9), .adSaldoValorConserto3(9), .adSaldoValorOutros3(9), .adSaldoValorBenef3(9), .adSaldoValorConsig3(10), .adSaldoValorDemo3(10), .adSaldoValorConserto3(10), .adSaldoValorOutros3(10), .adSaldoValorBenef3(10), .adSaldoValorConsig3(11), .adSaldoValorDemo3(11), .adSaldoValorConserto3(11), .adSaldoValorOutros3(11), .adSaldoValorBenef3(11), .adSaldoValorConsig3(12), .adSaldoValorDemo3(12), .adSaldoValorConserto3(12), .adSaldoValorOutros3(12), .adSaldoValorBenef3(12), .sProduto, .iAlmoxarifado, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno)
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 89844
   
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89845
    
    Do While lErro <> AD_SQL_SEM_DADOS
            
        'Quantitade Inicial
        sComandoSQL(2) = "SELECT Almoxarifado "
        'Tabela, Filtro, Ordem
        sComandoSQL(2) = sComandoSQL(2) & "FROM SldMesEstAlm1 WHERE Ano = ? AND Produto = ? AND Almoxarifado = ? ORDER BY Produto"

        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
        '----------------------------------------------------------------------------------------
        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, iAlmoxarifado, iAno + 1, tSldMesEstAlm1.sProduto, tSldMesEstAlm1.iAlmoxarifado)
        If lErro <> AD_SQL_SUCESSO Then gError 89846
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89847
        
        'Meses
        For iIndice = 1 To 12
            dValorAcumuladaBenef3 = dValorAcumuladaBenef3 + tSldMesEstAlm1.adSaldoValorBenef3(iIndice)
            dValorAcumuladaConserto3 = dValorAcumuladaConserto3 + tSldMesEstAlm1.adSaldoValorConserto3(iIndice)
            dValorAcumuladaConsig3 = dValorAcumuladaConsig3 + tSldMesEstAlm1.adSaldoValorConsig3(iIndice)
            dValorAcumuladaDemo3 = dValorAcumuladaDemo3 + tSldMesEstAlm1.adSaldoValorDemo3(iIndice)
            dValorAcumuladaOutras3 = dValorAcumuladaOutras3 + tSldMesEstAlm1.adSaldoValorOutros3(iIndice)
        Next
        
        '------------Atualiza SldMesEst para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
            
        If lErro = AD_SQL_SUCESSO Then
            
            'Monta comando SQL para UPDATE de SldMesEst
            sComandoSQL(3) = "UPDATE SldMesEstAlm1 SET ValorInicialConsig3 = ?, ValorInicialDemo3 = ?, ValorInicialConserto3 = ?, ValorInicialOutros3 = ?, ValorInicialBenef3 = ?"
            
            'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
            lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorAcumuladaConsig3, dValorAcumuladaDemo3, dValorAcumuladaConserto3, dValorAcumuladaOutras3, dValorAcumuladaBenef3)
            If lErro <> AD_SQL_SUCESSO Then gError 89849
        
        Else
        
            'Insere os dados no BD
            lErro = Comando_Executar(alComando(4), "INSERT INTO SldMesEstAlm1 (Almoxarifado, Ano, Produto, ValorInicialConsig3, ValorInicialDemo3, ValorInicialConserto3, ValorInicialOutros3, ValorInicialBenef3) VALUES (?,?,?,?,?,?,?,?)", tSldMesEstAlm1.iAlmoxarifado, iAno + 1, tSldMesEstAlm1.sProduto, dValorAcumuladaConsig3, dValorAcumuladaDemo3, dValorAcumuladaConserto3, dValorAcumuladaOutras3, dValorAcumuladaBenef3)
            If lErro <> AD_SQL_SUCESSO Then gError 40691
        
        End If
        
        
        'Atualiza tela de acompanhamento do Batch
        lErro = Rotina_CMP_AtualizaTelaBatch()
        If lErro <> SUCESSO Then gError 89850
        
        '-----------------Busca o Proximo -------------------
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89851
    
    Loop
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial = SUCESSO

    Exit Function

Erro_Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial:

    Rotina_CMP_Transfere_SldMesEstAlm1_ValorInicial = gErr

    Select Case gErr
                
        Case 89843
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89844, 89845
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM12", gErr, iAno)
        
        Case 89846, 89847, 89851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM11", gErr, iAno + 1, tSldMesEstAlm1.iAlmoxarifado, tSldMesEstAlm1.sProduto)
        
        Case 89848
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESESTALM1_NAO_CADASTRADO", gErr, iAno + 1, tSldMesEstAlm1.sProduto, tSldMesEstAlm1.iAlmoxarifado)
         
        Case 89849
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESESTALM1", gErr, iAno + 1, tSldMesEstAlm1.iAlmoxarifado, tSldMesEstAlm1.sProduto)
        
        Case 89850
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158698)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial(iAno As Integer) As Long
'Atualiza Valor Inicial para o Ano sequinte na tabela SaldoMesEstAlm2
'Chamada EM TRANSAÇÃO

Dim lErro As Long
Dim iMesFinal As Integer
Dim iIndice As Integer
Dim sComandoSQL(1 To 3) As String
Dim alComando(1 To 4) As Long
Dim tSldMesEstAlm2 As typeSldMesEstAlm2
Dim iAlmoxarifado As Integer
Dim dValorAcumuladaConsig As Double, dValorAcumuladaDemo As Double, dValorAcumuladaConserto As Double, dValorAcumuladaOutras As Double, dValorAcumuladaBenef As Double

On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69748
    Next

    'Le os Valores de Entrada e de Saida para todos os meses
    'Quantidade e valor inicial
    sComandoSQL(1) = "SELECT ValorInicialConsig, ValorInicialDemo, ValorInicialConserto, ValorInicialOutros, ValorInicialBenef, "
    'Quantidades e valores de entrada e de saida mensais
    For iIndice = 1 To 12
        sComandoSQL(1) = sComandoSQL(1) & "SaldoValorConsig" & CStr(iIndice) & ", " & "SaldoValorDemo" & CStr(iIndice) & ", " & "SaldoValorConserto" & CStr(iIndice) & ", " & "SaldoValorOutros" & CStr(iIndice) & ", " & "SaldoValorBenef" & CStr(iIndice) & ", "
    Next
    sComandoSQL(1) = sComandoSQL(1) & " Produto, "
    sComandoSQL(1) = sComandoSQL(1) & " Almoxarifado "
    'Tabela, Filtro, Ordem
    sComandoSQL(1) = sComandoSQL(1) & " FROM Produtos, SldMesEstAlm2 WHERE Produtos.Codigo = SldMesEstAlm2.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? ORDER BY Produtos.Codigo"
    
    With tSldMesEstAlm2
        .sProduto = String(STRING_PRODUTO, 0)
        
        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .dValorInicialConsig, .dValorInicialDemo, .dValorInicialConserto, .dValorInicialOutros, .dValorInicialBenef, .adSaldoValorConsig(1), .adSaldoValorDemo(1), .adSaldoValorConserto(1), .adSaldoValorOutros(1), .adSaldoValorBenef(1), .adSaldoValorConsig(2), .adSaldoValorDemo(2), .adSaldoValorConserto(2), .adSaldoValorOutros(2), .adSaldoValorBenef(2), .adSaldoValorConsig(3), .adSaldoValorDemo(3), .adSaldoValorConserto(3), .adSaldoValorOutros(3), .adSaldoValorBenef(3), .adSaldoValorConsig(4), .adSaldoValorDemo(4), .adSaldoValorConserto(4), .adSaldoValorOutros(4), .adSaldoValorBenef(4), .adSaldoValorConsig(5), .adSaldoValorDemo(5), .adSaldoValorConserto(5), .adSaldoValorOutros(5), .adSaldoValorBenef(5), .adSaldoValorConsig(6), .adSaldoValorDemo(6), .adSaldoValorConserto(6), .adSaldoValorOutros(6), .adSaldoValorBenef(6), _
        .adSaldoValorConsig(7), .adSaldoValorDemo(7), .adSaldoValorConserto(7), .adSaldoValorOutros(7), .adSaldoValorBenef(7), .adSaldoValorConsig(8), .adSaldoValorDemo(8), .adSaldoValorConserto(8), .adSaldoValorOutros(8), .adSaldoValorBenef(8), .adSaldoValorConsig(9), .adSaldoValorDemo(9), .adSaldoValorConserto(9), .adSaldoValorOutros(9), .adSaldoValorBenef(9), .adSaldoValorConsig(10), .adSaldoValorDemo(10), .adSaldoValorConserto(10), .adSaldoValorOutros(10), .adSaldoValorBenef(10), .adSaldoValorConsig(11), .adSaldoValorDemo(11), .adSaldoValorConserto(11), .adSaldoValorOutros(11), .adSaldoValorBenef(11), .adSaldoValorConsig(12), .adSaldoValorDemo(12), .adSaldoValorConserto(12), .adSaldoValorOutros(12), .adSaldoValorBenef(12), .sProduto, .iAlmoxarifado, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno)
    End With
    
    If lErro <> AD_SQL_SUCESSO Then gError 69749
   
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69750
    
    Do While lErro <> AD_SQL_SEM_DADOS
            
        'Quantitade Inicial
        sComandoSQL(2) = "SELECT Almoxarifado "
        'Tabela, Filtro, Ordem
        sComandoSQL(2) = sComandoSQL(2) & "FROM SldMesEstAlm2 WHERE Ano = ? AND Produto = ? AND Almoxarifado = ? ORDER BY Produto"

        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
        '----------------------------------------------------------------------------------------
        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, iAlmoxarifado, iAno + 1, tSldMesEstAlm2.sProduto, tSldMesEstAlm2.iAlmoxarifado)
        If lErro <> AD_SQL_SUCESSO Then gError 69751
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69752
        
   
        'Meses
        For iIndice = 1 To 12
            dValorAcumuladaBenef = dValorAcumuladaBenef + tSldMesEstAlm2.adSaldoValorBenef(iIndice)
            dValorAcumuladaConserto = dValorAcumuladaConserto + tSldMesEstAlm2.adSaldoValorConserto(iIndice)
            dValorAcumuladaConsig = dValorAcumuladaConsig + tSldMesEstAlm2.adSaldoValorConsig(iIndice)
            dValorAcumuladaDemo = dValorAcumuladaDemo + tSldMesEstAlm2.adSaldoValorDemo(iIndice)
            dValorAcumuladaOutras = dValorAcumuladaOutras + tSldMesEstAlm2.adSaldoValorOutros(iIndice)
        Next
        
        '------------Atualiza SldMesEst para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
            
        If lErro = AD_SQL_SUCESSO Then
            
            'Monta comando SQL para UPDATE de SldMesEst
            sComandoSQL(3) = "UPDATE SldMesEstAlm2 SET ValorInicialConsig = ?, ValorInicialDemo = ?, ValorInicialConserto = ?, ValorInicialOutros = ?, ValorInicialBenef = ?"
            
            'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
            lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorAcumuladaConsig, dValorAcumuladaDemo, dValorAcumuladaConserto, dValorAcumuladaOutras, dValorAcumuladaBenef)
            If lErro <> AD_SQL_SUCESSO Then gError 69754
        
        Else
        
            'Insere os dados no BD
            lErro = Comando_Executar(alComando(4), "INSERT INTO SldMesEstAlm2 (Almoxarifado, Ano, Produto, ValorInicialConsig, ValorInicialDemo, ValorInicialConserto, ValorInicialOutros, ValorInicialBenef) VALUES (?,?,?,?,?,?,?,?)", tSldMesEstAlm2.iAlmoxarifado, iAno + 1, tSldMesEstAlm2.sProduto, dValorAcumuladaConsig, dValorAcumuladaDemo, dValorAcumuladaConserto, dValorAcumuladaOutras, dValorAcumuladaBenef)
            If lErro <> AD_SQL_SUCESSO Then gError 40691
        
        End If
        
        'Atualiza tela de acompanhamento do Batch
        lErro = Rotina_CMP_AtualizaTelaBatch()
        If lErro <> SUCESSO Then gError 69755
        
        '-----------------Busca o Proximo -------------------
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69756
    
    Loop
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial = SUCESSO

    Exit Function

Erro_Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial:

    Rotina_CMP_Transfere_SldMesEstAlm2_ValorInicial = gErr

    Select Case gErr
                
        Case 69748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69749, 69750
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM22", gErr, iAno)
        
        Case 69751, 69752, 69756
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM21", gErr, iAno + 1, tSldMesEstAlm2.iAlmoxarifado, tSldMesEstAlm2.sProduto)
        
        Case 69753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESESTALM2_NAO_CADASTRADO", gErr, iAno + 1, tSldMesEstAlm2.sProduto, tSldMesEstAlm2.iAlmoxarifado)
         
        Case 69754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESESTALM2", gErr, iAno + 1, tSldMesEstAlm2.iAlmoxarifado, tSldMesEstAlm2.sProduto)
        
        Case 69755
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158699)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Rotina_CMP_TotalTransfereValorInicial(iAno As Integer, iMes As Integer, lTotalProdutos As Long) As Long
'Chamada EM TRANSAÇÃO
'Retorna o Número Total de Produtos que terao valores inicial tranferidos

Dim lComando As Long
Dim lErro As Long
Dim sProduto As String
Dim sComandoSQL(1 To 2) As String
Dim lSubTotal1 As Long, lSubTotal2 As Long

On Error GoTo Erro_Rotina_CMP_TotalTransfereValorInicial

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 69013
    
    'Soma a Quantidade de Produtos de SldMesEst que terão os valores iniciais transferidos
    sComandoSQL(1) = "SELECT COUNT(*) FROM Produtos, SldMesEst WHERE Produtos.Codigo = SldMesEst.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND SldMesEst.Ano = ?"
    
    lErro = Comando_Executar(lComando, sComandoSQL(1), lSubTotal1, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno)
    If lErro <> AD_SQL_SUCESSO Then gError 69014

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69015

    'Soma a Quantidade de Produtos de SldMesEstAlm que teram os valores iniciais transferidos
    sComandoSQL(2) = "SELECT COUNT(*) FROM Produtos, SldMesEstAlm WHERE Produtos.Codigo = SldMesEstAlm.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) and SldMesEstAlm.Ano = ?"
    
    lErro = Comando_Executar(lComando, sComandoSQL(2), lSubTotal2, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno)
    If lErro <> AD_SQL_SUCESSO Then gError 69022

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69023
    
    lTotalProdutos = lSubTotal1 + lSubTotal2
    
    'Fechamento comando
    Call Comando_Fechar(lComando)
    
    Rotina_CMP_TotalTransfereValorInicial = SUCESSO

    Exit Function

Erro_Rotina_CMP_TotalTransfereValorInicial:

    Rotina_CMP_TotalTransfereValorInicial = gErr

    Select Case gErr

        Case 69013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69014, 69015, 69016
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL(1))
        
        Case 69022, 69023, 69024
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sComandoSQL(2))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158700)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Private Function Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial(iAno As Integer) As Long
'Atualiza Valor Inicial e o CustoMedioProducaoInicial para o Ano sequinte na tabela SaldoMesEstAlm
'Chamada EM TRANSAÇÃO

Dim lErro As Long
Dim dValorInicialProxAno As Double
Dim iMesFinal As Integer
Dim iIndice As Integer
Dim sComandoSQL(1 To 3) As String
Dim alComando(1 To 4) As Long
Dim tSldMesEstAlm As typeSldMesEstAlm

On Error GoTo Erro_Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69026
    Next

    'Le os Valores de Entrada e de Saida para todos os meses
    sComandoSQL(1) = "SELECT "
    
    'valor inicial
    sComandoSQL(1) = sComandoSQL(1) & "ValorInicial, ValorInicialCusto, "
    
    'valores de entrada e de saida mensais
    For iIndice = 1 To 12
        sComandoSQL(1) = sComandoSQL(1) & "ValorEnt" & CStr(iIndice) & ", " & "ValorSai" & CStr(iIndice) & ", "
    Next
    
    For iIndice = 1 To 12
        sComandoSQL(1) = sComandoSQL(1) & "SaldoValorCusto" & CStr(iIndice) & ", "
    Next
    
    sComandoSQL(1) = sComandoSQL(1) & "Produto "
    sComandoSQL(1) = sComandoSQL(1) & ",Almoxarifado "
    
    'Tabela, Filtro, Ordem
    sComandoSQL(1) = sComandoSQL(1) & "FROM Produtos, SldMesEstAlm WHERE Produtos.Codigo = SldMesEstAlm.Produto AND (Produtos.Apropriacao = ? OR Produtos.Apropriacao = ?) AND Ano = ? ORDER BY Produtos.Codigo"
    
    With tSldMesEstAlm
        
        .sProduto = String(STRING_PRODUTO, 0)
        lErro = Comando_Executar(alComando(1), sComandoSQL(1), .dValorInicial, .dValorInicialCusto, .adValorEnt(1), .adValorSai(1), .adValorEnt(2), .adValorSai(2), .adValorEnt(3), .adValorSai(3), .adValorEnt(4), .adValorSai(4), .adValorEnt(5), .adValorSai(5), .adValorEnt(6), .adValorSai(6), .adValorEnt(7), .adValorSai(7), .adValorEnt(8), .adValorSai(8), .adValorEnt(9), .adValorSai(9), .adValorEnt(10), .adValorSai(10), .adValorEnt(11), .adValorSai(11), .adValorEnt(12), .adValorSai(12), _
        .adSaldoValorCusto(1), .adSaldoValorCusto(2), .adSaldoValorCusto(3), .adSaldoValorCusto(4), .adSaldoValorCusto(5), .adSaldoValorCusto(6), .adSaldoValorCusto(7), .adSaldoValorCusto(8), .adSaldoValorCusto(9), .adSaldoValorCusto(10), .adSaldoValorCusto(11), .adSaldoValorCusto(12), _
        .sProduto, .iAlmoxarifado, APROPR_CUSTO_REAL, APROPR_CUSTO_MEDIO_PRODUCAO, iAno)
        If lErro <> AD_SQL_SUCESSO Then gError 69027
    
    End With
   
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69028
    
    Do While lErro <> AD_SQL_SEM_DADOS
            
        'Valor Inicial
        sComandoSQL(2) = "SELECT ValorInicial "
        'Tabela, Filtro, Ordem
        sComandoSQL(2) = sComandoSQL(2) & "FROM SldMesEstAlm WHERE Ano = ? AND Almoxarifado = ? AND Produto = ? ORDER BY Produto"

        'Busca primeiro registro de SaldoMesEst para atualizar Valores Iniciais para o Ano sequinte
        '----------------------------------------------------------------------------------------
        lErro = Comando_ExecutarPos(alComando(2), sComandoSQL(2), 0, dValorInicialProxAno, iAno + 1, tSldMesEstAlm.iAlmoxarifado, tSldMesEstAlm.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 69029
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69030
        
   
        'Cálculo de saldo de valor
        '-------------------------
        
        dValorInicialProxAno = tSldMesEstAlm.dValorInicial
        
        'Meses
        For iIndice = 1 To 12
            dValorInicialProxAno = dValorInicialProxAno + tSldMesEstAlm.adValorEnt(iIndice) - tSldMesEstAlm.adValorSai(iIndice)
        Next
            
        'Meses
        For iIndice = 1 To 12
            tSldMesEstAlm.dValorInicialCusto = tSldMesEstAlm.dValorInicialCusto + tSldMesEstAlm.adSaldoValorCusto(iIndice)
        Next
            
        '------------Atualiza SldMesEstAlm para o Valor Inicial do próximo Ano e o CustoMedioProducaoInicial
        
        If lErro = AD_SQL_SUCESSO Then
        
            'Monta comando SQL para UPDATE de SldMesEstAlm
            sComandoSQL(3) = "UPDATE SldMesEstAlm SET ValorInicial = ?, ValorInicialCusto = ?"
            
            'Atualiza tabela SaldoMesEst, campos ValorEnt, ValorSai, ValorCons de iMes
            lErro = Comando_ExecutarPos(alComando(3), sComandoSQL(3), alComando(2), dValorInicialProxAno, tSldMesEstAlm.dValorInicialCusto)
            If lErro <> AD_SQL_SUCESSO Then gError 69032
        
        Else
        
            'Insere os dados no BD
            lErro = Comando_Executar(alComando(4), "INSERT INTO SldMesEstAlm (Ano, Almoxarifado, Produto, ValorInicial, ValorInicialCusto) VALUES (?,?,?,?,?)", iAno + 1, tSldMesEstAlm.iAlmoxarifado, tSldMesEstAlm.sProduto, dValorInicialProxAno, tSldMesEstAlm.dValorInicialCusto)
            If lErro <> AD_SQL_SUCESSO Then gError 40691
        
        End If
        
        'Atualiza tela de acompanhamento do Batch
        lErro = Rotina_CMP_AtualizaTelaBatch()
        If lErro <> SUCESSO Then gError 69033
        
        '-----------------Busca o Proximo -------------------
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69034
    
    Loop
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial = SUCESSO

    Exit Function

Erro_Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial:

    Rotina_CMP_Transfere_SldMesEstAlm_ValorInicial = gErr

    Select Case gErr
                
        Case 69026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69027, 69028, 69034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM2", gErr, iAno)
        
        Case 69029, 69030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM2", gErr, iAno + 1)
        
        Case 69031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_REGISTRO_SLDMESESTALM", gErr, iAno + 1, tSldMesEstAlm.sProduto, tSldMesEstAlm.iAlmoxarifado)
         
        Case 69032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESESTALM", gErr, iAno, tSldMesEstAlm.iAlmoxarifado, tSldMesEstAlm.sProduto)
        
        Case 69033
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158701)

    End Select
    
   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Rotina_CMP_TotalMovEstoque(iAno As Integer, iMes As Integer, lTotalMovs As Long) As Long
'Chamada EM TRANSAÇÃO
'retorna o número de movimentos de estoque que participam do processamento de custo de produção

Dim lComando As Long
Dim lErro As Long
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim iDiasMes As Integer
Dim sComandoSQL As String
Dim sProduto As String

On Error GoTo Erro_Rotina_CMP_TotalMovEstoque

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92606

    'Determinação de faixa de datas
    dtDataInicial = CDate("1/" & CStr(iMes) & "/" & CStr(iAno))
    iDiasMes = Dias_Mes(iMes, iAno)
    dtDataFinal = CDate(CStr(iDiasMes) & "/" & CStr(iMes) & "/" & CStr(iAno))

    lTotalMovs = 0

    sComandoSQL = "SELECT Count(*) FROM MovimentoEstoque, Produtos WHERE MovimentoEstoque.Produto = Produtos.Codigo AND Data >= ? AND Data <= ? AND Produtos.Apropriacao = ?"

    lErro = Comando_Executar(lComando, sComandoSQL, lTotalMovs, dtDataInicial, dtDataFinal, APROPR_CUSTO_REAL)
    If lErro <> AD_SQL_SUCESSO Then gError 92607

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92608

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Rotina_CMP_TotalMovEstoque = SUCESSO

    Exit Function

Erro_Rotina_CMP_TotalMovEstoque:

    Rotina_CMP_TotalMovEstoque = gErr

    Select Case gErr

        Case 92606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92607, 92608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158702)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function MovtoProcAntecipado_Verifica(tMovEstoque As typeItemMovEstoque, colMovsAntecipados As Collection, bAntecipado As Boolean) As Long
'identifica se o movimento foi processado antecipadamente
'Se foi, deve ser excluido do controle, seja ele uma colecao ou uma tabela no bd

Dim iIndice As Integer, sChave As String

On Error GoTo Erro_MovtoProcAntecipado_Verifica

    bAntecipado = False
            
'    If tMovEstoque.iTipoMov = MOV_EST_REQ_PRODUCAO Or tMovEstoque.iTipoMov = MOV_EST_REQ_PRODUCAO_BENEF3 Then
    
        sChave = "K" & CStr(tMovEstoque.lNumIntDoc)
        
        For iIndice = 1 To colMovsAntecipados.Count
        
            If colMovsAntecipados.Item(iIndice) = sChave Then
            
                bAntecipado = True
                Exit For
                
            End If
        
        Next
        
'    End If
        
    MovtoProcAntecipado_Verifica = SUCESSO
     
    Exit Function
    
Erro_MovtoProcAntecipado_Verifica:

    MovtoProcAntecipado_Verifica = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158703)
     
    End Select
     
    Exit Function

End Function

Function Insumos_Processa(ByVal objItemMovEst As ClassItemMovEstoque, colMovsAntecipados As Collection, alComando1() As Long, ByVal objEstoqueMes As ClassEstoqueMes, ByVal colEstoqueMesProduto As Collection, dtDataMovsAntecipados As Date) As Long
'Processa antecipadamente as requisicoes apropriadas à producao

Dim lErro As Long, objApropriacaoInsumo As ClassApropriacaoInsumosProd, dtDataInsumo As Date
Dim iIndice As Integer, sChave As String, bAntecipado As Boolean

On Error GoTo Erro_Insumos_Processa

    For Each objApropriacaoInsumo In objItemMovEst.colApropriacaoInsumo
    
        bAntecipado = False
        sChave = "K" & CStr(objApropriacaoInsumo.lNumIntReqProd)
        
        For iIndice = 1 To colMovsAntecipados.Count
        
            If colMovsAntecipados.Item(iIndice) = sChave Then
            
                bAntecipado = True
                Exit For
                
            End If
        
        Next
        
        If bAntecipado = False Then
        
            dtDataInsumo = DATA_NULA
            
            lErro = Movto_Processa_Custo(objApropriacaoInsumo.lNumIntReqProd, alComando1, False, objItemMovEst.dtData, objEstoqueMes, colEstoqueMesProduto, dtDataInsumo)
            If lErro <> SUCESSO Then gError 81902
            
            If dtDataInsumo = objItemMovEst.dtData Then
            
'                If objItemMovEst.dCusto = 0 Then
'                    MsgBox ("bbb")
'                End If
                
                lErro = MovtoProcAntecipado_Inclui(colMovsAntecipados, objApropriacaoInsumo.lNumIntReqProd, dtDataMovsAntecipados, objItemMovEst.dtData)
                If lErro <> SUCESSO Then gError 81903
                
            End If
        
        End If
                
    Next
    
    Insumos_Processa = SUCESSO
     
    Exit Function
    
Erro_Insumos_Processa:

    Insumos_Processa = gErr
     
    Select Case gErr
          
        Case 81902, 81903
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158704)
     
    End Select
     
    Exit Function

End Function

Private Function MovtoProcAntecipado_Inclui(colMovsAntecipados As Collection, ByVal lNumIntDoc As Long, dtDataMovsAntecipados As Date, ByVal dtDataMov As Date) As Long
'inclui no controle de movimentos processados antecipadamente o identificado por lNumIntDoc

Dim iIndice As Integer, sChave As String, bAchou As Boolean

On Error GoTo Erro_MovtoProcAntecipado_Inclui

    'quando troca de data pode esvaziar a colecao
    If dtDataMov <> dtDataMovsAntecipados Then
    
        dtDataMovsAntecipados = dtDataMov
        Set colMovsAntecipados = New Collection
    
    End If
    
    bAchou = False
    
    sChave = "K" & CStr(lNumIntDoc)
    
    For iIndice = 1 To colMovsAntecipados.Count
    
        If colMovsAntecipados.Item(iIndice) = sChave Then
        
            bAchou = True
            Exit For
        
        End If
    
    Next
    
    If bAchou = False Then Call colMovsAntecipados.Add(sChave)
    
    MovtoProcAntecipado_Inclui = SUCESSO
     
    Exit Function
    
Erro_MovtoProcAntecipado_Inclui:

    MovtoProcAntecipado_Inclui = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158705)
     
    End Select
     
    Exit Function

End Function

Private Function Movto_Processa_Custo(ByVal lNumIntDoc As Long, alComando1() As Long, bEstorno As Boolean, ByVal dtDataMovto As Date, ByVal objEstoqueMes As ClassEstoqueMes, ByVal colEstoqueMesProduto As Collection, dtDataMov As Date) As Long
'Processa o tratamento de custo para o movimento de estoque identificado por lNumIntDoc

Dim lErro As Long, lComando As Long
Dim tMovEstoque As typeItemMovEstoque, dtDataMovEst As Date
Dim objTipoMovEstoque As New ClassTipoMovEst
Dim objItemMovEst As New ClassItemMovEstoque

On Error GoTo Erro_Movto_Processa_Custo

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81906
    
    tMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tMovEstoque.sProduto = String(STRING_PRODUTO, 0)
    tMovEstoque.sOPCodigo = String(STRING_OPCODIGO, 0)
    tMovEstoque.sDocOrigem = String(STRING_DOCORIGEM, 0)
    
    'le o movimento de estoque
    With tMovEstoque
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, MovimentoEstoque.Codigo, NumIntDocEst, NumIntDocOrigem, Produto, CodigoOP, Quantidade, SiglaUM, Almoxarifado, TipoMov, Apropriacao, Data, Hora, HorasMaquina, Custo, Fornecedor, MovimentoEstoque.TipoNumIntDocOrigem, DocOrigem FROM MovimentoEstoque WHERE MovimentoEstoque.NumIntDoc = ?", .iFilialEmpresa, .lCodigo, .lNumIntDocEst, .lNumIntDocOrigem, .sProduto, .sOPCodigo, .dQuantidade, .sSiglaUM, .iAlmoxarifado, .iTipoMov, .iApropriacao, .dtData, .dHora, .lHorasMaquina, .dCusto, .lFornecedor, .iTipoNumIntDocOrigem, .sDocOrigem, lNumIntDoc)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 81910
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81907
    
    If lErro = AD_SQL_SUCESSO Then
        
        dtDataMov = tMovEstoque.dtData
        
        If bEstorno = True Or dtDataMovto = tMovEstoque.dtData Then
        
            'processa-lo e em seguida o seu estorno, se houver
            
            objItemMovEst.iFilialEmpresa = tMovEstoque.iFilialEmpresa
            objItemMovEst.lCodigo = tMovEstoque.lCodigo
            objItemMovEst.lNumIntDoc = lNumIntDoc
            objItemMovEst.lNumIntDocEst = tMovEstoque.lNumIntDocEst
            objItemMovEst.lNumIntDocOrigem = tMovEstoque.lNumIntDocOrigem
            objItemMovEst.sProduto = tMovEstoque.sProduto
            objItemMovEst.sOPCodigo = tMovEstoque.sOPCodigo
            objItemMovEst.dQuantidade = tMovEstoque.dQuantidade
            objItemMovEst.sSiglaUM = tMovEstoque.sSiglaUM
            objItemMovEst.iAlmoxarifado = tMovEstoque.iAlmoxarifado
            objItemMovEst.iTipoMov = tMovEstoque.iTipoMov
            objItemMovEst.iApropriacao = tMovEstoque.iApropriacao
            objItemMovEst.dtData = tMovEstoque.dtData
            objItemMovEst.dtHora = tMovEstoque.dHora
            objItemMovEst.lHorasMaquina = tMovEstoque.lHorasMaquina
            objItemMovEst.dCusto = tMovEstoque.dCusto
            objItemMovEst.lFornecedor = tMovEstoque.lFornecedor
            objItemMovEst.iTipoNumIntDocOrigem = tMovEstoque.iTipoNumIntDocOrigem
            objItemMovEst.sDocOrigem = tMovEstoque.sDocOrigem
            
            Set objItemMovEst.colApropriacaoInsumo = New Collection
            
            'guarda o custo anterior para poder guardar os valores em estoqueproduto pela diferença. Permite reexecução da apuração.
            objItemMovEst.dCustoAnt = objItemMovEst.dCusto
            
            lErro = CF("Estoque_ApuraCustoProducao", alComando1, objItemMovEst, objEstoqueMes, colEstoqueMesProduto)
            If lErro <> SUCESSO Then gError 81908
            
            'se o movto possui um estorno
            If bEstorno = False And objItemMovEst.lNumIntDocEst <> 0 Then
            
                lErro = Movto_Processa_Custo(objItemMovEst.lNumIntDocEst, alComando1, True, objItemMovEst.dtData, objEstoqueMes, colEstoqueMesProduto, dtDataMovEst)
                If lErro <> SUCESSO Then gError 81909
            
            End If
        
        End If
        
    End If
    
    Call Comando_Fechar(lComando)
    
    Movto_Processa_Custo = SUCESSO
     
    Exit Function
    
Erro_Movto_Processa_Custo:

    Movto_Processa_Custo = gErr
     
    Select Case gErr
          
        Case 81906
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 81907, 81910
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVEST_PROC_CUSTO", gErr)
        
        Case 81908, 81909
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158706)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function TransfFilialSaida_Processa(ByVal objItemMovEstEntTransf As ClassItemMovEstoque, colMovsAntecipados As Collection, alComando1() As Long, ByVal objEstoqueMes As ClassEstoqueMes, ByVal colEstoqueMesProduto As Collection, dtDataMovsAntecipados As Date) As Long
'Processa antecipadamente a saida por transferencia entre filiais associada a objItemMovEstEntTrans (o movto de entrada por transferencia)

Dim lErro As Long, objItemMovEstSaiTransf As New ClassItemMovEstoque, alComando(1 To 2) As Long, iIndice As Integer
Dim sChave As String, bAntecipado As Boolean, dtDataDummy As Date
Dim dtDataEmissao As Date, sSerie As String, lNumNotaFiscal As Long, dtData As Date, lNumIntDoc As Long

On Error GoTo Erro_TransfFilialSaida_Processa

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 106500
    Next
    
    If objItemMovEstEntTransf.lNumIntDocOrigem <> 0 And objItemMovEstEntTransf.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL Then
    
        sSerie = String(STRING_SERIE, 0)
        
        'obtem dados da nf de entrada
        lErro = Comando_Executar(alComando(1), "SELECT DataEmissao, Serie, NumNotaFiscal FROM NFiscal, ItensNFiscal WHERE ItensNFiscal.NumIntDoc = ? AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF", dtDataEmissao, sSerie, lNumNotaFiscal, objItemMovEstEntTransf.lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 106501
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106502
        If lErro <> AD_SQL_SUCESSO Then gError 106503
        
        If (objItemMovEstEntTransf.lNumIntDocOrigem = 395132 Or objItemMovEstEntTransf.lNumIntDocOrigem = 395133) And lNumNotaFiscal = 21066 Then
        
            lNumNotaFiscal = 21035
            dtDataEmissao = CDate("23/06/2015")
            
        End If
    
        Do While lErro = AD_SQL_SUCESSO
           
            'obtem movtos de estoque da saida
            lErro = Comando_Executar(alComando(2), "SELECT MovimentoEstoque.Data, MovimentoEstoque.NumIntDoc FROM NFiscal, ItensNFiscal, MovimentoEstoque, Produtos WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.NumNotaFiscal = ? AND NFiscal.DataEmissao = ? AND NFiscal.Cliente = ? AND MovimentoEstoque.TipoNumIntDocOrigem = ? AND MovimentoEstoque.NumIntDocOrigem = ItensNFiscal.NumIntDoc AND ItensNFiscal.Produto = Produtos.Codigo AND Produtos.Codigo = ?", _
                dtData, lNumIntDoc, lNumNotaFiscal, dtDataEmissao, gobjCRFAT.lCliEmp, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL, objItemMovEstEntTransf.sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 106504
    
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106505
            If lErro <> AD_SQL_SUCESSO Then gError 106506
    
            objItemMovEstSaiTransf.dtData = dtData
            objItemMovEstSaiTransf.lNumIntDoc = lNumIntDoc
    
            If objItemMovEstSaiTransf.dtData = objItemMovEstEntTransf.dtData Then
            
                bAntecipado = False
                sChave = "K" & CStr(objItemMovEstSaiTransf.lNumIntDoc)
                
                For iIndice = 1 To colMovsAntecipados.Count
                
                    If colMovsAntecipados.Item(iIndice) = sChave Then
                    
                        bAntecipado = True
                        Exit For
                        
                    End If
                
                Next
                
                If bAntecipado = False Then
                
                    dtDataDummy = DATA_NULA
                    
                    lErro = Movto_Processa_Custo(objItemMovEstSaiTransf.lNumIntDoc, alComando1, False, objItemMovEstSaiTransf.dtData, objEstoqueMes, colEstoqueMesProduto, dtDataDummy)
                    If lErro <> SUCESSO Then gError 106507
                    
                    lErro = MovtoProcAntecipado_Inclui(colMovsAntecipados, objItemMovEstSaiTransf.lNumIntDoc, dtDataMovsAntecipados, objItemMovEstSaiTransf.dtData)
                    If lErro <> SUCESSO Then gError 106508
                
                End If
        
            End If
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106509
        
        Loop
        
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    TransfFilialSaida_Processa = SUCESSO
     
    Exit Function
    
Erro_TransfFilialSaida_Processa:

    TransfFilialSaida_Processa = gErr
     
    Select Case gErr
    
        Case 106500
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 106501, 106502, 106509
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFO_MOVORIG", gErr)
        
        Case 106503
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFO_MOVORIG2", gErr)
        
        Case 106504, 106505
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFO_MOVORIG3", gErr)
        
        Case 106506
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFO_MOVORIG4", gErr, lNumNotaFiscal, dtDataEmissao, objItemMovEstEntTransf.sProduto)
        
        Case 106507, 106508
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158707)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
     
    Exit Function

End Function

Private Function CustoProd_Reprocessa_Contabilizacao(ByVal dtDataInicial As Date, ByVal dtDataFinal As Date) As Long
'Reprocessa a contabilizacao de todos os movimentos de estoque afetados pelo custo de producao
'Os movimentos de estoque sao tratados como um todo (identificado por filialempresa e codigo) e nao item a item

Dim lErro As Long, iFilialEmpresa As Integer, iFilialAnterior As Integer
Dim iOrigemLcto As Integer, lCodigo As Long, lCodigoAnterior As Long
Dim lNumIntDocOrigemCTB As Long
Dim iTipoNumIntDocOrigem As Integer, lNumIntDocOrigem As Long, lNumIntDoc As Long
Dim colExercicio As New Collection, lComando As Long

On Error GoTo Erro_CustoProd_Reprocessa_Contabilizacao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106525
    
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, MovimentoEstoque.Codigo, TipoNumIntDocOrigem, NumIntDocOrigem, NumIntDoc FROM MovimentoEstoque, Produtos WHERE MovimentoEstoque.Produto = Produtos.Codigo AND Data >= ? AND Data <= ? AND Produtos.Apropriacao = ? ORDER BY FilialEmpresa, MovimentoEstoque.Codigo", _
        iFilialEmpresa, lCodigo, iTipoNumIntDocOrigem, lNumIntDocOrigem, lNumIntDoc, dtDataInicial, dtDataFinal, APROPR_CUSTO_REAL)
    If lErro <> AD_SQL_SUCESSO Then gError 106526
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106527
    
    Do While lErro = AD_SQL_SUCESSO
    
        'se é o 1o item de um movimento de estoque
        If iFilialEmpresa <> iFilialAnterior Or lCodigo <> lCodigoAnterior Then
        
            
            'devolve o documento que originou o movimento de estoque.
            'Utilizado para descobrir os lançamentos contábeis associados e reprocessá-los.
            lErro = Retorna_Origem_Estoque_Contab(iTipoNumIntDocOrigem, lNumIntDocOrigem, lNumIntDoc, iOrigemLcto, lNumIntDocOrigemCTB)
            If lErro <> SUCESSO Then gError 83591
        
            'reprocessa a contabilização do movimento de estoque
            lErro = CF("Rotina_Reprocessamento_DocOrigem", iOrigemLcto, lNumIntDocOrigemCTB, colExercicio, iFilialEmpresa)
            If lErro <> SUCESSO Then gError 83592

            'guarda a identificacao do movto processado
            iFilialAnterior = iFilialEmpresa
            lCodigoAnterior = lCodigo

        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106528
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    CustoProd_Reprocessa_Contabilizacao = SUCESSO
     
    Exit Function
    
Erro_CustoProd_Reprocessa_Contabilizacao:

    CustoProd_Reprocessa_Contabilizacao = gErr
     
    Select Case gErr
          
        Case 106525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 106526, 106527, 106528
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
        
        Case 83591, 83592
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158708)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Private Function CustoProd_LimpaApropriacoesAutomaticas(ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, alComando() As Long) As Long
'elimina apropriacoes automaticas criadas em execucao anterior da rotina de custo

Dim lErro As Long, lNumIntDoc As Long

On Error GoTo Erro_CustoProd_LimpaApropriacoesAutomaticas

    lErro = Comando_Executar(alComando(5), "DELETE FROM ApropriacaoInsumosProd WHERE Automatico = 1 AND NumIntDocOrigem IN (SELECT NumIntDoc FROM MovimentoEstoque WHERE Data >= ? AND Data <= ?)", dtDataInicial, dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 106530
    
    CustoProd_LimpaApropriacoesAutomaticas = SUCESSO
     
    Exit Function
    
Erro_CustoProd_LimpaApropriacoesAutomaticas:

    CustoProd_LimpaApropriacoesAutomaticas = gErr
     
    Select Case gErr
          
        Case 106530
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_APROPRIACOES_AUTOMATICAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158709)
     
    End Select
     
    Exit Function

End Function

Private Function ProdEntradaInsumos_Processa(ByVal objItemMovEst As ClassItemMovEstoque, colMovsAntecipados As Collection, alComando1() As Long, ByVal colEstoqueMes As Collection, ByVal colEstoqueMesProduto As Collection, dtDataMovsAntecipados As Date, ByVal iOrdem As Integer, ByVal lNumIntDocPai As Long) As Long
'Processa antecipadamente as producoes de insumos utilizados para produzir objItemMovEst

Dim lErro As Long, alComando(1 To 4) As Long, iIndice As Integer, objProduto As New ClassProduto
Dim objApropriacaoInsumo As ClassApropriacaoInsumosProd, sChave As String, bAntecipado As Boolean
Dim tMovEstoque As typeItemMovEstoque, objItemMovEst2 As New ClassItemMovEstoque
Dim objTipoMovEstoque As New ClassTipoMovEst, iAchou As Integer
Dim colApropriacaoInsumo As Collection, objEstoqueMes2 As ClassEstoqueMes

On Error GoTo Erro_ProdEntradaInsumos_Processa

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 106868
    Next
    
    For Each objApropriacaoInsumo In objItemMovEst.colApropriacaoInsumo
    
        bAntecipado = False
        sChave = "K" & CStr(objApropriacaoInsumo.lNumIntReqProd)
        
        For iIndice = 1 To colMovsAntecipados.Count
        
            If colMovsAntecipados.Item(iIndice) = sChave Then
            
                bAntecipado = True
                Exit For
                
            End If
        
        Next
        
        If bAntecipado = False Then
        
            objProduto.sCodigo = objApropriacaoInsumo.sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 106869
            If lErro = 28030 Then gError 106870
            
            If objProduto.iCompras = PRODUTO_PRODUZIVEL And objProduto.iApropriacaoCusto = APROPR_CUSTO_REAL Then
            
                'pesquisar entradas nao estornadas que ainda nao tenham sido processadas e antecipa-las
                tMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
                tMovEstoque.sProduto = String(STRING_PRODUTO, 0)
                tMovEstoque.sOPCodigo = String(STRING_OPCODIGO, 0)
                'le o primeiro movimento de estoque
                With tMovEstoque
                    lErro = Comando_Executar(alComando(1), "SELECT FilialEmpresa, MovimentoEstoque.Codigo, NumIntDoc, NumIntDocEst, NumIntDocOrigem, Produto, CodigoOP, Quantidade, SiglaUM, Almoxarifado, TipoMov, MovimentoEstoque.Apropriacao, Data, Hora, MovimentoEstoque.HorasMaquina, Custo, Fornecedor, MovimentoEstoque.TipoNumIntDocOrigem FROM MovimentoEstoque, TiposMovimentoEstoque, TiposOrdemCusto, Produtos WHERE Produtos.Codigo = MovimentoEstoque.Produto AND MovimentoEstoque.TipoMov = TiposMovimentoEstoque.Codigo AND TiposMovimentoEstoque.OrdemCusto = TiposOrdemCusto.Codigo AND Data = ? AND Produtos.Apropriacao = ? AND NumIntDoc > ? AND Produtos.Codigo = ? AND Ordem = ? AND NumIntDocEst = 0 AND (Hora > ? OR Hora = 0) ORDER BY NumIntDoc", _
                        .iFilialEmpresa, .lCodigo, .lNumIntDoc, .lNumIntDocEst, .lNumIntDocOrigem, .sProduto, .sOPCodigo, .dQuantidade, .sSiglaUM, .iAlmoxarifado, .iTipoMov, .iApropriacao, .dtData, .dHora, .lHorasMaquina, .dCusto, .lFornecedor, .iTipoNumIntDocOrigem, objItemMovEst.dtData, APROPR_CUSTO_REAL, lNumIntDocPai, objApropriacaoInsumo.sProduto, iOrdem, HORA_NAO_ANTECIPAR)
                End With
                If lErro <> AD_SQL_SUCESSO Then gError 106871
            
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106872
    
                Do While lErro = AD_SQL_SUCESSO
                
                    objTipoMovEstoque.iCodigo = tMovEstoque.iTipoMov
                    
                    'ler os dados referentes ao tipo de movimento
                    lErro = CF("TiposMovEst_Le1", alComando(3), objTipoMovEstoque)
                    If lErro <> SUCESSO Then gError 81873
                
                    If objTipoMovEstoque.iAtualizaMovEstoque <> TIPOMOV_EST_ESTORNOMOV Then
                    
'                        '??? apenas p/debug, retirar
'                        If tMovEstoque.sOPCodigo = "8207" Or tMovEstoque.sOPCodigo = "8216" Then
'
'                            MsgBox ("ok")
'
'                        End If
            
                        'identifica se o movto foi processado antecipadamento
                        lErro = MovtoProcAntecipado_Verifica(tMovEstoque, colMovsAntecipados, bAntecipado)
                        If lErro <> SUCESSO Then gError 81904
                        
                        'se o processamento do movto nao foi antecipado
                        If bAntecipado = False Then
                        
                            'processa-lo e em seguida o seu estorno, se houver
                            
                            With objItemMovEst2
                                .iFilialEmpresa = tMovEstoque.iFilialEmpresa
                                .lCodigo = tMovEstoque.lCodigo
                                .lNumIntDoc = tMovEstoque.lNumIntDoc
                                .lNumIntDocEst = tMovEstoque.lNumIntDocEst
                                .lNumIntDocOrigem = tMovEstoque.lNumIntDocOrigem
                                .sProduto = tMovEstoque.sProduto
                                .sOPCodigo = tMovEstoque.sOPCodigo
                                .dQuantidade = tMovEstoque.dQuantidade
                                .sSiglaUM = tMovEstoque.sSiglaUM
                                .iAlmoxarifado = tMovEstoque.iAlmoxarifado
                                .iTipoMov = tMovEstoque.iTipoMov
                                .iApropriacao = tMovEstoque.iApropriacao
                                .dtData = tMovEstoque.dtData
                                .dtHora = tMovEstoque.dHora
                                .lHorasMaquina = tMovEstoque.lHorasMaquina
                                .dCusto = tMovEstoque.dCusto
                                .lFornecedor = tMovEstoque.lFornecedor
                                .iTipoNumIntDocOrigem = tMovEstoque.iTipoNumIntDocOrigem
                            End With
                            
                            lErro = MovtoProcAntecipado_Inclui(colMovsAntecipados, objItemMovEst2.lNumIntDoc, dtDataMovsAntecipados, objItemMovEst2.dtData)
                            If lErro <> SUCESSO Then gError 81903
                            
                            iAchou = 0
                            
                            For Each objEstoqueMes2 In colEstoqueMes
                                If objEstoqueMes2.iFilialEmpresa = objItemMovEst2.iFilialEmpresa Then
                                    iAchou = 1
                                    Exit For
                                End If
                            Next
                            
                            If iAchou = 0 Then gError 106914
                            
                            Set colApropriacaoInsumo = New Collection
                            
                            If objItemMovEst2.iTipoMov = MOV_EST_PRODUCAO Or objItemMovEst2.iTipoMov = MOV_EST_PRODUCAO_BENEF3 Then
                            
                                'Le as Apriações do Item
                                lErro = CF("ApropriacaoInsumo_Le_NumIntDocOrigem", tMovEstoque.lNumIntDoc, colApropriacaoInsumo)
                                If lErro <> SUCESSO Then gError 92554
                                            
                            End If
                                            
                            Set objItemMovEst2.colApropriacaoInsumo = colApropriacaoInsumo
                            
                            'processa movtos de producao entrada dos insumos para a data sendo processada
                            lErro = ProdEntradaInsumos_Processa(objItemMovEst2, colMovsAntecipados, alComando1, colEstoqueMes, colEstoqueMesProduto, dtDataMovsAntecipados, iOrdem, lNumIntDocPai)
                            If lErro <> SUCESSO Then gError 106510
                            
                            'processa movimentos de insumos do produto produzido
                            lErro = Insumos_Processa(objItemMovEst2, colMovsAntecipados, alComando1, objEstoqueMes2, colEstoqueMesProduto, dtDataMovsAntecipados)
                            If lErro <> SUCESSO Then gError 81905
                            
                            'guarda o custo anterior para poder guardar os valores em estoqueproduto pela diferença. Permite reexecução da apuração.
                            objItemMovEst2.dCustoAnt = objItemMovEst2.dCusto
                            
                            lErro = CF("Estoque_ApuraCustoProducao", alComando1, objItemMovEst2, objEstoqueMes2, colEstoqueMesProduto)
                            If lErro <> SUCESSO Then gError 92556
                        
                            'Atualiza tela de acompanhamento do Batch
                            lErro = Rotina_CMP_AtualizaTelaBatch()
                            If lErro <> SUCESSO Then gError 92557
        
                        End If
                        
                    End If
                    
                    lErro = Comando_BuscarProximo(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106873
                
                Loop
                
            End If
            
        End If
        
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    ProdEntradaInsumos_Processa = SUCESSO
     
    Exit Function
    
Erro_ProdEntradaInsumos_Processa:

    ProdEntradaInsumos_Processa = gErr
     
    Select Case gErr
    
        Case 106914
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE4", gErr, objEstoqueMes2.iFilialEmpresa, objEstoqueMes2.iAno, objEstoqueMes2.iMes)
        
        Case 106500
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158710)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
     
    Exit Function

End Function



