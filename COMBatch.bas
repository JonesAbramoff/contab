Attribute VB_Name = "COMBatch"
Option Explicit

Function Rotina_ReqComprasBaixar_Batch_Int(colReqComprasInfo As Collection) As Long
'Baixa as Requisições passadas por colReqComprasInfo

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 8) As Long
Dim lTransacao As Long
Dim tItemReqCompra As typeItemReqCompra
Dim tRequisicaoCompras As typeRequisicaoCompras
Dim lNumIntDoc As Long
Dim lItemConcorrencia As Long
Dim dQuantidadeCotar As Double
Dim objReqCompras As New ClassReqComprasInfo
Dim vbMesRes As VbMsgBoxResult

On Error GoTo Erro_Rotina_ReqComprasBaixar_Batch_Int

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 63376
    Next

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 63377

    For Each objReqCompras In colReqComprasInfo
       
        tRequisicaoCompras.sCcl = String(STRING_CCL, 0)
        tRequisicaoCompras.sOPCodigo = String(STRING_OPCODIGO, 0)
        tRequisicaoCompras.sDigitador = String(STRING_USUARIO_CODIGO, 0)
    
        TelaAcompanhaBatchCOM.dValorTotal = colReqComprasInfo.Count
        'Lê em Requisição Compras todos os dados da Requisição vinda em objReqCompras
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, FilialEmpresa, Codigo, Data, DataEnvio, DataLimite, DataBaixa,Urgente, Requisitante, Digitador, Ccl, OPCodigo ,FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao,TipoTributacao FROM RequisicaoCompraN WHERE FilialEmpresa = ? And Codigo = ? AND Status=0", 0, _
            tRequisicaoCompras.lNumIntDoc, tRequisicaoCompras.iFilialEmpresa, tRequisicaoCompras.lCodigo, tRequisicaoCompras.dtData, tRequisicaoCompras.dtDataEnvio, tRequisicaoCompras.dtDataLimite, tRequisicaoCompras.dtDataBaixa, tRequisicaoCompras.lUrgente, tRequisicaoCompras.lRequisitante, tRequisicaoCompras.sDigitador, tRequisicaoCompras.sCcl, tRequisicaoCompras.sOPCodigo, tRequisicaoCompras.iFilialCompra, tRequisicaoCompras.iTipoDestino, tRequisicaoCompras.lFornCliDestino, tRequisicaoCompras.iFilialDestino, tRequisicaoCompras.lObservacao, tRequisicaoCompras.iTipoTributacao, objReqCompras.iFilialEmpresa, objReqCompras.lCodRequisicao)
        If lErro <> AD_SQL_SUCESSO Then gError 63378

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 63379

        'Se não encontrou --> Erro
        If lErro <> AD_SQL_SUCESSO Then gError 63380
        
        'Faz "LockExclusive" em RequisiçãoCompra
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 63381

        tItemReqCompra.sProduto = String(STRING_PRODUTO, 0)
        tItemReqCompra.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        tItemReqCompra.sCcl = String(STRING_CCL, 0)
        tItemReqCompra.sContaContabil = String(STRING_CONTA, 0)
        tItemReqCompra.sUM = String(STRING_UM_SIGLA, 0)

        objReqCompras.lNumIntReq = tRequisicaoCompras.lNumIntDoc

        'Lê o Item da Requisição Compras
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc, ReqCompra, Produto, DescProduto, Status, Quantidade, QuantPedida, QuantRecebida, QuantCancelada, UM, Ccl, Almoxarifado, ContaContabil, CreditaICMS, CreditaIPI, Observacao, Fornecedor, Filial, Exclusivo,TipoTributacao FROM ItensReqCompraN WHERE ReqCompra = ? AND StatusBaixa=0", 0, _
            tItemReqCompra.lNumIntDoc, tItemReqCompra.lReqCompra, tItemReqCompra.sProduto, tItemReqCompra.sDescProduto, tItemReqCompra.iStatus, tItemReqCompra.dQuantidade, tItemReqCompra.dQuantPedida, tItemReqCompra.dQuantRecebida, tItemReqCompra.dQuantCancelada, tItemReqCompra.sUM, tItemReqCompra.sCcl, tItemReqCompra.iAlmoxarifado, tItemReqCompra.sContaContabil, tItemReqCompra.iCreditaICMS, tItemReqCompra.iCreditaIPI, tItemReqCompra.lObservacao, tItemReqCompra.lFornecedor, tItemReqCompra.iFilial, tItemReqCompra.iExclusivo, tItemReqCompra.iTipoTributacao, objReqCompras.lNumIntReq)
        If lErro <> AD_SQL_SUCESSO Then gError 63383

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 63384

        'Para cada Item de Requisição de Compras
        Do While lErro <> AD_SQL_SEM_DADOS

            'Fazer "LockExclusive" em Item Requisição Compras
            lErro = Comando_LockExclusive(alComando(2))
            If lErro <> AD_SQL_SUCESSO Then gError 63385

            'Pesquisa em ItemRCItemPC e ItensPedCompra vínculo entre o Item da Requisição e um Item de algum pedido não baixado
            lErro = Comando_Executar(alComando(1), "SELECT ItensPedCompra.NumIntDoc FROM ItemRCItemPC, ItensPedCompra, ItensReqCompra WHERE ItensPedCompra.NumIntDoc = ItemRCItemPC.ItemPC AND ItemRCItemPC.ItemRC = ItensReqCompra.NumIntDoc AND ItemRCItemPC.ItemRC = ?", lNumIntDoc, tItemReqCompra.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 63386

            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 63387

            'Se encontrou --> Erro
            If lErro = AD_SQL_SUCESSO Then gError 63388

            'Pesquisa em ItemRCItemConcorrencia vínculo entre o item da requisição e um item de algum pedido não baixado
            lErro = Comando_Executar(alComando(5), "SELECT ItemRCItemConcorrencia.ItemConcorrencia FROM ItemRCItemConcorrencia,ItensConcorrencia  WHERE ItemRCItemConcorrencia.ItemReqCompra = ? and ItensConcorrencia.NumIntDoc=ItemRCItemConcorrencia.ItemConcorrencia ", lItemConcorrencia, tItemReqCompra.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 63389

            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 63390

            'Se encontrou --> Erro
            If lErro = AD_SQL_SUCESSO Then gError 63391

            'Busca em CotaçãoProdutoItemRC as Cotações para o item
            lErro = Comando_ExecutarPos(alComando(3), "SELECT QuantidadeCotar FROM CotacaoProdutoItemRC WHERE ItemReqCompra = ?", 0, dQuantidadeCotar, tItemReqCompra.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 63392

            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 63393

            'Para cada registro encontrado
            Do While lErro = AD_SQL_SUCESSO

                'Exclui o Registro da tabela CotaçãoProdutoItemRC
                lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM CotacaoProdutoItemRC", alComando(3))
                If lErro <> AD_SQL_SUCESSO Then gError 63394

                'Busca a Próxima CotaçãoProdutoItemRC
                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 63395

            Loop

            tItemReqCompra.dQuantCancelada = tItemReqCompra.dQuantidade - tItemReqCompra.dQuantRecebida

            'Atualiza o item Requisição de Compra da tabela ItemRequisiçãoCompra
            lErro = Comando_ExecutarPos(alComando(6), "UPDATE ItensReqCompraN SET StatusBaixa = 1, QuantCancelada = ?", alComando(2), tItemReqCompra.dQuantCancelada)
            If lErro <> AD_SQL_SUCESSO Then gError 63397

            'Busca a próxima Cotação Produto Item Requisição Compras
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 63398

        Loop

        'Atualiza a Requisição de Compras
        lErro = Comando_ExecutarPos(alComando(7), "UPDATE RequisicaoCompraN SET Status = 1, DataBaixa = ?", alComando(0), gdtDataAtual)
        If lErro <> AD_SQL_SUCESSO Then gError 63399

    Next
    
    TelaAcompanhaBatchCOM.dValorAtual = TelaAcompanhaBatchCOM.dValorAtual + 1
    TelaAcompanhaBatchCOM.TotReg = TelaAcompanhaBatchCOM.dValorAtual
    TelaAcompanhaBatchCOM.ProgressBar1.Value = CInt((TelaAcompanhaBatchCOM.dValorAtual / TelaAcompanhaBatchCOM.dValorTotal) * 100)
                    
    DoEvents
        
    If TelaAcompanhaBatchCOM.iCancelaBatch = CANCELA_BATCH Then
        vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_CALCULO_PARAMETROS_PTOPEDIDO")
        If vbMesRes = vbYes Then gError 74944
        TelaAcompanhaBatchCOM.iCancelaBatch = 0
    End If
            
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 63400

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Rotina_ReqComprasBaixar_Batch_Int = SUCESSO

    Exit Function

Erro_Rotina_ReqComprasBaixar_Batch_Int:

    Rotina_ReqComprasBaixar_Batch_Int = gErr

    Select Case gErr

        Case 63376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 63377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 63378, 63379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr, objReqCompras.lCodRequisicao)

        Case 63380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objReqCompras.lCodRequisicao)
            
        Case 63381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISICAOCOMPRA", gErr, objReqCompras.lCodRequisicao)

        Case 63382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REQUISICAOCOMPRABAIXADA", gErr, objReqCompras.lCodRequisicao)

        Case 63383, 63384, 63398
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr, objReqCompras.lCodRequisicao)

        Case 63385
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSREQCOMPRA", gErr, objReqCompras.lCodRequisicao)

        Case 63386, 63387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)

        Case 63388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXA_ITEMRC_VINCULADO_ITEMPC_NAO_BAIXADO", gErr, objReqCompras.lCodRequisicao, lNumIntDoc)

        Case 63389, 63390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMRCITEMCONCORRENCIA", gErr, tItemReqCompra.lNumIntDoc)

        Case 63391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXA_ITEMRC_VINCULADO_ITEMCONCORRENCIA_NAO_BAIXADO", gErr, objReqCompras.lCodRequisicao, lItemConcorrencia)

        Case 63392, 63393, 63395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTOITEMRC", gErr, tItemReqCompra.lNumIntDoc)

        Case 63394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAOPRODUTOITEMRC", gErr, tItemReqCompra.lNumIntDoc)

        Case 63396
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSREQCOMPRABAIXADOS", gErr, tItemReqCompra.lNumIntDoc)

        Case 63397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSREQCOMPRAS", gErr, objReqCompras.lCodRequisicao)

        Case 63399
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REQUISICAOCOMPRA", gErr, objReqCompras.lNumIntReq)

        Case 63400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 74944
            'Erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154331)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'Este Bach calcula os Parametros para o Ponto de Pedido
Function ParametrosPtoPed_Calcula() As Long
'Calculas os Parametros de Ponto Pedido e atualiza a Tabela de Produto Filial

Dim lErro As Long
Dim tProdutoFilial As typeProdutoFilial
Dim objProdutoFilial As New ClassProdutoFilial
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lTransacao As Long
Dim objComprasConfig As New ClassComprasConfig
Dim lTotalProdutos As Long
Dim objProduto As New ClassProduto
Dim vbMesRes As VbMsgBoxResult

On Error GoTo Erro_ParametrosPtoPed_Calcula

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 64263

    'Abertura comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 64264

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 64265

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 69052

    'Pesquisa no BD o Total de Produto para o Batch
    lErro = Comando_Executar(lComando3, "SELECT COUNT(*) FROM ProdutosFilial, Produtos WHERE Produtos.Codigo = ProdutosFilial.Produto AND FilialEmpresa= ? AND Produtos.Compras = ?", lTotalProdutos, giFilialEmpresa, PRODUTO_COMPRAVEL)
    If lErro <> AD_SQL_SUCESSO Then gError 69053
        
    'Tenta selecionar Produto
    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69054
    
    TelaAcompanhaBatchCOM.dValorTotal = lTotalProdutos
    
    tProdutoFilial.sClasseABC = String(STRING_PRODUTOFILIAL_CLASSEABC, 0)
    tProdutoFilial.sProduto = String(STRING_PRODUTO, 0)

    '???? Verificar se posso incluir a condição de comprável
    'Pesquisa no BD ProdutoFilial
    lErro = Comando_ExecutarPos(lComando1, "SELECT Produto, Almoxarifado, Fornecedor, FilialForn, VisibilidadeAlmoxarifados, EstoqueSeguranca, ESAuto, EstoqueMaximo, TemPtoPedido, PontoPedido, PPAuto, ClasseABC, LoteEconomico, IntRessup, TempoRessup, TRAuto, TempoRessupMax, ConsumoMedio, CMAuto, ConsumoMedioMax, MesesConsumoMedio FROM ProdutosFilial WHERE FilialEmpresa= ? ORDER BY Produto", 0, tProdutoFilial.sProduto, tProdutoFilial.iAlmoxarifado, tProdutoFilial.lFornecedor, tProdutoFilial.iFilialForn, tProdutoFilial.iVisibilidadeAlmoxarifados, tProdutoFilial.dEstoqueSeguranca, tProdutoFilial.iESAuto, tProdutoFilial.dEstoqueMaximo, tProdutoFilial.iTemPtoPedido, _
    tProdutoFilial.dPontoPedido, tProdutoFilial.iPPAuto, tProdutoFilial.sClasseABC, tProdutoFilial.dLoteEconomico, tProdutoFilial.iIntRessup, tProdutoFilial.iTempoRessup, tProdutoFilial.iTRAuto, tProdutoFilial.dTempoRessupMax, tProdutoFilial.dConsumoMedio, tProdutoFilial.iCMAuto, tProdutoFilial.dConsumoMedioMax, tProdutoFilial.iMesesConsumoMedio, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 64266
        
    'Tenta selecionar Produto
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64267

    Do While lErro = AD_SQL_SUCESSO
            
        objProduto.sCodigo = tProdutoFilial.sProduto
        
        'Lê o Produto para saber se é Compravel
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 69103
        
        'Se não encontrou ---> Error
        If lErro = 28030 Then gError 69104
        
        'Se o Produto é de Compravel
        If objProduto.iCompras = PRODUTO_COMPRAVEL Then
            
            Set objProdutoFilial = New ClassProdutoFilial
        
            objProdutoFilial.sProduto = tProdutoFilial.sProduto
            objProdutoFilial.iFilialEmpresa = giFilialEmpresa
            objProdutoFilial.iAlmoxarifado = tProdutoFilial.iAlmoxarifado
            objProdutoFilial.lFornecedor = tProdutoFilial.lFornecedor
            objProdutoFilial.iFilialForn = tProdutoFilial.iFilialForn
            objProdutoFilial.iVisibilidadeAlmoxarifados = tProdutoFilial.iVisibilidadeAlmoxarifados
            objProdutoFilial.dEstoqueSeguranca = tProdutoFilial.dEstoqueSeguranca
            objProdutoFilial.iESCalculado = tProdutoFilial.iESAuto
            objProdutoFilial.dEstoqueMaximo = tProdutoFilial.dEstoqueMaximo
            objProdutoFilial.iTemPtoPedido = tProdutoFilial.iTemPtoPedido
            objProdutoFilial.dPontoPedido = tProdutoFilial.dPontoPedido
            objProdutoFilial.iPPCalculado = tProdutoFilial.iPPAuto
            objProdutoFilial.sClasseABC = tProdutoFilial.sClasseABC
            objProdutoFilial.dLoteEconomico = tProdutoFilial.dLoteEconomico
            objProdutoFilial.iIntRessup = tProdutoFilial.iIntRessup
            objProdutoFilial.iTempoRessup = tProdutoFilial.iTempoRessup
            objProdutoFilial.iTRCalculado = tProdutoFilial.iTRAuto
            objProdutoFilial.dTempoRessupMax = tProdutoFilial.dTempoRessupMax
            objProdutoFilial.dConsumoMedio = tProdutoFilial.dConsumoMedio
            objProdutoFilial.iCMCalculado = tProdutoFilial.iCMAuto
            objProdutoFilial.dConsumoMedioMax = tProdutoFilial.dConsumoMedioMax
            objProdutoFilial.iMesesConsumoMedio = tProdutoFilial.iMesesConsumoMedio
            
            'Calcula o Consumo Médio para este Produto
            lErro = Produto_Calcula_ConsumoMedio(objProdutoFilial)
            If lErro <> SUCESSO Then gError 64268
            
            'Calcula o Tempo de Ressuprimento para este Produto
            lErro = Produto_Calcula_TempoRessuprimento(objProdutoFilial)
            If lErro <> SUCESSO Then gError 64269
            
            'Calcula o Estoque de Segurança para este Produto
            lErro = Produto_Calcula_EstoqueSeguranca(objProdutoFilial)
            If lErro <> SUCESSO Then gError 64270
            
            'Calcula o Ponto de Pedido para este Produto
            lErro = Produto_Calcula_PontoPedido(objProdutoFilial)
            If lErro <> SUCESSO Then gError 64271
            
            'Atualiza a Tabela ProdutoFilial para os valores calculados
            'Utilizar o lComando2
            lErro = Comando_ExecutarPos(lComando2, "UPDATE ProdutosFilial SET EstoqueSeguranca = ?, PontoPedido= ?, TempoRessup = ?, ConsumoMedio = ?", lComando1, objProdutoFilial.dEstoqueSeguranca, objProdutoFilial.dPontoPedido, objProdutoFilial.iTempoRessup, objProdutoFilial.dConsumoMedio)
            If lErro <> AD_SQL_SUCESSO Then gError 64272
            
            TelaAcompanhaBatchCOM.dValorAtual = TelaAcompanhaBatchCOM.dValorAtual + 1
            TelaAcompanhaBatchCOM.TotReg = TelaAcompanhaBatchCOM.dValorAtual
            TelaAcompanhaBatchCOM.ProgressBar1.Value = CInt((TelaAcompanhaBatchCOM.dValorAtual / TelaAcompanhaBatchCOM.dValorTotal) * 100)
                    
            DoEvents
        
            If TelaAcompanhaBatchCOM.iCancelaBatch = CANCELA_BATCH Then
                vbMesRes = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_CALCULO_PARAMETROS_PTOPEDIDO")
                If vbMesRes = vbYes Then gError 69106
                TelaAcompanhaBatchCOM.iCancelaBatch = 0
            End If
                    
        End If
                    
        'Tenta selecionar ProdutoFilial
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64273
        
    Loop
    
    objComprasConfig.sCodigo = COMPRAS_CONFIG_DATA_CALCULO_PTO_PEDIDO
    objComprasConfig.iFilialEmpresa = EMPRESA_TODA
    objComprasConfig.sConteudo = gdtDataAtual
    
    'Atualiza a data de último cálculo no BD
    lErro = ComprasConfig_Atualiza_Conteudo_Trans(objComprasConfig)
    If lErro <> SUCESSO Then gError 64274
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    'Confirma a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 64275
   
    ParametrosPtoPed_Calcula = SUCESSO
    
    Exit Function
    
Erro_ParametrosPtoPed_Calcula:

    ParametrosPtoPed_Calcula = gErr
    
    Select Case gErr
        
        Case 64263
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 64264, 64265, 69052
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 64266, 64267, 64273, 69053, 69054
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL2", gErr)
        
        Case 64268, 64269, 64270, 64271, 64274, 69103, 69106 'Tratados nas rotinas chamadas
        
        Case 64272
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PRODUTOSFILIAL", gErr, objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)
        
        Case 64275
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 69104
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProdutoFilial.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154332)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Call Transacao_Rollback

    Exit Function

End Function

Function Produto_Calcula_ConsumoMedio(objProdutoFilial As ClassProdutoFilial) As Long
'Calcula o Consumo Medio para o ProdutoFilial

Dim lErro As Long
Dim objComprasConfig As New ClassComprasConfig
Dim iMesesConfig As Integer
Dim tEstoqueMes As typeEstoqueMes
Dim iContaMes As Integer
Dim objSldMesEst As New ClassSldMesEst
Dim dQuantidadeConsumidaTotal As Double
Dim lComando1 As Long

On Error GoTo Erro_Produto_Calcula_ConsumoMedio
    
    If objProdutoFilial.iCMCalculado = PRODUTOFILIAL_CALCULA_VALORES Then
        
        'Abertura comando
        lComando1 = Comando_Abrir()
        If lComando1 = 0 Then Error 64276

'        '???? Porque você não usa o gobjCOM para não precisar ler. Inclua esse campo nele e inclua esse código no select
'        '???? Onde ele está sendo gravado? Na tela de Configuração. Por favor quando souber me avise.
'        objComprasConfig.sCodigo = COMPRAS_CONFIG_MESES_CONSUMO_MEDIO
'        objComprasConfig.iFilialEmpresa = EMPRESA_TODA
'
'        'Lê o número de meses que serão calculados o Consumo Médio
'        lErro = CF("ComprasConfig_Le_Conteudo",objComprasConfig)
'        If lErro <> SUCESSO Then Error 64277
                
'        iMesesConfig = CInt(objComprasConfig.sConteudo)
 '??????
        iMesesConfig = objProdutoFilial.iMesesConsumoMedio
''''        iMesesConfig = gobjCOM.iMesesConsumoMedio
        
        'Le os Ultimos Meses Fechados
        lErro = Comando_Executar(lComando1, "SELECT FilialEmpresa, Ano, Mes FROM EstoqueMes WHERE Fechamento = ? AND FilialEmpresa = ? ORDER BY Ano DESC, Mes DESC", tEstoqueMes.iFilialEmpresa, tEstoqueMes.iAno, tEstoqueMes.iMes, ESTOQUEMES_FECHAMENTO_FECHADO, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 64278
        
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64279
        
        'Para cada Mês
        Do While (iContaMes < iMesesConfig And lErro = AD_SQL_SUCESSO)
                
            iContaMes = iContaMes + 1
            
            objSldMesEst.iAno = tEstoqueMes.iAno
            objSldMesEst.sProduto = objProdutoFilial.sProduto
            objSldMesEst.iFilialEmpresa = giFilialEmpresa
            
            'Le a Quantidade Consumida e a Quantidade de Venda
            lErro = CF("SldMesEst_Le", objSldMesEst)
            If lErro <> SUCESSO And lErro <> 25429 Then Error 64280
            
            dQuantidadeConsumidaTotal = dQuantidadeConsumidaTotal + (objSldMesEst.dQuantCons(tEstoqueMes.iMes) + objSldMesEst.dQuantVend(tEstoqueMes.iMes))

            'Tenta selecionar ProdutoFilial
            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64281
        
        Loop
        
        'Calcula o Consumo Médio
        '#########################
        'ALTERADO POR WAGNER - CONFIRMAR ALTERAÇÃO
        'objProdutoFilial.dConsumoMedio = dQuantidadeConsumidaTotal / (iContaMes + 1)
        If iContaMes > 0 Then
            objProdutoFilial.dConsumoMedio = dQuantidadeConsumidaTotal / (iContaMes)
        Else
            objProdutoFilial.dConsumoMedio = 0
        End If
        '#########################
    
    End If
    
    Call Comando_Fechar(lComando1)
    
    Produto_Calcula_ConsumoMedio = SUCESSO
    
    Exit Function
    
Erro_Produto_Calcula_ConsumoMedio:

    Produto_Calcula_ConsumoMedio = Err
    
    Select Case Err
    
        Case 64276
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 64277, 64280 'Tratados nas rotinas chamadas
        
        Case 64278, 64279, 64281
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154333)

    End Select

    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Produto_Calcula_TempoRessuprimento(objProdutoFilial As ClassProdutoFilial) As Long
'Calcula o Tempo de Ressuprimento para o ProdutoFilial

Dim lErro As Long
Dim iMesesRessup As Integer
Dim iNumeroComprasRessup As Integer
Dim tEstoqueMes As typeEstoqueMes
Dim dtLimiteMaior As Date
Dim dtLimiteMenor As Date
Dim iContaPedido As Integer
Dim dQuantRecebidaPC As Double
Dim dtDataEmissaoPC As Date
Dim dtDataEntradaNF As Date
Dim iDiasRecebimentoParcial As Integer
Dim iDiasRecebimentoTotal As Integer
Dim dValorParcial As Double
Dim lComando1 As Long
Dim lComando2 As Long
Dim objComprasConfig As New ClassComprasConfig
Dim iContaMes As Integer
Dim dQuantidadeRecebidaTotal As Double
Dim sSQL As String

On Error GoTo Erro_Produto_Calcula_TempoRessuprimento
    
    If objProdutoFilial.iTRCalculado = PRODUTOFILIAL_CALCULA_VALORES Then
    
        'Abertura comando
        lComando1 = Comando_Abrir()
        If lComando1 = 0 Then gError 64282

        lComando2 = Comando_Abrir()
        If lComando2 = 0 Then gError 64283
       
        iMesesRessup = gobjCOM.iMesesMediaTempoRessup
        iNumeroComprasRessup = gobjCOM.iNumComprasTempoRessup
    
        'Le os Ultimos Meses Fechados
        lErro = Comando_Executar(lComando1, "SELECT FilialEmpresa, Ano, Mes FROM EstoqueMes WHERE Fechamento = ? AND FilialEmpresa = ? ORDER BY Ano DESC, Mes DESC", tEstoqueMes.iFilialEmpresa, tEstoqueMes.iAno, tEstoqueMes.iMes, ESTOQUEMES_FECHAMENTO_FECHADO, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 64286
        
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64287
        
        'Se não encontrou nenhum mês fechado, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 67459
        
        'Calcula a ultima data Limite (Ultimo dia do Mes Lido /MesLido/AnoLido)
        If tEstoqueMes.iMes = 12 Then
            dtLimiteMaior = CDate("31/" & tEstoqueMes.iMes & "/" & tEstoqueMes.iAno)
        Else
            dtLimiteMaior = CDate("01/" & (tEstoqueMes.iMes + 1) & "/" & tEstoqueMes.iAno) - 1
        End If
        
        iContaMes = 1
        
        Do While (iContaMes < iMesesRessup And lErro = AD_SQL_SUCESSO)
                
            iContaMes = iContaMes + 1
        
            'Tenta selecionar ProdutoFilial
            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64288
        
        Loop
        
        'Calcula a Primeira data Limite (01/MesLido/AnoLido)
        dtLimiteMenor = CDate("01/" & tEstoqueMes.iMes & "/" & tEstoqueMes.iAno)
        
        'Vai pegar os últimos 6 meses e não só o último mês fechado
        dtLimiteMenor = DateAdd("m", -5, dtLimiteMenor)
        
        'Pedido -> NF
        'Comando SQL a executar abaixo
        sSQL = "SELECT SUM(ItensPedCompraN.QuantRecebida), MAX(PedidoCompraN.DataEmissao), MAX(NFiscal.DataEntrada) FROM ItemNFItemPC, ItensPedCompraN, ItensNFiscal, PedidoCompraN, NFiscal WHERE ItensPedCompraN.NumIntDoc = ItemNFItemPC.ItemPedCompra AND PedidoCompraN.NumIntDoc = ItensPedCompraN.PedCompra AND ItemNFItemPC.ItemNFiscal = ItensNFiscal.NumIntDoc AND ItensNFiscal.NumIntNF = NFiscal.NumIntDoc AND ItensPedCompraN.Produto = ? AND PedidoCompraN.FilialEmpresa = ? AND PedidoCompraN.DataEmissao >= ? AND  PedidoCompraN.DataEmissao <= ? GROUP BY PedidoCompraN.NumIntDoc ORDER BY PedidoCompraN.NumIntDoc"
        
        'Faz select em Todos os Pedidos de Compra (Baixados ou não)
        lErro = Comando_Executar(lComando2, sSQL, dQuantRecebidaPC, dtDataEmissaoPC, dtDataEntradaNF, objProdutoFilial.sProduto, giFilialEmpresa, dtLimiteMenor, dtLimiteMaior)
        If lErro <> AD_SQL_SUCESSO Then gError 64289
        
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64290
                
        Do While (iNumeroComprasRessup = 0 Or iContaPedido <= iNumeroComprasRessup) And lErro = AD_SQL_SUCESSO
                
            iContaPedido = iContaPedido + 1
            iDiasRecebimentoParcial = dtDataEntradaNF - dtDataEmissaoPC
            iDiasRecebimentoTotal = iDiasRecebimentoTotal + iDiasRecebimentoParcial
            dQuantidadeRecebidaTotal = dQuantidadeRecebidaTotal + dQuantRecebidaPC
            dValorParcial = dValorParcial + (dQuantRecebidaPC * iDiasRecebimentoParcial)
            
            'Tenta selecionar ProdutoFilial
            lErro = Comando_BuscarProximo(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 64291
                
        Loop

        'Calcula o Tempo de Ressuprimento
        '################################################
        'ALTERADO POR WAGNER
        'If iContaPedido > 0 Then 'Trocado o teste porque o Select retorna pedidos que ainda não foram recebidos para um produto em específico (o produto tem quantidade recebida para outro pedido na nota, mas não para esse)
        If dQuantidadeRecebidaTotal > DELTA_VALORMONETARIO Then
            If iDiasRecebimentoTotal > 0 Then objProdutoFilial.iTempoRessup = dValorParcial / dQuantidadeRecebidaTotal
        
            'Tempo mínimo = 1
            If objProdutoFilial.iTempoRessup < 1 Then objProdutoFilial.iTempoRessup = 1
        Else
            'Tempo Não calculado
            objProdutoFilial.iTempoRessup = 0
        End If
        '################################################
    
    End If
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    Produto_Calcula_TempoRessuprimento = SUCESSO
    
    Exit Function
    
Erro_Produto_Calcula_TempoRessuprimento:

    Produto_Calcula_TempoRessuprimento = gErr
    
    Select Case gErr
        
        Case 64282, 64283
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 64284, 64285 'Tratados nas Rotinas chamadas
        
        Case 64286, 64287, 64288
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", gErr)
        
        Case 64289, 64290, 64291
            Call Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQL)
        
        Case 67459
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_MES_FECHADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154334)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Produto_Calcula_EstoqueSeguranca(objProdutoFilial As ClassProdutoFilial) As Long
'Calcula o Estoque de Segurança

Dim lErro As Long
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim objProduto As New ClassProduto
Dim objComprasConfig As New ClassComprasConfig

'INSERIDO POR WAGNER
Dim dTempoRessupTotal As Double
Dim dConsumoMedioTotal As Double

On Error GoTo Erro_Produto_Calcula_EstoqueSeguranca
        
    'Se é Para calcular o Estoque de segurança
    If objProdutoFilial.iESCalculado = PRODUTOFILIAL_CALCULA_VALORES Then
        
        'Verifica se já foi Lido o Consumo Médio MAX e o Tempo de Ressuprimento MAX
        If objProdutoFilial.dConsumoMedioMax = 0 Or objProdutoFilial.dTempoRessupMax = 0 Then
                                
            objProduto.sCodigo = objProdutoFilial.sProduto
            
            'Lê o Produto para Pegar o Tipo de Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 64292
            
            'Se não encontrou --> Erro
            If lErro = 28030 Then Error 64293
            
            If objProduto.iTipo > 0 Then
                
                objTipoDeProduto.iTipo = objProduto.iTipo
                
                'Le o Tipo de Produto para ver se o Consumo Medio Max e o Tempo de Ressuprimento Max estão Preenchidos
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then Error 64294
                
                'Se não encontrou --> Erro
                If lErro = 22531 Then Error 64295
                
                objProdutoFilial.dConsumoMedioMax = objTipoDeProduto.dConsumoMedioMax
                objProdutoFilial.dTempoRessupMax = objTipoDeProduto.dTempoRessupMax
                
            End If
                        
            If objProdutoFilial.dConsumoMedioMax = 0 Or objProdutoFilial.dTempoRessupMax = 0 Then
                        
                objProdutoFilial.dConsumoMedioMax = gobjCOM.dConsumoMedioMax
                objProdutoFilial.dTempoRessupMax = gobjCOM.dTempoRessupMax
            
            End If
            
        End If
            
        'Calcula o Estoque de segurança
        '##########################################
        'ALTERADO POR WAGNER
        'Tempo de ressuprimento = 100% + % a Mais * TempoMédio
        'Consumo a mais  = 100% + % a Mais * TempoMédio /30 => Converte para consumo diário
        dTempoRessupTotal = (objProdutoFilial.dTempoRessupMax + 1) * objProdutoFilial.iTempoRessup
        dConsumoMedioTotal = ((objProdutoFilial.dConsumoMedioMax + 1) * objProdutoFilial.dConsumoMedio) / 30
        
        If dTempoRessupTotal = 0 Then
            objProdutoFilial.dEstoqueSeguranca = 0
        Else
            objProdutoFilial.dEstoqueSeguranca = dTempoRessupTotal * dConsumoMedioTotal
        End If
        '##########################################
        'objProdutoFilial.dEstoqueSeguranca = objProdutoFilial.dConsumoMedioMax * objProdutoFilial.dTempoRessupMax - objProdutoFilial.dConsumoMedio * objProdutoFilial.iTempoRessup
            
    End If
    
    Produto_Calcula_EstoqueSeguranca = SUCESSO
    
    Exit Function
    
Erro_Produto_Calcula_EstoqueSeguranca:

    Produto_Calcula_EstoqueSeguranca = Err
    
    Select Case Err
        
        Case 64292, 64294, 64296, 64297 'Tratados nas Rotinas chamadas
        
        Case 64293
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)
        
        Case 64295
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", Err, objTipoDeProduto.iTipo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154335)

    End Select

    Exit Function

End Function

Function Produto_Calcula_PontoPedido(objProdutoFilial As ClassProdutoFilial) As Long
'Calcula o Ponto de Pedido

On Error GoTo Erro_Produto_Calcula_PontoPedido
    
    'Verifica se é Para calcular o Ponto de Pedido
    If objProdutoFilial.iPPCalculado = PRODUTOFILIAL_CALCULA_VALORES Then
        
        'Calcula o Ponto de Pedido
        '###################################
        'ALTERADO POR WAGNER
        If objProdutoFilial.iTempoRessup > 0 Then
            objProdutoFilial.dPontoPedido = ((objProdutoFilial.dConsumoMedio * objProdutoFilial.iTempoRessup) / 30) + objProdutoFilial.dEstoqueSeguranca
        Else
            objProdutoFilial.dPontoPedido = 0
        End If
        '###################################
    
    End If
    
    Produto_Calcula_PontoPedido = SUCESSO
    
    Exit Function
    
Erro_Produto_Calcula_PontoPedido:

    Produto_Calcula_PontoPedido = Err
    
    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154336)

    End Select

    Exit Function
        
End Function

Function ComprasConfig_Atualiza_Conteudo_Trans(objComprasConfig As ClassComprasConfig) As Long
'Atualiza o Conteudo na Tabela de Compras Config para o Codigo e Filial Passados

Dim lErro As Long
Dim sConteudo As String
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long

On Error GoTo Erro_ComprasConfig_Atualiza_Conteudo_Trans

    'Abertura comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 64298

    'Abertura comando
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 64299

    sConteudo = String(STRING_CONTEUDO, 0)

    'Ler registo
    lErro = Comando_ExecutarPos(lComando1, "SELECT Conteudo FROM ComprasConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, objComprasConfig.sCodigo, objComprasConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 64300

    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64301

    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then Error 64302
    
    If sConteudo <> objComprasConfig.sConteudo Then
        
        'Atualiza o conteudo do código passado
        lErro = Comando_ExecutarPos(lComando2, "UPDATE ComprasConfig SET Conteudo = ?", lComando1, objComprasConfig.sConteudo)
        If lErro <> AD_SQL_SUCESSO Then Error 64303
    
    End If
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    ComprasConfig_Atualiza_Conteudo_Trans = SUCESSO

    Exit Function

Erro_ComprasConfig_Atualiza_Conteudo_Trans:

    ComprasConfig_Atualiza_Conteudo_Trans = Err

    Select Case Err
        
        Case 64298, 64299
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 64300, 64301
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRASCONFIG", Err, objComprasConfig.sCodigo)

        Case 64302
            Call Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_COMPRAS_CONFIG_NAO_ENCONTRADO", Err, objComprasConfig.sCodigo, objComprasConfig.iFilialEmpresa)

        Case 64303
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COMPRASCONFIG", Err, objComprasConfig.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154337)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function



