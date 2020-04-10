VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Dim giInclusao As Integer
Dim giExclusao As Integer
Dim giAlteracao As Integer
Dim giReserva As Integer
Dim gdtDataUltimoMov As Date
Dim gdHora As Double
Dim glCodigoEstorno As Long

Function MovEstoque_Grava(ByVal objMovEst As ClassMovEstoque) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim iIndice As Integer
Dim objItemME As ClassItemMovEstoque

On Error GoTo Erro_MovEstoque_Grava

    iIndice = 0
    For Each objItemME In objMovEst.colItens
        iIndice = iIndice + 1
        objItemME.iItem = iIndice
    Next

    If gcolModulo.Ativo(MODULO_ESTOQUE) = MODULO_ATIVO Then
    
        'se o movimento for de exclusao de recebimento
        If objMovEst.iTipoMov = MOV_EST_EXCLUSAO_RECEBIMENTO_MATERIAL Or objMovEst.iOperacao = EXCLUSAO_MOVIMENTO Then
        
            lErro = Estoque_ExcluiMovimento(objMovEst)
            If lErro <> SUCESSO Then gError 20416
        
        Else
            'se o movimento não está cadastrado ==> insere-o
            lErro = Estoque_InsereMovimento(objMovEst)
            If lErro <> SUCESSO Then gError 20417
            
        End If
    
    End If
    
    MovEstoque_Grava = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Grava:

    MovEstoque_Grava = gErr
    
    Select Case gErr
    
        Case 20416, 20417
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151603)
    
    End Select
    
    Exit Function

End Function

Function MovEstoque_Grava_Trans(ByVal objMovEst As ClassMovEstoque) As Long
'grava as informações relativas a um conjunto de movimentações de estoque

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_MovEstoque_Grava_Trans

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 20442

    lErro = CF("MovEstoque_Grava", objMovEst)
    If lErro <> SUCESSO Then Error 20443
    
    'Confirma a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 20444
    
    MovEstoque_Grava_Trans = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Grava_Trans:

    MovEstoque_Grava_Trans = Err
    
    Select Case Err
    
        Case 20442
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 20443
    
        Case 20444
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151604)
    
    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Function MovEstoque_Grava1(alComando() As Long, ByVal objMovEst As ClassMovEstoque) As Long
'grava as informações relativas a um conjunto de movimentações de estoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_MovEstoque_Grava1

    If gcolModulo.Ativo(MODULO_ESTOQUE) = MODULO_ATIVO Then
    
        'se o movimento for de exclusao de recebimento ==> erro
        If objMovEst.iTipoMov = MOV_EST_EXCLUSAO_RECEBIMENTO_MATERIAL Then
            Error 44283
        Else
            'se não for exclusão de recebimento
            
            'se o movimento não está cadastrado ==> insere-o
            lErro = Estoque_InsereMovimento1(alComando(), objMovEst)
            If lErro <> SUCESSO Then Error 44284
            
        End If
    
    End If
    
    MovEstoque_Grava1 = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Grava1:

    MovEstoque_Grava1 = Err
    
    Select Case Err
    
        Case 44283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_MOVIMENTOESTOQUE_INVALIDO", Err, objMovEst.iTipoMov)
        
        Case 44284
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151605)
    
    End Select
    
    Exit Function

End Function

Private Function Estoque_InsereMovimento(ByVal objMovEst As ClassMovEstoque) As Long
'insere as informações relativas a um conjunto de movimentações de estoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim alComando(1 To NUM_MAX_LCOMANDO_MOVESTOQUE) As Long
Dim iIndice As Integer

On Error GoTo Erro_Estoque_InsereMovimento

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 20379

    Next

    lErro = Estoque_InsereMovimento1(alComando(), objMovEst)
    If lErro <> SUCESSO Then Error 44282

    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Estoque_InsereMovimento = SUCESSO
    
    Exit Function
    
Erro_Estoque_InsereMovimento:

    Estoque_InsereMovimento = Err

    Select Case Err
            
        Case 20379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 44282
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151606)
    
    End Select
    
    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Estoque_InsereMovimento1(alComando() As Long, ByVal objMovEst As ClassMovEstoque) As Long
'insere as informações relativas a um conjunto de movimentações de estoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim objItemMovEst As New ClassItemMovEstoque
Dim objItemMovEst1 As New ClassItemMovEstoque
Dim lErro As Long, bTestouEstoqueMes As Boolean
Dim iIndice As Integer
Dim sConteudo As String
Dim lNumIntDoc As Long
Dim iStatus As Integer
Dim objProduto As New ClassProduto
Dim objEstoqueMes As New ClassEstoqueMes
Dim objEstoque As New ClassEstoque
Dim iRastro As Integer
Dim iEstoque As Integer
Dim objMATConfig As New ClassMATConfig
Dim dtDataMov As Date
Dim iOperacao As Integer
Dim colEstoqueMesProduto As New Collection

On Error GoTo Erro_Estoque_InsereMovimento1

    lErro = CF("Customizacoes_MovEstoque", objMovEst)
    If lErro <> SUCESSO Then gError 126712

    lErro = CF("MovimentoEstoque_Le_Codigo", objMovEst.iFilialEmpresa, objMovEst.lCodigo, dtDataMov)
    If lErro <> SUCESSO And lErro <> 36015 Then gError 36017
    
    If lErro = SUCESSO Then
        iStatus = MOVESTOQUE_CODIGO_JA_CADASTRADO
    End If

    'se se trata de uma alteração de data
    If lErro = SUCESSO And dtDataMov <> objMovEst.dtData Then
    
        'se for uma alteração de data ==> trata a alteração
        lErro = Estoque_AlteraMovimento(objMovEst, alComando)
        If lErro <> SUCESSO Then gError 89671
    
    End If
    
    gdtDataUltimoMov = DATA_NULA

    objMATConfig.iFilialEmpresa = EMPRESA_TODA
    objMATConfig.sCodigo = NUM_PROX_ITEM_MOV_ESTOQUE

    lErro = CF("MATConfig_Le_Lock", objMATConfig, alComando(11))
    If lErro <> SUCESSO Then gError 83544

    lNumIntDoc = CLng(objMATConfig.sConteudo)

    objMATConfig.iFilialEmpresa = EMPRESA_TODA
    objMATConfig.sCodigo = NUM_PROX_RASTREAMENTOMOVTO

    lErro = CF("MATConfig_Le_Lock", objMATConfig, alComando(59))
    If lErro <> SUCESSO Then gError 83545

    objEstoque.lNumIntDocRastroMovto = CLng(objMATConfig.sConteudo)

    objMATConfig.iFilialEmpresa = EMPRESA_TODA
    objMATConfig.sCodigo = NUM_PROX_RASTREAMENTOLOTE

    lErro = CF("MATConfig_Le_Lock", objMATConfig, alComando(60))
    If lErro <> SUCESSO Then gError 83546

    objEstoque.lNumIntDocRastroLote = CLng(objMATConfig.sConteudo)

    objMATConfig.iFilialEmpresa = EMPRESA_TODA
    objMATConfig.sCodigo = NUM_PROX_APROPRIACAOINSUMOSPROD

    lErro = CF("MATConfig_Le_Lock", objMATConfig, alComando(68))
    If lErro <> SUCESSO Then gError 83547

    objEstoque.lNumIntDocApropInsProd = CLng(objMATConfig.sConteudo)

    bTestouEstoqueMes = False
    
    'As flags giExclusao e giInclusao indicam se foi feito uma exclusao/inclusao de movimento
    giExclusao = 0
    giInclusao = 0
    giAlteracao = 0
    giReserva = 0
    glCodigoEstorno = 0
    
    iIndice = 1
    Do While iIndice > 0 And iIndice <= objMovEst.colItens.Count
            
        Set objItemMovEst = objMovEst.colItens(iIndice)
    
        'se o tipo de movimento estiver preenchido a nivel generico ==> colocar todos os itens com este tipo
        If objMovEst.iTipoMov <> 0 And objItemMovEst.iTipoMov = 0 Then objItemMovEst.iTipoMov = objMovEst.iTipoMov
            
        objProduto.sCodigo = objItemMovEst.sProduto
        
        'le o atributo controleestoque do produto em questão. Os atributos siglaUM e classeUM se deve a necessidade converter a unidade do movimento para unidade de estoque
        lErro = CF("Produto_Le_ControleEstoque", alComando(17), objProduto)
        If lErro <> SUCESSO Then gError 36036

        objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
        objItemMovEst.iClasseUM = objProduto.iClasseUM
        objItemMovEst.iControleEstoque = objProduto.iControleEstoque
        objItemMovEst.iApropriacaoProd = objProduto.iApropriacaoCusto
        objItemMovEst.dtDataRegistro = gdtDataHoje
        
        'testa se o produto é estocado. Se nao for, não faz nada
        If objItemMovEst.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And objItemMovEst.colItensRomaneioGrade.Count = 0 Then
                
            iEstoque = 1
                
            'se não se tratar de reserva ==> testa os fechamentos de mes de estoque
            If objItemMovEst.iTipoMov <> MOV_EST_UTILIZA_RESERVA And _
               objItemMovEst.iTipoMov <> MOV_EST_INCLUI_RESERVA And _
               objItemMovEst.iTipoMov <> MOV_EST_CANCELA_RESERVA Then
               
                If bTestouEstoqueMes = False Then
                
                    objEstoqueMes.iFilialEmpresa = objMovEst.iFilialEmpresa
                    objEstoqueMes.iAno = Year(objMovEst.dtData)
                    objEstoqueMes.iMes = Month(objMovEst.dtData)
                    
                    'verifica se o mes relativo a data do movimento em questão não está fechada
                    lErro = CF("EstoqueMes_Le_Lock", alComando(87), objEstoqueMes)
                    If lErro <> SUCESSO And lErro <> 41774 Then Error 20439
                    
                    If lErro = 41774 Then gError 36516
                    
                    'se o estoque do mes em questão estiver fechado ==> erro
'                    If objEstoqueMes.iFechamento = ESTOQUEMES_FECHAMENTO_FECHADO Then Error 36517
        
                    bTestouEstoqueMes = True
                
                    'preenche uma colecao com os produtos que tiveram gastos informados e que portanto não terão seu calculo feito com os demais produtos
                    lErro = CF("EstoqueMesProduto_Le", objEstoqueMes.iAno, objEstoqueMes.iMes, colEstoqueMesProduto)
                    If lErro <> SUCESSO Then gError 92890
                
                    Set objEstoque.colEstoqueMesProduto = colEstoqueMesProduto
                
                End If
                
            Else
                giReserva = 1
        
            End If
            
            objItemMovEst.iFilialEmpresa = objMovEst.iFilialEmpresa
            objItemMovEst.lCodigo = objMovEst.lCodigo
            objItemMovEst.dtData = objMovEst.dtData
            objItemMovEst.dtHora = objMovEst.dtHora
            If objItemMovEst.lCliente = 0 Then
                objItemMovEst.lCliente = objMovEst.lCliente
                objItemMovEst.iFilialCli = objMovEst.iFilialCli
            End If
            If objItemMovEst.lFornecedor = 0 Then
                objItemMovEst.lFornecedor = objMovEst.lFornecedor
                objItemMovEst.iFilialForn = objMovEst.iFilialForn
            End If
            objItemMovEst.sDocOrigem = objMovEst.sDocOrigem
            
            '##################################
            'Inserido por Wagner 22/12/04
            objItemMovEst.sObservacao = objMovEst.sObservacao
            objItemMovEst.lRequisitante = objMovEst.lRequisitante
            '##################################
            
            'excecutar a gravação do item
            lErro = Estoque_AtualizaItemMov(alComando, objItemMovEst, lNumIntDoc, iStatus, objProduto, objMovEst, objEstoqueMes, INCLUSAO_MOVIMENTO, objEstoque)
            If lErro <> SUCESSO Then gError 20375
        
            lNumIntDoc = lNumIntDoc + 1
            
        'se for um pai de grade
        ElseIf objItemMovEst.colItensRomaneioGrade.Count > 0 Then
        
            'se é um pai de grade insere na tabela MovimentoEstoqueGrade
            objItemMovEst.iFilialEmpresa = objMovEst.iFilialEmpresa
            objItemMovEst.lCodigo = objMovEst.lCodigo
            objItemMovEst.dtData = objMovEst.dtData
            objItemMovEst.dtHora = objMovEst.dtHora
            objItemMovEst.lCliente = objMovEst.lCliente
            objItemMovEst.lFornecedor = objMovEst.lFornecedor
            objItemMovEst.iFilialCli = objMovEst.iFilialCli
            objItemMovEst.iFilialForn = objMovEst.iFilialForn
            objItemMovEst.sDocOrigem = objMovEst.sDocOrigem
            
            '##################################
            'Inserido por Wagner 22/12/04
            objItemMovEst.sObservacao = objMovEst.sObservacao
            objItemMovEst.lRequisitante = objMovEst.lRequisitante
            '##################################
        
            'se o movimento de estoque ainda nao está cadastrado
            If iStatus <> MOVESTOQUE_CODIGO_JA_CADASTRADO Then
        
                objItemMovEst.lNumIntDoc = lNumIntDoc
                
                lNumIntDoc = lNumIntDoc + 1
                
                '############################################
                'Alterado por Wagner
                'Insere um registro na tabela de MovimentoEstoqueGrade
                lErro = Comando_Executar(alComando(8), "INSERT INTO MovimentoEstoqueGrade (FilialEmpresa, Codigo, NumIntDoc, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao, DataRegistro, Hora, Observacao, Requisitante) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo, objItemMovEst.lNumIntDoc, objItemMovEst.dCusto, objItemMovEst.iApropriacao, objItemMovEst.sProduto, objItemMovEst.sSiglaUM, objItemMovEst.dQuantidade, objItemMovEst.iAlmoxarifado, objItemMovEst.iTipoMov, objItemMovEst.lNumIntDocOrigem, objItemMovEst.iTipoNumIntDocOrigem, objItemMovEst.dtData, objItemMovEst.sCcl, objItemMovEst.lNumIntDocEst, objItemMovEst.lCliente, objItemMovEst.lFornecedor, objItemMovEst.sOPCodigo, objItemMovEst.sDocOrigem, objItemMovEst.sContaContabilEst, objItemMovEst.sContaContabilAplic, _
                    objItemMovEst.lHorasMaquina, objItemMovEst.dtDataInicioProducao, objItemMovEst.dtDataRegistro, CDbl(objItemMovEst.dtHora), objItemMovEst.sObservacao, objItemMovEst.lRequisitante)
                If lErro <> AD_SQL_SUCESSO Then gError 126667
                '############################################
                
                'coloca o lNumIntDoc do Pai de Grade nos filhos
                lErro = MovEstoque_Trata_Grade1(objMovEst, objItemMovEst, iIndice)
                If lErro <> SUCESSO Then gError 126668
                
            Else
                
                'se o movimento já estiver cadastrado
                objItemMovEst1.lNumIntDoc = objItemMovEst.lNumIntDoc
            
                'Verifica se o movimento a ser alterado está cadastrado
                lErro = CF("MovimentoEstoqueGrade_Le_NumIntDoc", alComando(83), objItemMovEst1)
                If lErro <> SUCESSO And lErro <> 117641 Then gError 117637
                    
                'se nao encontrou o movimento pai de grade
                If lErro = 117641 Then gError 117642
                    
                '############################################
                'Alterado por Wagner
                'atualiza a hora, hora de maquina e data inicio de producao do movimento
                lErro = Comando_ExecutarPos(alComando(73), "UPDATE MovimentoEstoqueGrade SET HorasMaquina=?, DataInicioProducao=?, Hora=?, Observacao=?, Requisitante=?", alComando(83), objItemMovEst.lHorasMaquina, objItemMovEst.dtDataInicioProducao, CDbl(objItemMovEst.dtHora), objItemMovEst.sObservacao, objItemMovEst.lRequisitante)
                If lErro <> AD_SQL_SUCESSO Then gError 117638
                '############################################
            
            End If
            
        End If
        iIndice = iIndice + 1
    Loop
    
    'se não for uma exclusao/estorno/inclusao/alteracao de rastreamento/alteração de producao entrada ou um pedido ==> erro. Ou seja, é uma tentativa de alterar os dados de movimentação já cadastrada.
    If giAlteracao = 0 And giExclusao = 0 And giInclusao = 0 And giReserva = 0 And iEstoque = 1 Then gError 71760

    lErro = Comando_ExecutarPos(alComando(12), "UPDATE MATConfig SET Conteudo=?", alComando(11), CStr(lNumIntDoc))
    If lErro <> AD_SQL_SUCESSO Then gError 20409

    lErro = Comando_ExecutarPos(alComando(77), "UPDATE MATConfig SET Conteudo=?", alComando(59), CStr(objEstoque.lNumIntDocRastroMovto))
    If lErro <> AD_SQL_SUCESSO Then gError 71735

    lErro = Comando_ExecutarPos(alComando(78), "UPDATE MATConfig SET Conteudo=?", alComando(60), CStr(objEstoque.lNumIntDocRastroLote))
    If lErro <> AD_SQL_SUCESSO Then gError 71736

    lErro = Comando_ExecutarPos(alComando(79), "UPDATE MATConfig SET Conteudo=?", alComando(68), CStr(objEstoque.lNumIntDocApropInsProd))
    If lErro <> AD_SQL_SUCESSO Then gError 71741

    If objEstoque.dtData <> DATA_NULA Then

        'Insere ou Atualiza registro em MATConfig com a data a partir da qual deve ser feito o reprocessamento de estoque.
        lErro = MATConfig_Grava_DataReproc(objEstoque.dtData, objMovEst.iFilialEmpresa)
        If lErro <> SUCESSO Then gError 83559

    End If

    Estoque_InsereMovimento1 = SUCESSO
    
    Exit Function
    
Erro_Estoque_InsereMovimento1:

    Estoque_InsereMovimento1 = gErr

    Select Case gErr
    
        Case 20375, 20376, 20439, 36017, 36036, 83544, 83545, 83546, 83547, 83559, 89671, 92890, 126666, 126667, 126668, 117637, 126712
    
        Case 20377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_STATUS_FECHADO", gErr, Month(objMovEst.dtData), Year(objMovEst.dtData))
            
        Case 20409
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MATCONFIG", gErr, NUM_PROX_ITEM_MOV_ESTOQUE)
        
        Case 36516
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
        
        Case 36517
            lErro = Rotina_Erro(vbOKOnly, "Erro_EstoqueMes_Fechado", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
        
        Case 71735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MATCONFIG", gErr, NUM_PROX_RASTREAMENTOMOVTO)
        
        Case 71736
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MATCONFIG", gErr, NUM_PROX_RASTREAMENTOLOTE)
        
        Case 71741
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MATCONFIG", gErr, NUM_PROX_APROPRIACAOINSUMOSPROD)
        
        Case 71760
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_SEM_ALTERACAO", gErr, objMovEst.iFilialEmpresa, objMovEst.lCodigo)
        
        Case 117638
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_MOVIMENTOESTOQUEGRADE", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)
        
        Case 117642
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUEGRADE_NAO_CADASTRADO", gErr, objItemMovEst.lNumIntDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151607)
    
    End Select
    
    Exit Function

End Function

Function Estoque_Reprocessamento(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque, ByVal iOperacao As Integer, ByVal colEstoqueMesProduto As Collection) As Long
'reprocessa o movimento de estoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iStatus As Integer
Dim objProduto As New ClassProduto
Dim objEstoqueMes As New ClassEstoqueMes
Dim objEstoque As New ClassEstoque
Dim objMovEst As ClassMovEstoque

On Error GoTo Erro_Estoque_Reprocessamento

    'verificar se os meses que envolvem o reprocessamentos estão fechados, se estiverem perguntar o que fazer.
    objEstoqueMes.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objEstoqueMes.iAno = Year(objItemMovEst.dtData)
    objEstoqueMes.iMes = Month(objItemMovEst.dtData)

    'verifica se o mes relativo a data inicio dos movimentos que serão reprocessados está fechado
    lErro = CF("EstoqueMes_Le", objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 36513 Then gError 83763

    If lErro = 36513 Then gError 92671

    objProduto.sCodigo = objItemMovEst.sProduto
    
    'le o atributo controleestoque do produto em questão. Os atributos siglaUM e classeUM se deve a necessidade converter a unidade do movimento para unidade de estoque
    lErro = CF("Produto_Le_ControleEstoque", alComando(17), objProduto)
    If lErro <> SUCESSO Then gError 83560

    objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
    objItemMovEst.iClasseUM = objProduto.iClasseUM
    objItemMovEst.iControleEstoque = objProduto.iControleEstoque
    objItemMovEst.iApropriacaoProd = objProduto.iApropriacaoCusto
        
    Set objEstoque.colEstoqueMesProduto = colEstoqueMesProduto
        
    'excecutar a gravação do item
    lErro = Estoque_AtualizaItemMov(alComando, objItemMovEst, lNumIntDoc, iStatus, objProduto, objMovEst, objEstoqueMes, iOperacao, objEstoque)
    If lErro <> SUCESSO Then gError 83561
        
    Estoque_Reprocessamento = SUCESSO
    
    Exit Function
    
Erro_Estoque_Reprocessamento:

    Estoque_Reprocessamento = gErr

    Select Case gErr
    
        Case 83560, 83561, 83763
    
        Case 92671
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE", Err, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151608)
    
    End Select
    
    Exit Function

End Function

Function Estoque_ReprocAjusteStd(alComando() As Long, ByVal objMovEst As ClassMovEstoque, lNumIntDoc As Long) As Long
'insere as informações relativas a um conjunto de movimentações de estoque
'utilizado no reprocessamento para inserir os novos movimentos de estoque de ajuste do custo std
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim objItemMovEst As New ClassItemMovEstoque
Dim lErro As Long
Dim iIndice As Integer
Dim iStatus As Integer
Dim objProduto As New ClassProduto
Dim objEstoqueMes As New ClassEstoqueMes
Dim objEstoque As New ClassEstoque

On Error GoTo Erro_Estoque_ReprocAjusteStd

    iIndice = 1
    Do While iIndice > 0 And iIndice <= objMovEst.colItens.Count
            
        Set objItemMovEst = objMovEst.colItens(iIndice)
    
        'se o tipo de movimento estiver preenchido a nivel generico ==> colocar todos os itens com este tipo
        If objMovEst.iTipoMov <> 0 And objItemMovEst.iTipoMov = 0 Then objItemMovEst.iTipoMov = objMovEst.iTipoMov
            
        objProduto.sCodigo = objItemMovEst.sProduto
        
        'le o atributo controleestoque do produto em questão. Os atributos siglaUM e classeUM se deve a necessidade converter a unidade do movimento para unidade de estoque
        lErro = CF("Produto_Le_ControleEstoque", alComando(17), objProduto)
        If lErro <> SUCESSO Then gError 83652

        objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
        objItemMovEst.iClasseUM = objProduto.iClasseUM
        objItemMovEst.iControleEstoque = objProduto.iControleEstoque
        objItemMovEst.iApropriacaoProd = objProduto.iApropriacaoCusto
        
        'testa se o produto é estocado. Se nao for, não faz nada
        If objItemMovEst.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And objItemMovEst.colItensRomaneioGrade.Count = 0 Then
                
            objItemMovEst.iFilialEmpresa = objMovEst.iFilialEmpresa
            objItemMovEst.lCodigo = objMovEst.lCodigo
            objItemMovEst.dtData = objMovEst.dtData
            objItemMovEst.lCliente = objMovEst.lCliente
            objItemMovEst.lFornecedor = objMovEst.lFornecedor
            objItemMovEst.sDocOrigem = objMovEst.sDocOrigem
            
            'excecutar a gravação do item
            lErro = Estoque_AtualizaItemMov(alComando, objItemMovEst, lNumIntDoc, iStatus, objProduto, objMovEst, objEstoqueMes, REPROCESSAMENTO_REFAZ, objEstoque)
            If lErro <> SUCESSO Then gError 83653
        
            lNumIntDoc = lNumIntDoc + 1
            
        End If
        iIndice = iIndice + 1
    Loop
    
    Estoque_ReprocAjusteStd = SUCESSO
    
    Exit Function
    
Erro_Estoque_ReprocAjusteStd:

    Estoque_ReprocAjusteStd = gErr

    Select Case gErr
    
        Case 83652, 83653
    
        Case 83654
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MATCONFIG", gErr, NUM_PROX_ITEM_MOV_ESTOQUE)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151609)
    
    End Select
    
    Exit Function

End Function

Private Function Estoque_ExcluiMovimento(ByVal objMovEst As ClassMovEstoque) As Long
'exclui as informações relativas a um conjunto de movimentações de estoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim objItemMovEst As New ClassItemMovEstoque
Dim lErro As Long
Dim alComando(1 To NUM_MAX_LCOMANDO_MOVESTOQUE) As Long
Dim iIndice As Integer
Dim iStatus As Integer
Dim objProduto As New ClassProduto
Dim objEstoqueMes As New ClassEstoqueMes
Dim lNumIntDoc As Long
Dim bTestouEstoqueMes As Boolean
Dim objEstoque As New ClassEstoque
Dim dtData As Date

On Error GoTo Erro_Estoque_ExcluiMovimento

    gdtDataUltimoMov = DATA_NULA

    bTestouEstoqueMes = False

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 20438

    Next

    lErro = CF("Customizacoes_MovEstoque", objMovEst)
    If lErro <> SUCESSO Then gError 126711

    lErro = CF("MovimentoEstoque_Le_Codigo", objMovEst.iFilialEmpresa, objMovEst.lCodigo, dtData)
    If lErro <> SUCESSO And lErro <> 36015 Then gError 20821
    
    If lErro = SUCESSO Then iStatus = MOVESTOQUE_CODIGO_JA_CADASTRADO

    For Each objItemMovEst In objMovEst.colItens
    
        objProduto.sCodigo = objItemMovEst.sProduto
        
        'se o tipo de movimento estiver preenchido a nivel generico ==> colocar todos os itens com este tipo
        If objMovEst.iTipoMov <> 0 And objItemMovEst.iTipoMov = 0 Then objItemMovEst.iTipoMov = objMovEst.iTipoMov
        
        'le o atributo controleestoque do produto em questão
        lErro = CF("Produto_Le_ControleEstoque", alComando(17), objProduto)
        If lErro <> SUCESSO Then gError 36037
    
        'testa se o produto é estocado. Se nao for, não faz nada
        'If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And objItemMovEst.colItensRomaneioGrade.Count = 0 Then
        'Se o movimento existe e é para excluir não importa se ele está ou não configurado para controlar estoque
    
            objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
            objItemMovEst.iClasseUM = objProduto.iClasseUM
    
            'se não se tratar de reserva ==> testa os fechamentos de mes de estoque
            If objItemMovEst.iTipoMov <> MOV_EST_UTILIZA_RESERVA And _
               objItemMovEst.iTipoMov <> MOV_EST_INCLUI_RESERVA And _
               objItemMovEst.iTipoMov <> MOV_EST_CANCELA_RESERVA Then
    
                If bTestouEstoqueMes = False Then
                
                    objEstoqueMes.iFilialEmpresa = objMovEst.iFilialEmpresa
                    objEstoqueMes.iAno = Year(objMovEst.dtData)
                    objEstoqueMes.iMes = Month(objMovEst.dtData)
                    
                    'verifica se o mes relativo a data do movimento em questão não está fechada
                    lErro = CF("EstoqueMes_Le_Lock", alComando(87), objEstoqueMes)
                    If lErro <> SUCESSO And lErro <> 41774 Then gError 20439
                    
                    If lErro = 41774 Then gError 20798
                    
                    'se o estoque do mes em questão estiver fechado ==> erro
'                    If objEstoqueMes.iFechamento = ESTOQUEMES_FECHAMENTO_FECHADO Then gError 20440
        
                    bTestouEstoqueMes = True
                
                End If
    
            End If
            
            'excecutar a gravação do item
            lErro = Estoque_AtualizaItemMov(alComando, objItemMovEst, lNumIntDoc, iStatus, objProduto, objMovEst, objEstoqueMes, EXCLUSAO_MOVIMENTO, objEstoque)
            If lErro <> SUCESSO Then gError 20441
            
'        ElseIf objItemMovEst.colItensRomaneioGrade.Count > 0 Then
'
'            lErro = Comando_ExecutarPos(alComando(16), "DELETE FROM MovimentoEstoque", lComando2)
'            If lErro <> AD_SQL_SUCESSO Then gError 20434
            
            
        'End If
    
    Next

    If objEstoque.dtData <> DATA_NULA Then

        'Insere ou Atualiza registro em MATConfig com a data a partir da qual deve ser feito o reprocessamento de estoque.
        lErro = MATConfig_Grava_DataReproc(objEstoque.dtData, objMovEst.iFilialEmpresa)
        If lErro <> SUCESSO Then gError 20822

    End If

    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Estoque_ExcluiMovimento = SUCESSO
    
    Exit Function
    
Erro_Estoque_ExcluiMovimento:

    Estoque_ExcluiMovimento = gErr

    Select Case gErr
    
        Case 20438
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 20439, 20441, 20821, 20822, 36037, 126711
    
        Case 20440
            lErro = Rotina_Erro(vbOKOnly, "Erro_EstoqueMes_Fechado", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
            
        Case 20798
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151610)
    
    End Select
    
    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Estoque_AlteraMovimento(ByVal objMovEst As ClassMovEstoque, alComando() As Long) As Long
'altera as informações relativas a um conjunto de movimentações de estoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim objItemMovEst As New ClassItemMovEstoque
Dim lErro As Long
Dim iIndice As Integer
Dim iStatus As Integer
Dim objProduto As New ClassProduto
Dim objEstoqueMes As New ClassEstoqueMes
Dim lNumIntDoc As Long
Dim bTestouEstoqueMes As Boolean
Dim objEstoque As New ClassEstoque
Dim objMovEstoque As New ClassMovEstoque
Dim iFase As Integer
Dim objMovEstoque1 As New ClassMovEstoque

On Error GoTo Erro_Estoque_AlteraMovimento

    gdtDataUltimoMov = DATA_NULA

    bTestouEstoqueMes = False

    'Lê o movto de estoque da NOta fiscal
    objMovEstoque.iFilialEmpresa = objMovEst.iFilialEmpresa
    
    objMovEstoque.lCodigo = objMovEst.lCodigo
    
    'Lê os movimento dos itens de nota fiscal
    lErro = CF("MovEstoqueItens_Le_Desc1", objMovEstoque)
    If lErro <> SUCESSO And lErro <> 89050 Then gError 89672
    
    If lErro = 89050 Then gError 89673
    
    'na primeira fase da alteração, os dados são excluidos dos saldos da data original
    'na segunda fase os dados são inseridos nos saldos na data atualizada e o movimento alterado
    For iFase = 1 To 2
            
        If iFase = 1 Then
            objMovEstoque.iOperacao = ALTERACAO_MOVIMENTO_FASE_EXCLUSAO
        Else
            
            'Lê o movto de estoque da NOta fiscal
            objMovEstoque1.iFilialEmpresa = objMovEst.iFilialEmpresa
            
            objMovEstoque1.lCodigo = objMovEst.lCodigo
            
            'Lê os movimento dos itens de nota fiscal
            lErro = CF("MovEstoqueItens_Le", objMovEstoque1)
            If lErro <> SUCESSO And lErro <> 89050 Then gError 89985
            
            If lErro = 89050 Then gError 89986
        
            Set objMovEstoque = objMovEstoque1
            objMovEstoque.iOperacao = ALTERACAO_MOVIMENTO_FASE_INCLUSAO
            For Each objItemMovEst In objMovEstoque.colItens
                objItemMovEst.dtData = objMovEst.dtData
                objItemMovEst.dtHora = objMovEst.dtHora
                objItemMovEst.iFilialEmpresa = objMovEst.iFilialEmpresa
                '######################################
                'Inserido por Wagner
                objItemMovEst.sObservacao = objMovEst.sObservacao
                objItemMovEst.lRequisitante = objMovEst.lRequisitante
                '######################################
            Next
            '######################################
            'Inserido por Wagner
            objMovEstoque.sObservacao = objMovEst.sObservacao
            objMovEstoque.lRequisitante = objMovEst.lRequisitante
            '######################################
            objMovEstoque.dtData = objMovEst.dtData
            objMovEstoque.dtHora = objMovEst.dtHora
        End If

        For Each objItemMovEst In objMovEstoque.colItens
    
            objProduto.sCodigo = objItemMovEst.sProduto
            
            'le o atributo controleestoque do produto em questão
            lErro = CF("Produto_Le_ControleEstoque", alComando(17), objProduto)
            If lErro <> SUCESSO Then gError 89674
        
            'testa se o produto é estocado. Se nao for, não faz nada
            If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And objItemMovEst.colItensRomaneioGrade.Count = 0 Then
        
                objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
                objItemMovEst.iClasseUM = objProduto.iClasseUM
        
                'se não se tratar de reserva ==> testa os fechamentos de mes de estoque
                If objItemMovEst.iTipoMov <> MOV_EST_UTILIZA_RESERVA And _
                   objItemMovEst.iTipoMov <> MOV_EST_INCLUI_RESERVA And _
                   objItemMovEst.iTipoMov <> MOV_EST_CANCELA_RESERVA Then
        
                    If bTestouEstoqueMes = False Then
                    
                        objEstoqueMes.iFilialEmpresa = objMovEst.iFilialEmpresa
                        objEstoqueMes.iAno = Year(objMovEst.dtData)
                        objEstoqueMes.iMes = Month(objMovEst.dtData)
                        
                        'verifica se o mes relativo a data do movimento em questão não está fechada
                        lErro = CF("EstoqueMes_Le_Lock", alComando(87), objEstoqueMes)
                        If lErro <> SUCESSO And lErro <> 41774 Then gError 89675
                        
                        If lErro = 41774 Then gError 89676
                        
                        bTestouEstoqueMes = True
                    
                    End If
        
                End If
                
                'excecutar a gravação do item
                lErro = Estoque_AtualizaItemMov(alComando, objItemMovEst, lNumIntDoc, MOVESTOQUE_CODIGO_JA_CADASTRADO, objProduto, objMovEst, objEstoqueMes, objMovEstoque.iOperacao, objEstoque)
                If lErro <> SUCESSO Then gError 89677
                
            End If
        
        Next

    Next

    If objEstoque.dtData <> DATA_NULA Then

        'Insere ou Atualiza registro em MATConfig com a data a partir da qual deve ser feito o reprocessamento de estoque.
        lErro = MATConfig_Grava_DataReproc(objEstoque.dtData, objMovEst.iFilialEmpresa)
        If lErro <> SUCESSO Then gError 89678

    End If

    
    Estoque_AlteraMovimento = SUCESSO
    
    Exit Function
    
Erro_Estoque_AlteraMovimento:

    Estoque_AlteraMovimento = gErr

    Select Case gErr
        
        Case 89672, 89674, 89675, 89677, 89678, 89985
        
        Case 89673
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO2", gErr, objMovEstoque.iFilialEmpresa, objMovEstoque.lCodigo)
        
        Case 89676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
            
        Case 20439, 20441, 20821, 20822, 36037
    
        Case 20440
            lErro = Rotina_Erro(vbOKOnly, "Erro_EstoqueMes_Fechado", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
            
        Case 89986
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO2", gErr, objMovEstoque.iFilialEmpresa, objMovEstoque.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151611)
    
    End Select
    
    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Estoque_AtualizaItemMov(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque, lNumIntDoc As Long, ByVal iStatus As Integer, ByVal objProduto As ClassProduto, ByVal objMovEst As ClassMovEstoque, ByVal objEstoqueMes As ClassEstoqueMes, ByVal iOperacao As Integer, ByVal objEstoque As ClassEstoque) As Long
'atualiza os saldos de estoque e grava o movimento de estoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim objTipoMovEstoque As New ClassTipoMovEst
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objSldDiaEst As New ClassSldDiaEst
Dim dCustoMedio As Double
Dim dCustoMedioConsig As Double
Dim dCustoMedioDemo As Double
Dim dCustoMedioConserto As Double
Dim dCustoMedioOutros As Double
Dim dCustoMedioBenef As Double
Dim dCustoMedioConsig3 As Double
Dim dCustoMedioDemo3 As Double
Dim dCustoMedioConserto3 As Double
Dim dCustoMedioOutros3 As Double
Dim dCustoMedioBenef3 As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim objItemMovEstConsig As New ClassItemMovEstoque
Dim iIndice As Integer
Dim colRastreamentoLoteSaldo As New Collection
Dim iEscaninho As Integer
Dim iDevolucao As Integer
Dim dQuantConsig3 As Double
Dim dQuantBenef3 As Double

On Error GoTo Erro_Estoque_AtualizaItemMov

    objTipoMovEstoque.iCodigo = objItemMovEst.iTipoMov
    
    'ler os dados referentes ao tipo de movimento
    lErro = CF("TiposMovEst_Le1", alComando(10), objTipoMovEstoque)
    If lErro <> SUCESSO Then gError 20368
    
    'se for um estorno ou uma exclusao ==> o custo será sempre informado já que será lido o movimento que ele está estornando e será utilizado este custo como o custo do movimento
    If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV Or iOperacao = EXCLUSAO_MOVIMENTO Or objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
        objTipoMovEstoque.iCustoInformado = TIPOMOV_EST_CUSTOINFORMADO
    End If
    
    'se é uma operação de exclusão e o movimento estiver estornado ==> exclui o movimento e seu estorno
    'não tem que fazer o resto do código, pois o movimento já tinha sido estornado e portanto já retirado a quantidade e valor dos saldos
    If iOperacao = EXCLUSAO_MOVIMENTO And objItemMovEst.lNumIntDocEst <> 0 Then
    
        lErro = Trata_Exclusao_Movimento_Estornado(alComando, objItemMovEst, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 89881
    
    Else
    
        If objItemMovEst.iAlmoxarifado = 0 Then
        
            objProdutoFilial.sProduto = objItemMovEst.sProduto
            objProdutoFilial.iFilialEmpresa = objItemMovEst.iFilialEmpresa
        
            lErro = CF("ProdutoFilial_Le1", alComando(18), objProdutoFilial)
            If lErro <> SUCESSO Then gError 36225
        
            objItemMovEst.iAlmoxarifado = objProdutoFilial.iAlmoxarifado
        
        End If
        
        lErro = Estoque_AtualizaItemMov1(alComando, objItemMovEst, objTipoMovEstoque, lNumIntDoc, iStatus, iOperacao, objProduto)
        If lErro <> SUCESSO Then gError 20435
        
        'verifica se a data se refere a um movimento retroativo e faz a validação se necessário
        lErro = Trata_Data_MovEstoque(iOperacao, objMovEst, objProduto, objEstoqueMes, objItemMovEst, objTipoMovEstoque, objEstoque, iStatus)
        If lErro <> SUCESSO Then gError 89856
        
        'se o movimento já estiver cadastrado e não for uma exclusao ou estorno ou reprocessamento ==> é uma alteracao
        If iStatus = MOVESTOQUE_CODIGO_JA_CADASTRADO And objTipoMovEstoque.iAtualizaMovEstoque <> TIPOMOV_EST_ESTORNOMOV And objTipoMovEstoque.iAtualizaMovEstoque <> TIPOMOV_EST_EXCLUIMOV And iOperacao <> EXCLUSAO_MOVIMENTO And _
           iOperacao <> REPROCESSAMENTO_DESFAZ And iOperacao <> REPROCESSAMENTO_REFAZ And iOperacao <> ALTERACAO_MOVIMENTO_FASE_EXCLUSAO And iOperacao <> ALTERACAO_MOVIMENTO_FASE_INCLUSAO And iOperacao <> APURACAO_CUSTO_PRODUCAO Then
                
            'processa a alteração dos movimentos de estoque
            lErro = Estoque_Processa_Alteracao(alComando, objItemMovEst, objTipoMovEstoque, objProduto, objEstoque)
            If lErro <> SUCESSO Then gError 71761
        
        Else
        
            'esta rotina está sendo pulada para apuração do custo de produção porque a atualização da quantidade já foi feita no momento em que o movimento de estoque foi cadastrado
            If iOperacao <> APURACAO_CUSTO_PRODUCAO Then
        
                lErro = CF("Estoque_AtualizaItemMov2", objItemMovEst, objTipoMovEstoque, objEstoqueProduto)
                If lErro <> SUCESSO Then gError 20436
                
            Else
            
                objEstoqueProduto.sProduto = objItemMovEst.sProduto
                objEstoqueProduto.iAlmoxarifado = objItemMovEst.iAlmoxarifado
            
            End If
                        
            'somente para os movimentos de entrada ou saida do estoque é necessário calcular o custo.
            'não calcula em caso de exclusão de movimento, estorno e desfaz reprocessamento. Neste caso será processado pelo custo do movimento sendo excluido.
            If (objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_ENTRADA Or objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_SAIDA Or objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD Or objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_NAO_ENTRADASAIDA) And iOperacao <> REPROCESSAMENTO_DESFAZ _
               And iOperacao <> ALTERACAO_MOVIMENTO_FASE_EXCLUSAO And iOperacao <> ALTERACAO_MOVIMENTO_FASE_INCLUSAO And objTipoMovEstoque.iAtualizaSoLote <> INVENTARIO_ATUALIZA_SO_LOTE Then
            
                'calcula o custo do movimento, se necessário
                lErro = Estoque_Custo(alComando, objItemMovEst, objTipoMovEstoque, objProduto, dCustoMedio, objEstoqueMes, iOperacao, objEstoque)
                If lErro <> SUCESSO Then gError 36527
            
                'recalcula o custo médio dos produtos nossos em poder de terceiros, se necessário
                lErro = Recalcula_CustoMedio_NossoEm3(alComando, objItemMovEst, objTipoMovEstoque, dCustoMedioConsig, dCustoMedioDemo, dCustoMedioConserto, dCustoMedioOutros, dCustoMedioBenef)
                If lErro <> SUCESSO Then gError 20794
            
                'recalcula o custo médio dos produtos de terceiros em nosso poder, se necessário
                lErro = Recalcula_CustoMedio_Terceiros(alComando, objItemMovEst, objTipoMovEstoque, dCustoMedioConsig3, dCustoMedioDemo3, dCustoMedioConserto3, dCustoMedioOutros3, dCustoMedioBenef3)
                If lErro <> SUCESSO Then gError 20795
            
            End If
            
            lErro = CF("Estoque_AtualizaItemMov3", objItemMovEst, objTipoMovEstoque, objEstoqueProduto, objSldDiaEst, iOperacao)
            If lErro <> SUCESSO Then gError 20437
            
            lErro = CF("Estoque_AtualizaItemMov4", objItemMovEst, objTipoMovEstoque, objEstoqueProduto, objSldDiaEst)
            If lErro <> SUCESSO Then gError 60913
            
            If iOperacao <> ALTERACAO_MOVIMENTO_FASE_EXCLUSAO And iOperacao <> ALTERACAO_MOVIMENTO_FASE_INCLUSAO And objTipoMovEstoque.iAtualizaSoLote <> INVENTARIO_ATUALIZA_SO_LOTE Then
            
                'atualiza a tabela de EstoqueProduto
                lErro = CF("EstoqueProduto_Atualiza", alComando(2), alComando(3), objEstoqueProduto, objTipoMovEstoque, objItemMovEst, alComando(29), iOperacao, dQuantConsig3, dQuantBenef3)
                If lErro <> SUCESSO Then gError 20370
                
            End If
            
            'não faz tratamento de material consignado no reprocessamento, exclusão ou estorno
            If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO1 And (dQuantConsig3 > 0 Or dQuantBenef3 > 0) And objTipoMovEstoque.iAtualizaMovEstoque <> TIPOMOV_EST_EXCLUIMOV And objTipoMovEstoque.iAtualizaMovEstoque <> TIPOMOV_EST_ESTORNOMOV _
               And iOperacao <> REPROCESSAMENTO_DESFAZ And iOperacao <> REPROCESSAMENTO_REFAZ And iOperacao <> EXCLUSAO_MOVIMENTO And iOperacao <> ALTERACAO_MOVIMENTO_FASE_EXCLUSAO And iOperacao <> ALTERACAO_MOVIMENTO_FASE_INCLUSAO And iOperacao <> APURACAO_CUSTO_PRODUCAO Then
            
                'trata da venda de material consignado. Devem ser geradas transferencias de consignado para disponivel, com atualizacao de custo, e venda de consignado (do material que foi transferido) e venda normal do material nosso disponivel.
                lErro = Trata_MaterialConsignado(objMovEst, objItemMovEst, dQuantConsig3, objSldDiaEst, alComando(29), alComando(74), objTipoMovEstoque, objProduto, alComando, dQuantBenef3)
                If lErro <> SUCESSO Then gError 60953
                
            End If
            
            'se for uma saida e não for uma transferencia, não for a baixa de um recebimento, exclusão , estorno ou alteração de data ==> tenta gerar rastreamento FIFO
            If objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_SAIDA And objTipoMovEstoque.iTransferencia <> TIPOMOV_EST_TRANSFERENCIA And objTipoMovEstoque.iCodigo <> MOV_EST_BAIXA_RECEBIMENTO_MATERIAL _
               And objTipoMovEstoque.iAtualizaMovEstoque <> TIPOMOV_EST_EXCLUIMOV And objTipoMovEstoque.iAtualizaMovEstoque <> TIPOMOV_EST_ESTORNOMOV And iOperacao <> REPROCESSAMENTO_DESFAZ And iOperacao <> REPROCESSAMENTO_REFAZ And iOperacao <> EXCLUSAO_MOVIMENTO _
               And iOperacao <> ALTERACAO_MOVIMENTO_FASE_EXCLUSAO And iOperacao <> ALTERACAO_MOVIMENTO_FASE_INCLUSAO And objTipoMovEstoque.iInventario = TIPOMOV_EST_NAOINVENTARIO And iOperacao <> APURACAO_CUSTO_PRODUCAO Then
            
                'Testa se o movimento é uma devolução de material (saida). Se for iDevolucao retorna 1.
                lErro = Testa_TipoMovEstoque_Devolucao(objItemMovEst, iDevolucao)
                If lErro <> SUCESSO Then gError 83455
    
                'se não for uma devolucao de material, pode gerar FIFO. A devolução funcionará como as entradas, ou seja, permitirá que o usuário informe os lotes.
'                If iDevolucao = 0 Then
    
                    Call Escaninho_Retorna(objEstoqueProduto, iEscaninho)
                
                    'Gera os rastreamentos automaticamente seguindo a ordem FIFO se necessário.
                    lErro = CF("Rastreamento_FIFO_Gera", alComando(74), objItemMovEst, iEscaninho, objProduto)
                    If lErro <> SUCESSO Then gError 71775
                
'                End If
                
            End If
                
            If iOperacao <> ALTERACAO_MOVIMENTO_FASE_EXCLUSAO And iOperacao <> ALTERACAO_MOVIMENTO_FASE_INCLUSAO And iOperacao <> APURACAO_CUSTO_PRODUCAO Then
                
                lErro = Atualiza_RastreamentoLoteSaldo(objItemMovEst, objTipoMovEstoque, colRastreamentoLoteSaldo, objItemMovEst.colRastreamentoMovto)
                If lErro <> SUCESSO Then gError 71695
                
                'atualiza a tabela de RastreamentoLoteSaldo
                lErro = RastreamentoLoteSaldo_Atualiza(alComando(), colRastreamentoLoteSaldo, objItemMovEst, objEstoque, objProduto, objTipoMovEstoque)
                If lErro <> SUCESSO Then gError 71732
            
            End If
            
            'somente para os movimentos de entrada ou saida do estoque são gravados movimentos. Ex: Os movimentos de reserva não gravam.
            If objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_ENTRADA Or objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_SAIDA Or (objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD And objItemMovEst.dCusto <> 0) Or objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_NAO_ENTRADASAIDA Then
                        
                If objTipoMovEstoque.iAtualizaSoLote <> INVENTARIO_ATUALIZA_SO_LOTE Then
                    
                    'atualiza a tabela de saldos diários de movimentação
                    lErro = CF("SldDiaEst_Atualiza", alComando(4), alComando(5), objSldDiaEst)
                    If lErro <> SUCESSO Then gError 20371
                    
                    'atualiza a tabela de saldos mensais de movimentação
                    lErro = CF("SldMesEst_Atualiza", alComando(6), alComando(7), objSldDiaEst, dCustoMedio)
                    If lErro <> SUCESSO Then gError 20372
                    
                    'atualiza a tabela de saldos mensais de movimentação 1
                    lErro = CF("SldMesEst1_Atualiza", alComando(30), alComando(31), objSldDiaEst, objTipoMovEstoque, dCustoMedioConsig3, dCustoMedioDemo3, dCustoMedioConserto3, dCustoMedioOutros3, dCustoMedioBenef3)
                    If lErro <> SUCESSO Then gError 60958
                    
                    'atualiza a tabela de saldos mensais de movimentação 2
                    lErro = CF("SldMesEst2_Atualiza", alComando(32), alComando(33), objSldDiaEst, objTipoMovEstoque, dCustoMedioConsig, dCustoMedioDemo, dCustoMedioConserto, dCustoMedioOutros, dCustoMedioBenef)
                    If lErro <> SUCESSO Then gError 60959
                    
                    'atualiza a tabela de saldos diários de movimentação por almoxarifado
                    lErro = CF("SldDiaEstAlm_Atualiza", alComando(23), alComando(24), objSldDiaEst, objItemMovEst.iAlmoxarifado)
                    If lErro <> SUCESSO Then gError 44108
                    
                    'atualiza a tabela de saldos mensais de movimentação por almoxarifado
                    lErro = CF("SldMesEstAlm_Atualiza", alComando(25), alComando(26), objSldDiaEst, objItemMovEst.iAlmoxarifado)
                    If lErro <> SUCESSO Then gError 44109
                    
                    'atualiza a tabela de saldos mensais de movimentação por almoxarifado 1
                    lErro = CF("SldMesEstAlm1_Atualiza", alComando(34), alComando(35), objSldDiaEst, objItemMovEst.iAlmoxarifado, objTipoMovEstoque)
                    If lErro <> SUCESSO Then gError 60964
                    
                    'atualiza a tabela de saldos mensais de movimentação por almoxarifado 2
                    lErro = CF("SldMesEstAlm2_Atualiza", alComando(36), alComando(37), objSldDiaEst, objItemMovEst.iAlmoxarifado, objTipoMovEstoque)
                    If lErro <> SUCESSO Then gError 60965
                    
                    'atualiza as informações do fornecedor e filial
                    If iOperacao = REPROCESSAMENTO_REFAZ Then
                        lErro = CF("MovEst_Atualiza_CliForn_Filial", objItemMovEst, True, alComando(97), alComando(98), alComando(99), alComando(100), alComando(101), alComando(102))
                        If lErro <> SUCESSO Then gError 60959
                    Else
                        lErro = CF("MovEst_Atualiza_CliForn_Filial", objItemMovEst, False, alComando(97), alComando(98), alComando(99), alComando(100), alComando(101), alComando(102))
                        If lErro <> SUCESSO Then gError 60959
                    End If

                   'atualiza a tabela de saldos diários de movimentação
                    lErro = CF("SldDiaEstTerc_Atualiza", alComando(91), alComando(92), objSldDiaEst, objItemMovEst.lCliente, objItemMovEst.iFilialCli, objItemMovEst.lFornecedor, objItemMovEst.iFilialForn)
                    If lErro <> SUCESSO Then gError 20371
                    
                    'atualiza a tabela de saldos mensais de movimentação 1
                    lErro = CF("SldMesEst1Terc_Atualiza", alComando(93), alComando(94), alComando(103), objSldDiaEst, objTipoMovEstoque, objItemMovEst.lCliente, objItemMovEst.iFilialCli, objItemMovEst.lFornecedor, objItemMovEst.iFilialForn, dCustoMedioConsig3, dCustoMedioDemo3, dCustoMedioConserto3, dCustoMedioOutros3, dCustoMedioBenef3)
                    If lErro <> SUCESSO Then gError 60958
                    
                    'atualiza a tabela de saldos mensais de movimentação 2
                    lErro = CF("SldMesEst2Terc_Atualiza", alComando(95), alComando(96), alComando(104), objSldDiaEst, objTipoMovEstoque, objItemMovEst.lCliente, objItemMovEst.iFilialCli, objItemMovEst.lFornecedor, objItemMovEst.iFilialForn, dCustoMedioConsig, dCustoMedioDemo, dCustoMedioConserto, dCustoMedioOutros, dCustoMedioBenef)
                    If lErro <> SUCESSO Then gError 60959
                
                End If
                
                If iOperacao <> REPROCESSAMENTO_DESFAZ And iOperacao <> REPROCESSAMENTO_REFAZ And iOperacao <> APURACAO_CUSTO_PRODUCAO Or (iOperacao = REPROCESSAMENTO_REFAZ And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD) Then
                
                    'ponto bom para validar estoque para ver se ficou negativo em alguma data
                    lErro = CF("EstoqueNegativo_Verifica", iOperacao, objItemMovEst, objTipoMovEstoque)
                    If lErro <> SUCESSO Then gError 89880
                    
                    If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Or iOperacao = EXCLUSAO_MOVIMENTO Then
                        
                        lErro = Trata_Exclusao(alComando(), objItemMovEst, objTipoMovEstoque)
                        If lErro <> SUCESSO Then gError 89880
                        
                    ElseIf iOperacao = ALTERACAO_MOVIMENTO_FASE_EXCLUSAO Or iOperacao = ALTERACAO_MOVIMENTO_FASE_INCLUSAO Then
                    
                        'se estiver na fase de inclusao da alteração de data ==> altera a data do movimento
                        'se estiver na fase de exclusao da alteração de data ==> não faz nada em relação ao movimento
                        If iOperacao = ALTERACAO_MOVIMENTO_FASE_INCLUSAO Then
                        
                            'altera a data/hora do movimento de estoque passado como parametro
                            lErro = CF("MovimentoEstoque_Alterar", objItemMovEst)
                            If lErro <> SUCESSO Then gError 89684
                        
'                            'processa a alteração dos movimentos de estoque e do rastreamento
'                            lErro = Estoque_Processa_Alteracao(alComando, objItemMovEst, objTipoMovEstoque, objProduto, objEstoque)
'                            If lErro <> SUCESSO Then gError 89857
                        
                        End If
                        
                        giAlteracao = 1
                    
                    Else
                        
                        If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV Then objItemMovEst.dCusto = -objItemMovEst.dCusto
                        
                        'inserir na tabela MovimentoEstoque a movimentação
                        lErro = CF("MovimentoEstoque_Inserir", alComando(8), objItemMovEst)
                        If lErro <> SUCESSO Then gError 20373
                        
                        'inserir na tabela RastreamentoMovto a movimentação
                        lErro = RastreamentoMovto_Inserir(alComando(66), objItemMovEst.lNumIntDoc, objEstoque, objItemMovEst.colRastreamentoMovto, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE)
                        If lErro <> SUCESSO Then gError 71733
                        
                        'inserir na tabela ApropriacaoInsumosprod e RastreamentoMovto a movimentação dos insumos
                        lErro = ApropriacaoInsumosProd_Inserir(alComando(66), alComando(67), alComando(61), objItemMovEst, objEstoque, objTipoMovEstoque)
                        If lErro <> SUCESSO Then gError 71737
                        
                        'indica que ocorreu pelo menos uma inserção de movimento de estoque
                        giInclusao = 1
                        
                    End If
                
                'se for a fase de refazer o reprocessamento ou de apuração do custo de produção ==> atualiza o custo e apropriacao do movimento de estoque
                ElseIf iOperacao = REPROCESSAMENTO_REFAZ Or iOperacao = APURACAO_CUSTO_PRODUCAO Then
                
                    If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV Then objItemMovEst.dCusto = -objItemMovEst.dCusto
                    
                    'processa a alteração dos dados do movimento de estoque para o reprocessamento
                    lErro = MovEstoque_Altera_Reproc(alComando, objItemMovEst)
                    If lErro <> SUCESSO Then gError 83760
                
                End If
                
            End If
        
        End If
    
    End If
    
    If iOperacao <> REPROCESSAMENTO_DESFAZ And iOperacao <> REPROCESSAMENTO_REFAZ And iOperacao <> APURACAO_CUSTO_PRODUCAO Then
        lErro = CF("Estoque_TestaIntegridade", objItemMovEst.iFilialEmpresa)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    Estoque_AtualizaItemMov = SUCESSO
    
    Exit Function
    
Erro_Estoque_AtualizaItemMov:

    Estoque_AtualizaItemMov = gErr
    
    Select Case gErr
    
        Case 20368, 20370, 20371, 20372, 20373, 20435, 20436, 20437, 20794, 20795, 36225, 36527, 44108, 44109, 60913, 60953, 60958, 60959, 60964, 60965, 71695, 71732, 71733, 71737, 71761, 71775, 83455, 83760, 89684, 89856, 89881, 89857, 89880
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151612)
            
    End Select
        
    Exit Function
        
End Function
    
Private Function Trata_Data_MovEstoque(ByVal iOperacao As Integer, ByVal objMovEst As ClassMovEstoque, ByVal objProduto As ClassProduto, ByVal objEstoqueMes As ClassEstoqueMes, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst, ByVal objEstoque As ClassEstoque, ByVal iStatus As Integer) As Long
'verifica se a data se refere a um movimento retroativo e faz a validação se necessário

' *** Função alterada em 03/08/2001 por Luiz Gustavo de Freitas Nogueira ***
' *** Verifica também se a hora do movimento é retroativa ***

Dim dtData As Date
Dim dtData1 As Date
Dim objItemMovEst1 As New ClassItemMovEstoque
Dim lErro As Long

On Error GoTo Erro_Trata_Data_MovEstoque

    'se não for reprocessamento ==> verificar se o lançamento é retroativo
    If iOperacao <> REPROCESSAMENTO_DESFAZ And iOperacao <> REPROCESSAMENTO_REFAZ And iOperacao <> APURACAO_CUSTO_PRODUCAO Then
    
        If gdtDataUltimoMov = DATA_NULA Then
    
            'le a data mais recente de movimento cadastrada
            lErro = CF("MovimentoEstoque_Le_Ultima_Data", objMovEst.iFilialEmpresa, objItemMovEst1)
            If lErro <> SUCESSO And lErro <> 36029 Then gError 36992
        
            'guarda a data para possivel teste, dependendo do movimento
            If lErro = SUCESSO Then
                gdtDataUltimoMov = objItemMovEst1.dtData
                gdHora = objItemMovEst1.dtHora
            Else
                gdtDataUltimoMov = objMovEst.dtData
                gdHora = objMovEst.dtHora
            End If
    
        End If
    
        'se não se tratar de um produto apropriado pelo custo standard e se não se tratar de um produto apropriado pelo custo de produção cujo mes ainda não foi apurado
        'se a data do movimento menor do que a data do ultimo movimento (ou seja lancamento retroativo) ou for uma exclusão e a data é igual ==> guarda a data que indica a necessidade de reprocessar a partir dela
        'se a data do movimento for igual a data do último movimento, mas a hora foir anterior à hora do último movimento = > guarda a data também pois será necessário reprocessar
        'mesmo que a exclusão ocorra na mesma data do ultimo movimento ela necessita de reprocessamento já que ela afeta os lançamentos da mesma data posteriores a ela.
        If objProduto.iApropriacaoCusto <> APROPR_CUSTO_STANDARD And (objProduto.iApropriacaoCusto <> APROPR_CUSTO_REAL Or objEstoqueMes.iCustoProdApurado <> CUSTO_APURADO) Then
        
            'Teste alterado em 03/08/2001 => Ver comentário no início da função
            If objItemMovEst.dtData < gdtDataUltimoMov Or ((objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV Or iOperacao = EXCLUSAO_MOVIMENTO Or objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV) And objItemMovEst.dtData = gdtDataUltimoMov) Or (objItemMovEst.dtData = gdtDataUltimoMov And objItemMovEst.dtHora < gdHora) Then
                dtData1 = CDate("01/" & CStr(Month(objItemMovEst.dtData)) & "/" & CStr(Year(objItemMovEst.dtData)))
                objEstoque.dtData = dtData1
            End If
        
        End If
        
        'se for um movimento de ajuste standard
        If objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD Then
        
            dtData = CDate("01/" & CStr(Month(objItemMovEst.dtData)) & "/" & CStr(Year(objItemMovEst.dtData)))
        
            'se a data do ajuste do custo standard for menor do que o ultimo movimento de estoque
            If dtData <= gdtDataUltimoMov Then
            
                'verifica se há movimento para o produto que está tendo seu custo standard ajustado a partir da primeiro dia do mes do ajuste.
                'Se tiver ==> tem que reprocessar a partir do primeiro dia do mes do ajuste
                lErro = CF("MovimentoEstoque_Le_ProdutoData", objItemMovEst.iFilialEmpresa, objItemMovEst.sProduto, dtData)
                If lErro <> SUCESSO And lErro <> 83596 Then gError 83598
            
                If lErro = SUCESSO Then objEstoque.dtData = dtData
                
            End If
                
        End If
    
        If objTipoMovEstoque.iValidaDataUltMov = TIPOMOV_EST_VALIDADATAULTMOV And (iStatus <> MOVESTOQUE_CODIGO_JA_CADASTRADO) Then

            'valida a data do movimento de estoque. Não pode ser menor do que a data do ultimo movimento cadastrado
            If objItemMovEst.dtData < gdtDataUltimoMov Then gError 36991
        
        End If
        
    End If

    Trata_Data_MovEstoque = SUCESSO
    
    Exit Function
    
Erro_Trata_Data_MovEstoque:

    Trata_Data_MovEstoque = gErr
    
    Select Case gErr
    
        Case 36992, 83598
        
        Case 36991
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_DATA", gErr, CStr(gdtDataUltimoMov), CStr(objItemMovEst.dtData))
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151613)
            
    End Select
        
    Exit Function
    
End Function

Private Function Trata_Exclusao(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'Exclui o movimento de estoque das tabelas: MovimentoEstoque, RastreamentoMovto e ApropriacaoInsumosProd

Dim lErro As Long

On Error GoTo Erro_Trata_Exclusao

    'excluir da tabela RastreamentoMovto a movimentação
    lErro = RastreamentoMovto_Excluir(alComando(69), alComando(70), objItemMovEst.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE)
    If lErro <> SUCESSO Then gError 71744
    
    'excluir da tabela ApropriacaoInsumosprod e RastreamentoMovto a movimentação dos insumos
    lErro = ApropriacaoInsumosProd_Excluir(alComando(69), alComando(70), alComando(71), alComando(72), objItemMovEst.lNumIntDoc, objTipoMovEstoque)
    If lErro <> SUCESSO Then gError 71745

    'excluir da tabela MovimentoEstoque a movimentação
    lErro = CF("MovimentoEstoque_Excluir", alComando(16), alComando(14), objItemMovEst)
    If lErro <> SUCESSO Then gError 20433

    'excluir da tabela MovimentoEstoque a movimentação
    lErro = CF("LogMovEstoque_Inserir", objItemMovEst, EXCLUSAO_MOVIMENTO)
    If lErro <> SUCESSO Then gError 126166

    'indica que ocorreu pelo menos uma exclusao de movimento de estoque
    giExclusao = 1

    Trata_Exclusao = SUCESSO
    
    Exit Function
    
Erro_Trata_Exclusao:

    Trata_Exclusao = gErr
    
    Select Case gErr
    
        Case 20433, 71744, 71745, 126166
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151614)
            
    End Select
        
    Exit Function

End Function

Private Function Trata_Exclusao_Movimento_Estornado(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'trata da exclusao de movimento que já está estornado

Dim objTipoMovEstoque1 As New ClassTipoMovEst
Dim objItemMovEst1 As New ClassItemMovEstoque
Dim lErro As Long

On Error GoTo Erro_Trata_Exclusao_Movimento_Estornado

    'Carrega os dados do movimento de estorno
    lErro = CF("MovimentoEstoque_Le_NumIntDoc", alComando(14), objItemMovEst)
    If lErro <> SUCESSO And lErro <> 83748 Then gError 89884

    If lErro <> SUCESSO Then gError 89888

    'Exclui o movimento de estoque das tabelas: MovimentoEstoque, RastreamentoMovto e ApropriacaoInsumosProd
    lErro = Trata_Exclusao(alComando, objItemMovEst, objTipoMovEstoque)
    If lErro <> SUCESSO Then gError 89883

    objItemMovEst1.lNumIntDoc = objItemMovEst.lNumIntDocEst

    'Carrega os dados do movimento de estorno
    lErro = CF("MovimentoEstoque_Le_NumIntDoc", alComando(14), objItemMovEst1)
    If lErro <> SUCESSO And lErro <> 83748 Then gError 89884

    If lErro <> SUCESSO Then gError 89889

    objTipoMovEstoque1.iCodigo = objItemMovEst1.iTipoMov
    
    'ler os dados referentes ao tipo de movimento de estorno
    lErro = CF("TiposMovEst_Le1", alComando(10), objTipoMovEstoque1)
    If lErro <> SUCESSO Then gError 89885
    
    'Exclui o estorno do movimento de estoque das tabelas: MovimentoEstoque, RastreamentoMovto e ApropriacaoInsumosProd
    lErro = Trata_Exclusao(alComando, objItemMovEst1, objTipoMovEstoque1)
    If lErro <> SUCESSO Then gError 89886

    Trata_Exclusao_Movimento_Estornado = SUCESSO
    
    Exit Function
    
Erro_Trata_Exclusao_Movimento_Estornado:

    Trata_Exclusao_Movimento_Estornado = gErr
    
    Select Case gErr
    
        Case 89882
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 89883, 89884, 89885, 89886
        
        Case 89888
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO1", gErr, objItemMovEst.lNumIntDoc)
                
        Case 89889
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO1", gErr, objItemMovEst1.lNumIntDoc)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151615)
            
    End Select

    Exit Function

End Function


Private Function Testa_TipoMovEstoque_Devolucao(ByVal objItemMovEstoque As ClassItemMovEstoque, iDevolucao As Integer) As Long
'Testa se o movimento é uma devolução de material (saida). Se for iDevolucao retorna 1.

Dim lErro As Long

On Error GoTo Erro_Testa_TipoMovEstoque_Devolucao

    If objItemMovEstoque.iTipoMov = MOV_EST_NF_DEV_COMPRA Or objItemMovEstoque.iTipoMov = MOV_EST_NF_SAIDA_OUTRAS_DEV_MAT_TERC Or objItemMovEstoque.iTipoMov = MOV_EST_NF_SAIDA_DEV_CONSERTO Or objItemMovEstoque.iTipoMov = MOV_EST_NF_SAIDA_DEV_DEMO Or objItemMovEstoque.iTipoMov = MOV_EST_NF_SAIDA_DEV_CONSIG Then
        iDevolucao = 1
    Else
        iDevolucao = 0
    End If

    Testa_TipoMovEstoque_Devolucao = SUCESSO

    Exit Function

Erro_Testa_TipoMovEstoque_Devolucao:

    Testa_TipoMovEstoque_Devolucao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151616)

    End Select

    Exit Function

End Function
    
'Private Function Estoque_Custo(alComando() As Long, objItemMovEst As ClassItemMovEstoque, objTipoMovEstoque As ClassTipoMovEst, objProduto As ClassProduto, dCustoMedio As Double, objEstoqueMes As ClassEstoqueMes, ByVal iOperacao As Integer) As Long
''calcula o custo do movimento, se necessário
'
'Dim lErro As Long
'Dim dCustoStandard As Double
'Dim dCustoRealProducao As Double
'Dim dCustoMedioProducao As Double
'Dim dQuantCustoAtual As Double
'Dim dQuantConsertoAtual As Double
'Dim dQuantDemoAtual As Double
'Dim dQuantConsigAtual As Double
'Dim dQuantBenefAtual As Double
'Dim dQuantOutrosAtual As Double
'
'On Error GoTo Erro_Estoque_Custo
'
'    'guarda junto ao movimento de estoque como ele é apropriado
'    If objTipoMovEstoque.iCustoInformado = TIPOMOV_EST_CUSTOINFORMADO Then
'        objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO
'    ElseIf objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO Then
'        objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO
'    ElseIf objProduto.iApropriacaoCusto = APROPR_CUSTO_STANDARD Then
'        objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD
'    ElseIf objTipoMovEstoque.iAtualizaCMRProd = TIPOMOV_EST_ATUALIZACMRPDICIONA Then
'        objItemMovEst.iApropriacao = APROPR_CUSTO_REAL
'    Else
'        objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO
'    End If
'
'    'se for um produto nosso saindo de conserto em terceiros ==> o custo é estimado pelo custo médio de produtos em conserto e impacta no custo médio do estoque
'    If (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'
'        lErro = CF("SldMesEst2_Le_CustoMedioConserto",alComando(38), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
'        If lErro <> SUCESSO Then gError 60966
'
'        objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
'
'        If objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
'
'            'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
'            lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
'            If lErro <> SUCESSO Then gError 71503
'
'        Else
'            dCustoMedio = 0
'        End If
'
'    'se for um produto nosso saindo de demonstração em terceiros ==> o custo é estimado pelo custo médio de produtos em demonstração e impacta no custo médio do estoque
'    ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'
'        lErro = CF("SldMesEst2_Le_CustoMedioDemo",alComando(39), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
'        If lErro <> SUCESSO Then gError 60969
'
'        objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
'
'        If objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
'
'            'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
'            lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
'            If lErro <> SUCESSO Then gError 71504
'
'        Else
'            dCustoMedio = 0
'        End If
'
'    'se for um produto nosso saindo de outros em terceiros ==> o custo é estimado pelo custo médio de outros produtos em terceiros e impacta no custo médio do estoque
'    ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'
'        lErro = CF("SldMesEst2_Le_CustoMedioOutros",alComando(40), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
'        If lErro <> SUCESSO Then gError 60972
'
'        objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
'
'        If objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
'
'            'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
'            lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
'            If lErro <> SUCESSO Then gError 71505
'
'        Else
'            dCustoMedio = 0
'        End If
'
'    'se for um produto nosso saindo de beneficiamento em terceiros ==> o custo é estimado pelo custo médio de produtos em beneficiamento em terceiros e impacta no custo médio do estoque
'    ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'
'        lErro = CF("SldMesEst2_Le_CustoMedioBenef",alComando(41), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
'        If lErro <> SUCESSO Then gError 60975
'
'        objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
'
'        If objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
'
'            'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
'            lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
'            If lErro <> SUCESSO Then gError 71506
'
'        Else
'            dCustoMedio = 0
'        End If
'
'    'se for um produto nosso saindo de consignação em terceiros ==> o custo é estimado pelo custo médio de produtos em consignação em terceiros e impacta no custo médio do estoque
'    ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
'
'        lErro = CF("SldMesEst2_Le_CustoMedioConsig",alComando(42), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
'        If lErro <> SUCESSO Then gError 60978
'
'        objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
'
'        If objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
'
'            'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
'            lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
'            If lErro <> SUCESSO Then gError 71507
'
'        Else
'            dCustoMedio = 0
'        End If
'
'    'se for um produto de terceiros em consignação saindo do estoque (não inclui as devoluções que são apropriadas pelo custo informado) ==> o custo é estimado pelo custo médio de produtos de terceiros em consignação
'    ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
'
'        lErro = CF("SldMesEst1_Le_CustoMedioConsig3",alComando(29), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
'        If lErro <> SUCESSO Then gError 71526
'
'        objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
'
'        dCustoMedio = 0
'
'    'se for um produto de terceiros em beneficiamento saindo do estoque (não inclui as devoluções que são apropriadas pelo custo informado) ==> o custo é estimado pelo custo médio de produtos de terceiros em beneficiamento
'    ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
'
'        lErro = CF("SldMesEst1_Le_CustoMedioBenef3",alComando(90), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
'        If lErro <> SUCESSO Then gError 89858
'
'        objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
'
'        dCustoMedio = 0
'
'    'se é um produto nosso em conserto em terceiros e se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material em Conserto nosso que existe no estoque
'    ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD And objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
'
'        'calcula a quantidade de material em conserto nosso em poder de terceiros.
'        lErro = CF("SldMesEstAlm2_Le_QuantConserto",alComando(43), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), objItemMovEst.iAlmoxarifado, dQuantConsertoAtual)
'        If lErro <> SUCESSO Then gError 71529
'
'         objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantConsertoAtual
'
'    'se é um produto nosso em demonstracao em terceiros e se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material em Conserto nosso que existe no estoque
'    ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD And objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
'
'        'calcula a quantidade de material em demonstração nosso em poder de terceiros.
'        lErro = CF("SldMesEstAlm2_Le_QuantDemo",alComando(44), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), objItemMovEst.iAlmoxarifado, dQuantDemoAtual)
'        If lErro <> SUCESSO Then gError 71530
'
'         objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantDemoAtual
'
'    'se é um produto nosso em consignação em terceiros e se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material em Conserto nosso que existe no estoque
'    ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD And objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
'
'        'calcula a quantidade de material em consignação nosso em poder de terceiros.
'        lErro = CF("SldMesEstAlm2_Le_QuantConsig",alComando(45), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), objItemMovEst.iAlmoxarifado, dQuantConsigAtual)
'        If lErro <> SUCESSO Then gError 71531
'
'         objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantConsigAtual
'
'    'se é um produto nosso outros em terceiros e se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material em Conserto nosso que existe no estoque
'    ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD And objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
'
'        'calcula a quantidade de material outros nosso em poder de terceiros.
'        lErro = CF("SldMesEstAlm2_Le_QuantOutros",alComando(46), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), objItemMovEst.iAlmoxarifado, dQuantOutrosAtual)
'        If lErro <> SUCESSO Then gError 71532
'
'         objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantOutrosAtual
'
'    'se é um produto nosso em beneficiamento em terceiros e se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material em Conserto nosso que existe no estoque
'    ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD And objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
'
'        'calcula a quantidade de material em beneficiamento nosso em poder de terceiros.
'        lErro = CF("SldMesEstAlm2_Le_QuantBenef",alComando(47), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), objItemMovEst.iAlmoxarifado, dQuantBenefAtual)
'        If lErro <> SUCESSO Then gError 71533
'
'         objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantBenefAtual
'
'    'se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material nosso que existe no estoque
'    ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD Then
'
'            'retorna a quantidade do produto em estoque que influencia o custo
'            lErro = QuantEstCusto_Atual(alComando(48), objItemMovEst, dQuantCustoAtual)
'            If lErro <> SUCESSO Then gError 71528
'
'            objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantCustoAtual
'
'
'    'se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica que a movimentação incide no calculo do custo médio ==> calcula o novo custo medio
'    ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO And objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO Then
'
'        'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
'        lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
'        If lErro <> SUCESSO Then gError 20374
'
'    'se for material apropriado pelo custo médio e atualiza o custo médio ==> deve ler o custo médio de um dos escaninhos e atualizar o custo médio
'    ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO And objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
'
'        'se o movimento é valorizado pelo custo médio de material em consignação de terceiros
'        If objTipoMovEstoque.iCustoMedio = TIPOMOV_EST_CUSTOMEDIO_CONSIG3 Then
'
'            lErro = CF("SldMesEst1_Le_CustoMedioConsig3",alComando(29), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
'            If lErro <> SUCESSO Then gError 20785
'
'            objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
'
'            'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
'            lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
'            If lErro <> SUCESSO Then gError 20786
'
'        'se o movimento é valorizado pelo custo médio de de material nossso em consignação
'        ElseIf objTipoMovEstoque.iCustoMedio = TIPOMOV_EST_CUSTOMEDIO_CONSIG Then
'
'            lErro = CF("SldMesEst2_Le_CustoMedioConsig",alComando(42), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
'            If lErro <> SUCESSO Then gError 20787
'
'            objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
'
'            'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
'            lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
'            If lErro <> SUCESSO Then gError 20788
'
'        End If
'
'    'se a apropriacao do produto é pelo custo médio e se o custo não é informado ==> apropria pelo custo médio
'    ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO And objTipoMovEstoque.iCustoInformado = TIPOMOV_EST_CUSTONAOINFORMADO Then
'
'        lErro = CF("SldMesEst_Le_CustoMedio",alComando(21), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
'        If lErro <> SUCESSO Then gError 36525
'
'        objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
'
'        dCustoMedio = 0
'
'    'se o produto é para ser apropriado pelo custo standard e o custo não é informado ==> o custo apropriado deve ser o standard
'    ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD And objTipoMovEstoque.iCustoInformado = TIPOMOV_EST_CUSTONAOINFORMADO Then
'
'        lErro = CF("SldMesEst_Le_CustoStandard",alComando(22), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoStandard)
'        If lErro <> SUCESSO Then gError 36526
'
'        'se o custo standard ainda não tiver sido fornecido ==> erro
'        If dCustoStandard = 0 Then gError 60905
'
'        objItemMovEst.dCusto = dCustoStandard * objItemMovEst.dQuantidadeEst
'
'    ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_REAL Or objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO Then
'
'        'se o custo de producao ainda nao estiver apurado ou estiver refazendo os movimentos no reprocessamento ==> o custo vai ficar zerado pois sera apurado posteriormente
'        If objEstoqueMes.iCustoProdApurado = CUSTO_NAO_APURADO Or iOperacao = REPROCESSAMENTO_REFAZ Then
'
'            objItemMovEst.dCusto = 0
'
'        Else
'
'            If objItemMovEst.iApropriacao = APROPR_CUSTO_REAL Then
'
'                'le os dados de custo real de Produção de SldMesEst relativos ao Ano, FilialEmpresa, Produto e mes passados como parametro
'                lErro = CF("SldMesEst_Le_CustoRealProducao",alComando(27), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoRealProducao)
'                If lErro <> SUCESSO Then gError 60911
'
'                objItemMovEst.dCusto = dCustoRealProducao * objItemMovEst.dQuantidadeEst
'
'
'            Else
'
'                'Calcula o Custo Médio de Produção do Produto sProduto
'                lErro = CF("Calcula_CustoMedioProducao",objItemMovEst.iFilialEmpresa, objItemMovEst.sProduto, Year(objItemMovEst.dtData), Month(objItemMovEst.dtData), dCustoMedioProducao)
'                If lErro <> SUCESSO And lErro <> 25433 And lErro <> 55052 Then gError 60912
'
'                'Se não encontrou o SldMesEst
'                If lErro = 25433 Then gError 69863
'
'                'Se não encontrou o SldMesEst
'                If lErro = 55052 Then gError 69864
'
'                objItemMovEst.dCusto = dCustoMedioProducao * objItemMovEst.dQuantidadeEst
'
'            End If
'
'        End If
'
'    End If
'
'    Estoque_Custo = SUCESSO
'
'    Exit Function
'
'Erro_Estoque_Custo:
'
'    Estoque_Custo = gErr
'
'    Select Case gErr
'
'        Case 20374, 20785, 20786, 20787, 20788, 36525, 36526, 60911, 60912, 60917, 60921, 60966, 60969, 60972, 60975, 60978, 71503, 71504, 71505, 71506, 71507, 71526, 71528, 71529, 71530, 71531, 71532, 71533
'
'        Case 60905
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTO_STANDARD_NAO_INFORMADO", gErr, objItemMovEst.sProduto, Month(objItemMovEst.dtData), Year(objItemMovEst.dtData))
'
'        Case 69863
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO", gErr, objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto)
'
'        Case 69864
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO1", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.sProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151617)
'
'    End Select
'
'    Exit Function
'
'End Function
    
Private Function Estoque_Custo(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst, ByVal objProduto As ClassProduto, dCustoMedio As Double, ByVal objEstoqueMes As ClassEstoqueMes, ByVal iOperacao As Integer, ByVal objEstoque As ClassEstoque) As Long
'calcula o custo do movimento, se necessário

Dim lErro As Long
Dim dCustoStandard As Double
Dim dCustoRealProducao As Double
Dim dCustoMedioProducao As Double
Dim dQuantCustoAtual As Double
Dim dQuantConsertoAtual As Double
Dim dQuantDemoAtual As Double
Dim dQuantConsigAtual As Double
Dim dQuantBenefAtual As Double
Dim dQuantOutrosAtual As Double
Dim dCustoUnitarioMovOrig As Double
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iNotaDev As Integer
Dim iTransfer As Integer

On Error GoTo Erro_Estoque_Custo

    'guarda junto ao movimento de estoque como ele é apropriado
    If objTipoMovEstoque.iCustoInformado = TIPOMOV_EST_CUSTOINFORMADO Then
        objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO
    ElseIf objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO Then
        objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO
    ElseIf objProduto.iApropriacaoCusto = APROPR_CUSTO_STANDARD Then
        objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD
    ElseIf objTipoMovEstoque.iAtualizaCMRProd = TIPOMOV_EST_ATUALIZACMRPDICIONA Then
        objItemMovEst.iApropriacao = APROPR_CUSTO_REAL
    Else
        objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO '??? Jones 30/05/2002
    End If

    
    If InStr(1, objItemMovEst.sDocOrigem, " ") > 0 Then
    
        objTipoDocInfo.sSigla = Mid(objItemMovEst.sDocOrigem, 1, InStr(1, objItemMovEst.sDocOrigem, " ") - 1)
        
        lErro = CF("TiposDocInfo_Le_Sigla", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 133361 Then gError 133362
    
        'se se tratar de um movimento oriundo de uma nota fiscal de devolucao ===>
        'pegar o custo da nota fiscal de origem, se houver
        If lErro = SUCESSO And (UCase(objTipoDocInfo.sNomeTelaNFiscal) = "NFISCALDEV" Or UCase(objTipoDocInfo.sNomeTelaNFiscal) = "NFISCALENTDEV") Then
        
            lErro = MovEst_ObtemCustoMovOrig1(objItemMovEst, dCustoUnitarioMovOrig)
            If lErro <> SUCESSO And lErro <> 133366 And lErro <> 133369 Then gError 133371
            
            If lErro = SUCESSO Then
                'se a nf de saida ainda nao foi registrada ou se o produto for produzido o custo pode nao estar disponivel ainda
                If Abs(dCustoUnitarioMovOrig) > DELTA_VALORMONETARIO Then
                    objItemMovEst.dCusto = dCustoUnitarioMovOrig * objItemMovEst.dQuantidadeEst
                    iNotaDev = 1
                End If
            End If
    
        End If

    End If
    
    If objItemMovEst.iTipoMov = 205 Then
    
        lErro = Estoque_Custo_Trata_TipoMov205(objItemMovEst)
        If lErro <> SUCESSO Then gError 133362
    
        iNotaDev = 1
    
    End If

    If iNotaDev = 0 Then

        'se é uma entrada onde o custo pode ter que vir de uma saida de outra filial da mesma empresa
        If gobjCRFAT.lFornEmp <> 0 And (objItemMovEst.lFornecedor = gobjCRFAT.lFornEmp Or objItemMovEst.lCliente = gobjCRFAT.lCliEmp) And objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_ENTRADA And _
            objTipoMovEstoque.iCodigoOrig = 0 And objTipoMovEstoque.iAtualizaSoLote = 0 And objTipoMovEstoque.iAtualizaRecebIndisp = 0 And objTipoMovEstoque.iInventario = 0 Then
            
            lErro = MovEst_ObtemCustoMovOrig(objItemMovEst, dCustoUnitarioMovOrig)
            If lErro <> SUCESSO And lErro <> 81897 And lErro <> 81900 Then gError 81893
            
            If lErro = SUCESSO Then
                'se a nf de saida ainda nao foi registrada ou se o produto for produzido o custo pode nao estar disponivel ainda
                If Abs(dCustoUnitarioMovOrig) > DELTA_VALORMONETARIO Then
                    objItemMovEst.dCusto = dCustoUnitarioMovOrig * objItemMovEst.dQuantidadeEst
                    iTransfer = 1
                End If
                    
            End If
            
        End If
        
        If iTransfer = 0 Then
        
            'se for um produto nosso saindo de conserto em terceiros ==> o custo é estimado pelo custo médio de produtos em conserto e impacta no custo médio do estoque
            If (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
        
                lErro = CF("SldMesEst2_Le_CustoMedioConserto", alComando(38), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                If lErro <> SUCESSO Then gError 60966
                
                objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                If objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
                
                    'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
                    lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
                    If lErro <> SUCESSO Then gError 71503
                    
                Else
                    dCustoMedio = 0
                End If
                
            'se for um produto nosso saindo de demonstração em terceiros ==> o custo é estimado pelo custo médio de produtos em demonstração e impacta no custo médio do estoque
            ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
        
                lErro = CF("SldMesEst2_Le_CustoMedioDemo", alComando(39), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                If lErro <> SUCESSO Then gError 60969
                
                objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                If objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
                
                    'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
                    lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
                    If lErro <> SUCESSO Then gError 71504
               
                Else
                    dCustoMedio = 0
                End If
               
            'se for um produto nosso saindo de outros em terceiros ==> o custo é estimado pelo custo médio de outros produtos em terceiros e impacta no custo médio do estoque
            ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
        
                lErro = CF("SldMesEst2_Le_CustoMedioOutros", alComando(40), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                If lErro <> SUCESSO Then gError 60972
                
                objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                If objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
                
                    'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
                    lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
                    If lErro <> SUCESSO Then gError 71505
                
                Else
                    dCustoMedio = 0
                End If
                
            'se for um produto nosso saindo de beneficiamento em terceiros ==> o custo é estimado pelo custo médio de produtos em beneficiamento em terceiros e impacta no custo médio do estoque
            ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
        
                lErro = CF("SldMesEst2_Le_CustoMedioBenef", alComando(41), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                If lErro <> SUCESSO Then gError 60975
                
                objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                If objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
                
                    'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
                    lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
                    If lErro <> SUCESSO Then gError 71506
                
                Else
                    dCustoMedio = 0
                End If
                
            'se for um produto nosso saindo de consignação em terceiros ==> o custo é estimado pelo custo médio de produtos em consignação em terceiros e impacta no custo médio do estoque
            ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
        
                lErro = CF("SldMesEst2_Le_CustoMedioConsig", alComando(42), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                If lErro <> SUCESSO Then gError 60978
                
                objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                If objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
                
                    'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
                    lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
                    If lErro <> SUCESSO Then gError 71507
                
                Else
                    dCustoMedio = 0
                End If
                
            'se for um produto de terceiros em consignação saindo do estoque (não inclui as devoluções que são apropriadas pelo custo informado) ==> o custo é estimado pelo custo médio de produtos de terceiros em consignação
            ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
        
                lErro = CF("SldMesEst1_Le_CustoMedioConsig3", alComando(29), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                If lErro <> SUCESSO Then gError 71526
                
                objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                dCustoMedio = 0
        
            'se for um produto de terceiros em beneficiamento saindo do estoque (não inclui as devoluções que são apropriadas pelo custo informado) ==> o custo é estimado pelo custo médio de produtos de terceiros em beneficiamento
            ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
        
                lErro = CF("SldMesEst1_Le_CustoMedioBenef3", alComando(90), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                If lErro <> SUCESSO Then gError 89858
                
                objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                dCustoMedio = 0
        
            'se for um produto de terceiros em outros saindo do estoque (não inclui as devoluções que são apropriadas pelo custo informado) ==> o custo é estimado pelo custo médio de produtos de terceiros em beneficiamento
            ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
        
                lErro = CF("SldMesEst1_Le_CustoMedioOutros3", alComando(105), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                If lErro <> SUCESSO Then gError 89858
                
                objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                dCustoMedio = 0
                
            'se for um produto de terceiros em demo saindo do estoque (não inclui as devoluções que são apropriadas pelo custo informado) ==> o custo é estimado pelo custo médio de produtos de terceiros em beneficiamento
            ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
        
                lErro = CF("SldMesEst1_Le_CustoMedioDemo3", alComando(106), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                If lErro <> SUCESSO Then gError 89858
                
                objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                dCustoMedio = 0
                
            'se for um produto de terceiros em conserto saindo do estoque (não inclui as devoluções que são apropriadas pelo custo informado) ==> o custo é estimado pelo custo médio de produtos de terceiros em beneficiamento
            ElseIf (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO Or objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD Or (objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO And objEstoqueMes.iCustoProdApurado = CUSTO_APURADO)) And objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO And objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
        
                lErro = CF("SldMesEst1_Le_CustoMedioConserto3", alComando(107), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                If lErro <> SUCESSO Then gError 89858
                
                objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                dCustoMedio = 0
                
            'se é um produto nosso em conserto em terceiros e se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material em Conserto nosso que existe no estoque
            ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD And objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
        
                'calcula a quantidade de material em conserto nosso em poder de terceiros.
                lErro = CF("SldMesEstAlm2_Le_QuantConserto", alComando(43), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), objItemMovEst.iAlmoxarifado, dQuantConsertoAtual)
                If lErro <> SUCESSO Then gError 71529
        
                 objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantConsertoAtual
        
            'se é um produto nosso em demonstracao em terceiros e se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material em Conserto nosso que existe no estoque
            ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD And objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
        
                'calcula a quantidade de material em demonstração nosso em poder de terceiros.
                lErro = CF("SldMesEstAlm2_Le_QuantDemo", alComando(44), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), objItemMovEst.iAlmoxarifado, dQuantDemoAtual)
                If lErro <> SUCESSO Then gError 71530
        
                 objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantDemoAtual
        
            'se é um produto nosso em consignação em terceiros e se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material em Conserto nosso que existe no estoque
            ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD And objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
        
                'calcula a quantidade de material em consignação nosso em poder de terceiros.
                lErro = CF("SldMesEstAlm2_Le_QuantConsig", alComando(45), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), objItemMovEst.iAlmoxarifado, dQuantConsigAtual)
                If lErro <> SUCESSO Then gError 71531
        
                 objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantConsigAtual
        
            'se é um produto nosso outros em terceiros e se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material em Conserto nosso que existe no estoque
            ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD And objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
        
                'calcula a quantidade de material outros nosso em poder de terceiros.
                lErro = CF("SldMesEstAlm2_Le_QuantOutros", alComando(46), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), objItemMovEst.iAlmoxarifado, dQuantOutrosAtual)
                If lErro <> SUCESSO Then gError 71532
        
                 objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantOutrosAtual
        
            'se é um produto nosso em beneficiamento em terceiros e se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material em Conserto nosso que existe no estoque
            ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD And objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
        
                'calcula a quantidade de material em beneficiamento nosso em poder de terceiros.
                lErro = CF("SldMesEstAlm2_Le_QuantBenef", alComando(47), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), objItemMovEst.iAlmoxarifado, dQuantBenefAtual)
                If lErro <> SUCESSO Then gError 71533
        
                 objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantBenefAtual
        
            'se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica ajuste do custo standard ==> calcula o custo baseado no custo standard unitário que foi passado como parametro e na quantidade de material nosso que existe no estoque
            ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD Then
        
                'retorna a quantidade do produto em estoque que influencia o custo
                lErro = QuantEstCusto_Atual(alComando(48), objItemMovEst, dQuantCustoAtual)
                If lErro <> SUCESSO Then gError 71528
        
                objItemMovEst.dCusto = objItemMovEst.dCusto * dQuantCustoAtual
        
            'se a apropriacao do produto é pelo custo informado e se o tipo do movimento indica que a movimentação incide no calculo do custo médio ==> calcula o novo custo medio
            ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_INFORMADO And objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO And objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO Then
            
                'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
                lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
                If lErro <> SUCESSO Then gError 20374
                
            'se for material apropriado pelo custo médio e atualiza o custo médio ==> deve ler o custo médio de um dos escaninhos e atualizar o custo médio
            ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO And objTipoMovEstoque.iAtualizaCusto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
            
                'se o movimento é valorizado pelo custo médio de material em consignação de terceiros
                If objTipoMovEstoque.iCustoMedio = TIPOMOV_EST_CUSTOMEDIO_CONSIG3 Then
                
                    lErro = CF("SldMesEst1_Le_CustoMedioConsig3", alComando(29), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                    If lErro <> SUCESSO Then gError 20785
                    
                    objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                    'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
                    lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
                    If lErro <> SUCESSO Then gError 20786
                
                'se o movimento é valorizado pelo custo médio de de material nossso em consignação
                ElseIf objTipoMovEstoque.iCustoMedio = TIPOMOV_EST_CUSTOMEDIO_CONSIG Then
                
                    lErro = CF("SldMesEst2_Le_CustoMedioConsig", alComando(42), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                    If lErro <> SUCESSO Then gError 20787
                    
                    objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                    
                    'devolve o novo custo médio do produto em questão em função da movimentação ocorrida
                    lErro = Custo_Medio_Calculo(alComando(20), objItemMovEst, dCustoMedio, objTipoMovEstoque)
                    If lErro <> SUCESSO Then gError 20788
                
                End If
                
            'se a apropriacao do produto é pelo custo médio e se o custo não é informado ==> apropria pelo custo médio
            ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO And objTipoMovEstoque.iCustoInformado = TIPOMOV_EST_CUSTONAOINFORMADO Then
                
                lErro = CF("SldMesEst_Le_CustoMedio", alComando(21), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                If lErro <> SUCESSO Then gError 36525
                
                objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                
                dCustoMedio = 0
                
            'se o produto é para ser apropriado pelo custo standard e o custo não é informado ==> o custo apropriado deve ser o standard
            ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_STANDARD And objTipoMovEstoque.iCustoInformado = TIPOMOV_EST_CUSTONAOINFORMADO Then
            
                lErro = CF("SldMesEst_Le_CustoStandard", alComando(22), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoStandard)
                If lErro <> SUCESSO Then gError 36526
                
                'se o custo standard ainda não tiver sido fornecido ==> erro
                If dCustoStandard = 0 Then gError 60905
                
                objItemMovEst.dCusto = dCustoStandard * objItemMovEst.dQuantidadeEst
            
            ElseIf objItemMovEst.iApropriacao = APROPR_CUSTO_REAL Or objItemMovEst.iApropriacao = APROPR_CUSTO_MEDIO_PRODUCAO Then
            
                'se o custo de producao ainda nao estiver apurado ou estiver refazendo os movimentos no reprocessamento ==> o custo vai ficar zerado pois sera apurado posteriormente
                If objEstoqueMes.iCustoProdApurado = CUSTO_NAO_APURADO Then
                
                    objItemMovEst.dCusto = 0
                    
                Else
                
                    If objItemMovEst.iApropriacao = APROPR_CUSTO_REAL Then
                
                        'calcula o custo real de produção em função do custo dos insumos requisitados, do rateamento dos custos indiretos em função das horas trabalhadas,
                        'do rateamento dos custos diretos em função dos custos dos insumos materias primas (sem incluir embalagens e subprodutos)
                        'a Produção (pois somente produções vão passar por aqui) precisa ter insumos associados ou a OP ter sido baixada.
                        lErro = Calcula_Custo_Real_Producao(objItemMovEst, dCustoRealProducao, objEstoqueMes, objProduto, objEstoque.colEstoqueMesProduto)
                        If lErro <> SUCESSO Then gError 60911
                        
                        objItemMovEst.dCusto = dCustoRealProducao
                        
                        
                    Else
                        
        '                'Calcula o Custo Médio de Produção do Produto sProduto
        '                lErro = CF("Calcula_CustoMedioProducao",objItemMovEst.iFilialEmpresa, objItemMovEst.sProduto, Year(objItemMovEst.dtData), Month(objItemMovEst.dtData), dCustoMedioProducao)
        '                If lErro <> SUCESSO And lErro <> 25433 And lErro <> 55052 Then gError 60912
        '
        '                'Se não encontrou o SldMesEst
        '                If lErro = 25433 Then gError 69863
        '
        '                'Se não encontrou o SldMesEst
        '                If lErro = 55052 Then gError 69864
        '
        '                objItemMovEst.dCusto = dCustoMedioProducao * objItemMovEst.dQuantidadeEst
                        
                        
                        lErro = CF("SldMesEst_Le_CustoMedio", alComando(21), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
                        If lErro <> SUCESSO Then gError 92650
                        
                        objItemMovEst.dCusto = dCustoMedio * objItemMovEst.dQuantidadeEst
                        
                        dCustoMedio = 0
                        
                    End If
                    
                End If
            
            End If

        End If

    End If

    Estoque_Custo = SUCESSO

    Exit Function
    
Erro_Estoque_Custo:

    Estoque_Custo = gErr
    
    Select Case gErr
    
        Case 20374, 20785, 20786, 20787, 20788, 36525, 36526, 60911, 60912, 60917, 60921, 60966, 60969, 60972, 60975, 60978, 71503, 71504, 71505, 71506, 71507, 71526, 71528, 71529, 71530, 71531, 71532, 71533, 81893, 92650, 133362, 133371
            
        Case 60905
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTO_STANDARD_NAO_INFORMADO", gErr, objItemMovEst.sProduto, Month(objItemMovEst.dtData), Year(objItemMovEst.dtData))
            
        Case 69863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO", gErr, objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto)
        
        Case 69864
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO1", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151618)
        
    End Select
        
    Exit Function

End Function
    
Private Function Custo_Medio_Calculo(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dCustoMedio As Double, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'devolve o novo custo médio do produto em questão em função da movimentação ocorrida

Dim objSldMesEst As New ClassSldMesEst
Dim lErro As Long
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim dCustoMedioAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_Custo_Medio_Calculo

    dCustoMedio = 0

    objSldMesEst.iAno = Year(objItemMovEst.dtData)
    objSldMesEst.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldMesEst.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEst_Le_QuantidadeCusto", lComando, objSldMesEst)
    If lErro <> SUCESSO Then Error 60916

    dQuantAtual = objSldMesEst.dQuantInicialCusto
    dValorAtual = objSldMesEst.dValorInicialCusto

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantAtual = dQuantAtual + objSldMesEst.dSaldoQuantCusto(iIndice)
        dValorAtual = dValorAtual + objSldMesEst.dSaldoValorCusto(iIndice)
        
    Next
    
    If objTipoMovEstoque.iAtualizaCusto = TIPOMOV_EST_ATUALIZACUSTOADICIONA Then
    
        If ((dQuantAtual + objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual + objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual + objItemMovEst.dQuantidadeEst)
            
        Else
        
            dCustoMedio = 0
            
        End If
            
    Else
    
        If ((dQuantAtual - objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual - objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual - objItemMovEst.dQuantidadeEst)
        
        Else
        
            dCustoMedio = 0
            
        End If
    
    End If
    
    Custo_Medio_Calculo = SUCESSO

    Exit Function
    
Erro_Custo_Medio_Calculo:

    Custo_Medio_Calculo = Err
    
    Select Case Err
    
        Case 60916
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151619)
        
    End Select
        
    Exit Function

End Function

Private Function Estoque_AtualizaItemMov1(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst, ByVal lNumIntDoc As Long, ByVal iStatus As Integer, ByVal iOperacao As Integer, ByVal objProduto As ClassProduto) As Long
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim iStatus_Rastreamento As Integer
Dim iTipoMov As Integer
Dim lNumIntDocEst As Integer
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objApropriacaoInsumosProd As ClassApropriacaoInsumosProd
Dim objItemMovEst1 As ClassItemMovEstoque
Dim objItemOP As ClassItemOP

On Error GoTo Erro_Estoque_AtualizaItemMov1
    
    'se for reprocessamento desfaz ==> pula este trecho
    If iOperacao <> REPROCESSAMENTO_DESFAZ Then
    
        'se for um estorno ou uma exclusao ==> le o movimento original para verificar se existe e não está estornado
        If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV Or iOperacao = EXCLUSAO_MOVIMENTO Or objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
        
            iTipoMov = objItemMovEst.iTipoMov
            
            'se é uma operação de apuração de custo de produção ou refaz o reprocessamento e estou com o movimento de estorno ==> pegar o custo do movimento que foi estornado para ser este o custo do movimento estornador
            If iOperacao = APURACAO_CUSTO_PRODUCAO Or iOperacao = REPROCESSAMENTO_REFAZ Then
                Set objItemMovEst1 = objItemMovEst
                Set objItemMovEst = New ClassItemMovEstoque
                objItemMovEst.lNumIntDoc = objItemMovEst1.lNumIntDocEst
            End If
            
            'Verifica se o movimento a ser estornado/excluido está cadastrado
            lErro = CF("MovimentoEstoque_Le_Estorno", alComando(14), objItemMovEst)
            If lErro <> SUCESSO And lErro <> 20428 Then gError 20410
            
            'se o movimento não estiver cadastrado ==> erro
            If lErro = 20428 Then gError 20429
            
            If objTipoMovEstoque.iAtualizaOP <> TIPOMOV_EST_NAOATUALIZAOP Then
                objItemMovEst.dQuantidadeOP = objItemMovEst.dQuantidade
                
                'Se é uma exclusão de produção e produziu a mais não pode voltar toda quantidade para QuantOP
                'só o que faltar para produzir (se faltar algo)
                If iOperacao = EXCLUSAO_MOVIMENTO And (objTipoMovEstoque.iCodigo = MOV_EST_PRODUCAO Or objTipoMovEstoque.iCodigo = MOV_EST_PRODUCAO_BENEF3 Or objTipoMovEstoque.iCodigo = MOV_EST_PRODUCAO_OUTROS) Then
                
                    Set objItemOP = New ClassItemOP
                    
                    objItemOP.lNumIntDoc = objItemMovEst.lNumIntDocOrigem
                    
                    lErro = CF("ItemOP_Le_NumIntDoc", objItemOP)
                    If lErro <> SUCESSO And lErro <> 33226 Then gError ERRO_SEM_MENSAGEM
                    
                    If lErro = SUCESSO Then
                        'Produziu a mais
                        If objItemOP.dQuantidade - objItemOP.dQuantidadeProd < objItemMovEst.dQuantidadeOP Then
                            objItemMovEst.dQuantidadeOP = objItemOP.dQuantidade - objItemOP.dQuantidadeProd
                            'Zera caso mesmo após a exclusão a produção seja maior que o projetado
                            If objItemMovEst.dQuantidadeOP < 0 Then objItemMovEst.dQuantidadeOP = 0
                        End If
                    End If
                
                End If
                
                
                objItemMovEst.dQuantidadeOPEst = objItemMovEst.dQuantidadeOPEst
            End If
            
            'se é uma operação de apuração de custo de produção ou refaz o reprocessamento e estou com o movimento de estorno ==> pegar o custo do movimento que foi estornado para ser este o custo do movimento estornador
            If iOperacao = APURACAO_CUSTO_PRODUCAO Or iOperacao = REPROCESSAMENTO_REFAZ Then
            
                objItemMovEst1.dCusto = objItemMovEst.dCusto
                Set objItemMovEst = objItemMovEst1
                
            Else
            
                'se for um estorno ==> gera o novo numero de estorno, se ainda não tiver sido gerado para esta operação
                If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV Then
            
                    If glCodigoEstorno = 0 Then
            
                        'Gera o próximo número de movimento de estoque
                        lErro = CF("MovEstoque_Automatico_EmTransacao", objItemMovEst.iFilialEmpresa, glCodigoEstorno)
                        If lErro <> SUCESSO Then gError 83661
                        
                    End If
                    
                    'coloca o codigo do estorno em objItemMovEst
                    objItemMovEst.lCodigo = glCodigoEstorno
    
                End If
                
                objItemMovEst.iTipoMov = iTipoMov
                
                Set objItemMovEst.colRastreamentoMovto = New Collection
                
                'le os rastreamentos relativos ao movimento de estoque que está sendo estornado/excluido e carrega-os em colRastreamento. Isto é feito pois o rastreamento poderia ter sido alterado indevidamente no momento do estorno
                lErro = RastreamentoMovto_Le(alComando(84), objItemMovEst.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, objItemMovEst.colRastreamentoMovto)
                If lErro <> SUCESSO Then gError 71809
                            
                Set objItemMovEst.colApropriacaoInsumo = New Collection
                
                'le as apropriacoes de insumo de producao relativos ao movimento de estoque que está sendo estornado/excluido e carrega-os em colApropriacaoInsumosProd. Isto é feito pois a apropriacao poderia ter sido alterada indevidamente no momento do estorno
                lErro = ApropriacaoInsumosProd_Le(alComando(85), alComando(86), objItemMovEst.lNumIntDoc, objItemMovEst.colApropriacaoInsumo, objTipoMovEstoque)
                If lErro <> SUCESSO Then gError 71814
            
                'se for um estorno, verifica se o movimento já não está estornado
                If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV And iOperacao <> EXCLUSAO_MOVIMENTO Then
                      
                    'se o movimento já estiver estornado ==> erro
                    If objItemMovEst.lNumIntDocEst <> 0 Then gError 20430
                
                    'o numero interno do documento que está sendo estornado será a origem do estorno
                    objItemMovEst.lNumIntDocOrigem = objItemMovEst.lNumIntDoc
                    objItemMovEst.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE
                    
                    'guarda no movimento de estorno o numero do movimento que ele está estornando
                    objItemMovEst.lNumIntDocEst = objItemMovEst.lNumIntDoc
        
                    'numero interno do movimento de estorno
                    objItemMovEst.lNumIntDoc = lNumIntDoc
                
                    'grava no movimento que será estornado o numero do documento de estorno
                    lErro = CF("MovimentoEstoque_Grava_Estorno", alComando(15), alComando(14), objItemMovEst)
                    If lErro <> SUCESSO Then gError 20432
        
                End If
            
            End If
            
            'se for um exclusao e o documento já estiver estornado ou for o estornador ==> não altera as estatisticas de valor e quantidade, somente exclui os registros
            If iOperacao = EXCLUSAO_MOVIMENTO And objItemMovEst.lNumIntDocEst <> 0 Then
            
                objItemMovEst.dQuantidadeEst = 0
                objItemMovEst.dQuantidadeOPEst = 0
                objItemMovEst.dCusto = 0
                  
                If Not (objItemMovEst.colRastreamentoMovto Is Nothing) Then
            
                    'zera os rastreamentos
                    For Each objRastreamentoMovto In objItemMovEst.colRastreamentoMovto
                        objRastreamentoMovto.dQuantidade = 0
                    Next

                End If
                  
                  
            End If
            
        Else
        
            If iStatus <> MOVESTOQUE_CODIGO_JA_CADASTRADO And iOperacao <> APURACAO_CUSTO_PRODUCAO And iOperacao <> REPROCESSAMENTO_REFAZ Then
            
                'numero interno do movimento que está sendo inserido
                objItemMovEst.lNumIntDoc = lNumIntDoc
                
                'se for um movimento de entrada de material transferido e se o NumIntDocOrigem não estiver preenchido (no caso de transferencia de material consignado numa venda de nf a origem da transf é a nf) ==> guardar em lNumIntDocOrigem o numero do documento associado de Saida da transferencia (origem da transferencia)
                If objTipoMovEstoque.iTransferencia = TIPOMOV_EST_TRANSFERENCIA And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_ENTRADA And objItemMovEst.lNumIntDocOrigem = 0 Then objItemMovEst.lNumIntDocOrigem = lNumIntDoc + 1
                
            End If
                
        End If
    
    End If
    
    If iOperacao = REPROCESSAMENTO_REFAZ And objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD Then
        'numero interno do movimento que está sendo inserido
        objItemMovEst.lNumIntDoc = lNumIntDoc
    End If
    
    'transforma a quantidade do movimento na quantidade de estoque
    lErro = Estoque_Transforma_UM(alComando(), objItemMovEst)
    If lErro <> SUCESSO Then gError 36108
    
    'se for uma exclusão, estorno ou a fase de desfazer os movimentos no reprocessamento ou a fase de excluir os saldos da data original de uma alteração de data ==> coloca o valor do movimento com o sinal contrário ao que o movimento estava
    If iOperacao = EXCLUSAO_MOVIMENTO Or iOperacao = REPROCESSAMENTO_DESFAZ Or objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV Or objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Or iOperacao = ALTERACAO_MOVIMENTO_FASE_EXCLUSAO Then

        'se estiver no reprocessamento executando a fase desfaz e for um movimento de estorno ==> deixar os movimentos com seus valores originais que é o inverso do que foi feito quando o estorno foi processado
        'caso contrário inverter os valores
        If Not (iOperacao = REPROCESSAMENTO_DESFAZ And objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV) Then

            objItemMovEst.dQuantidadeEst = -objItemMovEst.dQuantidadeEst
            objItemMovEst.dQuantidadeOPEst = -objItemMovEst.dQuantidadeOPEst
            objItemMovEst.dCusto = -objItemMovEst.dCusto
    
'??? Esse código parece ser desnecessário - Mário
'            'se o movimento nao for um estorno de um produto apropriado pelo custo de produção ==> vai ser apropriado pelo custo informado
'            'se o movimento for um estorno de um produto apropriado pelo custo de produção ==> o movimento será apropriado pelo mesmo custo que o movimento que ele está estornando, ou seja, custo real de produção ou custo médio de produção
'            If objTipoMovEstoque.iAtualizaMovEstoque <> TIPOMOV_EST_ESTORNOMOV Or objProduto.iApropriacaoCusto <> APROPR_CUSTO_REAL Then
'                objTipoMovEstoque.iCustoInformado = TIPOMOV_EST_CUSTOINFORMADO
'            End If
            
            If Not (objItemMovEst.colRastreamentoMovto Is Nothing) Then
            
                'altera o sinal das quantidades para que possam subtrair dos totalizadores de lote
                For Each objRastreamentoMovto In objItemMovEst.colRastreamentoMovto
                    objRastreamentoMovto.dQuantidadeEst = -objRastreamentoMovto.dQuantidadeEst
                Next

            End If

        End If

    End If
    
    Estoque_AtualizaItemMov1 = SUCESSO
    
    Exit Function
    
Erro_Estoque_AtualizaItemMov1:

    Estoque_AtualizaItemMov1 = gErr
    
    Select Case gErr
    
        Case 20410, 20432, 36108, 71809, 71814, 83661
            
        Case 20429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_NAO_CADASTRADO", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)
        
        Case 20430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151620)
        
    End Select
        
    Exit Function
    
End Function

Function Estoque_AtualizaItemMov2(ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst, ByVal objEstoqueProduto As ClassEstoqueProduto) As Long
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long

On Error GoTo Erro_Estoque_AtualizaItemMov2
    
    objEstoqueProduto.sProduto = objItemMovEst.sProduto
    objEstoqueProduto.iAlmoxarifado = objItemMovEst.iAlmoxarifado
    
    'se a movimentação é referente a um conserto
    If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConserto = objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantConserto3 = objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se é uma movimentação referente a um conserto
    If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConserto = -objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantConserto3 = -objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a uma demonstração
    If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantDemo = objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantDemo3 = objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a uma demonstração
    If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantDemo = -objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantDemo3 = -objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a material em consignação
    If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConsig = objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantConsig3 = objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a material em consignação
    If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConsig = -objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantConsig3 = -objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a outras movimentações de material
    If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantOutras = objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantOutras3 = objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a outras movimentações de material
    If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantOutras = -objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantOutras3 = -objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se a movimentação é referente a material em beneficiamento
    If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantBenef = objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantBenef3 = objItemMovEst.dQuantidadeEst
        End If
    End If
    
    'se é uma movimentação referente a um conserto
    If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantBenef = -objItemMovEst.dQuantidadeEst
        Else
            objEstoqueProduto.dQuantBenef3 = -objItemMovEst.dQuantidadeEst
        End If
    End If
    
    Estoque_AtualizaItemMov2 = SUCESSO
    
    Exit Function
    
Erro_Estoque_AtualizaItemMov2:

    Estoque_AtualizaItemMov2 = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151621)
        
    End Select
        
    Exit Function
    
End Function
    
Function Estoque_AtualizaItemMov3(ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst, ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal objSldDiaEst As ClassSldDiaEst, ByVal iOperacao As Integer) As Long
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long

On Error GoTo Erro_Estoque_AtualizaItemMov3
    
    'se a movimentação é referente a outras movimentações de material indisponivel
    If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
        objEstoqueProduto.dQuantInd = objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a outras movimentações de material indisponivel
    If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
        objEstoqueProduto.dQuantInd = -objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a nosso material disponivel
    If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
        objEstoqueProduto.dQuantDispNossa = objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a nosso material disponivel
    If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
        objEstoqueProduto.dQuantDispNossa = -objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a material defeituoso
    If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
        objEstoqueProduto.dQuantDefeituosa = objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a material defeituoso
    If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
        objEstoqueProduto.dQuantDefeituosa = -objItemMovEst.dQuantidadeEst
    End If
    
    'se a movimentação é referente a material recebido e indisponível
    If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP Then
        objEstoqueProduto.dQuantRecIndl = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dQuantEntRecIndl = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorEntRecIndl = objItemMovEst.dCusto
    End If
    
    'se a movimentação é referente a material recebido e indisponível
    If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP Then
        objEstoqueProduto.dQuantRecIndl = -objItemMovEst.dQuantidadeEst
        objSldDiaEst.dQuantSaiRecIndl = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorSaiRecIndl = objItemMovEst.dCusto
    End If
    
    'se não for uma operação de reprocessamento
    If iOperacao <> REPROCESSAMENTO_DESFAZ And iOperacao <> REPROCESSAMENTO_REFAZ Then
        'se a movimentação é referente a material em ordem de producao
        If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_ADICIONAOP Then
            objEstoqueProduto.dQuantOP = objItemMovEst.dQuantidadeOPEst
        End If
        
        'se a movimentação é referente a material em ordem de producao
        If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_SUBTRAIOP Then
            objEstoqueProduto.dQuantOP = -objItemMovEst.dQuantidadeOPEst
        End If
    
    End If
    
    objSldDiaEst.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldDiaEst.dtData = objItemMovEst.dtData
    objSldDiaEst.sProduto = objItemMovEst.sProduto
    
    'se a movimentação se refere a uma entrada de material no estoque
    If (objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_ENTRADA Or objTipoMovEstoque.iCodigo = MOV_EST_AJUSTE_CUSTO_STD_NOSSO) And objTipoMovEstoque.iCodigo <> MOV_EST_MAT_NOSSO_PARA_BENEF_ENTRADA And objTipoMovEstoque.sEntradaSaidaCMP = objTipoMovEstoque.sEntradaOuSaida Then
'        If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'            objSldDiaEst.dQuantEntrada = -objItemMovEst.dQuantidadeEst
'            objSldDiaEst.dValorEntrada = -objItemMovEst.dCusto
'        Else
            objSldDiaEst.dQuantEntrada = objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorEntrada = objItemMovEst.dCusto
'        End If
    'se a movimentação se refere a uma saida de material do estoque
    ElseIf objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_SAIDA And objTipoMovEstoque.iCodigo <> MOV_EST_MAT_NOSSO_PARA_BENEF_SAIDA And objTipoMovEstoque.sEntradaSaidaCMP = objTipoMovEstoque.sEntradaOuSaida Then
'        If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'            objSldDiaEst.dQuantSaida = -objItemMovEst.dQuantidadeEst
'            objSldDiaEst.dValorSaida = -objItemMovEst.dCusto
'        Else
            objSldDiaEst.dQuantSaida = objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorSaida = objItemMovEst.dCusto
'        End If
    End If
    
    'se a movimentação se refere a consumo de material
    If objTipoMovEstoque.iAtualizaConsumo = TIPOMOV_EST_ADICIONACONSUMO Then
        objSldDiaEst.dQuantCons = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorCons = objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a consumo de material
    If objTipoMovEstoque.iAtualizaConsumo = TIPOMOV_EST_SUBTRAICONSUMO Then
        objSldDiaEst.dQuantCons = -objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorCons = -objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a venda de material
    If objTipoMovEstoque.iAtualizaVenda = TIPOMOV_EST_ADICIONAVENDA Then
        objSldDiaEst.dQuantVend = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorVend = objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a venda de material
    If objTipoMovEstoque.iAtualizaVenda = TIPOMOV_EST_SUBTRAIVENDA Then
        objSldDiaEst.dQuantVend = -objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorVend = -objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a venda de material em consignação de terceiros
    If objTipoMovEstoque.iAtualizaVendaConsig3 = TIPOMOV_EST_ADICIONAVENDACONSIG3 Then
        objSldDiaEst.dQuantVendConsig3 = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorVendConsig3 = objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a venda de material em consignação de terceiros
    If objTipoMovEstoque.iAtualizaVendaConsig3 = TIPOMOV_EST_SUBTRAIVENDACONSIG3 Then
        objSldDiaEst.dQuantVendConsig3 = -objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorVendConsig3 = -objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a compra de material
    If objTipoMovEstoque.iAtualizaCompra = TIPOMOV_EST_ADICIONACOMPRA Then
        objSldDiaEst.dQuantComp = objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorComp = objItemMovEst.dCusto
    End If
    
    'se a movimentação se refere a consumo de material
    If objTipoMovEstoque.iAtualizaCompra = TIPOMOV_EST_SUBTRAICOMPRA Then
        objSldDiaEst.dQuantComp = -objItemMovEst.dQuantidadeEst
        objSldDiaEst.dValorComp = -objItemMovEst.dCusto
    End If
    
    Estoque_AtualizaItemMov3 = SUCESSO
    
    Exit Function
    
Erro_Estoque_AtualizaItemMov3:

    Estoque_AtualizaItemMov3 = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151622)
        
    End Select
        
    Exit Function
    
End Function
    
Function Estoque_AtualizaItemMov4(ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst, ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal objSldDiaEst As ClassSldDiaEst) As Long
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long

On Error GoTo Erro_Estoque_AtualizaItemMov4
    
    'se a movimentação se refere a uma entrada de material no estoque
    If objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_ENTRADA Or objTipoMovEstoque.iCodigo = MOV_EST_AJUSTE_CUSTO_STD_NOSSO Then
    
        'se é para atualizar os saldos que impactam no calculo do custo
        If objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOADICIONA Then
            objSldDiaEst.dQuantEntCusto = objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorEntCusto = objItemMovEst.dCusto
        ElseIf objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOSUBTRAI Then
            objSldDiaEst.dQuantEntCusto = -objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorEntCusto = -objItemMovEst.dCusto
        End If
        
    'se a movimentação se refere a uma saida de material do estoque
    ElseIf objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_SAIDA Then
    
        If objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOADICIONA Then
            objSldDiaEst.dQuantSaiCusto = objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorSaiCusto = objItemMovEst.dCusto
        ElseIf objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOSUBTRAI Then
            objSldDiaEst.dQuantSaiCusto = -objItemMovEst.dQuantidadeEst
            objSldDiaEst.dValorSaiCusto = -objItemMovEst.dCusto
        End If
        
    End If
  
    'se é um produto de terceiros
    If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
        
        'se a movimentação é referente a material consignado
        If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiConsig3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiConsig3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConsig3 = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntConsig3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntConsig3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorConsig3 = objItemMovEst.dCusto
'            End If
               
        ElseIf objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntConsig3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntConsig3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConsig3 = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiConsig3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiConsig3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorConsig3 = -objItemMovEst.dCusto
'            End If
        End If
        
        'se a movimentação é referente a uma demonstração
        If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiDemo3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiDemo3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorDemo3 = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntDemo3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntDemo3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorDemo3 = objItemMovEst.dCusto
'            End If
            
        ElseIf objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntDemo3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntDemo3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorDemo3 = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiDemo3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiDemo3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorDemo3 = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a um conserto
        If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiConserto3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiConserto3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConserto3 = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntConserto3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntConserto3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorConserto3 = objItemMovEst.dCusto
'            End If
            
        ElseIf objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntConserto3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntConserto3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConserto3 = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiConserto3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiConserto3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorConserto3 = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a outras movimentações de material
        If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiOutros3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiOutros3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorOutras3 = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntOutros3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntOutros3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorOutras3 = objItemMovEst.dCusto
'            End If
            
        ElseIf objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntOutros3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntOutros3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorOutras3 = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiOutros3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiOutros3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorOutras3 = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a material em beneficiamento
        If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiBenef3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiBenef3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorBenef3 = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntBenef3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntBenef3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorBenef3 = objItemMovEst.dCusto
'            End If
        ElseIf objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntBenef3 = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntBenef3 = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorBenef3 = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiBenef3 = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiBenef3 = objItemMovEst.dCusto
                objEstoqueProduto.dValorBenef3 = -objItemMovEst.dCusto
'            End If
        End If
     
    End If
           
    'se é um produto nosso
    If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
        
        'se a movimentação é referente a material consignado
        If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiConsig = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiConsig = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConsig = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntConsig = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntConsig = objItemMovEst.dCusto
                objEstoqueProduto.dValorConsig = objItemMovEst.dCusto
'            End If
            
        ElseIf objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntConsig = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntConsig = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConsig = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiConsig = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiConsig = objItemMovEst.dCusto
                objEstoqueProduto.dValorConsig = -objItemMovEst.dCusto
'            End If
        End If
        
        'se a movimentação é referente a uma demonstração
        If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiDemo = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiDemo = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorDemo = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntDemo = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntDemo = objItemMovEst.dCusto
                objEstoqueProduto.dValorDemo = objItemMovEst.dCusto
'            End If
        ElseIf objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntDemo = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntDemo = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorDemo = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiDemo = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiDemo = objItemMovEst.dCusto
                objEstoqueProduto.dValorDemo = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a um conserto
        If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiConserto = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiConserto = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConserto = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntConserto = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntConserto = objItemMovEst.dCusto
                objEstoqueProduto.dValorConserto = objItemMovEst.dCusto
'            End If
        ElseIf objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntConserto = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntConserto = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorConserto = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiConserto = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiConserto = objItemMovEst.dCusto
                objEstoqueProduto.dValorConserto = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a outras movimentações de material
        If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiOutros = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiOutros = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorOutras = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntOutros = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntOutros = objItemMovEst.dCusto
                objEstoqueProduto.dValorOutras = objItemMovEst.dCusto
'            End If
        ElseIf objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntOutros = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntOutros = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorOutras = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiOutros = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiOutros = objItemMovEst.dCusto
                objEstoqueProduto.dValorOutras = -objItemMovEst.dCusto
'            End If
        End If
     
        'se a movimentação é referente a material em beneficiamento
        If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era SUBTRAI que foi trocada para ADICIONA para facilitar a exclusao
'                objSldDiaEst.dQuantSaiBenef = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorSaiBenef = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorBenef = objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantEntBenef = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorEntBenef = objItemMovEst.dCusto
                objEstoqueProduto.dValorBenef = objItemMovEst.dCusto
'            End If
        ElseIf objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
'            If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_EXCLUIMOV Then
'                'o Atualiza original era ADICIONA que foi trocada para SUBTRAI para facilitar a exclusao
'                objSldDiaEst.dQuantEntBenef = -objItemMovEst.dQuantidadeEst
'                objSldDiaEst.dValorEntBenef = -objItemMovEst.dCusto
'                objEstoqueProduto.dValorBenef = -objItemMovEst.dCusto
'            Else
                objSldDiaEst.dQuantSaiBenef = objItemMovEst.dQuantidadeEst
                objSldDiaEst.dValorSaiBenef = objItemMovEst.dCusto
                objEstoqueProduto.dValorBenef = -objItemMovEst.dCusto
'            End If
        End If
     
    End If
           
    Estoque_AtualizaItemMov4 = SUCESSO
    
    Exit Function
    
Erro_Estoque_AtualizaItemMov4:

    Estoque_AtualizaItemMov4 = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151623)
        
    End Select
        
    Exit Function
    
End Function
    
'Private Function Estoque_AtualizaItemMov5(objItemMovEst As ClassItemMovEstoque, objTipoMovEstoque As ClassTipoMovEst, objEstoqueProduto As ClassEstoqueProduto, objSldDiaEst As ClassSldDiaEst) As Long
''ATENCAO: Esta Funcao tem que ser chamada dentro de transacao
'
'Dim lErro As Long
'
'On Error GoTo Erro_Estoque_AtualizaItemMov5
'
'    'se a movimentação se refere a uma entrada de material no estoque
'    If objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_ENTRADA Or objTipoMovEstoque.iCodigo = MOV_EST_AJUSTE_CUSTO_STD_NOSSO Then
'
'        If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_ADICIONAOP Then
'            objSldDiaEst.dQuantEntOP = objItemMovEst.dQuantidadeEst
'        ElseIf objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_SUBTRAIOP Then
'            objSldDiaEst.dQuantEntOP = -objItemMovEst.dQuantidadeEst
'        End If
'
'        If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
'            objSldDiaEst.dQuantEntIndOutras = objItemMovEst.dQuantidadeEst
'        ElseIf objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
'            objSldDiaEst.dQuantEntIndOutras = -objItemMovEst.dQuantidadeEst
'        End If
'
'        If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
'            objSldDiaEst.dQuantEntDispNossa = objItemMovEst.dQuantidadeEst
'        ElseIf objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
'            objSldDiaEst.dQuantEntDispNossa = -objItemMovEst.dQuantidadeEst
'        End If
'
'        If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
'            objSldDiaEst.dQuantEntDefeituosa = objItemMovEst.dQuantidadeEst
'        ElseIf objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
'            objSldDiaEst.dQuantEntDefeituosa = -objItemMovEst.dQuantidadeEst
'        End If
'
'        If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP Then
'            objSldDiaEst.dQuantEntRecIndl = objItemMovEst.dQuantidadeEst
'        ElseIf objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP Then
'            objSldDiaEst.dQuantEntRecIndl = -objItemMovEst.dQuantidadeEst
'        End If
'
'    'se a movimentação se refere a uma saida de material do estoque
'    ElseIf objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_SAIDA Then
'
'        If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_ADICIONAOP Then
'            objSldDiaEst.dQuantEntOP = objItemMovEst.dQuantidadeEst
'        ElseIf objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_SUBTRAIOP Then
'            objSldDiaEst.dQuantEntOP = -objItemMovEst.dQuantidadeEst
'        End If
'
'        If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
'            objSldDiaEst.dQuantEntIndOutras = objItemMovEst.dQuantidadeEst
'        ElseIf objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
'            objSldDiaEst.dQuantEntIndOutras = -objItemMovEst.dQuantidadeEst
'        End If
'
'        If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
'            objSldDiaEst.dQuantEntDispNossa = objItemMovEst.dQuantidadeEst
'        ElseIf objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
'            objSldDiaEst.dQuantEntDispNossa = -objItemMovEst.dQuantidadeEst
'        End If
'
'        If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
'            objSldDiaEst.dQuantEntDefeituosa = objItemMovEst.dQuantidadeEst
'        ElseIf objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
'            objSldDiaEst.dQuantEntDefeituosa = -objItemMovEst.dQuantidadeEst
'        End If
'
'        If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP Then
'            objSldDiaEst.dQuantEntRecIndl = objItemMovEst.dQuantidadeEst
'        ElseIf objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP Then
'            objSldDiaEst.dQuantEntRecIndl = -objItemMovEst.dQuantidadeEst
'        End If
'
'
'
'
'    End If
'
'    Estoque_AtualizaItemMov4 = SUCESSO
'
'    Exit Function
'
'Erro_Estoque_AtualizaItemMov5:
'
'    Estoque_AtualizaItemMov5 = Err
'
'    Select Case Err
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151624)
'
'    End Select
'
'    Exit Function
'
'End Function
    
    
Private Function MovEstoque_Verifica_Data(ByVal objMovEst As ClassMovEstoque) As Long
'verifica se a data do conjunto de movimentos é maior ou igual a ultima data de movimento cadastrada
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim lCodigo As Long
Dim objItemMovEst As New ClassItemMovEstoque

On Error GoTo Erro_MovEstoque_Verifica_Data

    lErro = CF("MovimentoEstoque_Le_Ultima_Data", objMovEst.iFilialEmpresa, objItemMovEst)
    If lErro <> SUCESSO And lErro <> 36029 Then Error 36025

    If lErro = SUCESSO Then
        If objItemMovEst.dtData > objMovEst.dtData Then Error 36030
    End If

    MovEstoque_Verifica_Data = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Verifica_Data:

    MovEstoque_Verifica_Data = Err
    
    Select Case Err
    
        Case 36025
    
        Case 36030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_DATA", Err, CStr(objItemMovEst.dtData), CStr(objMovEst.dtData))
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151625)
    
    End Select
    
    Exit Function

End Function

Private Function Estoque_Transforma_UM(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque) As Long
' transforma a quantidade do movimento na quantidade de estoque

Dim lErro As Long
Dim objUnidadeMedida As New ClassUnidadeDeMedida
Dim dQuantidade As Double
Dim objRastreamentoMovto As ClassRastreamentoMovto

On Error GoTo Erro_Estoque_Transforma_UM

    If objItemMovEst.sSiglaUM = objItemMovEst.sSiglaUMEst Then
    
        objItemMovEst.dQuantidadeEst = objItemMovEst.dQuantidade
        objItemMovEst.dQuantidadeOPEst = objItemMovEst.dQuantidadeOP
    
        If Not (objItemMovEst.colRastreamentoMovto Is Nothing) Then
    
            'guarda a unidade de medida com que a quantidade de rastreamento foi cadastrada
            For Each objRastreamentoMovto In objItemMovEst.colRastreamentoMovto
                objRastreamentoMovto.sSiglaUM = objItemMovEst.sSiglaUM
                objRastreamentoMovto.dQuantidadeEst = objRastreamentoMovto.dQuantidade
            Next
    
        End If
    
    Else
    
        objUnidadeMedida.iClasse = objItemMovEst.iClasseUM
        objUnidadeMedida.sSigla = objItemMovEst.sSiglaUM
    
        lErro = CF("UM_Le1", alComando(18), objUnidadeMedida)
        If lErro <> SUCESSO Then Error 36112
    
        dQuantidade = objUnidadeMedida.dQuantidade

        objUnidadeMedida.sSigla = objItemMovEst.sSiglaUMEst
    
        lErro = CF("UM_Le1", alComando(18), objUnidadeMedida)
        If lErro <> SUCESSO Then Error 36113
    
        objItemMovEst.dQuantidadeEst = (objItemMovEst.dQuantidade * dQuantidade) / objUnidadeMedida.dQuantidade
        objItemMovEst.dQuantidadeOPEst = (objItemMovEst.dQuantidadeOP * dQuantidade) / objUnidadeMedida.dQuantidade
        
        If Not (objItemMovEst.colRastreamentoMovto Is Nothing) Then
            'converte as quantidades de rastreamento da unidade da tela para a unidade de estoque com que serão armazenadas
            For Each objRastreamentoMovto In objItemMovEst.colRastreamentoMovto
                objRastreamentoMovto.dQuantidadeEst = objRastreamentoMovto.dQuantidade * (dQuantidade / objUnidadeMedida.dQuantidade)
                objRastreamentoMovto.sSiglaUM = objItemMovEst.sSiglaUM
            Next
        End If

    End If
    
    Estoque_Transforma_UM = SUCESSO
    
    Exit Function
    
Erro_Estoque_Transforma_UM:

    Estoque_Transforma_UM = Err
    
    Select Case Err
    
        Case 36112, 36113
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151626)
    
    End Select
    
    Exit Function

End Function

Private Function Trata_MaterialConsignado(ByVal objMovEst As ClassMovEstoque, ByVal objItemMovEst As ClassItemMovEstoque, ByVal dQuantConsig3 As Double, ByVal objSldDiaEst As ClassSldDiaEst, ByVal lComando As Long, ByVal lComando1 As Long, ByVal objTipoMovEstoque As ClassTipoMovEst, ByVal objProduto As ClassProduto, alComando() As Long, ByVal dQuantBenef3 As Double) As Long
'trata da venda de material consignado/beneficiado. Devem ser geradas transferencias de consignado/beneficiado para disponivel, com atualizacao de custo, e venda de consignado/beneficiado (do material que foi transferido) e venda normal do material nosso disponivel.
'se o material vendido for todo consignado ==> não será gerado movimento de estoque de venda normal

Dim lErro As Long
Dim objItemMovEst1 As ClassItemMovEstoque
Dim dCustoMedio As Double
Dim dQuantDisp As Double
Dim dQuantTotal As Double
Dim dFator As Double

On Error GoTo Erro_Trata_MaterialConsignado

    'zera o totalizador de venda pois o proximo movimento a ser processado é a transferencia de consignado para disponivel
    objSldDiaEst.dQuantVend = 0
    objSldDiaEst.dValorVend = 0
    objSldDiaEst.dQuantSaiCusto = 0
    objSldDiaEst.dValorSaiCusto = 0
    objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS
    dQuantTotal = objItemMovEst.dQuantidadeEst
    objItemMovEst.sSiglaUM = objItemMovEst.sSiglaUMEst

    'Se houve venda de material beneficiado de 3os ==> significa que ela vem antes da venda de material consignado e portanto é o movimento
    'de estoque que está valendo com seu calculo de custo.

    If dQuantBenef3 > 0 Then

        objSldDiaEst.dQuantSaiBenef3 = dQuantBenef3
        objSldDiaEst.dValorSaiBenef3 = objItemMovEst.dCusto
        objSldDiaEst.dQuantSaida = dQuantBenef3
        objSldDiaEst.dValorSaida = objItemMovEst.dCusto
        objItemMovEst.iTipoMov = MOV_EST_SAIDA_TRANSF_BENEF_TERC
        
        objItemMovEst.dQuantidade = dQuantBenef3
        objItemMovEst.dQuantidadeEst = dQuantBenef3
        
        objTipoMovEstoque.iCodigo = objItemMovEst.iTipoMov
    
        'ler os dados referentes ao tipo de movimento
        lErro = CF("TiposMovEst_Le1", alComando(10), objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 83340
        
        'Gera os rastreamentos automaticamente seguindo a ordem FIFO se necessário.
        lErro = CF("Rastreamento_FIFO_Gera", lComando1, objItemMovEst, ESCANINHO_3_EM_BENEF, objProduto)
        If lErro <> SUCESSO Then gError 71776
        
        Set objItemMovEst1 = objMovEst.colItens.Add(0, MOV_EST_ENTRADA_TRANSF_DISP4, objItemMovEst.dCusto, TIPOMOV_EST_CUSTOINFORMADO, objItemMovEst.sProduto, objItemMovEst.sProdutoDesc, objItemMovEst.sSiglaUM, objItemMovEst.dQuantidade, objItemMovEst.iAlmoxarifado, "", objItemMovEst.lNumIntDocOrigem, objItemMovEst.sCcl, objItemMovEst.lNumIntDocEst, objItemMovEst.sOPCodigo, objItemMovEst.sProdutoOP, objItemMovEst.sContaContabilAplic, objItemMovEst.sContaContabilEst, 0, Nothing, Nothing, DATA_NULA)
        
        objItemMovEst1.iTipoNumIntDocOrigem = objItemMovEst.iTipoNumIntDocOrigem
        Set objItemMovEst1.colRastreamentoMovto = objItemMovEst.colRastreamentoMovto
        
        Set objItemMovEst1 = objMovEst.colItens.Add(0, MOV_EST_NF_VENDA_MAT_BENEF3, 0, objItemMovEst.iApropriacao, objItemMovEst.sProduto, objItemMovEst.sProdutoDesc, objItemMovEst.sSiglaUM, objItemMovEst.dQuantidade, objItemMovEst.iAlmoxarifado, "", objItemMovEst.lNumIntDocOrigem, objItemMovEst.sCcl, objItemMovEst.lNumIntDocEst, objItemMovEst.sOPCodigo, objItemMovEst.sProdutoOP, objItemMovEst.sContaContabilAplic, objItemMovEst.sContaContabilEst, 0, Nothing, Nothing, DATA_NULA)
        
        objItemMovEst1.iTipoNumIntDocOrigem = objItemMovEst.iTipoNumIntDocOrigem
        Set objItemMovEst1.colRastreamentoMovto = objItemMovEst.colRastreamentoMovto


    End If
    
    If dQuantConsig3 > 0 Then
    
        'se houve venda de material consignado mas antes houve a venda do material beneficiado ==> o custo do material consignado precisa ser calculado principalmente para a entrada de material disponivel transferido
        If dQuantBenef3 > 0 Then
        
            'calcula o custo médio de material de terceiros em consignação pois o proximo movimento a ser processado é a transferencia de material consignado de terceiros para disponível
            lErro = CF("SldMesEst1_Le_CustoMedioConsig3", alComando(29), objItemMovEst.iFilialEmpresa, Year(objItemMovEst.dtData), objItemMovEst.sProduto, Month(objItemMovEst.dtData), dCustoMedio)
            If lErro <> SUCESSO Then gError 60952
    
            objItemMovEst.dCusto = dQuantConsig3 * dCustoMedio
    
        Else
    
            'se não houve venda de material beneficiado é sinal que este movimento de consignacao esta sendo processado agora e o seu custo já foi apurado.
    
            objSldDiaEst.dQuantSaiConsig3 = dQuantConsig3
            objSldDiaEst.dValorSaiConsig3 = objItemMovEst.dCusto
            objSldDiaEst.dQuantSaida = dQuantConsig3
            objSldDiaEst.dValorSaida = objItemMovEst.dCusto
        
        End If
        
        objItemMovEst.iTipoMov = MOV_EST_SAIDA_TRANSF_CONSIG_TERC
        objItemMovEst.dQuantidade = dQuantConsig3
        objItemMovEst.dQuantidadeEst = dQuantConsig3
        
        objTipoMovEstoque.iCodigo = objItemMovEst.iTipoMov
    
        'ler os dados referentes ao tipo de movimento
        lErro = CF("TiposMovEst_Le1", alComando(10), objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 83340
        
        'Gera os rastreamentos automaticamente seguindo a ordem FIFO se necessário.
        lErro = CF("Rastreamento_FIFO_Gera", lComando1, objItemMovEst, ESCANINHO_3_EM_CONSIGNACAO, objProduto)
        If lErro <> SUCESSO Then gError 71776
        
        Set objItemMovEst1 = objMovEst.colItens.Add(0, MOV_EST_ENTRADA_TRANSF_DISP1, objItemMovEst.dCusto, TIPOMOV_EST_CUSTOINFORMADO, objItemMovEst.sProduto, objItemMovEst.sProdutoDesc, objItemMovEst.sSiglaUM, objItemMovEst.dQuantidade, objItemMovEst.iAlmoxarifado, "", objItemMovEst.lNumIntDocOrigem, objItemMovEst.sCcl, objItemMovEst.lNumIntDocEst, objItemMovEst.sOPCodigo, objItemMovEst.sProdutoOP, objItemMovEst.sContaContabilAplic, objItemMovEst.sContaContabilEst, 0, Nothing, Nothing, DATA_NULA)
        
        objItemMovEst1.iTipoNumIntDocOrigem = objItemMovEst.iTipoNumIntDocOrigem
        Set objItemMovEst1.colRastreamentoMovto = objItemMovEst.colRastreamentoMovto
        
        Set objItemMovEst1 = objMovEst.colItens.Add(0, MOV_EST_NF_VENDA_MAT_CONSIG, 0, objItemMovEst.iApropriacao, objItemMovEst.sProduto, objItemMovEst.sProdutoDesc, objItemMovEst.sSiglaUM, objItemMovEst.dQuantidade, objItemMovEst.iAlmoxarifado, "", objItemMovEst.lNumIntDocOrigem, objItemMovEst.sCcl, objItemMovEst.lNumIntDocEst, objItemMovEst.sOPCodigo, objItemMovEst.sProdutoOP, objItemMovEst.sContaContabilAplic, objItemMovEst.sContaContabilEst, 0, Nothing, Nothing, DATA_NULA)
        
        objItemMovEst1.iTipoNumIntDocOrigem = objItemMovEst.iTipoNumIntDocOrigem
        Set objItemMovEst1.colRastreamentoMovto = objItemMovEst.colRastreamentoMovto
    
    End If
    
    If dQuantTotal <> dQuantConsig3 + dQuantBenef3 Then
        
        Set objItemMovEst1 = objMovEst.colItens.Add(0, MOV_EST_NF_VENDA, 0, objItemMovEst.iApropriacao, objItemMovEst.sProduto, objItemMovEst.sProdutoDesc, objItemMovEst.sSiglaUM, (dQuantTotal - dQuantConsig3 - dQuantBenef3), objItemMovEst.iAlmoxarifado, "", objItemMovEst.lNumIntDocOrigem, objItemMovEst.sCcl, objItemMovEst.lNumIntDocEst, objItemMovEst.sOPCodigo, objItemMovEst.sProdutoOP, objItemMovEst.sContaContabilAplic, objItemMovEst.sContaContabilEst, 0, Nothing, Nothing, DATA_NULA)
        
        objItemMovEst1.iTipoNumIntDocOrigem = objItemMovEst.iTipoNumIntDocOrigem
        
    End If

    Trata_MaterialConsignado = SUCESSO

    Exit Function

Erro_Trata_MaterialConsignado:
    
    Trata_MaterialConsignado = gErr
    
    Select Case gErr
    
        Case 71776, 83340
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151627)
        
    End Select
    
    Exit Function

End Function

Private Function QuantEstCusto_Atual(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dQuantCustoAtual As Double) As Long
'retorna a quantidade do produto em estoque que influencia o custo

Dim objSldMesEstAlm As New ClassSldMesEstAlm
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_QuantEstCusto_Atual

    objSldMesEstAlm.iAno = Year(objItemMovEst.dtData)
    objSldMesEstAlm.iAlmoxarifado = objItemMovEst.iAlmoxarifado
    objSldMesEstAlm.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEstAlm_Le_QuantidadeCusto", lComando, objSldMesEstAlm)
    If lErro <> SUCESSO Then gError 71527

    dQuantCustoAtual = objSldMesEstAlm.dQuantInicialCusto

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantCustoAtual = dQuantCustoAtual + objSldMesEstAlm.dSaldoQuantCusto(iIndice)
        
    Next
    
    QuantEstCusto_Atual = SUCESSO

    Exit Function
    
Erro_QuantEstCusto_Atual:

    QuantEstCusto_Atual = gErr
    
    Select Case gErr
    
        Case 71527
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151628)
        
    End Select
        
    Exit Function

End Function

Private Function Custo_MedioConsig3_Calculo(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dCustoMedio As Double, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'calcula o novo custo médio dos produtos de terceiros em consignação em função da movimentação ocorrida

Dim objSldMesEst1 As New ClassSldMesEst1
Dim lErro As Long
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim dCustoMedioAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_Custo_MedioConsig3_Calculo

    dCustoMedio = 0

    objSldMesEst1.iAno = Year(objItemMovEst.dtData)
    objSldMesEst1.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldMesEst1.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEst1_Le_QuantValorConsig3", lComando, objSldMesEst1)
    If lErro <> SUCESSO Then gError 71614

    dQuantAtual = objSldMesEst1.dQuantInicialConsig3
    dValorAtual = objSldMesEst1.dValorInicialConsig3

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantAtual = dQuantAtual + objSldMesEst1.dSaldoQuantConsig3(iIndice)
        dValorAtual = dValorAtual + objSldMesEst1.dSaldoValorConsig3(iIndice)
        
    Next
    
    If objTipoMovEstoque.iAtualizaCustoConsig3 = TIPOMOV_EST_ATUALIZACUSTOADICIONA Then
    
        If ((dQuantAtual + objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual + objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual + objItemMovEst.dQuantidadeEst)
            
        Else
        
            dCustoMedio = 0
            
        End If
            
    Else
    
        If ((dQuantAtual - objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual - objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual - objItemMovEst.dQuantidadeEst)
        
        Else
        
            dCustoMedio = 0
            
        End If
    
    End If
    
    Custo_MedioConsig3_Calculo = SUCESSO

    Exit Function
    
Erro_Custo_MedioConsig3_Calculo:

    Custo_MedioConsig3_Calculo = gErr
    
    Select Case gErr
    
        Case 71614
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151629)
        
    End Select
        
    Exit Function

End Function

Private Function Custo_MedioDemo3_Calculo(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dCustoMedio As Double, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'calcula o novo custo médio do produto de terceiros em demonstração em função da movimentação ocorrida

Dim objSldMesEst1 As New ClassSldMesEst1
Dim lErro As Long
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim dCustoMedioAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_Custo_MedioDemo3_Calculo

    dCustoMedio = 0

    objSldMesEst1.iAno = Year(objItemMovEst.dtData)
    objSldMesEst1.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldMesEst1.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEst1_Le_QuantValorDemo3", lComando, objSldMesEst1)
    If lErro <> SUCESSO Then gError 71615

    dQuantAtual = objSldMesEst1.dQuantInicialDemo3
    dValorAtual = objSldMesEst1.dValorInicialDemo3

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantAtual = dQuantAtual + objSldMesEst1.dSaldoQuantDemo3(iIndice)
        dValorAtual = dValorAtual + objSldMesEst1.dSaldoValorDemo3(iIndice)
        
    Next
    
    If objTipoMovEstoque.iAtualizaCustoDemo3 = TIPOMOV_EST_ATUALIZACUSTOADICIONA Then
    
        If ((dQuantAtual + objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual + objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual + objItemMovEst.dQuantidadeEst)
            
        Else
        
            dCustoMedio = 0
            
        End If
            
    Else
    
        If ((dQuantAtual - objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual - objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual - objItemMovEst.dQuantidadeEst)
        
        Else
        
            dCustoMedio = 0
            
        End If
    
    End If
    
    Custo_MedioDemo3_Calculo = SUCESSO

    Exit Function
    
Erro_Custo_MedioDemo3_Calculo:

    Custo_MedioDemo3_Calculo = gErr
    
    Select Case gErr
    
        Case 71615
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151630)
        
    End Select
        
    Exit Function

End Function

Private Function Custo_MedioConserto3_Calculo(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dCustoMedio As Double, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'calcula o novo custo médio do produto de terceiros em conserto em função da movimentação ocorrida

Dim objSldMesEst1 As New ClassSldMesEst1
Dim lErro As Long
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim dCustoMedioAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_Custo_MedioConserto3_Calculo

    dCustoMedio = 0

    objSldMesEst1.iAno = Year(objItemMovEst.dtData)
    objSldMesEst1.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldMesEst1.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEst1_Le_QuantValorConserto3", lComando, objSldMesEst1)
    If lErro <> SUCESSO Then gError 71616

    dQuantAtual = objSldMesEst1.dQuantInicialConserto3
    dValorAtual = objSldMesEst1.dValorInicialConserto3

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantAtual = dQuantAtual + objSldMesEst1.dSaldoQuantConserto3(iIndice)
        dValorAtual = dValorAtual + objSldMesEst1.dSaldoValorConserto3(iIndice)
        
    Next
    
    If objTipoMovEstoque.iAtualizaCustoConserto3 = TIPOMOV_EST_ATUALIZACUSTOADICIONA Then
    
        If ((dQuantAtual + objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual + objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual + objItemMovEst.dQuantidadeEst)
            
        Else
        
            dCustoMedio = 0
            
        End If
            
    Else
    
        If ((dQuantAtual - objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual - objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual - objItemMovEst.dQuantidadeEst)
        
        Else
        
            dCustoMedio = 0
            
        End If
    
    End If
    
    Custo_MedioConserto3_Calculo = SUCESSO

    Exit Function
    
Erro_Custo_MedioConserto3_Calculo:

    Custo_MedioConserto3_Calculo = gErr
    
    Select Case gErr
    
        Case 71616
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151631)
        
    End Select
        
    Exit Function

End Function

Private Function Custo_MedioOutros3_Calculo(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dCustoMedio As Double, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'calcula o novo custo médio de outros produtos de terceiros em função da movimentação ocorrida

Dim objSldMesEst1 As New ClassSldMesEst1
Dim lErro As Long
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim dCustoMedioAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_Custo_MedioOutros3_Calculo

    dCustoMedio = 0

    objSldMesEst1.iAno = Year(objItemMovEst.dtData)
    objSldMesEst1.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldMesEst1.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEst1_Le_QuantValorOutros3", lComando, objSldMesEst1)
    If lErro <> SUCESSO Then gError 71617

    dQuantAtual = objSldMesEst1.dQuantInicialOutros3
    dValorAtual = objSldMesEst1.dValorInicialOutros3

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantAtual = dQuantAtual + objSldMesEst1.dSaldoQuantOutros3(iIndice)
        dValorAtual = dValorAtual + objSldMesEst1.dSaldoValorOutros3(iIndice)
        
    Next
    
    If objTipoMovEstoque.iAtualizaCustoOutros3 = TIPOMOV_EST_ATUALIZACUSTOADICIONA Then
    
        If ((dQuantAtual + objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual + objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual + objItemMovEst.dQuantidadeEst)
            
        Else
        
            dCustoMedio = 0
            
        End If
            
    Else
    
        If ((dQuantAtual - objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual - objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual - objItemMovEst.dQuantidadeEst)
        
        Else
        
            dCustoMedio = 0
            
        End If
    
    End If
    
    Custo_MedioOutros3_Calculo = SUCESSO

    Exit Function
    
Erro_Custo_MedioOutros3_Calculo:

    Custo_MedioOutros3_Calculo = gErr
    
    Select Case gErr
    
        Case 71617
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151632)
        
    End Select
        
    Exit Function

End Function

Private Function Custo_MedioBenef3_Calculo(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dCustoMedio As Double, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'calcula o novo custo médio do produto de terceiros em conserto em função da movimentação ocorrida

Dim objSldMesEst1 As New ClassSldMesEst1
Dim lErro As Long
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim dCustoMedioAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_Custo_MedioBenef3_Calculo

    dCustoMedio = 0

    objSldMesEst1.iAno = Year(objItemMovEst.dtData)
    objSldMesEst1.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldMesEst1.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEst1_Le_QuantValorBenef3", lComando, objSldMesEst1)
    If lErro <> SUCESSO Then gError 71618

    dQuantAtual = objSldMesEst1.dQuantInicialBenef3
    dValorAtual = objSldMesEst1.dValorInicialBenef3

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantAtual = dQuantAtual + objSldMesEst1.dSaldoQuantBenef3(iIndice)
        dValorAtual = dValorAtual + objSldMesEst1.dSaldoValorBenef3(iIndice)
        
    Next
    
    If objTipoMovEstoque.iAtualizaCustoBenef3 = TIPOMOV_EST_ATUALIZACUSTOADICIONA Then
    
        If ((dQuantAtual + objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual + objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual + objItemMovEst.dQuantidadeEst)
            
        Else
        
            dCustoMedio = 0
            
        End If
            
    Else
    
        If ((dQuantAtual - objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual - objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual - objItemMovEst.dQuantidadeEst)
        
        Else
        
            dCustoMedio = 0
            
        End If
    
    End If
    
    Custo_MedioBenef3_Calculo = SUCESSO

    Exit Function
    
Erro_Custo_MedioBenef3_Calculo:

    Custo_MedioBenef3_Calculo = gErr
    
    Select Case gErr
    
        Case 71618
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151633)
        
    End Select
        
    Exit Function

End Function

Private Function Custo_MedioConsig_Calculo(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dCustoMedio As Double, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'calcula o novo custo médio dos produtos nossos em consignação em função da movimentação ocorrida

Dim objSldMesEst2 As New ClassSldMesEst2
Dim lErro As Long
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim dCustoMedioAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_Custo_MedioConsig_Calculo

    dCustoMedio = 0

    objSldMesEst2.iAno = Year(objItemMovEst.dtData)
    objSldMesEst2.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldMesEst2.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEst2_Le_QuantValorConsig", lComando, objSldMesEst2)
    If lErro <> SUCESSO Then gError 71630

    dQuantAtual = objSldMesEst2.dQuantInicialConsig
    dValorAtual = objSldMesEst2.dValorInicialConsig

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantAtual = dQuantAtual + objSldMesEst2.dSaldoQuantConsig(iIndice)
        dValorAtual = dValorAtual + objSldMesEst2.dSaldoValorConsig(iIndice)
        
    Next
    
    If objTipoMovEstoque.iAtualizaCustoConsig = TIPOMOV_EST_ATUALIZACUSTOADICIONA Then
    
        If ((dQuantAtual + objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual + objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual + objItemMovEst.dQuantidadeEst)
            
        Else
        
            dCustoMedio = 0
            
        End If
            
    Else
    
        If ((dQuantAtual - objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual - objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual - objItemMovEst.dQuantidadeEst)
        
        Else
        
            dCustoMedio = 0
            
        End If
    
    End If
    
    Custo_MedioConsig_Calculo = SUCESSO

    Exit Function
    
Erro_Custo_MedioConsig_Calculo:

    Custo_MedioConsig_Calculo = gErr
    
    Select Case gErr
    
        Case 71630
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151634)
        
    End Select
        
    Exit Function

End Function

Private Function Custo_MedioDemo_Calculo(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dCustoMedio As Double, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'calcula o novo custo médio dos produtos nossos em demo em função da movimentação ocorrida

Dim objSldMesEst2 As New ClassSldMesEst2
Dim lErro As Long
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim dCustoMedioAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_Custo_MedioDemo_Calculo

    dCustoMedio = 0

    objSldMesEst2.iAno = Year(objItemMovEst.dtData)
    objSldMesEst2.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldMesEst2.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEst2_Le_QuantValorDemo", lComando, objSldMesEst2)
    If lErro <> SUCESSO Then gError 71631

    dQuantAtual = objSldMesEst2.dQuantInicialDemo
    dValorAtual = objSldMesEst2.dValorInicialDemo

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantAtual = dQuantAtual + objSldMesEst2.dSaldoQuantDemo(iIndice)
        dValorAtual = dValorAtual + objSldMesEst2.dSaldoValorDemo(iIndice)
        
    Next
    
    If objTipoMovEstoque.iAtualizaCustoDemo = TIPOMOV_EST_ATUALIZACUSTOADICIONA Then
    
        If ((dQuantAtual + objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual + objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual + objItemMovEst.dQuantidadeEst)
            
        Else
        
            dCustoMedio = 0
            
        End If
            
    Else
    
        If ((dQuantAtual - objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual - objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual - objItemMovEst.dQuantidadeEst)
        
        Else
        
            dCustoMedio = 0
            
        End If
    
    End If
    
    Custo_MedioDemo_Calculo = SUCESSO

    Exit Function
    
Erro_Custo_MedioDemo_Calculo:

    Custo_MedioDemo_Calculo = gErr
    
    Select Case gErr
    
        Case 71631
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151635)
        
    End Select
        
    Exit Function

End Function

Private Function Custo_MedioConserto_Calculo(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dCustoMedio As Double, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'calcula o novo custo médio dos produtos nossos em Conserto em função da movimentação ocorrida

Dim objSldMesEst2 As New ClassSldMesEst2
Dim lErro As Long
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim dCustoMedioAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_Custo_MedioConserto_Calculo

    dCustoMedio = 0

    objSldMesEst2.iAno = Year(objItemMovEst.dtData)
    objSldMesEst2.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldMesEst2.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEst2_Le_QuantValorConserto", lComando, objSldMesEst2)
    If lErro <> SUCESSO Then gError 71632

    dQuantAtual = objSldMesEst2.dQuantInicialConserto
    dValorAtual = objSldMesEst2.dValorInicialConserto

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantAtual = dQuantAtual + objSldMesEst2.dSaldoQuantConserto(iIndice)
        dValorAtual = dValorAtual + objSldMesEst2.dSaldoValorConserto(iIndice)
        
    Next
    
    If objTipoMovEstoque.iAtualizaCustoConserto = TIPOMOV_EST_ATUALIZACUSTOADICIONA Then
    
        If ((dQuantAtual + objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual + objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual + objItemMovEst.dQuantidadeEst)
            
        Else
        
            dCustoMedio = 0
            
        End If
            
    Else
    
        If ((dQuantAtual - objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual - objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual - objItemMovEst.dQuantidadeEst)
        
        Else
        
            dCustoMedio = 0
            
        End If
    
    End If
    
    Custo_MedioConserto_Calculo = SUCESSO

    Exit Function
    
Erro_Custo_MedioConserto_Calculo:

    Custo_MedioConserto_Calculo = gErr
    
    Select Case gErr
    
        Case 71632
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151636)
        
    End Select
        
    Exit Function

End Function

Private Function Custo_MedioOutros_Calculo(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dCustoMedio As Double, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'calcula o novo custo médio dos produtos nossos em Outros em função da movimentação ocorrida

Dim objSldMesEst2 As New ClassSldMesEst2
Dim lErro As Long
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim dCustoMedioAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_Custo_MedioOutros_Calculo

    dCustoMedio = 0

    objSldMesEst2.iAno = Year(objItemMovEst.dtData)
    objSldMesEst2.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldMesEst2.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEst2_Le_QuantValorOutros", lComando, objSldMesEst2)
    If lErro <> SUCESSO Then gError 71633

    dQuantAtual = objSldMesEst2.dQuantInicialOutros
    dValorAtual = objSldMesEst2.dValorInicialOutros

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantAtual = dQuantAtual + objSldMesEst2.dSaldoQuantOutros(iIndice)
        dValorAtual = dValorAtual + objSldMesEst2.dSaldoValorOutros(iIndice)
        
    Next
    
    If objTipoMovEstoque.iAtualizaCustoOutros = TIPOMOV_EST_ATUALIZACUSTOADICIONA Then
    
        If ((dQuantAtual + objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual + objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual + objItemMovEst.dQuantidadeEst)
            
        Else
        
            dCustoMedio = 0
            
        End If
            
    Else
    
        If ((dQuantAtual - objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual - objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual - objItemMovEst.dQuantidadeEst)
        
        Else
        
            dCustoMedio = 0
            
        End If
    
    End If
    
    Custo_MedioOutros_Calculo = SUCESSO

    Exit Function
    
Erro_Custo_MedioOutros_Calculo:

    Custo_MedioOutros_Calculo = gErr
    
    Select Case gErr
    
        Case 71633
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151637)
        
    End Select
        
    Exit Function

End Function

Private Function Custo_MedioBenef_Calculo(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, dCustoMedio As Double, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'calcula o novo custo médio dos produtos nossos em Benef em função da movimentação ocorrida

Dim objSldMesEst2 As New ClassSldMesEst2
Dim lErro As Long
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim dCustoMedioAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_Custo_MedioBenef_Calculo

    dCustoMedio = 0

    objSldMesEst2.iAno = Year(objItemMovEst.dtData)
    objSldMesEst2.iFilialEmpresa = objItemMovEst.iFilialEmpresa
    objSldMesEst2.sProduto = objItemMovEst.sProduto

    lErro = CF("SldMesEst2_Le_QuantValorBenef", lComando, objSldMesEst2)
    If lErro <> SUCESSO Then gError 71634

    dQuantAtual = objSldMesEst2.dQuantInicialBenef
    dValorAtual = objSldMesEst2.dValorInicialBenef

    For iIndice = 1 To Month(objItemMovEst.dtData)
    
        dQuantAtual = dQuantAtual + objSldMesEst2.dSaldoQuantBenef(iIndice)
        dValorAtual = dValorAtual + objSldMesEst2.dSaldoValorBenef(iIndice)
        
    Next
    
    If objTipoMovEstoque.iAtualizaCustoBenef = TIPOMOV_EST_ATUALIZACUSTOADICIONA Then
    
        If ((dQuantAtual + objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual + objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual + objItemMovEst.dQuantidadeEst)
            
        Else
        
            dCustoMedio = 0
            
        End If
            
    Else
    
        If ((dQuantAtual - objItemMovEst.dQuantidadeEst) <> 0) Then
    
            dValorAtual = dValorAtual - objItemMovEst.dCusto
            dCustoMedio = dValorAtual / (dQuantAtual - objItemMovEst.dQuantidadeEst)
        
        Else
        
            dCustoMedio = 0
            
        End If
    
    End If
    
    Custo_MedioBenef_Calculo = SUCESSO

    Exit Function
    
Erro_Custo_MedioBenef_Calculo:

    Custo_MedioBenef_Calculo = gErr
    
    Select Case gErr
    
        Case 71634
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151638)
        
    End Select
        
    Exit Function

End Function

Private Function Recalcula_CustoMedio_NossoEm3(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst, dCustoMedioConsig As Double, dCustoMedioDemo As Double, dCustoMedioConserto As Double, dCustoMedioOutros As Double, dCustoMedioBenef As Double) As Long
'recalcula o custo médio dos produtos nossos em poder de terceiros, se necessário

Dim lErro As Long

On Error GoTo Erro_Recalcula_CustoMedio_NossoEm3

    If objTipoMovEstoque.iAtualizaCustoConsig <> TIPOMOV_EST_NAOATUALIZACUSTO Then
    
        'Atualiza o custo médio do material em consignação
        lErro = Custo_MedioConsig_Calculo(alComando(49), objItemMovEst, dCustoMedioConsig, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71635
        
    Else
        dCustoMedioConsig = 0
    End If
        
    If objTipoMovEstoque.iAtualizaCustoDemo <> TIPOMOV_EST_NAOATUALIZACUSTO Then
    
        'Atualiza o custo médio do material em demo
        lErro = Custo_MedioDemo_Calculo(alComando(50), objItemMovEst, dCustoMedioDemo, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71636
        
    Else
        dCustoMedioDemo = 0
    End If

    If objTipoMovEstoque.iAtualizaCustoConserto <> TIPOMOV_EST_NAOATUALIZACUSTO Then
    
        'Atualiza o custo médio do material em conserto
        lErro = Custo_MedioConserto_Calculo(alComando(51), objItemMovEst, dCustoMedioConserto, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71637
        
    Else
        dCustoMedioConserto = 0
    End If

    If objTipoMovEstoque.iAtualizaCustoOutros <> TIPOMOV_EST_NAOATUALIZACUSTO Then
    
        'Atualiza o custo médio do  material outros
        lErro = Custo_MedioOutros_Calculo(alComando(52), objItemMovEst, dCustoMedioOutros, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71638
        
    Else
        dCustoMedioOutros = 0
    End If

    If objTipoMovEstoque.iAtualizaCustoBenef <> TIPOMOV_EST_NAOATUALIZACUSTO Then
    
        'Atualiza o custo médio do material em beneficiamento
        lErro = Custo_MedioBenef_Calculo(alComando(53), objItemMovEst, dCustoMedioBenef, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71639
        
    Else
        dCustoMedioBenef = 0
    End If

    Recalcula_CustoMedio_NossoEm3 = SUCESSO

    Exit Function
    
Erro_Recalcula_CustoMedio_NossoEm3:

    Recalcula_CustoMedio_NossoEm3 = gErr
    
    Select Case gErr
    
        Case 71635, 71636, 71637, 71638, 71639
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151639)
        
    End Select
        
    Exit Function

End Function

Private Function Recalcula_CustoMedio_Terceiros(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst, dCustoMedioConsig3 As Double, dCustoMedioDemo3 As Double, dCustoMedioConserto3 As Double, dCustoMedioOutros3 As Double, dCustoMedioBenef3 As Double) As Long
'recalcula o custo médio dos produtos de terceiros em nosso poder, se necessário

Dim lErro As Long

On Error GoTo Erro_Recalcula_CustoMedio_Terceiros

    If objTipoMovEstoque.iAtualizaCustoConsig3 <> TIPOMOV_EST_NAOATUALIZACUSTO Then
    
        'Atualiza o custo médio do material de terceiros em consignação
        lErro = Custo_MedioConsig3_Calculo(alComando(54), objItemMovEst, dCustoMedioConsig3, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71640
        
    Else
        dCustoMedioConsig3 = 0
    End If
        
    If objTipoMovEstoque.iAtualizaCustoDemo3 <> TIPOMOV_EST_NAOATUALIZACUSTO Then
    
        'Atualiza o custo médio do material de terceiros em demo
        lErro = Custo_MedioDemo3_Calculo(alComando(55), objItemMovEst, dCustoMedioDemo3, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71641
        
    Else
        dCustoMedioDemo3 = 0
    End If

    If objTipoMovEstoque.iAtualizaCustoConserto3 <> TIPOMOV_EST_NAOATUALIZACUSTO Then
    
        'Atualiza o custo médio do material de terceiros em conserto
        lErro = Custo_MedioConserto3_Calculo(alComando(56), objItemMovEst, dCustoMedioConserto3, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71642
        
    Else
        dCustoMedioConserto3 = 0
    End If

    If objTipoMovEstoque.iAtualizaCustoOutros3 <> TIPOMOV_EST_NAOATUALIZACUSTO Then
    
        'Atualiza o custo médio do  material outros de terceiros
        lErro = Custo_MedioOutros3_Calculo(alComando(57), objItemMovEst, dCustoMedioOutros3, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71643
        
    Else
        dCustoMedioOutros3 = 0
    End If

    If objTipoMovEstoque.iAtualizaCustoBenef3 <> TIPOMOV_EST_NAOATUALIZACUSTO Then
    
        'Atualiza o custo médio do material de terceiros em beneficiamento
        lErro = Custo_MedioBenef3_Calculo(alComando(58), objItemMovEst, dCustoMedioBenef3, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71644
        
    Else
        dCustoMedioBenef3 = 0
    End If

    Recalcula_CustoMedio_Terceiros = SUCESSO

    Exit Function
    
Erro_Recalcula_CustoMedio_Terceiros:

    Recalcula_CustoMedio_Terceiros = gErr
    
    Select Case gErr
    
        Case 71640, 71641, 71642, 71643, 71644
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151640)
        
    End Select
        
    Exit Function

End Function

Private Function Transforma_TipoMovEst_Exclusao(ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'Transforma as flags do tipo de movimento de estoque para ficar compativel com a exclusao

On Error GoTo Erro_Transforma_TipoMovEst_Exclusao

    objTipoMovEstoque.iValidaDataUltMov = TIPOMOV_EST_NAO_VALIDADATAULTMOVEST

    If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
        objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO1
    ElseIf objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
        objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO
    End If

    If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
        objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO
    ElseIf objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
        objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO
    End If

    If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
        objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO
    ElseIf objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
        objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO
    End If

    If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
        objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS
    ElseIf objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
        objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS
    End If

    If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
        objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF
    ElseIf objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
        objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF
    End If

    If objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOADICIONA Then
        objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOSUBTRAI
    ElseIf objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOSUBTRAI Then
        objTipoMovEstoque.iAtualizaSaldoCusto = TIPOMOV_EST_ATUALIZASALDOCUSTOADICIONA
    End If

    If objTipoMovEstoque.iAtualizaCompra = TIPOMOV_EST_ADICIONACOMPRA Then
        objTipoMovEstoque.iAtualizaCompra = TIPOMOV_EST_SUBTRAICOMPRA
    ElseIf objTipoMovEstoque.iAtualizaCompra = TIPOMOV_EST_SUBTRAICOMPRA Then
        objTipoMovEstoque.iAtualizaCompra = TIPOMOV_EST_ADICIONACOMPRA
    End If
    
    If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
        objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS
    ElseIf objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
        objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS
    End If

    If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
        objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP
    ElseIf objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
        objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP
    End If

    If objTipoMovEstoque.iAtualizaConsumo = TIPOMOV_EST_ADICIONACONSUMO Then
        objTipoMovEstoque.iAtualizaConsumo = TIPOMOV_EST_SUBTRAICONSUMO
    ElseIf objTipoMovEstoque.iAtualizaConsumo = TIPOMOV_EST_SUBTRAICONSUMO Then
        objTipoMovEstoque.iAtualizaConsumo = TIPOMOV_EST_ADICIONACONSUMO
    End If

    If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
        objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA
    ElseIf objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
        objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA
    End If

    If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP Then
        objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP
    ElseIf objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP Then
        objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP
    End If

    If objTipoMovEstoque.iAtualizaVenda = TIPOMOV_EST_ADICIONAVENDA Then
        objTipoMovEstoque.iAtualizaVenda = TIPOMOV_EST_SUBTRAIVENDA
    ElseIf objTipoMovEstoque.iAtualizaVenda = TIPOMOV_EST_SUBTRAIVENDA Then
        objTipoMovEstoque.iAtualizaVenda = TIPOMOV_EST_ADICIONAVENDA
    End If

    If objTipoMovEstoque.iAtualizaVendaConsig3 = TIPOMOV_EST_ADICIONAVENDACONSIG3 Then
        objTipoMovEstoque.iAtualizaVendaConsig3 = TIPOMOV_EST_SUBTRAIVENDACONSIG3
    ElseIf objTipoMovEstoque.iAtualizaVendaConsig3 = TIPOMOV_EST_SUBTRAIVENDACONSIG3 Then
        objTipoMovEstoque.iAtualizaVendaConsig3 = TIPOMOV_EST_ADICIONAVENDACONSIG3
    End If

    If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_ADICIONAOP Then
        objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_SUBTRAIOP
    ElseIf objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_SUBTRAIOP Then
        objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_ADICIONAOP
    End If

    Transforma_TipoMovEst_Exclusao = SUCESSO
    
    Exit Function
    
Erro_Transforma_TipoMovEst_Exclusao:

    Transforma_TipoMovEst_Exclusao = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151641)
        
    End Select
        
    Exit Function

End Function

Private Function Atualiza_RastreamentoLoteSaldo(ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst, ByVal colRastreamentoLoteSaldo As Collection, ByVal colRastreamentoMovto As Collection) As Long
'coloca os valores do lote dentro de escaninho totalizador apropriado

Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objRastreamentoLoteSaldo As ClassRastreamentoLoteSaldo
Dim iAchou As Integer

On Error GoTo Erro_Atualiza_RastreamentoLoteSaldo

    If Not (colRastreamentoMovto Is Nothing) Then
    
        For Each objRastreamentoMovto In colRastreamentoMovto
            
            For Each objRastreamentoLoteSaldo In colRastreamentoLoteSaldo
            
                If objRastreamentoLoteSaldo.sProduto = objRastreamentoMovto.sProduto And UCase(objRastreamentoMovto.sLote) = UCase(objRastreamentoLoteSaldo.sLote) And objRastreamentoMovto.iFilialOP = objRastreamentoLoteSaldo.iFilialOP And objRastreamentoLoteSaldo.iAlmoxarifado = objItemMovEst.iAlmoxarifado Then
                    iAchou = 1
                    Exit For
                End If
                
            Next
            
            If iAchou = 0 Then
                Set objRastreamentoLoteSaldo = New ClassRastreamentoLoteSaldo
                objRastreamentoLoteSaldo.sProduto = objItemMovEst.sProduto
                objRastreamentoLoteSaldo.iAlmoxarifado = objItemMovEst.iAlmoxarifado
                objRastreamentoLoteSaldo.iFilialOP = objRastreamentoMovto.iFilialOP
                objRastreamentoLoteSaldo.sLote = objRastreamentoMovto.sLote
                objRastreamentoLoteSaldo.sLocalizacao = objRastreamentoMovto.sLocalizacao
                If Not (objRastreamentoMovto.objInfoUsu Is Nothing) Then
                    Set objRastreamentoLoteSaldo.objInfoUsu = objRastreamentoMovto.objInfoUsu
                End If
                colRastreamentoLoteSaldo.Add objRastreamentoLoteSaldo
            Else
                iAchou = 0
            End If
            
            'se a movimentação é referente a um conserto
            If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    objRastreamentoLoteSaldo.dQuantConserto = objRastreamentoLoteSaldo.dQuantConserto + objRastreamentoMovto.dQuantidadeEst
                Else
                    objRastreamentoLoteSaldo.dQuantConserto3 = objRastreamentoLoteSaldo.dQuantConserto3 + objRastreamentoMovto.dQuantidadeEst
                End If
            End If
            
            'se é uma movimentação referente a um conserto
            If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    objRastreamentoLoteSaldo.dQuantConserto = objRastreamentoLoteSaldo.dQuantConserto - objRastreamentoMovto.dQuantidadeEst
                Else
                    objRastreamentoLoteSaldo.dQuantConserto3 = objRastreamentoLoteSaldo.dQuantConserto3 - objRastreamentoMovto.dQuantidadeEst
                End If
            End If
            
            'se a movimentação é referente a uma demonstração
            If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    objRastreamentoLoteSaldo.dQuantDemo = objRastreamentoLoteSaldo.dQuantDemo + objRastreamentoMovto.dQuantidadeEst
                Else
                    objRastreamentoLoteSaldo.dQuantDemo3 = objRastreamentoLoteSaldo.dQuantDemo3 + objRastreamentoMovto.dQuantidadeEst
                End If
            End If
            
            'se a movimentação é referente a uma demonstração
            If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    objRastreamentoLoteSaldo.dQuantDemo = objRastreamentoLoteSaldo.dQuantDemo - objRastreamentoMovto.dQuantidadeEst
                Else
                    objRastreamentoLoteSaldo.dQuantDemo3 = objRastreamentoLoteSaldo.dQuantDemo3 - objRastreamentoMovto.dQuantidadeEst
                End If
            End If
            
            'se a movimentação é referente a material em consignação
            If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    objRastreamentoLoteSaldo.dQuantConsig = objRastreamentoLoteSaldo.dQuantConsig + objRastreamentoMovto.dQuantidadeEst
                Else
                    objRastreamentoLoteSaldo.dQuantConsig3 = objRastreamentoLoteSaldo.dQuantConsig3 + objRastreamentoMovto.dQuantidadeEst
                End If
            End If
            
            'se a movimentação é referente a material em consignação
            If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    objRastreamentoLoteSaldo.dQuantConsig = objRastreamentoLoteSaldo.dQuantConsig - objRastreamentoMovto.dQuantidadeEst
                Else
                    objRastreamentoLoteSaldo.dQuantConsig3 = objRastreamentoLoteSaldo.dQuantConsig3 - objRastreamentoMovto.dQuantidadeEst
                End If
            End If
            
            'se a movimentação é referente a outras movimentações de material
            If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    objRastreamentoLoteSaldo.dQuantOutras = objRastreamentoLoteSaldo.dQuantOutras + objRastreamentoMovto.dQuantidadeEst
                Else
                    objRastreamentoLoteSaldo.dQuantOutras3 = objRastreamentoLoteSaldo.dQuantOutras3 + objRastreamentoMovto.dQuantidadeEst
                End If
            End If
            
            'se a movimentação é referente a outras movimentações de material
            If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    objRastreamentoLoteSaldo.dQuantOutras = objRastreamentoLoteSaldo.dQuantOutras - objRastreamentoMovto.dQuantidadeEst
                Else
                    objRastreamentoLoteSaldo.dQuantOutras3 = objRastreamentoLoteSaldo.dQuantOutras3 - objRastreamentoMovto.dQuantidadeEst
                End If
            End If
            
            'se a movimentação é referente a material em beneficiamento
            If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    objRastreamentoLoteSaldo.dQuantBenef = objRastreamentoLoteSaldo.dQuantBenef + objRastreamentoMovto.dQuantidadeEst
                Else
                    objRastreamentoLoteSaldo.dQuantBenef3 = objRastreamentoLoteSaldo.dQuantBenef3 + objRastreamentoMovto.dQuantidadeEst
                End If
            End If
            
            'se é uma movimentação referente a um conserto
            If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
                If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                    objRastreamentoLoteSaldo.dQuantBenef = objRastreamentoLoteSaldo.dQuantBenef - objRastreamentoMovto.dQuantidadeEst
                Else
                    objRastreamentoLoteSaldo.dQuantBenef3 = objRastreamentoLoteSaldo.dQuantBenef3 - objRastreamentoMovto.dQuantidadeEst
                End If
            End If
    
            'se a movimentação é referente a outras movimentações de material indisponivel
            If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
                objRastreamentoLoteSaldo.dQuantIndOutras = objRastreamentoLoteSaldo.dQuantIndOutras + objRastreamentoMovto.dQuantidadeEst
            End If
            
            'se a movimentação é referente a outras movimentações de material indisponivel
            If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
                objRastreamentoLoteSaldo.dQuantIndOutras = objRastreamentoLoteSaldo.dQuantIndOutras - objRastreamentoMovto.dQuantidadeEst
            End If
            
            'se a movimentação é referente a nosso material disponivel
            If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
                objRastreamentoLoteSaldo.dQuantDispNossa = objRastreamentoLoteSaldo.dQuantDispNossa + objRastreamentoMovto.dQuantidadeEst
            End If
            
            'se a movimentação é referente a nosso material disponivel
            If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
                objRastreamentoLoteSaldo.dQuantDispNossa = objRastreamentoLoteSaldo.dQuantDispNossa - objRastreamentoMovto.dQuantidadeEst
            End If
            
            'se a movimentação é referente a material defeituoso
            If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
                objRastreamentoLoteSaldo.dQuantDefeituosa = objRastreamentoLoteSaldo.dQuantDefeituosa + objRastreamentoMovto.dQuantidadeEst
            End If
            
            'se a movimentação é referente a material defeituoso
            If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
                objRastreamentoLoteSaldo.dQuantDefeituosa = objRastreamentoLoteSaldo.dQuantDefeituosa - objRastreamentoMovto.dQuantidadeEst
            End If
            
            'se a movimentação é referente a material recebido e indisponível
            If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP Then
                objRastreamentoLoteSaldo.dQuantRecIndl = objRastreamentoLoteSaldo.dQuantRecIndl + objRastreamentoMovto.dQuantidadeEst
            End If
            
            'se a movimentação é referente a material recebido e indisponível
            If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP Then
                objRastreamentoLoteSaldo.dQuantRecIndl = objRastreamentoLoteSaldo.dQuantRecIndl - objRastreamentoMovto.dQuantidadeEst
            End If
            
            'se a movimentação é referente a material em ordem de producao
            If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_ADICIONAOP Then
                objRastreamentoLoteSaldo.dQuantOP = objRastreamentoLoteSaldo.dQuantOP + objRastreamentoMovto.dQuantidadeEst
            End If
            
            'se a movimentação é referente a material em ordem de producao
            If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_SUBTRAIOP Then
                objRastreamentoLoteSaldo.dQuantOP = objRastreamentoLoteSaldo.dQuantOP - objRastreamentoMovto.dQuantidadeEst
            End If
            
        Next

    End If

    Atualiza_RastreamentoLoteSaldo = SUCESSO
    
    Exit Function
    
Erro_Atualiza_RastreamentoLoteSaldo:

    Atualiza_RastreamentoLoteSaldo = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151642)
        
    End Select
        
    Exit Function

End Function

Private Function RastreamentoLoteSaldo_Atualiza(alComando() As Long, ByVal colRastreamentoLoteSaldo As Collection, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objEstoque As ClassEstoque, ByVal objProduto As ClassProduto, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'Atualiza as tabelas de RastreamentoLote e RastreamentoLoteSaldo utilizando os valores passados como parametro
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objRastreamentoLoteSaldo As ClassRastreamentoLoteSaldo
Dim iStatus As Integer
Dim lNumIntDocLote As Long
Dim dtDataValidade As Date
Dim iIndice As Integer
Dim tRastreamentoLoteSaldo As typeRastreamentoLoteSaldo
Dim tRastreamentoLoteSaldo1 As typeRastreamentoLoteSaldo
Dim sLote As String
Dim dQuantConsig3 As Double
Dim objRastroLote As ClassRastreamentoLote

On Error GoTo Erro_RastreamentoLoteSaldo_Atualiza

    For Each objRastreamentoLoteSaldo In colRastreamentoLoteSaldo
    
        sLote = String(STRING_LOTE_RASTREAMENTO, 0)
        
        'Le a tabela de rastreamentolote para o produto/lote/filial em questão
        lErro = Comando_ExecutarPos(alComando(61), "SELECT Status, NumIntDoc, Lote FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", 0, iStatus, lNumIntDocLote, sLote, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.sLote, objRastreamentoLoteSaldo.iFilialOP)
        If lErro <> AD_SQL_SUCESSO Then gError 71702
    
        'leitura do registro referente ao produto/lote/filial em questão
        lErro = Comando_BuscarPrimeiro(alComando(61))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71703
        
        If lErro = AD_SQL_SUCESSO Then
        
            'lock do produto/lote/filial
            lErro = Comando_LockExclusive(alComando(61))
            If lErro <> AD_SQL_SUCESSO Then gError 71704
    
            If iStatus <> RATREAMENTOLOTE_STATUS_ABERTO Then gError 71705
            
        Else
        
            'se não for uma produção ==> não pode gerar lote
            If objItemMovEst.iTipoMov <> MOV_EST_PRODUCAO And objItemMovEst.iTipoMov <> MOV_EST_PRODUCAO_BENEF3 And objItemMovEst.iTipoMov <> MOV_EST_PRODUCAO_OUTROS Then gError 71706
        
            dtDataValidade = objItemMovEst.dtData + objProduto.iPrazoValidade
        
            lNumIntDocLote = objEstoque.lNumIntDocRastroLote
            objEstoque.lNumIntDocRastroLote = objEstoque.lNumIntDocRastroLote + 1
            objRastreamentoLoteSaldo.lNumIntDocLote = lNumIntDocLote
        
            lErro = Comando_Executar(alComando(62), "INSERT INTO RastreamentoLote (NumIntDoc, Produto, Lote, FilialOP, Status, DataValidade, DataEntrada, DataFabricacao, Localizacao) VALUES (?,?,?,?,?,?,?,?,?)", lNumIntDocLote, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.sLote, objRastreamentoLoteSaldo.iFilialOP, RATREAMENTOLOTE_STATUS_ABERTO, dtDataValidade, objItemMovEst.dtData, objItemMovEst.dtData, objRastreamentoLoteSaldo.sLocalizacao)
            If lErro <> AD_SQL_SUCESSO Then gError 71707
    
            sLote = objRastreamentoLoteSaldo.sLote
            
            Set objRastroLote = New ClassRastreamentoLote
            
            objRastroLote.sProduto = objRastreamentoLoteSaldo.sProduto
            objRastroLote.sCodigo = objRastreamentoLoteSaldo.sLote
            objRastroLote.iFilialOP = objRastreamentoLoteSaldo.iFilialOP
            Set objRastroLote.objInfoUsu = objRastreamentoLoteSaldo.objInfoUsu
            
            lErro = CF("RastreamentoLote_Grava_Customizado", objRastroLote)
            If lErro <> SUCESSO Then gError 196436
        
        End If
        
        objRastreamentoLoteSaldo.sLote = sLote
        objRastreamentoLoteSaldo.lNumIntDocLote = lNumIntDocLote
        
        If Not objItemMovEst.colRastreamentoMovto Is Nothing Then
            For Each objRastreamentoMovto In objItemMovEst.colRastreamentoMovto
                If objRastreamentoMovto.iFilialOP = objRastreamentoLoteSaldo.iFilialOP And objRastreamentoMovto.sProduto = objRastreamentoLoteSaldo.sProduto And UCase(objRastreamentoMovto.sLote) = UCase(objRastreamentoLoteSaldo.sLote) Then
                    objRastreamentoMovto.lNumIntDocLote = lNumIntDocLote
                    Exit For
                End If
            Next
        End If
        objRastreamentoLoteSaldo.iAlmoxarifado = objItemMovEst.iAlmoxarifado
        
        'Le a tabela de rastreamentolotesaldo para o produto/almoxarifado/lote em questão
        lErro = Comando_ExecutarPos(alComando(63), "SELECT QuantDispNossa, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantBenef, QuantBenef3, QuantOP FROM RastreamentoLoteSaldo WHERE Produto = ? AND Almoxarifado = ? AND NumIntDocLote=?", 0, tRastreamentoLoteSaldo.dQuantDispNossa, tRastreamentoLoteSaldo.dQuantRecIndl, tRastreamentoLoteSaldo.dQuantIndOutras, tRastreamentoLoteSaldo.dQuantDefeituosa, tRastreamentoLoteSaldo.dQuantConsig3, tRastreamentoLoteSaldo.dQuantConsig, tRastreamentoLoteSaldo.dQuantDemo3, tRastreamentoLoteSaldo.dQuantDemo, tRastreamentoLoteSaldo.dQuantConserto3, tRastreamentoLoteSaldo.dQuantConserto, tRastreamentoLoteSaldo.dQuantOutras3, tRastreamentoLoteSaldo.dQuantOutras, tRastreamentoLoteSaldo.dQuantBenef, tRastreamentoLoteSaldo.dQuantBenef3, tRastreamentoLoteSaldo.dQuantOP, _
                                        objRastreamentoLoteSaldo.sProduto, objItemMovEst.iAlmoxarifado, lNumIntDocLote)
        If lErro <> AD_SQL_SUCESSO Then gError 71708
    
        'leitura do registro referente ao estoque/almoxarifado em questão
        lErro = Comando_BuscarPrimeiro(alComando(63))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71709
        
        If lErro = AD_SQL_SUCESSO Then
        
            'lock do estoque/almoxarifado
            lErro = Comando_LockExclusive(alComando(63))
            If lErro <> AD_SQL_SUCESSO Then gError 71710
            
            If gobjMAT.iAceitaEstoqueNegativo = DESMARCADO Then
            
                If tRastreamentoLoteSaldo.dQuantConserto + objRastreamentoLoteSaldo.dQuantConserto < -QTDE_ESTOQUE_DELTA2 Then gError 71711
                
                If tRastreamentoLoteSaldo.dQuantConserto3 + objRastreamentoLoteSaldo.dQuantConserto3 < -QTDE_ESTOQUE_DELTA2 Then gError 71712
                
                If tRastreamentoLoteSaldo.dQuantConsig + objRastreamentoLoteSaldo.dQuantConsig < -QTDE_ESTOQUE_DELTA2 Then gError 71713
                
                If tRastreamentoLoteSaldo.dQuantConsig3 + objRastreamentoLoteSaldo.dQuantConsig3 < -QTDE_ESTOQUE_DELTA2 Then gError 71714
                
                If tRastreamentoLoteSaldo.dQuantDefeituosa + objRastreamentoLoteSaldo.dQuantDefeituosa < -QTDE_ESTOQUE_DELTA2 Then gError 71715
                
                If tRastreamentoLoteSaldo.dQuantDemo + objRastreamentoLoteSaldo.dQuantDemo < -QTDE_ESTOQUE_DELTA2 Then gError 71716
                
                If tRastreamentoLoteSaldo.dQuantDemo3 + objRastreamentoLoteSaldo.dQuantDemo3 < -QTDE_ESTOQUE_DELTA2 Then gError 71717
                
                If tRastreamentoLoteSaldo.dQuantDispNossa + objRastreamentoLoteSaldo.dQuantDispNossa < -QTDE_ESTOQUE_DELTA2 Then gError 83339
                        
                If tRastreamentoLoteSaldo.dQuantIndOutras + objRastreamentoLoteSaldo.dQuantIndOutras < -QTDE_ESTOQUE_DELTA2 Then gError 71719
                
                If tRastreamentoLoteSaldo.dQuantOutras + objRastreamentoLoteSaldo.dQuantOutras < -QTDE_ESTOQUE_DELTA2 Then gError 71720
                
                If tRastreamentoLoteSaldo.dQuantOutras3 + objRastreamentoLoteSaldo.dQuantOutras3 < -QTDE_ESTOQUE_DELTA2 Then gError 71721
                
                If tRastreamentoLoteSaldo.dQuantBenef + objRastreamentoLoteSaldo.dQuantBenef < -QTDE_ESTOQUE_DELTA2 Then gError 71722
                
                If tRastreamentoLoteSaldo.dQuantBenef3 + objRastreamentoLoteSaldo.dQuantBenef3 < -QTDE_ESTOQUE_DELTA2 Then gError 71723
                
            End If
            
            lErro = RastreamentoLoteSaldo_Arredonda(objRastreamentoLoteSaldo, tRastreamentoLoteSaldo, tRastreamentoLoteSaldo1)
            If lErro <> AD_SQL_SUCESSO Then gError 83785
                
            'Atualiza a tabela de rastreamentosaldolote para o produto/almoxarifado/NumIntDocLote em questão
            lErro = Comando_ExecutarPos(alComando(64), "UPDATE RastreamentoLoteSaldo SET QuantDispNossa = ?, QuantRecIndl = ?, QuantIndOutras = ?, QuantDefeituosa = ?, QuantConsig3 = ?, QuantConsig = ?, QuantDemo3 = ?, QuantDemo = ?, QuantConserto3 = ?, QuantConserto = ?, QuantOutras3 = ?, QuantOutras = ?, QuantReservada = ?, QuantReservadaConsig = ?, QuantBenef = ?, QuantBenef3 = ?, QuantOP = ?", alComando(63), _
            tRastreamentoLoteSaldo1.dQuantDispNossa, tRastreamentoLoteSaldo1.dQuantRecIndl, tRastreamentoLoteSaldo1.dQuantIndOutras, tRastreamentoLoteSaldo1.dQuantDefeituosa, tRastreamentoLoteSaldo1.dQuantConsig3, tRastreamentoLoteSaldo1.dQuantConsig, tRastreamentoLoteSaldo1.dQuantDemo3, tRastreamentoLoteSaldo1.dQuantDemo, tRastreamentoLoteSaldo1.dQuantConserto3, tRastreamentoLoteSaldo1.dQuantConserto, tRastreamentoLoteSaldo1.dQuantOutras3, tRastreamentoLoteSaldo1.dQuantOutras, tRastreamentoLoteSaldo1.dQuantReservada, tRastreamentoLoteSaldo1.dQuantReservadaConsig, tRastreamentoLoteSaldo1.dQuantBenef, tRastreamentoLoteSaldo1.dQuantBenef3, tRastreamentoLoteSaldo1.dQuantOP)
            If lErro <> AD_SQL_SUCESSO Then gError 71726
        
        Else
        
            lErro = Comando_Executar(alComando(65), "INSERT INTO RastreamentoLoteSaldo (Produto, Almoxarifado, NumIntDocLote, QuantDispNossa, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantReservada, QuantReservadaConsig, QuantBenef, QuantBenef3, QuantOP) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, objRastreamentoLoteSaldo.lNumIntDocLote, objRastreamentoLoteSaldo.dQuantDispNossa, objRastreamentoLoteSaldo.dQuantRecIndl, objRastreamentoLoteSaldo.dQuantIndOutras, objRastreamentoLoteSaldo.dQuantDefeituosa, objRastreamentoLoteSaldo.dQuantConsig3, objRastreamentoLoteSaldo.dQuantConsig, objRastreamentoLoteSaldo.dQuantDemo3, objRastreamentoLoteSaldo.dQuantDemo, objRastreamentoLoteSaldo.dQuantConserto3, objRastreamentoLoteSaldo.dQuantConserto, objRastreamentoLoteSaldo.dQuantOutras3, objRastreamentoLoteSaldo.dQuantOutras, objRastreamentoLoteSaldo.dQuantReservada, objRastreamentoLoteSaldo.dQuantReservadaConsig, objRastreamentoLoteSaldo.dQuantBenef, objRastreamentoLoteSaldo.dQuantBenef3, objRastreamentoLoteSaldo.dQuantOP)
            If lErro <> AD_SQL_SUCESSO Then gError 71727

        End If


    Next
    
    '####################################################################
    'Inserido por Wagner 14/03/2006
    'Se for rastreamento por série tem que pegar o lote inicial
    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE And Not (objItemMovEst.colRastreamentoMovto Is Nothing) Then
        For Each objRastreamentoLoteSaldo In colRastreamentoLoteSaldo
            For Each objRastreamentoMovto In objItemMovEst.colRastreamentoMovto
                objRastreamentoMovto.lNumIntDocLoteSerieIni = objRastreamentoLoteSaldo.lNumIntDocLote
            Next
            Exit For
        Next
    End If
    '####################################################################

    RastreamentoLoteSaldo_Atualiza = SUCESSO

    Exit Function

Erro_RastreamentoLoteSaldo_Atualiza:

    RastreamentoLoteSaldo_Atualiza = gErr

    Select Case gErr

        Case 71702, 71703
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE2", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.sLote, objRastreamentoLoteSaldo.iFilialOP)

        Case 71704
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RASTREAMENTOLOTE", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.sLote, objRastreamentoLoteSaldo.iFilialOP)

        Case 71705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTOLOTE_NAO_ABERTO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.sLote, objRastreamentoLoteSaldo.iFilialOP)

        Case 71706
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_MOV_NAO_PRODUCAO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.sLote, objRastreamentoLoteSaldo.iFilialOP)

        Case 71707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RASTREAMENTOLOTE", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.sLote, objRastreamentoLoteSaldo.iFilialOP)

        Case 71708, 71709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTESALDO", gErr, objRastreamentoLoteSaldo.sProduto, objItemMovEst.iAlmoxarifado, sLote)

        Case 71710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RASTREAMENTOLOTESALDO", gErr, objRastreamentoLoteSaldo.sProduto, objItemMovEst.iAlmoxarifado, sLote)

        Case 71711
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSERTO_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantConserto)
            
        Case 71712
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSERTO3_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantConserto3)
            
        Case 71713
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSIG_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantConsig)
            
        Case 71714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSIG3_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantConsig3)
            
        Case 71715
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEFEITUOSO_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantDefeituosa)
            
        Case 71716
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEMO_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantDemo)
            
        Case 71717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEMO3_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantDemo3)
            
        Case 71718, 83339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DISPONIVEL_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantDispNossa)
            
        Case 71719
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_INDISPONIVEL_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantIndOutras)
            
        Case 71720
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_OUTRAS_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantOutras)
            
        Case 71721
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_OUTRAS3_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantOutras3)
            
        Case 71722
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_BENEF_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantBenef)
            
        Case 71723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_BENEF3_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantBenef3)
        
        Case 71724
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_OP_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, tRastreamentoLoteSaldo.dQuantOP)
        
        Case 71725, 196436
        
        Case 71726
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RASTREAMENTOLOTESALDO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote)
        
        Case 71727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RASTREAMENTOLOTESALDO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote)
        
        Case 71731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DISP_CONSIG3_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantDispNossa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151643)

    End Select

    Exit Function

End Function

Private Function RastreamentoLoteSaldo_Arredonda(ByVal objRastreamentoLoteSaldo As ClassRastreamentoLoteSaldo, tRastreamentoLoteSaldo As typeRastreamentoLoteSaldo, tRastreamentoLoteSaldo1 As typeRastreamentoLoteSaldo) As Long
'auxiliar a RastreamentoLoteSaldo_Atualiza(), para eliminar residuo por arredondamento

Dim lErro As Long

On Error GoTo Erro_RastreamentoLoteSaldo_Arredonda

    tRastreamentoLoteSaldo1.dQuantConserto = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantConserto + objRastreamentoLoteSaldo.dQuantConserto)
    tRastreamentoLoteSaldo1.dQuantConserto3 = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantConserto3 + objRastreamentoLoteSaldo.dQuantConserto3)
    tRastreamentoLoteSaldo1.dQuantConsig = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantConsig + objRastreamentoLoteSaldo.dQuantConsig)
    tRastreamentoLoteSaldo1.dQuantConsig3 = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantConsig3 + objRastreamentoLoteSaldo.dQuantConsig3)
    tRastreamentoLoteSaldo1.dQuantDefeituosa = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantDefeituosa + objRastreamentoLoteSaldo.dQuantDefeituosa)
    tRastreamentoLoteSaldo1.dQuantDemo = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantDemo + objRastreamentoLoteSaldo.dQuantDemo)
    tRastreamentoLoteSaldo1.dQuantDemo3 = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantDemo3 + objRastreamentoLoteSaldo.dQuantDemo3)
    tRastreamentoLoteSaldo1.dQuantDispNossa = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantDispNossa + objRastreamentoLoteSaldo.dQuantDispNossa)
    tRastreamentoLoteSaldo1.dQuantIndOutras = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantIndOutras + objRastreamentoLoteSaldo.dQuantIndOutras)
    tRastreamentoLoteSaldo1.dQuantOutras = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantOutras + objRastreamentoLoteSaldo.dQuantOutras)
    tRastreamentoLoteSaldo1.dQuantOutras3 = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantOutras3 + objRastreamentoLoteSaldo.dQuantOutras3)
    tRastreamentoLoteSaldo1.dQuantBenef = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantBenef + objRastreamentoLoteSaldo.dQuantBenef)
    tRastreamentoLoteSaldo1.dQuantBenef3 = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantBenef3 + objRastreamentoLoteSaldo.dQuantBenef3)
    tRastreamentoLoteSaldo1.dQuantOP = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantOP + objRastreamentoLoteSaldo.dQuantOP)
    tRastreamentoLoteSaldo1.dQuantReservada = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantReservada + objRastreamentoLoteSaldo.dQuantReservada)
    tRastreamentoLoteSaldo1.dQuantReservadaConsig = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantReservadaConsig + objRastreamentoLoteSaldo.dQuantReservadaConsig)
    tRastreamentoLoteSaldo1.dQuantRecIndl = Arredonda_Estoque(tRastreamentoLoteSaldo.dQuantRecIndl + objRastreamentoLoteSaldo.dQuantRecIndl)

    RastreamentoLoteSaldo_Arredonda = SUCESSO
    
    Exit Function
    
Erro_RastreamentoLoteSaldo_Arredonda:

    RastreamentoLoteSaldo_Arredonda = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151644)

    End Select

    Exit Function

End Function

Private Function RastreamentoMovto_Inserir(ByVal lComando1 As Long, ByVal lNumIntDocOrigem As Long, ByVal objEstoque As ClassEstoque, ByVal colRastreamentoMovto As Collection, ByVal iTipoDocOrigem As Integer) As Long
'inserir um rastreamento movimento
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim objRastreamentoMovto As ClassRastreamentoMovto

On Error GoTo Erro_RastreamentoMovto_Inserir
    
    If Not (colRastreamentoMovto Is Nothing) Then
        
        For Each objRastreamentoMovto In colRastreamentoMovto
        
            'Não aceita rastreamento sem o lote (tem que verificar porque não achou o numintdoc)
            If objRastreamentoMovto.lNumIntDocLote = 0 Then gError 206781
            
            'Insere um registro na tabela de RastreamentoMovto
            lErro = Comando_Executar(lComando1, "INSERT INTO RastreamentoMovto (NumIntDoc, NumIntDocLote, TipoDocOrigem, NumIntDocOrigem, Produto, Quantidade, SiglaUM, NumIntDocLoteSerieIni, QuantidadeFaixaSerie) VALUES (?,?,?,?,?,?,?,?,?)", objEstoque.lNumIntDocRastroMovto, objRastreamentoMovto.lNumIntDocLote, iTipoDocOrigem, lNumIntDocOrigem, objRastreamentoMovto.sProduto, objRastreamentoMovto.dQuantidade, objRastreamentoMovto.sSiglaUM, objRastreamentoMovto.lNumIntDocLoteSerieIni, objRastreamentoMovto.dQuantidadeSerie)
            If lErro <> AD_SQL_SUCESSO Then gError 71734
    
            objEstoque.lNumIntDocRastroMovto = objEstoque.lNumIntDocRastroMovto + 1
    
        Next
    
    End If
    
    RastreamentoMovto_Inserir = SUCESSO

    Exit Function

Erro_RastreamentoMovto_Inserir:

    RastreamentoMovto_Inserir = gErr

    Select Case gErr

        Case 71734, 206781
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RASTREAMENTOMOVTO", gErr, objRastreamentoMovto.sProduto, objRastreamentoMovto.lNumIntDocLote)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151645)

    End Select

    Exit Function

End Function

Private Function ApropriacaoInsumosProd_Inserir(ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objEstoque As ClassEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'inserir na tabela ApropriacaoInsumosprod e RastreamentoMovto a movimentação dos insumos

Dim lErro As Long
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim objApropriacaoInsumosProd As ClassApropriacaoInsumosProd
Dim sProduto As String
Dim sLote As String
Dim iStatus As Integer
Dim lNumIntDocLote As Long

On Error GoTo Erro_ApropriacaoInsumosProd_Inserir

    If Not (objItemMovEst.colApropriacaoInsumo Is Nothing) Then

        If objTipoMovEstoque.iCodigo = MOV_EST_PRODUCAO Or objTipoMovEstoque.iCodigo = MOV_EST_PRODUCAO_BENEF3 Then
    
            For Each objApropriacaoInsumosProd In objItemMovEst.colApropriacaoInsumo
            
                'Insere um registro na tabela de RastreamentoMovto
                lErro = Comando_Executar(lComando1, "INSERT INTO ApropriacaoInsumosProd (NumIntDoc, NumIntDocOrigem, NumIntReqProd, Produto, Quantidade, Automatico) VALUES (?,?,?,?,?,?)", objEstoque.lNumIntDocApropInsProd, objItemMovEst.lNumIntDoc, objApropriacaoInsumosProd.lNumIntReqProd, objApropriacaoInsumosProd.sProduto, objApropriacaoInsumosProd.dQuantidade, objApropriacaoInsumosProd.iAutomatico)
                If lErro <> AD_SQL_SUCESSO Then gError 71742
                    
                For Each objRastreamentoMovto In objApropriacaoInsumosProd.colRastreamentoMovto
                    
                    sProduto = String(STRING_PRODUTO, 0)
                    sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    
                    'Le a tabela de rastreamentolote para o produto/lote/filial em questão
                    lErro = Comando_ExecutarPos(lComando2, "SELECT Status, NumIntDoc, Lote FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", 0, iStatus, lNumIntDocLote, sLote, objRastreamentoMovto.sProduto, objRastreamentoMovto.sLote, objRastreamentoMovto.iFilialOP)
                    If lErro <> AD_SQL_SUCESSO Then gError 71751
                    
                    'leitura do registro de rastreamentolote
                    lErro = Comando_BuscarPrimeiro(lComando2)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71752
                            
                    'se o lote não está cadastrado ==> erro
                    If lErro = AD_SQL_SEM_DADOS Then gError 71753
                    
                    objRastreamentoMovto.lNumIntDocLote = lNumIntDocLote
                            
                Next
                    
                'inserir na tabela RastreamentoMovto os rastros dos insumos
                lErro = RastreamentoMovto_Inserir(lComando, objEstoque.lNumIntDocApropInsProd, objEstoque, objApropriacaoInsumosProd.colRastreamentoMovto, TIPO_RASTREAMENTO_MOVTO_APROP_PRODUCAO_ENTRADA)
                If lErro <> SUCESSO Then gError 71743
            
                objEstoque.lNumIntDocApropInsProd = objEstoque.lNumIntDocApropInsProd + 1
                
            Next
    
        End If

    End If
    
    ApropriacaoInsumosProd_Inserir = SUCESSO

    Exit Function

Erro_ApropriacaoInsumosProd_Inserir:

    ApropriacaoInsumosProd_Inserir = gErr

    Select Case gErr

        Case 71742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_APROPRIACAOINSUMOSPROD", gErr, objApropriacaoInsumosProd.sProduto)

        Case 71743
        
        Case 71751, 71752
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE2", gErr, objRastreamentoMovto.sProduto, objRastreamentoMovto.sLote, objRastreamentoMovto.iFilialOP)
        
        Case 71753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_CADASTRADO_RASTREAMENTO", gErr, objRastreamentoMovto.sLote, objRastreamentoMovto.sProduto, objRastreamentoMovto.iFilialOP)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151646)

    End Select

    Exit Function

End Function

Private Function RastreamentoMovto_Excluir(ByVal lComando As Long, ByVal lComando1 As Long, ByVal lNumIntDocOrigem As Long, ByVal iTipoDocOrigem As Integer) As Long
'excluir os rastreamentos relativos a um movimento de estoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim sProduto As String
Dim lNumIntDocLote As Long

On Error GoTo Erro_RastreamentoMovto_Excluir

    sProduto = String(STRING_PRODUTO, 0)

    'Selecionar registros na tabela de RastreamentoMovto associados ao movimento que está sendo excluido
    lErro = Comando_ExecutarPos(lComando, "SELECT Produto, NumIntDocLote FROM RastreamentoMovto WHERE TipoDocOrigem=? AND NumIntDocOrigem=?", 0, sProduto, lNumIntDocLote, iTipoDocOrigem, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 71746

    'leitura do registro referente ao RastreamentoMovto
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71747

    Do While lErro = AD_SQL_SUCESSO

        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM RastreamentoMovto", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 71748

        'leitura do proximo registro referente ao RastreamentoMovto
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71749

    Loop

    RastreamentoMovto_Excluir = SUCESSO

    Exit Function

Erro_RastreamentoMovto_Excluir:

    RastreamentoMovto_Excluir = gErr

    Select Case gErr

        Case 71746, 71747, 71749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOMOVTO", gErr, lNumIntDocOrigem, iTipoDocOrigem)

        Case 71748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RASTREAMENTOMOVTO", gErr, sProduto, lNumIntDocLote)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151647)

    End Select

    Exit Function

End Function

Private Function ApropriacaoInsumosProd_Excluir(ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lNumIntDocOrigem As Long, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'excluir da tabela ApropriacaoInsumosprod e RastreamentoMovto a movimentação dos insumos

Dim lNumIntDoc As Long
Dim lErro As Long
Dim sSQL As String

On Error GoTo Erro_ApropriacaoInsumosProd_Excluir

    If objTipoMovEstoque.iCodigo = MOV_EST_PRODUCAO Or objTipoMovEstoque.iCodigo = MOV_EST_PRODUCAO_BENEF3 Then
        sSQL = "SELECT NumIntDoc FROM ApropriacaoInsumosProd WHERE NumIntDocOrigem=?"
    ElseIf objTipoMovEstoque.iCodigo = MOV_EST_REQ_PRODUCAO Or objTipoMovEstoque.iCodigo = MOV_EST_REQ_PRODUCAO_BENEF3 Then
        sSQL = "SELECT NumIntDoc FROM ApropriacaoInsumosProd WHERE NumIntReqProd=?"
    End If

    If objTipoMovEstoque.iCodigo = MOV_EST_PRODUCAO Or objTipoMovEstoque.iCodigo = MOV_EST_PRODUCAO_BENEF3 Or _
       objTipoMovEstoque.iCodigo = MOV_EST_REQ_PRODUCAO Or objTipoMovEstoque.iCodigo = MOV_EST_REQ_PRODUCAO_BENEF3 Then
        
        'Selecionar registros na tabela de ApropriacaoInsumosProd gerado a partir do lNumIntDocOrigem
        lErro = Comando_ExecutarPos(lComando2, sSQL, 0, lNumIntDoc, lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 71754
    
        'leitura do registro referente ao RastreamentoMovto
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71755
        
        Do While lErro = AD_SQL_SUCESSO
        
            'excluir da tabela RastreamentoMovto a movimentação
            lErro = RastreamentoMovto_Excluir(lComando, lComando1, lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_APROP_PRODUCAO_ENTRADA)
            If lErro <> SUCESSO Then gError 71756
        
            lErro = Comando_ExecutarPos(lComando3, "DELETE FROM ApropriacaoInsumosProd", lComando2)
            If lErro <> AD_SQL_SUCESSO Then gError 71757
               
            'leitura do proximo registro referente a ApropriacãoInsumosProd
            lErro = Comando_BuscarProximo(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71758
    
        Loop

    End If

    ApropriacaoInsumosProd_Excluir = SUCESSO

    Exit Function

Erro_ApropriacaoInsumosProd_Excluir:

    ApropriacaoInsumosProd_Excluir = gErr

    Select Case gErr

        Case 71754, 71755, 71758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APROPRIACAOINSUMOSPROD", gErr, lNumIntDocOrigem)

        Case 71756
        
        Case 71757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_APROPRIACAOINSUMOSPROD", gErr, lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151648)

    End Select

    Exit Function

End Function

Private Function Estoque_Processa_Alteracao(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst, ByVal objProduto As ClassProduto, ByVal objEstoque As ClassEstoque) As Long
'processa a alteração dos dados possiveis de movimento de estoque (rastreamento, alguns dados do movimento como horasmaquina e datainicioproducao, e data e hora do movimento de estoque)

Dim lErro As Long
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim colRastreamentoLoteSaldo As New Collection
Dim colRastreamentoMovto As New Collection
Dim dFator As Double
Dim objItemMovEst1 As New ClassItemMovEstoque

On Error GoTo Erro_Estoque_Processa_Alteracao

    objItemMovEst1.lNumIntDoc = objItemMovEst.lNumIntDoc

    'Verifica se o movimento a ser alterado está cadastrado
    lErro = CF("MovimentoEstoque_Le_NumIntDoc1", alComando(83), objItemMovEst1)
    If lErro <> SUCESSO Then gError 71762
        
    '####################################
    'Alterado por Wagner
    'atualiza a hora, hora de maquina e data inicio de producao do movimento
    lErro = Comando_ExecutarPos(alComando(73), "UPDATE MovimentoEstoque SET HorasMaquina=?, DataInicioProducao=?, Hora=?, Observacao=?, Requisitante = ?", alComando(83), objItemMovEst.lHorasMaquina, objItemMovEst.dtDataInicioProducao, CDbl(objItemMovEst.dtHora), objItemMovEst.sObservacao, objItemMovEst.lRequisitante)
    If lErro <> AD_SQL_SUCESSO Then gError 71763
    '####################################
    
    'se o produto trabalha com rastreamento ==> altera os rastreamentos
    If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM And objItemMovEst.colItensRomaneioGrade.Count = 0 Then
    
        'carrega os rastreamentos do movimento em colRastreamentoMovto e simultaneamente exclui seus dados do bd
        lErro = RastreamentoMovto_Excluir1(alComando(75), alComando(76), TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, objItemMovEst.lNumIntDoc, colRastreamentoMovto)
        If lErro <> SUCESSO Then gError 71768
    
        'altera o sinal das quantidades para que possam subtrair dos totalizadores de lote
        For Each objRastreamentoMovto In colRastreamentoMovto
    
            ' devolve o fator de conversao de sSiglaUM1(Origem)  em sSiglaUM2(destino)
            lErro = UM_Transforma(alComando(), objItemMovEst.iClasseUM, objRastreamentoMovto.sSiglaUM, objItemMovEst.sSiglaUMEst, dFator)
            If lErro <> SUCESSO Then gError 83224
            
            'objRastreamentoMovto.dQuantidadeEst = -objRastreamentoMovto.dQuantidade * dFator
            'UM_Transforma, ao contrário de UM_Conversao, divide dQtde2 pela dQtde1 ao contrário da função mais usada (UM_Conversao) que divide dQtde1 pela dQtde2
            'Por isso o sinal estava errado
            objRastreamentoMovto.dQuantidadeEst = -objRastreamentoMovto.dQuantidade / dFator
            
        Next
    
        'Retira os saldos dos rastreamentos antigos que foram excluidos/alterados
        Set colRastreamentoLoteSaldo = New Collection
    
        'Retira os saldos dos rastreamentos antigos que foram excluidos/alterados
        lErro = Atualiza_RastreamentoLoteSaldo(objItemMovEst, objTipoMovEstoque, colRastreamentoLoteSaldo, colRastreamentoMovto)
        If lErro <> SUCESSO Then gError 71764
    
'        'atualiza a tabela de RastreamentoLoteSaldo
'        lErro = RastreamentoLoteSaldo_Atualiza1(alComando(), colRastreamentoLoteSaldo, colRastreamentoMovto, objItemMovEst)
'        If lErro <> SUCESSO Then gError 71801
'
'        'Atualiza os saldos dos rastreamentos com os dados novos/alterados
'        Set colRastreamentoLoteSaldo = New Collection
    
        'Atualiza os saldos dos rastreamentos com os dados novos/alterados
        lErro = Atualiza_RastreamentoLoteSaldo(objItemMovEst, objTipoMovEstoque, colRastreamentoLoteSaldo, objItemMovEst.colRastreamentoMovto)
        If lErro <> SUCESSO Then gError 71769

        'atualiza a tabela de RastreamentoLoteSaldo
        lErro = RastreamentoLoteSaldo_Atualiza(alComando(), colRastreamentoLoteSaldo, objItemMovEst, objEstoque, objProduto, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71770
        
        'inserir na tabela RastreamentoMovto a movimentação
        lErro = RastreamentoMovto_Inserir(alComando(66), objItemMovEst.lNumIntDoc, objEstoque, objItemMovEst.colRastreamentoMovto, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE)
        If lErro <> SUCESSO Then gError 71771
        
        'excluir da tabela ApropriacaoInsumosprod e RastreamentoMovto a movimentação dos insumos
        lErro = ApropriacaoInsumosProd_Excluir(alComando(69), alComando(70), alComando(71), alComando(72), objItemMovEst.lNumIntDoc, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71772
        
        'inserir na tabela ApropriacaoInsumosprod e RastreamentoMovto a movimentação dos insumos
        lErro = ApropriacaoInsumosProd_Inserir(alComando(66), alComando(67), alComando(61), objItemMovEst, objEstoque, objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 71773
            
    End If

    giAlteracao = 1

    Estoque_Processa_Alteracao = SUCESSO

    Exit Function
    
Erro_Estoque_Processa_Alteracao:

    Estoque_Processa_Alteracao = gErr
    
    Select Case gErr
    
        Case 71762, 71764, 71768, 71769, 71770, 71771, 71772, 71773, 71801, 83224
    
        Case 71763, 89670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MOVIMENTOESTOQUE", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151649)
    
    End Select
    
    Exit Function

End Function

Private Function RastreamentoMovto_Excluir1(ByVal lComando As Long, ByVal lComando1 As Long, ByVal iTipoDocOrigem As Integer, ByVal lNumIntDocOrigem As Long, ByVal colRastreamentoMovto As Collection) As Long

Dim lErro As Long
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim sProduto As String
Dim lNumIntDocLote As Long
Dim lNumIntDoc As Long
Dim dQuantidade As Double
Dim sLote As String
Dim sSiglaUM As String
Dim lComando2 As Long
Dim iFilialOP As Integer

On Error GoTo Erro_RastreamentoMovto_Excluir1

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 92001

    sProduto = String(STRING_PRODUTO, 0)
    sSiglaUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)

    'Selecionar registros na tabela de RastreamentoMovto associados ao movimento que está sendo excluido
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntDocLote, Produto, Quantidade, SiglaUM FROM RastreamentoMovto WHERE TipoDocOrigem=? AND NumIntDocOrigem=?", 0, lNumIntDoc, lNumIntDocLote, sProduto, dQuantidade, sSiglaUM, iTipoDocOrigem, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 71766

    'leitura do registro referente ao RastreamentoMovto
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71767

    Do While lErro = AD_SQL_SUCESSO
        
        Set objRastreamentoMovto = New ClassRastreamentoMovto
        
        objRastreamentoMovto.lNumIntDoc = lNumIntDoc
        objRastreamentoMovto.lNumIntDocLote = lNumIntDocLote
        objRastreamentoMovto.sProduto = sProduto
        objRastreamentoMovto.dQuantidade = dQuantidade
        objRastreamentoMovto.iTipoDocOrigem = iTipoDocOrigem
        objRastreamentoMovto.lNumIntDocOrigem = lNumIntDocOrigem
        objRastreamentoMovto.sSiglaUM = sSiglaUM
        
        colRastreamentoMovto.Add objRastreamentoMovto
        
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM RastreamentoMovto", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 71768

        sLote = String(STRING_LOTE_RASTREAMENTO, 0)

        'Selecionar o registro na tabela de RastreamentoLote associados ao lote está sendo excluido
        lErro = Comando_Executar(lComando2, "SELECT Lote, FilialOP FROM RastreamentoLote WHERE NumIntDoc=?", sLote, iFilialOP, lNumIntDocLote)
        If lErro <> AD_SQL_SUCESSO Then gError 92002
    
        'leitura do registro referente ao RastreamentoMovto
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92003

        'se não encontrar o lote no BD == > erro
        If lErro = AD_SQL_SEM_DADOS Then gError 92004

        objRastreamentoMovto.sLote = sLote
        objRastreamentoMovto.iFilialOP = iFilialOP

        'leitura do proximo registro referente ao RastreamentoMovto
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71769

    Loop

    Call Comando_Fechar(lComando2)

    RastreamentoMovto_Excluir1 = SUCESSO

    Exit Function

Erro_RastreamentoMovto_Excluir1:

    RastreamentoMovto_Excluir1 = gErr

    Select Case gErr

        Case 71766, 71767, 71769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOMOVTO", gErr, lNumIntDocOrigem, iTipoDocOrigem)

        Case 71768
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RASTREAMENTOMOVTO", gErr, sProduto, lNumIntDocLote)

        Case 92001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92002, 92003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)

        Case 92004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO2", gErr, lNumIntDocLote)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151650)

    End Select

    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Public Function Rastreamento_FIFO_Gera(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque, ByVal iEscaninho As Integer, ByVal objProduto As ClassProduto) As Long
    Rastreamento_FIFO_Gera = CF("Rastreamento_FIFO_Gera2", objItemMovEst, iEscaninho, objProduto)
End Function

Public Function Rastreamento_FIFO_Gera2(ByVal objItemMovEst As ClassItemMovEstoque, ByVal iEscaninho As Integer, ByVal objProduto As ClassProduto) As Long
'Gera na ordem de Data de Entrada descrecente os Movimentos dos Lotes

Dim lErro As Long
Dim tRastreamentoLoteSaldo As typeRastreamentoLoteSaldo
Dim objRastreamentoMovto As New ClassRastreamentoMovto
Dim dQuantidadeEscaninho As Double
Dim dQuantidadeTotal As Double
Dim objMATConfig As New ClassMATConfig
Dim lComando As Long

'##################################
'Inserido por Wagner 09/03/2006
Dim lNumIntDocLote As Long
Dim sSerieIni As String
'##################################

On Error GoTo Erro_Rastreamento_FIFO_Gera2

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 199100

    objMATConfig.iFilialEmpresa = EMPRESA_TODA
    objMATConfig.sCodigo = HABILITA_FIFO_NFISCAIS

    lErro = CF("MATConfig_Le", objMATConfig)
    If lErro <> SUCESSO And lErro <> 89118 Then gError 87450

    If lErro = 89118 Then gError 89170
    
    lErro = FIFO_Testa(objItemMovEst.iFilialEmpresa, objMATConfig)
    If lErro <> SUCESSO Then gError 87450

    If objMATConfig.sConteudo = CStr(HABILITA_FIFO_NF) Then
    
        'se é um produto com rastreamento
        If objProduto.iRastro <> PRODUTO_RASTRO_NENHUM Then
    
            'se a colecao de rastreamentos não tiver sido inicializada ==> inicializa
            If objItemMovEst.colRastreamentoMovto Is Nothing Then Set objItemMovEst.colRastreamentoMovto = New Collection
    
            'se a colecao de rastreamentos estiver zerada ==> tem que gerar os rastreamentos.
            If objItemMovEst.colRastreamentoMovto.Count = 0 Then
    
                tRastreamentoLoteSaldo.sProduto = String(STRING_PRODUTO, 0)
                tRastreamentoLoteSaldo.sLote = String(STRING_LOTE_RASTREAMENTO, 0) 'Alterado por Wagner 09/03/2006
                
                'Lê o RastreamentoLoteSaldo, RastreamentoLote em ordem decrescente de Data
                lErro = Comando_Executar(lComando, "SELECT NumIntDocLote, QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantEmpenhada, QuantPedida, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantOP, QuantBenef, QuantBenef3, Lote, FilialOP FROM RastreamentoLoteSaldo, RastreamentoLote WHERE RastreamentoLote.NumIntDoc = RastreamentoLoteSaldo.NumIntDocLote AND RastreamentoLoteSaldo.Produto = ? AND Almoxarifado = ? AND Status = ? ORDER BY DataEntrada, RastreamentoLote.NumIntDoc", tRastreamentoLoteSaldo.lNumIntDocLote, tRastreamentoLoteSaldo.dQuantDispNossa, tRastreamentoLoteSaldo.dQuantReservada, tRastreamentoLoteSaldo.dQuantReservadaConsig, tRastreamentoLoteSaldo.dQuantEmpenhada, _
                tRastreamentoLoteSaldo.dQuantPedida, tRastreamentoLoteSaldo.dQuantRecIndl, tRastreamentoLoteSaldo.dQuantIndOutras, tRastreamentoLoteSaldo.dQuantDefeituosa, tRastreamentoLoteSaldo.dQuantConsig3, tRastreamentoLoteSaldo.dQuantConsig, tRastreamentoLoteSaldo.dQuantDemo3, tRastreamentoLoteSaldo.dQuantDemo, tRastreamentoLoteSaldo.dQuantConserto3, tRastreamentoLoteSaldo.dQuantConserto, tRastreamentoLoteSaldo.dQuantOutras3, tRastreamentoLoteSaldo.dQuantOutras, tRastreamentoLoteSaldo.dQuantOP, tRastreamentoLoteSaldo.dQuantBenef, tRastreamentoLoteSaldo.dQuantBenef3, tRastreamentoLoteSaldo.sLote, tRastreamentoLoteSaldo.iFilialOP, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, RATREAMENTOLOTE_STATUS_ABERTO)
                If lErro <> AD_SQL_SUCESSO Then gError 78492
                       
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78493
                    
                'Se ainda não alcançou a quantidade e ainda existe lote a ser acumulado
                Do While objItemMovEst.dQuantidadeEst > dQuantidadeTotal And lErro = AD_SQL_SUCESSO
                
                    Set objRastreamentoMovto = New ClassRastreamentoMovto
                    
                    'Depende do escaninho
                    Select Case iEscaninho
                        
                        'Acumula uma quantidade Total e a quantidade no Lote-escaninho atual
                        Case ESCANINHO_NOSSO
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDispNossa
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDispNossa
                        Case ESCANINHO_3_EM_CONSIGNACAO
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConsig3
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConsig3
                        Case ESCANINHO_NOSSO_EM_CONSIGNACAO
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConsig
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConsig
                        Case ESCANINHO_3_EM_DEMO
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDemo3
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDemo3
                        Case ESCANINHO_NOSSO_EM_DEMO
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDemo
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDemo
                        Case ESCANINHO_3_EM_CONSERTO
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConserto3
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConserto3
                        Case ESCANINHO_NOSSO_EM_CONSERTO
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConserto
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConserto
                        Case ESCANINHO_3_EM_OUTROS
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantOutras3
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantOutras3
                        Case ESCANINHO_NOSSO_EM_OUTROS
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantOutras
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantOutras
                        Case ESCANINHO_3_EM_BENEF
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantBenef3
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantBenef3
                        Case ESCANINHO_NOSSO_EM_BENEF
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantBenef
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantBenef
                        Case ESCANINHO_DEFEITUOSO
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDefeituosa
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDefeituosa
                        Case ESCANINHO_INDISPONIVEL
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantIndOutras
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantIndOutras
                        Case ESCANINHO_RECEBIDO_INDISPONIVEL
                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantRecIndl
                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantRecIndl
    
                    End Select
                            
                    'Se a quantidade for maior
                    If dQuantidadeTotal > objItemMovEst.dQuantidadeEst Then
                        'Pega somente o necessário
                        objRastreamentoMovto.dQuantidadeEst = objItemMovEst.dQuantidadeEst - (dQuantidadeTotal - dQuantidadeEscaninho)
                    Else
                        'Se for menor ou igual --> Pega toda a quantidade do lote
                        objRastreamentoMovto.dQuantidadeEst = dQuantidadeEscaninho
                    End If
                    
                    If objRastreamentoMovto.dQuantidadeEst <> 0 Then
                    
                        'Armazena os dados do Lote
                        objRastreamentoMovto.sSiglaUM = objItemMovEst.sSiglaUM
                        objRastreamentoMovto.dQuantidade = objRastreamentoMovto.dQuantidadeEst * (objItemMovEst.dQuantidade / objItemMovEst.dQuantidadeEst)
                        objRastreamentoMovto.iFilialOP = tRastreamentoLoteSaldo.iFilialOP
                        objRastreamentoMovto.sLote = tRastreamentoLoteSaldo.sLote
                        objRastreamentoMovto.lNumIntDocLote = tRastreamentoLoteSaldo.lNumIntDocLote
                        objRastreamentoMovto.sProduto = objItemMovEst.sProduto
                                               
                        'coloca na coleção
                        objItemMovEst.colRastreamentoMovto.Add objRastreamentoMovto
                    
                    End If
                    
                    'Busca o próximo Lote
                    lErro = Comando_BuscarProximo(lComando)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78494
                            
                Loop
                
                'Se não tem quantidade suficiente ---> Erro
                If objItemMovEst.dQuantidadeEst > dQuantidadeTotal Then gError 78495
                
                '############################################
                'Inserido por Wagner 09/03/2006
                'Se for Rastreado por Série tem que gravar a Série Inicial e a Qtd de Séries
                If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
                    For Each objRastreamentoMovto In objItemMovEst.colRastreamentoMovto
                        If Len(Trim(sSerieIni)) = 0 Or sSerieIni > objRastreamentoMovto.sLote Then
                            lNumIntDocLote = objRastreamentoMovto.lNumIntDocLote
                            sSerieIni = objRastreamentoMovto.sLote
                        End If
                    Next
                    For Each objRastreamentoMovto In objItemMovEst.colRastreamentoMovto
                        objRastreamentoMovto.lNumIntDocLoteSerieIni = lNumIntDocLote
                        objRastreamentoMovto.dQuantidadeSerie = objItemMovEst.dQuantidadeEst
                    Next
                End If
                '############################################
        
            End If
            
        End If
        
    End If
        
    'Fechamento comando
    Call Comando_Fechar(lComando)
        
    Rastreamento_FIFO_Gera2 = SUCESSO
        
    Exit Function
    
Erro_Rastreamento_FIFO_Gera2:

    Rastreamento_FIFO_Gera2 = gErr
    
    Select Case gErr
            
        Case 78492, 78493, 78494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTESALDO2", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
                
        Case 78495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_INSUF_RASTREAMENTO", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, iEscaninho)
        
        Case 87450
        
        Case 89170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MATCONFIG_NAO_CADASTRADO", gErr, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 151651)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Private Sub Escaninho_Retorna(ByVal objEstoqueProduto As ClassEstoqueProduto, iEscaninho As Integer)

    If objEstoqueProduto.dQuantBenef < 0 Then
        iEscaninho = ESCANINHO_NOSSO_EM_BENEF
    ElseIf objEstoqueProduto.dQuantBenef3 < 0 Then
        iEscaninho = ESCANINHO_3_EM_BENEF
    ElseIf objEstoqueProduto.dQuantConserto < 0 Then
        iEscaninho = ESCANINHO_NOSSO_EM_CONSERTO
    ElseIf objEstoqueProduto.dQuantConserto3 < 0 Then
        iEscaninho = ESCANINHO_3_EM_CONSERTO
    ElseIf objEstoqueProduto.dQuantConsig < 0 Then
        iEscaninho = ESCANINHO_NOSSO_EM_CONSIGNACAO
    ElseIf objEstoqueProduto.dQuantConsig3 < 0 Then
        iEscaninho = ESCANINHO_3_EM_CONSIGNACAO
    ElseIf objEstoqueProduto.dQuantDefeituosa < 0 Then
        iEscaninho = ESCANINHO_DEFEITUOSO
    ElseIf objEstoqueProduto.dQuantDemo < 0 Then
        iEscaninho = ESCANINHO_NOSSO_EM_DEMO
    ElseIf objEstoqueProduto.dQuantDemo3 < 0 Then
        iEscaninho = ESCANINHO_3_EM_DEMO
    ElseIf objEstoqueProduto.dQuantDispNossa < 0 Then
        iEscaninho = ESCANINHO_NOSSO
    ElseIf objEstoqueProduto.dQuantInd < 0 Then
        iEscaninho = ESCANINHO_INDISPONIVEL
    ElseIf objEstoqueProduto.dQuantRecIndl < 0 Then
        iEscaninho = ESCANINHO_RECEBIDO_INDISPONIVEL
    ElseIf objEstoqueProduto.dQuantOutras < 0 Then
        iEscaninho = ESCANINHO_NOSSO_EM_OUTROS
    ElseIf objEstoqueProduto.dQuantOutras3 < 0 Then
        iEscaninho = ESCANINHO_3_EM_OUTROS
    ElseIf objEstoqueProduto.dQuantOutras < 0 Then
        iEscaninho = ESCANINHO_NOSSO_EM_OUTROS
    End If

End Sub

'Private Function RastreamentoLoteSaldo_Atualiza1(alComando() As Long, ByVal colRastreamentoLoteSaldo As Collection, ByVal colRastreamentoMovto As Collection, ByVal objItemMovEst As ClassItemMovEstoque) As Long
''Atualiza a tabela de RastreamentoLoteSaldo utilizando os valores passados como parametro
''ATENCAO: Esta Funcao tem que ser chamada dentro de transacao
'
'Dim lErro As Long
'Dim objRastreamentoMovto As ClassRastreamentoMovto
'Dim objRastreamentoLoteSaldo As ClassRastreamentoLoteSaldo
'Dim iStatus As Integer
'Dim lNumIntDocLote As Long
'Dim iIndice As Integer
'Dim tRastreamentoLoteSaldo As typeRastreamentoLoteSaldo
'Dim tRastreamentoLoteSaldo1 As typeRastreamentoLoteSaldo
'Dim sLote As String
'
'On Error GoTo Erro_RastreamentoLoteSaldo_Atualiza1
'
'    If Not (colRastreamentoMovto Is Nothing) Then
'
'        For iIndice = 1 To colRastreamentoMovto.Count
'
'            Set objRastreamentoMovto = colRastreamentoMovto.Item(iIndice)
'            Set objRastreamentoLoteSaldo = colRastreamentoLoteSaldo.Item(iIndice)
'
'            sLote = String(STRING_LOTE_RASTREAMENTO, 0)
'
'            'Le a tabela de rastreamentolote para o produto/lote/filial em questão
'            lErro = Comando_ExecutarPos(alComando(80), "SELECT Status, NumIntDoc, Lote FROM RastreamentoLote WHERE NumIntDoc=?", 0, iStatus, lNumIntDocLote, sLote, objRastreamentoMovto.lNumIntDocLote)
'            If lErro <> AD_SQL_SUCESSO Then gError 71780
'
'            'leitura do registro referente ao produto/lote/filial em questão
'            lErro = Comando_BuscarPrimeiro(alComando(80))
'            If lErro <> AD_SQL_SUCESSO Then gError 71781
'
'            'lock do produto/lote/filial
'            lErro = Comando_LockExclusive(alComando(80))
'            If lErro <> AD_SQL_SUCESSO Then gError 71782
'
'            If iStatus <> RATREAMENTOLOTE_STATUS_ABERTO Then gError 71783
'
'            objRastreamentoLoteSaldo.sLote = sLote
'            objRastreamentoLoteSaldo.lNumIntDocLote = lNumIntDocLote
'            objRastreamentoMovto.lNumIntDocLote = lNumIntDocLote
'            objRastreamentoLoteSaldo.iAlmoxarifado = objItemMovEst.iAlmoxarifado
'
'            'Le a tabela de rastreamentolotesaldo para o produto/almoxarifado/lote em questão
'            lErro = Comando_ExecutarPos(alComando(81), "SELECT QuantDispNossa, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantBenef, QuantBenef3, QuantOP FROM RastreamentoLoteSaldo WHERE Produto = ? AND Almoxarifado = ? AND NumIntDocLote=?", 0, tRastreamentoLoteSaldo.dQuantDispNossa, tRastreamentoLoteSaldo.dQuantRecIndl, tRastreamentoLoteSaldo.dQuantIndOutras, tRastreamentoLoteSaldo.dQuantDefeituosa, tRastreamentoLoteSaldo.dQuantConsig3, tRastreamentoLoteSaldo.dQuantConsig, tRastreamentoLoteSaldo.dQuantDemo3, tRastreamentoLoteSaldo.dQuantDemo, tRastreamentoLoteSaldo.dQuantConserto3, tRastreamentoLoteSaldo.dQuantConserto, tRastreamentoLoteSaldo.dQuantOutras3, tRastreamentoLoteSaldo.dQuantOutras, tRastreamentoLoteSaldo.dQuantBenef, tRastreamentoLoteSaldo.dQuantBenef3, tRastreamentoLoteSaldo.dQuantOP, _
'                                            objRastreamentoMovto.sProduto, objItemMovEst.iAlmoxarifado, lNumIntDocLote)
'            If lErro <> AD_SQL_SUCESSO Then gError 71784
'
'            'leitura do registro referente ao estoque/almoxarifado em questão
'            lErro = Comando_BuscarPrimeiro(alComando(81))
'            If lErro <> AD_SQL_SUCESSO Then gError 71785
'
'            'lock do estoque/almoxarifado
'            lErro = Comando_LockExclusive(alComando(81))
'            If lErro <> AD_SQL_SUCESSO Then gError 71786
'
'            If tRastreamentoLoteSaldo.dQuantConserto + objRastreamentoLoteSaldo.dQuantConserto < 0 Then gError 71787
'
'            If tRastreamentoLoteSaldo.dQuantConserto3 + objRastreamentoLoteSaldo.dQuantConserto3 < 0 Then gError 71788
'
'            If tRastreamentoLoteSaldo.dQuantConsig + objRastreamentoLoteSaldo.dQuantConsig < 0 Then gError 71789
'
'            If tRastreamentoLoteSaldo.dQuantConsig3 + objRastreamentoLoteSaldo.dQuantConsig3 < 0 Then gError 71790
'
'            If tRastreamentoLoteSaldo.dQuantDefeituosa + objRastreamentoLoteSaldo.dQuantDefeituosa < 0 Then gError 71791
'
'            If tRastreamentoLoteSaldo.dQuantDemo + objRastreamentoLoteSaldo.dQuantDemo < 0 Then gError 71792
'
'            If tRastreamentoLoteSaldo.dQuantDemo3 + objRastreamentoLoteSaldo.dQuantDemo3 < 0 Then gError 71793
'
'            If tRastreamentoLoteSaldo.dQuantDispNossa + objRastreamentoLoteSaldo.dQuantDispNossa < 0 Then gError 71794
'
'            If tRastreamentoLoteSaldo.dQuantIndOutras + objRastreamentoLoteSaldo.dQuantIndOutras < 0 Then gError 71795
'
'            If tRastreamentoLoteSaldo.dQuantOutras + objRastreamentoLoteSaldo.dQuantOutras < 0 Then gError 71796
'
'            If tRastreamentoLoteSaldo.dQuantOutras3 + objRastreamentoLoteSaldo.dQuantOutras3 < 0 Then gError 71797
'
'            If tRastreamentoLoteSaldo.dQuantBenef + objRastreamentoLoteSaldo.dQuantBenef < 0 Then gError 71798
'
'            If tRastreamentoLoteSaldo.dQuantBenef3 + objRastreamentoLoteSaldo.dQuantBenef3 < 0 Then gError 71799
'
'            lErro = RastreamentoLoteSaldo_Zera_Delta(objRastreamentoLoteSaldo, tRastreamentoLoteSaldo, tRastreamentoLoteSaldo1)
'            If lErro <> AD_SQL_SUCESSO Then gError 184415
'
'            'Atualiza a tabela de rastreamentosaldolote para o produto/almoxarifado/NumIntDocLote em questão
'            lErro = Comando_ExecutarPos(alComando(82), "UPDATE RastreamentoLoteSaldo SET QuantDispNossa = ?, QuantRecIndl = ?, QuantIndOutras = ?, QuantDefeituosa = ?, QuantConsig3 = ?, QuantConsig = ?, QuantDemo3 = ?, QuantDemo = ?, QuantConserto3 = ?, QuantConserto = ?, QuantOutras3 = ?, QuantOutras = ?, QuantReservada = ?, QuantReservadaConsig = ?, QuantBenef = ?, QuantBenef3 = ?, QuantOP = ?", alComando(81), _
'                tRastreamentoLoteSaldo1.dQuantDispNossa, tRastreamentoLoteSaldo1.dQuantRecIndl, tRastreamentoLoteSaldo1.dQuantIndOutras, tRastreamentoLoteSaldo1.dQuantDefeituosa, tRastreamentoLoteSaldo1.dQuantConsig3, tRastreamentoLoteSaldo1.dQuantConsig, tRastreamentoLoteSaldo1.dQuantDemo3, tRastreamentoLoteSaldo1.dQuantDemo, tRastreamentoLoteSaldo1.dQuantConserto3, tRastreamentoLoteSaldo1.dQuantConserto, tRastreamentoLoteSaldo1.dQuantOutras3, tRastreamentoLoteSaldo1.dQuantOutras, tRastreamentoLoteSaldo1.dQuantReservada, tRastreamentoLoteSaldo1.dQuantReservadaConsig, tRastreamentoLoteSaldo1.dQuantBenef, tRastreamentoLoteSaldo1.dQuantBenef3, tRastreamentoLoteSaldo1.dQuantOP)
'            If lErro <> AD_SQL_SUCESSO Then gError 71800
'
'        Next
'
'    End If
'
'    RastreamentoLoteSaldo_Atualiza1 = SUCESSO
'
'    Exit Function
'
'Erro_RastreamentoLoteSaldo_Atualiza1:
'
'    RastreamentoLoteSaldo_Atualiza1 = gErr
'
'    Select Case gErr
'
'        Case 71780, 71781
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE2", gErr, objRastreamentoMovto.sProduto, objRastreamentoMovto.sLote, objRastreamentoMovto.iFilialOP)
'
'        Case 71782
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RASTREAMENTOLOTE", gErr, objRastreamentoMovto.sProduto, objRastreamentoMovto.sLote, objRastreamentoMovto.iFilialOP)
'
'        Case 71783
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTOLOTE_NAO_ABERTO", gErr, objRastreamentoMovto.sProduto, objRastreamentoMovto.sLote, objRastreamentoMovto.iFilialOP)
'
'        Case 71784, 71785
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTESALDO", gErr, objRastreamentoLoteSaldo.sProduto, objItemMovEst.iAlmoxarifado, sLote)
'
'        Case 71786
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RASTREAMENTOLOTESALDO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote)
'
'        Case 71787
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSERTO_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantConserto)
'
'        Case 71788
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSERTO3_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantConserto3)
'
'        Case 71789
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSIG_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantConsig)
'
'        Case 71790
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSIG3_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantConsig3)
'
'        Case 71791
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEFEITUOSO_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantDefeituosa)
'
'        Case 71792
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEMO_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantDemo)
'
'        Case 71793
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEMO3_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantDemo3)
'
'        Case 71794
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DISPONIVEL_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantDispNossa)
'
'        Case 71795
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_INDISPONIVEL_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantIndOutras)
'
'        Case 71796
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_OUTRAS_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantOutras)
'
'        Case 71797
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_OUTRAS3_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantOutras3)
'
'        Case 71798
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_BENEF_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantBenef)
'
'        Case 71799
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_BENEF3_RASTRO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote, tRastreamentoLoteSaldo.dQuantBenef3)
'
'        Case 71800
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RASTREAMENTOLOTESALDO", gErr, objRastreamentoLoteSaldo.sProduto, objRastreamentoLoteSaldo.iAlmoxarifado, sLote)
'
'        Case 184415
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151652)
'
'    End Select
'
'    Exit Function
'
'End Function

Function MovimentoEstoque_Le_NumIntDoc1(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque) As Long
'Verifica se o movimento a ser estornado está cadastrado
    
Dim lErro As Long
Dim lNumIntDocEst As Long
Dim dCusto As Double

On Error GoTo Erro_MovimentoEstoque_Le_NumIntDoc1

    'le o numintdoc de estorno do movimento de estoque passado como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDocEst, Custo FROM MovimentoEstoque WHERE NumIntDoc = ?", 0, lNumIntDocEst, dCusto, objItemMovEst.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 71802

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71803

    'se o movimento não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 71805

    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 71804

    objItemMovEst.lNumIntDocEst = lNumIntDocEst
    objItemMovEst.dCusto = dCusto
        
    MovimentoEstoque_Le_NumIntDoc1 = SUCESSO
    
    Exit Function

Erro_MovimentoEstoque_Le_NumIntDoc1:

    MovimentoEstoque_Le_NumIntDoc1 = gErr

    Select Case gErr

        Case 71802, 71803
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)

        Case 71804
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MOVIMENTOESTOQUE", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)

        Case 71805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO1", gErr, objItemMovEst.lNumIntDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151653)

    End Select

    Exit Function

End Function

Private Function RastreamentoMovto_Le(ByVal lComando As Long, ByVal lNumIntDocOrigem As Long, ByVal iTipoDocOrigem As Integer, ByVal colRastreamentoMovto As Collection) As Long
'Le os rastreamentos relativos a um movimento de estoque

Dim lErro As Long
Dim objRastreamentoMovto As ClassRastreamentoMovto
Dim sProduto As String
Dim lNumIntDocLote As Long
Dim lNumIntDoc As Long
Dim dQuantidade As Double
Dim sLote As String
Dim iFilialOP As Integer
Dim sSiglaUM As String

On Error GoTo Erro_RastreamentoMovto_Le

    sProduto = String(STRING_PRODUTO, 0)
    sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    sSiglaUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
    
    'Selecionar registros na tabela de RastreamentoMovto associados ao movimento que está sendo excluido
    lErro = Comando_Executar(lComando, "SELECT RastreamentoMovto.NumIntDoc, RastreamentoMovto.NumIntDocLote, RastreamentoMovto.Produto, RastreamentoMovto.Quantidade, RastreamentoMovto.SiglaUM, RastreamentoLote.Lote, RastreamentoLote.FilialOP FROM RastreamentoMovto, RastreamentoLote WHERE RastreamentoMovto.NumIntDocLote = RastreamentoLote.NumIntDoc AND RastreamentoMovto.TipoDocOrigem=? AND RastreamentoMovto.NumIntDocOrigem=?", lNumIntDoc, lNumIntDocLote, sProduto, dQuantidade, sSiglaUM, sLote, iFilialOP, iTipoDocOrigem, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 71806

    'leitura do registro referente ao RastreamentoMovto
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71807

    Do While lErro = AD_SQL_SUCESSO
        
        Set objRastreamentoMovto = New ClassRastreamentoMovto
        
        objRastreamentoMovto.lNumIntDoc = lNumIntDoc
        objRastreamentoMovto.lNumIntDocLote = lNumIntDocLote
        objRastreamentoMovto.sProduto = sProduto
        objRastreamentoMovto.dQuantidade = dQuantidade
        objRastreamentoMovto.sSiglaUM = sSiglaUM
        objRastreamentoMovto.iTipoDocOrigem = iTipoDocOrigem
        objRastreamentoMovto.lNumIntDocOrigem = lNumIntDocOrigem
        objRastreamentoMovto.sLote = sLote
        objRastreamentoMovto.iFilialOP = iFilialOP
        
        colRastreamentoMovto.Add objRastreamentoMovto
        
        'leitura do proximo registro referente ao RastreamentoMovto
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71808

    Loop

    RastreamentoMovto_Le = SUCESSO

    Exit Function

Erro_RastreamentoMovto_Le:

    RastreamentoMovto_Le = gErr

    Select Case gErr

        Case 71806, 71807, 71808
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOMOVTO", gErr, lNumIntDocOrigem, iTipoDocOrigem)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151654)

    End Select

    Exit Function

End Function

Private Function ApropriacaoInsumosProd_Le(ByVal lComando As Long, ByVal lComando1 As Long, ByVal lNumIntDocOrigem As Long, ByVal colApropriacaoInsumosProd As Collection, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'carregar as apropriacoes de insumo de producao do movimento de estoque em questao

Dim lNumIntDoc As Long
Dim lErro As Long
Dim tApropriacaoInsumosProd As typeApropricacaoInsumosProd
Dim objApropriacaoInsumosProd As ClassApropriacaoInsumosProd

On Error GoTo Erro_ApropriacaoInsumosProd_Le

    If objTipoMovEstoque.iCodigo = MOV_EST_PRODUCAO Or objTipoMovEstoque.iCodigo = MOV_EST_PRODUCAO_BENEF3 Then

        tApropriacaoInsumosProd.sProduto = String(STRING_PRODUTO, 0)

        'Selecionar registros na tabela de ApropriacaoInsumosProd gerado a partir do lNumIntDocOrigem
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumIntReqProd, Produto, Quantidade, Automatico FROM ApropriacaoInsumosProd WHERE NumIntDocOrigem=?", 0, tApropriacaoInsumosProd.lNumIntDoc, tApropriacaoInsumosProd.lNumIntReqProd, tApropriacaoInsumosProd.sProduto, tApropriacaoInsumosProd.dQuantidade, tApropriacaoInsumosProd.iAutomatico, lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 71810
    
        'leitura do registro referente ao RastreamentoMovto
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71811
        
        Do While lErro = AD_SQL_SUCESSO
        
            Set objApropriacaoInsumosProd = New ClassApropriacaoInsumosProd
            
            objApropriacaoInsumosProd.lNumIntDoc = tApropriacaoInsumosProd.lNumIntDoc
            objApropriacaoInsumosProd.lNumIntReqProd = tApropriacaoInsumosProd.lNumIntReqProd
            objApropriacaoInsumosProd.lNumIntDocOrigem = lNumIntDocOrigem
            objApropriacaoInsumosProd.sProduto = tApropriacaoInsumosProd.sProduto
            objApropriacaoInsumosProd.dQuantidade = tApropriacaoInsumosProd.dQuantidade
            objApropriacaoInsumosProd.iAutomatico = tApropriacaoInsumosProd.iAutomatico
            
            colApropriacaoInsumosProd.Add objApropriacaoInsumosProd
            
            Set objApropriacaoInsumosProd.colRastreamentoMovto = New Collection
            
            'le os rastreamentos relativos a uma apropricao
            lErro = RastreamentoMovto_Le(lComando1, tApropriacaoInsumosProd.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_APROP_PRODUCAO_ENTRADA, objApropriacaoInsumosProd.colRastreamentoMovto)
            If lErro <> SUCESSO Then gError 71812
        
            'leitura do proximo registro referente ao RastreamentoMovto
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71813
    
        Loop

    End If

    ApropriacaoInsumosProd_Le = SUCESSO

    Exit Function

Erro_ApropriacaoInsumosProd_Le:

    ApropriacaoInsumosProd_Le = gErr

    Select Case gErr

        Case 71810, 71811, 71813
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_APROPRIACAOINSUMOSPROD", gErr, lNumIntDocOrigem)

        Case 71812
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151655)

    End Select

    Exit Function

End Function

Private Function UM_Transforma(alComando() As Long, ByVal iClasseUM As Integer, ByVal sSiglaUM1 As String, ByVal sSiglaUM2 As String, dFator As Double) As Long
' devolve o fator de conversao de sSiglaUM1(Origem)  em sSiglaUM2(destino)

Dim lErro As Long
Dim objUnidadeMedida As New ClassUnidadeDeMedida
Dim dQuantidade As Double

On Error GoTo Erro_UM_Transforma

    If sSiglaUM1 = sSiglaUM2 Then
    
        dFator = 1
    
    Else
    
        objUnidadeMedida.iClasse = iClasseUM
        objUnidadeMedida.sSigla = sSiglaUM1
    
        lErro = CF("UM_Le1", alComando(18), objUnidadeMedida)
        If lErro <> SUCESSO Then gError 83222
    
        dQuantidade = objUnidadeMedida.dQuantidade

        objUnidadeMedida.sSigla = sSiglaUM2
    
        lErro = CF("UM_Le1", alComando(18), objUnidadeMedida)
        If lErro <> SUCESSO Then gError 83223
    
        dFator = objUnidadeMedida.dQuantidade / dQuantidade
    
    End If
    
    UM_Transforma = SUCESSO
    
    Exit Function
    
Erro_UM_Transforma:

    UM_Transforma = gErr
    
    Select Case gErr
    
        Case 83222, 83223
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151656)
    
    End Select
    
    Exit Function

End Function

Private Function MATConfig_Grava_DataReproc(ByVal dtData As Date, ByVal iFilialEmpresa As Integer) As Long
'Insere ou Atualiza registro em MATConfig com a data a partir da qual deve ser feito o reprocessamento de estoque.

Dim sConteudo As String
Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim dtDataBD As Date

On Error GoTo Erro_MATConfig_Grava_DataReproc

    'verifica se a data passada como parametro
    'esta zerada, se estiver, nao grava...
    'correcao por tulio em 18/10/02
    'tulio181002
    If dtData = CDate(0) Then
    
      'retorna sucesso para nao desviar o fluxo da rotina
      'chamadora e sai sem fazer coisa alguma...
      MATConfig_Grava_DataReproc = SUCESSO
      Exit Function
      
    
    End If
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 83554
    Next

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Conteudo FROM MATConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, DATA_REPROCESSAMENTO, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 83555
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83556
    
    If lErro = AD_SQL_SUCESSO Then
    
        dtDataBD = CDate(sConteudo)
    
        'se a data passada como parametro for menor do que a data do banco de dados ==> atualiza a data
        If dtData < dtDataBD Then
    
            sConteudo = CStr(dtData)
    
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE MATConfig SET Conteudo = ?", alComando(1), sConteudo)
            If lErro <> AD_SQL_SUCESSO Then gError 83557
    
        End If
    
    Else
    
        sConteudo = CStr(dtData)
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO MATConfig (Codigo, FilialEmpresa, Descricao, Tipo, Conteudo) VALUES (?,?,?,?,?)", DATA_REPROCESSAMENTO, iFilialEmpresa, DATA_REPROCESSAMENTO_DESCR, CONFIG_TIPO_DATA, sConteudo)
        If lErro <> AD_SQL_SUCESSO Then gError 83558
        
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    MATConfig_Grava_DataReproc = SUCESSO
    
    Exit Function

Erro_MATConfig_Grava_DataReproc:

    MATConfig_Grava_DataReproc = gErr
    
    Select Case gErr
    
        Case 83554
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 83555, 83556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG1", gErr, DATA_REPROCESSAMENTO, iFilialEmpresa)
        
        Case 83557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MATCONFIG1", gErr, DATA_REPROCESSAMENTO, iFilialEmpresa)
        
        Case 83558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MATCONFIG", gErr, DATA_REPROCESSAMENTO, iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151657)
        
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function MovEstoque_Altera_Reproc(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque) As Long
'processa a alteração dos dados do movimento de estoque para o reprocessamento

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_MovEstoque_Altera_Reproc

    If objItemMovEst.iTipoMov = MOV_EST_ACRES_INVENT_DISPONIVEL_NOSSA_FX_ZERA Or _
        objItemMovEst.iTipoMov = MOV_EST_DECR_INVENT_DISPONIVEL_NOSSA_FX_ZERA Then
        'Esse movimento tem que zerar o saldo logo, altera a quantidade e potencialmente o tipo de movimento (acréscimo/decréscimo)
        lErro = Comando_Executar(alComando(89), "UPDATE MovimentoEstoque SET Quantidade = ?, TipoMov = ?, Custo=?, Apropriacao=? WHERE NumIntDoc = ?", objItemMovEst.dQuantidade, objItemMovEst.iTipoMov, objItemMovEst.dCusto, objItemMovEst.iApropriacao, objItemMovEst.lNumIntDoc)
    Else
        'atualiza a o numero de documento de estorno no documento que está sendo estornado
        lErro = Comando_Executar(alComando(89), "UPDATE MovimentoEstoque SET Custo=?, Apropriacao=? WHERE NumIntDoc = ?", objItemMovEst.dCusto, objItemMovEst.iApropriacao, objItemMovEst.lNumIntDoc)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 83760

    MovEstoque_Altera_Reproc = SUCESSO

    Exit Function
    
Erro_MovEstoque_Altera_Reproc:

    MovEstoque_Altera_Reproc = gErr
    
    Select Case gErr
    
        Case 83759
    
        Case 83760
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MOVIMENTOESTOQUE", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151658)
    
    End Select
    
    Exit Function

End Function

Private Function Calcula_Custo_Real_Producao(ByVal objItemMovEst As ClassItemMovEstoque, dCustoRealProducao As Double, ByVal objEstoqueMes As ClassEstoqueMes, ByVal objProduto As ClassProduto, ByVal colEstoqueMesProduto As Collection) As Long
'calcula o custo real de produção em função do custo dos insumos requisitados, do rateamento dos custos indiretos em função das horas trabalhadas,
'do rateamento dos custos diretos em função dos custos dos insumos materias primas (sem incluir embalagens e subprodutos)
'a Produção (pois somente produções vão passar por aqui) precisa ter insumos associados ou a OP ter sido baixada

Dim dCustoInsumos As Double
Dim dCustosDiretos As Double
Dim dCustosIndiretos As Double
Dim dCustoMatPrim As Double
Dim lErro As Long
Dim iCalculouCustoMatPrim As Integer
Dim objEstoqueMesProduto As ClassEstoqueMesProduto
Dim iAchou As Integer
Dim dQuantFator1 As Double, dQuantFator2 As Double, dQuantFator3 As Double, dQuantFator4 As Double, dQuantFator5 As Double, dQuantFator6 As Double
Dim dCustoFator1 As Double, dCustoFator2 As Double, dCustoFator3 As Double, dCustoFator4 As Double, dCustoFator5 As Double, dCustoFator6 As Double
Dim dOutrosGastosDiretos As Double, dCustoFatorOutros As Double, dCustoProd As Double
Dim alComando(1 To 3) As Long
Dim iIndice As Integer, bRegistroNovo As Boolean

On Error GoTo Erro_Calcula_Custo_Real_Producao

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201487
    Next
    
    For Each objEstoqueMesProduto In colEstoqueMesProduto
        If objEstoqueMesProduto.iFilialEmpresa = objItemMovEst.iFilialEmpresa And objEstoqueMesProduto.sProduto = objItemMovEst.sProduto Then
            iAchou = 1
            Exit For
        End If
    Next

    dCustoRealProducao = 0
    
    lErro = CF("Calcula_Custo_MatPrim", objItemMovEst, objProduto, dCustoInsumos, dCustoMatPrim, iCalculouCustoMatPrim)
    If lErro <> SUCESSO Then gError 94572
    
    If iCalculouCustoMatPrim = 1 Then
            
        lErro = Comando_ExecutarPos(alComando(1), "SELECT QuantFator1, QuantFator2, QuantFator3, QuantFator4, QuantFator5, QuantFator6 FROM CustoFabrMovto WHERE NumIntMovEst = ?", 0, _
            dQuantFator1, dQuantFator2, dQuantFator3, dQuantFator4, dQuantFator5, dQuantFator6, objItemMovEst.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 201484
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201485
        
        bRegistroNovo = (lErro <> AD_SQL_SUCESSO)
        
        If iAchou = 0 Then
    
            dCustosDiretos = 0
            dOutrosGastosDiretos = Arredonda_Moeda(objEstoqueMes.dGastosDiretos - objEstoqueMes.dCustoFator1 - objEstoqueMes.dCustoFator2 - objEstoqueMes.dCustoFator3 - objEstoqueMes.dCustoFator4 - objEstoqueMes.dCustoFator5 - objEstoqueMes.dCustoFator6)
            
            If objEstoqueMes.dQuantFator1 <> 0 And objEstoqueMes.dCustoFator1 <> 0 And dQuantFator1 <> 0 Then
                dCustoFator1 = (objEstoqueMes.dCustoFator1 * (dQuantFator1 / objEstoqueMes.dQuantFator1))
                dCustosDiretos = dCustosDiretos + dCustoFator1
            End If
             
            If objEstoqueMes.dQuantFator2 <> 0 And objEstoqueMes.dCustoFator2 <> 0 And dQuantFator2 <> 0 Then
                dCustoFator2 = (objEstoqueMes.dCustoFator2 * (dQuantFator2 / objEstoqueMes.dQuantFator2))
                dCustosDiretos = dCustosDiretos + dCustoFator2
            End If
             
            If objEstoqueMes.dQuantFator3 <> 0 And objEstoqueMes.dCustoFator3 <> 0 And dQuantFator3 <> 0 Then
                dCustoFator3 = (objEstoqueMes.dCustoFator3 * (dQuantFator3 / objEstoqueMes.dQuantFator3))
                dCustosDiretos = dCustosDiretos + dCustoFator3
            End If
             
            If objEstoqueMes.dQuantFator4 <> 0 And objEstoqueMes.dCustoFator4 <> 0 And dQuantFator4 <> 0 Then
                dCustoFator4 = (objEstoqueMes.dCustoFator4 * (dQuantFator4 / objEstoqueMes.dQuantFator4))
                dCustosDiretos = dCustosDiretos + dCustoFator4
            End If
             
            If objEstoqueMes.dQuantFator5 <> 0 And objEstoqueMes.dCustoFator5 <> 0 And dQuantFator5 <> 0 Then
                dCustoFator5 = (objEstoqueMes.dCustoFator5 * (dQuantFator5 / objEstoqueMes.dQuantFator5))
                dCustosDiretos = dCustosDiretos + dCustoFator5
            End If
             
            If objEstoqueMes.dQuantFator6 <> 0 And objEstoqueMes.dCustoFator6 <> 0 And dQuantFator6 <> 0 Then
                dCustoFator6 = (objEstoqueMes.dCustoFator6 * (dQuantFator6 / objEstoqueMes.dQuantFator6))
                dCustosDiretos = dCustosDiretos + dCustoFator6
            End If
             
            If objEstoqueMes.dCustoFator1 = 0 And objEstoqueMes.dCustoFator2 = 0 And objEstoqueMes.dCustoFator3 = 0 And objEstoqueMes.dCustoFator4 = 0 And objEstoqueMes.dCustoFator5 = 0 And objEstoqueMes.dCustoFator6 = 0 Then
            
                If objEstoqueMes.lHorasMaquina <> 0 Then
                    
                    dCustoFatorOutros = (objItemMovEst.lHorasMaquina / objEstoqueMes.lHorasMaquina) * dOutrosGastosDiretos
                
                End If
        
            Else
            
                dCustoFatorOutros = (dCustoFator1 + dCustoFator2 + dCustoFator3 + dCustoFator4 + dCustoFator5 + dCustoFator6) / (objEstoqueMes.dCustoFator1 + objEstoqueMes.dCustoFator2 + objEstoqueMes.dCustoFator3 + objEstoqueMes.dCustoFator4 + objEstoqueMes.dCustoFator5 + objEstoqueMes.dCustoFator6) * dOutrosGastosDiretos
            
            End If
            
            dCustosDiretos = dCustosDiretos + dCustoFatorOutros
        
            If objEstoqueMes.dCustoMatPrim <> 0 Then
                dCustosIndiretos = (dCustoMatPrim / objEstoqueMes.dCustoMatPrim) * objEstoqueMes.dGastosIndiretos
            End If
            
            dCustoRealProducao = dCustoInsumos + dCustosDiretos + dCustosIndiretos
                
        Else
        
            dCustoProd = objItemMovEst.dQuantidadeEst * (objEstoqueMesProduto.dGasto / objEstoqueMesProduto.dQuantidade)
            dCustoRealProducao = dCustoProd + dCustoInsumos
        
        End If
        
        If bRegistroNovo Then
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO CustoFabrMovto (NumIntMovEst, QuantFator1, QuantFator2, QuantFator3, QuantFator4, QuantFator5, QuantFator6, Versao, QuantFator, UnidadeMed, CustoFatorOutros, CustoFator1, CustoFator2, CustoFator3, CustoFator4, CustoFator5, CustoFator6, CustoInsumos, CustoProd, CustosIndiretos) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objItemMovEst.lNumIntDoc, 0, 0, 0, 0, 0, 0, "", 0, "", dCustoFatorOutros, dCustoFator1, dCustoFator2, dCustoFator3, dCustoFator4, dCustoFator5, dCustoFator6, dCustoInsumos, dCustoProd, dCustosIndiretos)
            If lErro <> AD_SQL_SUCESSO Then gError 201486
        
        Else
        
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE CustoFabrMovto SET CustoFatorOutros = ?, CustoFator1 = ?, CustoFator2 = ?, CustoFator3 = ?, CustoFator4 = ?, CustoFator5 = ?, CustoFator6 = ?, CustoInsumos = ?, CustoProd = ?, CustosIndiretos = ?", alComando(1), _
                dCustoFatorOutros, dCustoFator1, dCustoFator2, dCustoFator3, dCustoFator4, dCustoFator5, dCustoFator6, dCustoInsumos, dCustoProd, dCustosIndiretos)
            If lErro <> AD_SQL_SUCESSO Then gError 201486
        
        End If
    Else
    
        dCustoRealProducao = 0
        
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Calcula_Custo_Real_Producao = SUCESSO
    
    Exit Function

Erro_Calcula_Custo_Real_Producao:

    Calcula_Custo_Real_Producao = gErr

    Select Case gErr
    
        Case 92504, 94572
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151659)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Calcula_Custo_MatPrim(ByVal objItemMovEst As ClassItemMovEstoque, ByVal objProduto As ClassProduto, dCustoInsumos As Double, dCustoMatPrim As Double, iCalculouCustoMatPrim As Integer) As Long
'calcula o custo real de produção em função do custo dos insumos requisitados, do rateamento dos custos indiretos em função das horas trabalhadas,
'do rateamento dos custos diretos em função dos custos dos insumos materias primas (sem incluir embalagens e subprodutos)
'a Produção (pois somente produções vão passar por aqui) precisa ter insumos associados ou a OP ter sido baixada.

Dim objApropriacaoInsumo As ClassApropriacaoInsumosProd
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objItemMovEst1 As ClassItemMovEstoque
Dim dFator As Double
Dim dCustosDiretos As Double
Dim dCustosIndiretos As Double
Dim lErro As Long
Dim colItemMovEst As New Collection
Dim dQuantidadeEst As Double

On Error GoTo Erro_Calcula_Custo_MatPrim

    
    iCalculouCustoMatPrim = 0

    'se o produto tiver apropriação de insumo ==> pode calcular o custo a partir destes dados
    If objItemMovEst.colApropriacaoInsumo.Count > 0 Then
                
        lErro = CF("Calcula_Custo_MatPrim_Apropriacao", objItemMovEst, dCustoMatPrim, dCustoInsumos)
        If lErro <> SUCESSO Then gError 92664
        
        iCalculouCustoMatPrim = 1
    
    Else
    
        'se o produto não tem apropriação de insumos associado ==> a O.P. precisa estar baixada
        'calcular os insumos pelo rateamento entre as varias produções associadas a OP
    
        objOrdemProducao.iFilialEmpresa = objItemMovEst.iFilialEmpresa
        objOrdemProducao.sCodigo = objItemMovEst.sOPCodigo
    
        'pesquisa se a O.P. associada está baixada
        lErro = CF("OrdemProducao_Le1", objOrdemProducao)
        If lErro <> SUCESSO And lErro <> 94578 And lErro <> 94579 Then gError 92513

        'se a O.P. não estiver cadastrada ==> erro
        If lErro = 94579 Then gError 92514

        'se a O.P. estiver baixada
        If lErro = 94578 Then
    
            'pesquisa as produções para este produto associada a esta O.P.
            lErro = CF("MovimentoEstoque_Le_Producoes", objItemMovEst, colItemMovEst)
            If lErro <> SUCESSO Then gError 92516
        
            'pesquisa o custo das materias primas e insumos associadas à produção
            lErro = CF("MovEstoque_Le_CustoMatPrim", objItemMovEst, dCustoMatPrim, dCustoInsumos, colItemMovEst, MOSTRA_MENSAGEM_ERRO_QUANTAPROPRIADA)
            If lErro <> SUCESSO And lErro <> 92489 Then gError 92519
        
            iCalculouCustoMatPrim = 1
        
'            'Se tiver mais de uma produção para a O.P. associada a produção em questão ==> calcula o custo das mat. primas e insumos pelo rateio da quantidade produzida neste movimento pela quantidade total produzida pelas diversas produções associadas a O.P.
'            If colItemMovEst.Count > 1 Then
'
'                For Each objItemMovEst1 In colItemMovEst
'
'                    lErro = CF("UM_Conversao_Trans",objProduto.iClasseUM, objItemMovEst1.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
'                    If lErro <> SUCESSO Then gError 92512
'
'                    dQuantidadeEst = dQuantidadeEst + objItemMovEst1.dQuantidade * dFator
'
'                Next
'
'                'rateia o custo da matéria prima com relação a quantidade total produzida para o produto levando em consideração todas as produções para a O.P. em questão
'                dCustoMatPrim = dCustoMatPrim * objItemMovEst.dQuantidadeEst / dQuantidadeEst
'
'                'rateia o custo dos insumos com relação a quantidade total produzida para o produto levando em consideração todas as produções para a O.P. em questão
'
'                dCustoInsumos = dCustoInsumos * objItemMovEst.dQuantidadeEst / dQuantidadeEst
'
'            End If
            
        End If
            
    End If
    
    Calcula_Custo_MatPrim = SUCESSO
    
    Exit Function

Erro_Calcula_Custo_MatPrim:

    Calcula_Custo_MatPrim = gErr

    Select Case gErr
    
        Case 92504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 92513, 92516, 92519, 92664, ERRO_SEM_MENSAGEM
            
        Case 92514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE_MOVESTOQUE", gErr, objItemMovEst.sOPCodigo, objItemMovEst.sProduto)
        
        Case 92515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_NAO_BAIXADA_MOVESTOQUE", gErr, objItemMovEst.sOPCodigo, objItemMovEst.sProduto)

        Case 92517
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OUTRA_PRODUCAO_MESMA_OP", gErr, objItemMovEst.sProduto, objItemMovEst.sOPCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151660)
    
    End Select

    Exit Function

End Function

Function Estoque_ApuraCustoProducao(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque, ByVal objEstoqueMes As ClassEstoqueMes, ByVal colEstoqueMesProduto As Collection) As Long
'apura o custo de produção
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim lNumIntDoc As Long
Dim iStatus As Integer
Dim objProduto As New ClassProduto
Dim objEstoque As New ClassEstoque
Dim objMovEst As ClassMovEstoque

On Error GoTo Erro_Estoque_ApuraCustoProducao

    objProduto.sCodigo = objItemMovEst.sProduto
    
    'le o atributo controleestoque do produto em questão. Os atributos siglaUM e classeUM se deve a necessidade converter a unidade do movimento para unidade de estoque
    lErro = CF("Produto_Le_ControleEstoque", alComando(17), objProduto)
    If lErro <> SUCESSO Then gError 92535

    objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
    objItemMovEst.iClasseUM = objProduto.iClasseUM
        
    If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
        
        Set objEstoque.colEstoqueMesProduto = colEstoqueMesProduto
        
        'excecutar a gravação do item
        lErro = Estoque_AtualizaItemMov(alComando, objItemMovEst, lNumIntDoc, MOVESTOQUE_CODIGO_JA_CADASTRADO, objProduto, objMovEst, objEstoqueMes, APURACAO_CUSTO_PRODUCAO, objEstoque)
        If lErro <> SUCESSO Then gError 92536
        
    End If
        
    Estoque_ApuraCustoProducao = SUCESSO
    
    Exit Function
    
Erro_Estoque_ApuraCustoProducao:

    Estoque_ApuraCustoProducao = gErr

    Select Case gErr
    
        Case 92535, 92536
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151661)
    
    End Select
    
    Exit Function

End Function

Function MovEst_ObtemCustoMovOrig(ByVal objItemMovEst As ClassItemMovEstoque, dCustoUnitarioMovOrig As Double) As Long
'Retorna em dCustoUnitarioMovOrig o custo unitario (correspondente a unidades de estoque) do movimento de estoque de saida associado ao que origina objItemMovEst
'Esta funcao é especifica para o tratamento de notas fiscais de transferencia de material da propria empresa entre filiais onde o custo de entrada deve ser igual ao de saida em vez do valor da nfiscal

Dim lErro As Long, alComando(1 To 2) As Long, iIndice As Integer
Dim dtDataEmissao As Date, sSerie As String, lNumNotaFiscal As Long
Dim dCusto As Double
Dim dQuantidade As Double
Dim sSiglaUM As String
Dim sSiglaUMEstoque As String
Dim iClasseUM As Integer
Dim dFator As Double

On Error GoTo Erro_MovEst_ObtemCustoMovOrig
        
    dCustoUnitarioMovOrig = 0
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81894
    Next
    
    If objItemMovEst.lNumIntDocOrigem <> 0 And objItemMovEst.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL Then
    
        sSerie = String(STRING_SERIE, 0)
        
        'obtem dados da nf de entrada
        lErro = Comando_Executar(alComando(1), "SELECT DataEmissao, Serie, NumNotaFiscal FROM NFiscal, ItensNFiscal WHERE ItensNFiscal.NumIntDoc = ? AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF", dtDataEmissao, sSerie, lNumNotaFiscal, objItemMovEst.lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 81895
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81896
        If lErro <> AD_SQL_SUCESSO Then gError 81897
    
        sSiglaUM = String(STRING_UM_SIGLA, 0)
        sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
        
        'busca a nf de saida da outra filialempresa
        lErro = Comando_Executar(alComando(2), "SELECT MovimentoEstoque.Custo, MovimentoEstoque.Quantidade, MovimentoEstoque.SiglaUM, ClasseUM, SiglaUMEstoque FROM NFiscal, ItensNFiscal, MovimentoEstoque, Produtos WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.NumNotaFiscal = ? AND NFiscal.DataEmissao = ? AND NFiscal.Cliente = ? AND MovimentoEstoque.TipoNumIntDocOrigem = ? AND MovimentoEstoque.NumIntDocOrigem = ItensNFiscal.NumIntDoc AND ItensNFiscal.Produto = Produtos.Codigo AND Produtos.Codigo = ? AND MovimentoEstoque.FilialEmpresa <> ?", _
            dCusto, dQuantidade, sSiglaUM, iClasseUM, sSiglaUMEstoque, lNumNotaFiscal, dtDataEmissao, gobjCRFAT.lCliEmp, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL, objItemMovEst.sProduto, objItemMovEst.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 81898

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81899
        If lErro <> AD_SQL_SUCESSO Then gError 81900

        'Converte a unidade de medida da quantidade do Movimento de Estoque para a UM de Estoque
        lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUM, sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 81901
        
        dQuantidade = dQuantidade * dFator
        
        If Abs(dQuantidade) > QTDE_ESTOQUE_DELTA Then
            dCustoUnitarioMovOrig = dCusto / dQuantidade
        Else
            dCustoUnitarioMovOrig = 0
        End If
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    MovEst_ObtemCustoMovOrig = SUCESSO
     
    Exit Function
    
Erro_MovEst_ObtemCustoMovOrig:

    MovEst_ObtemCustoMovOrig = gErr
     
    Select Case gErr
          
        Case 81894
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 81895, 81896
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFO_MOVORIG", gErr)
        
        Case 81898, 81899
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFO_MOVORIG3", gErr)
        
        Case 81897, 81900, 81901
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151662)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function Jones_Acerta_SaldosEst_3112() As Long

Dim lErro As Long, iFilialEmpresa As Integer, sProduto As String, iAlmoxarifado As Integer, iIndice As Integer
Dim tSldMesEst_2001 As typeSldMesEst, tSldMesEst_2002 As typeSldMesEst
Dim tSldMesEst1_2001 As typeSldMesEst1, tSldMesEst1_2002 As typeSldMesEst1
Dim tSldMesEstAlm_2001 As typeSldMesEstAlm
Dim tSldMesEstAlm1_2001 As typeSldMesEstAlm1
Dim tSldMesEstAlm2_2001 As typeSldMesEstAlm2
Dim tSldDiaEst As typeSldDiaEst, tSldDiaEstAlm As typeSldDiaEstAlm
Dim alComando(1 To 64) As Long, lTransacao As Long, iMes As Integer

On Error GoTo Erro_Jones_Acerta_SaldosEst_3112

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 81923

    Next
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 81924

    sProduto = String(STRING_PRODUTO, 0)
    
'    lErro = Comando_Executar(alComando(1), "SELECT FilialEmpresa, Produto, Almoxarifado FROM ProdutosFilial WHERE Produto IN ('S000105','0000364','M000030') ORDER BY FilialEmpresa, Produto", iFilialEmpresa, sProduto, iAlmoxarifado)
    lErro = Comando_Executar(alComando(1), "SELECT FilialEmpresa, Produto, Almoxarifado FROM ProdutosFilial ORDER BY FilialEmpresa, Produto", iFilialEmpresa, sProduto, iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 81925
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81926
    
    'loop de produtosfilial
    Do While lErro = AD_SQL_SUCESSO
     
        lErro = Jones_Acerta_SaldosEst2(alComando, iFilialEmpresa, sProduto, iAlmoxarifado)
        If lErro <> SUCESSO Then gError 81927
        
        lErro = Jones_Acerta_SaldosEst1(alComando, iFilialEmpresa, sProduto, iAlmoxarifado)
        If lErro <> SUCESSO Then gError 81928
        
        lErro = Jones_Acerta_SaldosEst(alComando, iFilialEmpresa, sProduto, iAlmoxarifado)
        If lErro <> SUCESSO Then gError 81929
        
        lErro = Jones_Acerta_SaldosEstCusto(alComando, iFilialEmpresa, sProduto, iAlmoxarifado)
        If lErro <> SUCESSO Then gError 81929
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81930
        
    Loop
        
    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 81931
    
    Jones_Acerta_SaldosEst_3112 = SUCESSO
     
    Exit Function
    
Erro_Jones_Acerta_SaldosEst_3112:

    Jones_Acerta_SaldosEst_3112 = gErr
     
    Select Case gErr
          
        Case 81923 To 81931
            Call Rotina_Erro(vbOKOnly, "ERRO_JONES", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151663)
     
    End Select
     
    Call Transacao_Rollback
    
    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Jones_Acerta_SaldosEst2(alComando() As Long, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal iAlmoxarifado As Integer) As Long
'Acerta Sldmesest2 e Sldmesestalm2(material nosso em 3os) em 2001 para chegar ao saldo inicial desejado em 2002

Dim lErro As Long, iMes As Integer
Dim sSelectSldMesEst2_2001 As String, sSelectSldMesEst2_2002 As String
Dim tSldMesEst2_2001 As typeSldMesEst2, tSldMesEst2_2002 As typeSldMesEst2
Dim dQuantConsig As Double, dValorConsig As Double, dQuantDemo As Double, dValorDemo As Double, dQuantConserto As Double, dValorConserto As Double, dQuantOutros As Double, dValorOutros As Double, dQuantBenef As Double, dValorBenef As Double
Dim dDifQuantConsig As Double, dDifValorConsig As Double, dDifQuantDemo As Double, dDifValorDemo As Double, dDifQuantConserto As Double, dDifValorConserto As Double, dDifQuantOutros As Double, dDifValorOutros As Double, dDifQuantBenef As Double, dDifValorBenef As Double

On Error GoTo Erro_Jones_Acerta_SaldosEst2

    'le sldmesest2 2001
    sSelectSldMesEst2_2001 = "SELECT QuantInicialConsig, ValorInicialConsig, QuantInicialDemo, ValorInicialDemo, QuantInicialConserto, ValorInicialConserto, QuantInicialOutros, ValorInicialOutros, QuantInicialBenef, ValorInicialBenef"
    
    For iMes = 1 To 12
    
        sSelectSldMesEst2_2001 = sSelectSldMesEst2_2001 & ", SaldoQuantConsig" & CStr(iMes) & ", SaldoValorConsig" & CStr(iMes) & ", SaldoQuantDemo" & CStr(iMes) & ", SaldoValorDemo" & CStr(iMes) & ", SaldoQuantConserto" & CStr(iMes) & ", SaldoValorConserto" & CStr(iMes) & ", SaldoQuantOutros" & CStr(iMes) & ", SaldoValorOutros" & CStr(iMes) & ", SaldoQuantBenef" & CStr(iMes) & ", SaldoValorBenef" & CStr(iMes)
    
    Next
    
    sSelectSldMesEst2_2001 = sSelectSldMesEst2_2001 & " FROM SldMesEst2 WHERE FilialEmpresa = ? AND Produto = ? AND Ano = 2001"
    With tSldMesEst2_2001
        lErro = Comando_Executar(alComando(2), sSelectSldMesEst2_2001, .dQuantInicialConsig, .dValorInicialConsig, .dQuantInicialDemo, .dValorInicialDemo, .dQuantInicialConserto, .dValorInicialConserto, .dQuantInicialOutros, .dValorInicialOutros, .dQuantInicialBenef, .dValorInicialBenef, _
            .adSaldoQuantConsig(1), .adSaldoValorConsig(1), .adSaldoQuantDemo(1), .adSaldoValorDemo(1), .adSaldoQuantConserto(1), .adSaldoValorConserto(1), .adSaldoQuantOutros(1), .adSaldoValorOutros(1), .adSaldoQuantBenef(1), .adSaldoValorBenef(1), _
            .adSaldoQuantConsig(2), .adSaldoValorConsig(2), .adSaldoQuantDemo(2), .adSaldoValorDemo(2), .adSaldoQuantConserto(2), .adSaldoValorConserto(2), .adSaldoQuantOutros(2), .adSaldoValorOutros(2), .adSaldoQuantBenef(2), .adSaldoValorBenef(2), _
            .adSaldoQuantConsig(3), .adSaldoValorConsig(3), .adSaldoQuantDemo(3), .adSaldoValorDemo(3), .adSaldoQuantConserto(3), .adSaldoValorConserto(3), .adSaldoQuantOutros(3), .adSaldoValorOutros(3), .adSaldoQuantBenef(3), .adSaldoValorBenef(3), _
            .adSaldoQuantConsig(4), .adSaldoValorConsig(4), .adSaldoQuantDemo(4), .adSaldoValorDemo(4), .adSaldoQuantConserto(4), .adSaldoValorConserto(4), .adSaldoQuantOutros(4), .adSaldoValorOutros(4), .adSaldoQuantBenef(4), .adSaldoValorBenef(4), _
            .adSaldoQuantConsig(5), .adSaldoValorConsig(5), .adSaldoQuantDemo(5), .adSaldoValorDemo(5), .adSaldoQuantConserto(5), .adSaldoValorConserto(5), .adSaldoQuantOutros(5), .adSaldoValorOutros(5), .adSaldoQuantBenef(5), .adSaldoValorBenef(5), _
            .adSaldoQuantConsig(6), .adSaldoValorConsig(6), .adSaldoQuantDemo(6), .adSaldoValorDemo(6), .adSaldoQuantConserto(6), .adSaldoValorConserto(6), .adSaldoQuantOutros(6), .adSaldoValorOutros(6), .adSaldoQuantBenef(6), .adSaldoValorBenef(6), _
            .adSaldoQuantConsig(7), .adSaldoValorConsig(7), .adSaldoQuantDemo(7), .adSaldoValorDemo(7), .adSaldoQuantConserto(7), .adSaldoValorConserto(7), .adSaldoQuantOutros(7), .adSaldoValorOutros(7), .adSaldoQuantBenef(7), .adSaldoValorBenef(7), _
            .adSaldoQuantConsig(8), .adSaldoValorConsig(8), .adSaldoQuantDemo(8), .adSaldoValorDemo(8), .adSaldoQuantConserto(8), .adSaldoValorConserto(8), .adSaldoQuantOutros(8), .adSaldoValorOutros(8), .adSaldoQuantBenef(8), .adSaldoValorBenef(8), _
            .adSaldoQuantConsig(9), .adSaldoValorConsig(9), .adSaldoQuantDemo(9), .adSaldoValorDemo(9), .adSaldoQuantConserto(9), .adSaldoValorConserto(9), .adSaldoQuantOutros(9), .adSaldoValorOutros(9), .adSaldoQuantBenef(9), .adSaldoValorBenef(9), _
            .adSaldoQuantConsig(10), .adSaldoValorConsig(10), .adSaldoQuantDemo(10), .adSaldoValorDemo(10), .adSaldoQuantConserto(10), .adSaldoValorConserto(10), .adSaldoQuantOutros(10), .adSaldoValorOutros(10), .adSaldoQuantBenef(10), .adSaldoValorBenef(10), _
            .adSaldoQuantConsig(11), .adSaldoValorConsig(11), .adSaldoQuantDemo(11), .adSaldoValorDemo(11), .adSaldoQuantConserto(11), .adSaldoValorConserto(11), .adSaldoQuantOutros(11), .adSaldoValorOutros(11), .adSaldoQuantBenef(11), .adSaldoValorBenef(11), _
            .adSaldoQuantConsig(12), .adSaldoValorConsig(12), .adSaldoQuantDemo(12), .adSaldoValorDemo(12), .adSaldoQuantConserto(12), .adSaldoValorConserto(12), .adSaldoQuantOutros(12), .adSaldoValorOutros(12), .adSaldoQuantBenef(12), .adSaldoValorBenef(12), _
            iFilialEmpresa, sProduto)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 81913
    
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81914
    
    If lErro = AD_SQL_SUCESSO Then
    
        'le 2002 do produto filial
        With tSldMesEst2_2002
            lErro = Comando_Executar(alComando(3), "SELECT QuantInicialConsig, ValorInicialConsig, QuantInicialDemo, ValorInicialDemo, QuantInicialConserto, ValorInicialConserto, QuantInicialOutros, ValorInicialOutros, QuantInicialBenef, ValorInicialBenef FROM SldMesEst2 WHERE FilialEmpresa = ? AND Produto = ? AND Ano = 2002", _
                .dQuantInicialConsig, .dValorInicialConsig, .dQuantInicialDemo, .dValorInicialDemo, .dQuantInicialConserto, .dValorInicialConserto, .dQuantInicialOutros, .dValorInicialOutros, .dQuantInicialBenef, .dValorInicialBenef, iFilialEmpresa, sProduto)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 81911
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81912
        
        If lErro = AD_SQL_SUCESSO Then
    
            With tSldMesEst2_2001
                
                dQuantConsig = .dQuantInicialConsig + .adSaldoQuantConsig(1) + .adSaldoQuantConsig(2) + .adSaldoQuantConsig(3) + .adSaldoQuantConsig(4) + .adSaldoQuantConsig(5) + .adSaldoQuantConsig(6) + .adSaldoQuantConsig(7) + .adSaldoQuantConsig(8) + .adSaldoQuantConsig(9) + .adSaldoQuantConsig(10) + .adSaldoQuantConsig(11) + .adSaldoQuantConsig(12)
                dValorConsig = .dValorInicialConsig + .adSaldoValorConsig(1) + .adSaldoValorConsig(2) + .adSaldoValorConsig(3) + .adSaldoValorConsig(4) + .adSaldoValorConsig(5) + .adSaldoValorConsig(6) + .adSaldoValorConsig(7) + .adSaldoValorConsig(8) + .adSaldoValorConsig(9) + .adSaldoValorConsig(10) + .adSaldoValorConsig(11) + .adSaldoValorConsig(12)
                dQuantDemo = .dQuantInicialDemo + .adSaldoQuantDemo(1) + .adSaldoQuantDemo(2) + .adSaldoQuantDemo(3) + .adSaldoQuantDemo(4) + .adSaldoQuantDemo(5) + .adSaldoQuantDemo(6) + .adSaldoQuantDemo(7) + .adSaldoQuantDemo(8) + .adSaldoQuantDemo(9) + .adSaldoQuantDemo(10) + .adSaldoQuantDemo(11) + .adSaldoQuantDemo(12)
                dValorDemo = .dValorInicialDemo + .adSaldoValorDemo(1) + .adSaldoValorDemo(2) + .adSaldoValorDemo(3) + .adSaldoValorDemo(4) + .adSaldoValorDemo(5) + .adSaldoValorDemo(6) + .adSaldoValorDemo(7) + .adSaldoValorDemo(8) + .adSaldoValorDemo(9) + .adSaldoValorDemo(10) + .adSaldoValorDemo(11) + .adSaldoValorDemo(12)
                dQuantConserto = .dQuantInicialConserto + .adSaldoQuantConserto(1) + .adSaldoQuantConserto(2) + .adSaldoQuantConserto(3) + .adSaldoQuantConserto(4) + .adSaldoQuantConserto(5) + .adSaldoQuantConserto(6) + .adSaldoQuantConserto(7) + .adSaldoQuantConserto(8) + .adSaldoQuantConserto(9) + .adSaldoQuantConserto(10) + .adSaldoQuantConserto(11) + .adSaldoQuantConserto(12)
                dValorConserto = .dValorInicialConserto + .adSaldoValorConserto(1) + .adSaldoValorConserto(2) + .adSaldoValorConserto(3) + .adSaldoValorConserto(4) + .adSaldoValorConserto(5) + .adSaldoValorConserto(6) + .adSaldoValorConserto(7) + .adSaldoValorConserto(8) + .adSaldoValorConserto(9) + .adSaldoValorConserto(10) + .adSaldoValorConserto(11) + .adSaldoValorConserto(12)
                dQuantOutros = .dQuantInicialOutros + .adSaldoQuantOutros(1) + .adSaldoQuantOutros(2) + .adSaldoQuantOutros(3) + .adSaldoQuantOutros(4) + .adSaldoQuantOutros(5) + .adSaldoQuantOutros(6) + .adSaldoQuantOutros(7) + .adSaldoQuantOutros(8) + .adSaldoQuantOutros(9) + .adSaldoQuantOutros(10) + .adSaldoQuantOutros(11) + .adSaldoQuantOutros(12)
                dValorOutros = .dValorInicialOutros + .adSaldoValorOutros(1) + .adSaldoValorOutros(2) + .adSaldoValorOutros(3) + .adSaldoValorOutros(4) + .adSaldoValorOutros(5) + .adSaldoValorOutros(6) + .adSaldoValorOutros(7) + .adSaldoValorOutros(8) + .adSaldoValorOutros(9) + .adSaldoValorOutros(10) + .adSaldoValorOutros(11) + .adSaldoValorOutros(12)
                dQuantBenef = .dQuantInicialBenef + .adSaldoQuantBenef(1) + .adSaldoQuantBenef(2) + .adSaldoQuantBenef(3) + .adSaldoQuantBenef(4) + .adSaldoQuantBenef(5) + .adSaldoQuantBenef(6) + .adSaldoQuantBenef(7) + .adSaldoQuantBenef(8) + .adSaldoQuantBenef(9) + .adSaldoQuantBenef(10) + .adSaldoQuantBenef(11) + .adSaldoQuantBenef(12)
                dValorBenef = .dValorInicialBenef + .adSaldoValorBenef(1) + .adSaldoValorBenef(2) + .adSaldoValorBenef(3) + .adSaldoValorBenef(4) + .adSaldoValorBenef(5) + .adSaldoValorBenef(6) + .adSaldoValorBenef(7) + .adSaldoValorBenef(8) + .adSaldoValorBenef(9) + .adSaldoValorBenef(10) + .adSaldoValorBenef(11) + .adSaldoValorBenef(12)
            
            End With
            
            With tSldMesEst2_2002
            
                dDifQuantConsig = .dQuantInicialConsig - dQuantConsig
                dDifValorConsig = .dValorInicialConsig - dValorConsig
                dDifQuantDemo = .dQuantInicialDemo - dQuantDemo
                dDifValorDemo = .dValorInicialDemo - dValorDemo
                dDifQuantConserto = .dQuantInicialConserto - dQuantConserto
                dDifValorConserto = .dValorInicialConserto - dValorConserto
                dDifQuantOutros = .dQuantInicialOutros - dQuantOutros
                dDifValorOutros = .dValorInicialOutros - dValorOutros
                dDifQuantBenef = .dQuantInicialBenef - dQuantBenef
                dDifValorBenef = .dValorInicialBenef - dValorBenef
            
            End With
            
            'se houver diferenca
            If Abs(dDifQuantConsig) > QTDE_ESTOQUE_DELTA Or _
                Abs(dDifValorConsig) > DELTA_VALORMONETARIO Or _
                Abs(dDifQuantDemo) > QTDE_ESTOQUE_DELTA Or _
                Abs(dDifValorDemo) > DELTA_VALORMONETARIO Or _
                Abs(dDifQuantConserto) > QTDE_ESTOQUE_DELTA Or _
                Abs(dDifValorConserto) > DELTA_VALORMONETARIO Or _
                Abs(dDifQuantOutros) > QTDE_ESTOQUE_DELTA Or _
                Abs(dDifValorOutros) > DELTA_VALORMONETARIO Or _
                Abs(dDifQuantBenef) > QTDE_ESTOQUE_DELTA Or _
                Abs(dDifValorBenef) > DELTA_VALORMONETARIO Then

                'modificar quant/valor inicial em sldmesest2
                lErro = Comando_Executar(alComando(4), "UPDATE SldMesEst2 SET QuantInicialConsig = QuantInicialConsig + ?, ValorInicialConsig = ValorInicialConsig + ?, QuantInicialDemo = QuantInicialDemo + ?, ValorInicialDemo = ValorInicialDemo + ?, QuantInicialConserto = QuantInicialConserto + ?, ValorInicialConserto = ValorInicialConserto + ?, QuantInicialOutros = QuantInicialOutros + ?, ValorInicialOutros = ValorInicialOutros + ?, QuantInicialBenef = QuantInicialBenef + ?, ValorInicialBenef = ValorInicialBenef + ? WHERE FilialEmpresa = ? AND Produto = ? AND Ano = 2001", _
                    dDifQuantConsig, dDifValorConsig, dDifQuantDemo, dDifValorDemo, dDifQuantConserto, dDifValorConserto, dDifQuantOutros, dDifValorOutros, dDifQuantBenef, dDifValorBenef, iFilialEmpresa, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81915
                
                'modificar quant/valor inicial em sldmesestalm2 no almoxarifado padrao
                lErro = Comando_Executar(alComando(5), "UPDATE SldMesEstAlm2 SET QuantInicialConsig = QuantInicialConsig + ?, ValorInicialConsig = ValorInicialConsig + ?, QuantInicialDemo = QuantInicialDemo + ?, ValorInicialDemo = ValorInicialDemo + ?, QuantInicialConserto = QuantInicialConserto + ?, ValorInicialConserto = ValorInicialConserto + ?, QuantInicialOutros = QuantInicialOutros + ?, ValorInicialOutros = ValorInicialOutros + ?, QuantInicialBenef = QuantInicialBenef + ?, ValorInicialBenef = ValorInicialBenef + ? WHERE Almoxarifado = ? AND Produto = ? AND Ano = 2001", _
                    dDifQuantConsig, dDifValorConsig, dDifQuantDemo, dDifValorDemo, dDifQuantConserto, dDifValorConserto, dDifQuantOutros, dDifValorOutros, dDifQuantBenef, dDifValorBenef, iAlmoxarifado, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81916
                
                lErro = Comando_Executar(alComando(16), "UPDATE EstoqueProduto SET QuantInicialConsig = QuantInicialConsig + ?, ValorInicialConsig = ValorInicialConsig + ?, QuantInicialDemo = QuantInicialDemo + ?, ValorInicialDemo = ValorInicialDemo + ?, QuantInicialConserto = QuantInicialConserto + ?, ValorInicialConserto = ValorInicialConserto + ?, QuantInicialOutras = QuantInicialOutras + ?, ValorInicialOutras = ValorInicialOutras + ?, QuantInicialBenef = QuantInicialBenef + ?, ValorInicialBenef = ValorInicialBenef + ?, DataUltAtualizacao = ?, Exportado = ?  WHERE Almoxarifado = ? AND Produto = ?", _
                    dDifQuantConsig, dDifValorConsig, dDifQuantDemo, dDifValorDemo, dDifQuantConserto, dDifValorConserto, dDifQuantOutros, dDifValorOutros, dDifQuantBenef, dDifValorBenef, Date, 0, iAlmoxarifado, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81916
                
            End If
        
        End If
        
    End If

    Jones_Acerta_SaldosEst2 = SUCESSO
     
    Exit Function
    
Erro_Jones_Acerta_SaldosEst2:

    Jones_Acerta_SaldosEst2 = gErr
     
    Select Case gErr
          
        Case 81911 To 81916
            Call Rotina_Erro(vbOKOnly, "ERRO_JONES", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151664)
     
    End Select
     
    Exit Function

End Function
        
Private Function Jones_Acerta_SaldosEst1(alComando() As Long, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal iAlmoxarifado As Integer) As Long
'Acerta Sldmesest1 e Sldmesestalm1(material de 3os em nosso poder) em 2001 para chegar ao saldo inicial desejado em 2002

Dim lErro As Long, iMes As Integer
Dim sSelectSldMesEst1_2001 As String, sSelectSldMesEst1_2002 As String
Dim tSldMesEst1_2001 As typeSldMesEst1, tSldMesEst1_2002 As typeSldMesEst1
Dim dQuantConsig3 As Double, dValorConsig3 As Double, dQuantDemo3 As Double, dValorDemo3 As Double, dQuantConserto3 As Double, dValorConserto3 As Double, dQuantOutros3 As Double, dValorOutros3 As Double, dQuantBenef3 As Double, dValorBenef3 As Double
Dim dDifQuantConsig3 As Double, dDifValorConsig3 As Double, dDifQuantDemo3 As Double, dDifValorDemo3 As Double, dDifQuantConserto3 As Double, dDifValorConserto3 As Double, dDifQuantOutros3 As Double, dDifValorOutros3 As Double, dDifQuantBenef3 As Double, dDifValorBenef3 As Double
Dim dDifQuant As Double, dDifValor As Double

On Error GoTo Erro_Jones_Acerta_SaldosEst1

    'le sldMesEst1 2001
    sSelectSldMesEst1_2001 = "SELECT QuantInicialConsig3, ValorInicialConsig3, QuantInicialDemo3, ValorInicialDemo3, QuantInicialConserto3, ValorInicialConserto3, QuantInicialOutros3, ValorInicialOutros3, QuantInicialBenef3, ValorInicialBenef3"
    
    For iMes = 1 To 12
    
        sSelectSldMesEst1_2001 = sSelectSldMesEst1_2001 & ", SaldoQuantConsig3" & CStr(iMes) & ", SaldoValorConsig3" & CStr(iMes) & ", SaldoQuantDemo3" & CStr(iMes) & ", SaldoValorDemo3" & CStr(iMes) & ", SaldoQuantConserto3" & CStr(iMes) & ", SaldoValorConserto3" & CStr(iMes) & ", SaldoQuantOutros3" & CStr(iMes) & ", SaldoValorOutros3" & CStr(iMes) & ", SaldoQuantBenef3" & CStr(iMes) & ", SaldoValorBenef3" & CStr(iMes)
    
    Next
    
    sSelectSldMesEst1_2001 = sSelectSldMesEst1_2001 & " FROM SldMesEst1 WHERE FilialEmpresa = ? AND Produto = ? AND Ano = 2001"
    With tSldMesEst1_2001
        lErro = Comando_Executar(alComando(6), sSelectSldMesEst1_2001, .dQuantInicialConsig3, .dValorInicialConsig3, .dQuantInicialDemo3, .dValorInicialDemo3, .dQuantInicialConserto3, .dValorInicialConserto3, .dQuantInicialOutros3, .dValorInicialOutros3, .dQuantInicialBenef3, .dValorInicialBenef3, _
            .adSaldoQuantConsig3(1), .adSaldoValorConsig3(1), .adSaldoQuantDemo3(1), .adSaldoValorDemo3(1), .adSaldoQuantConserto3(1), .adSaldoValorConserto3(1), .adSaldoQuantOutros3(1), .adSaldoValorOutros3(1), .adSaldoQuantBenef3(1), .adSaldoValorBenef3(1), _
            .adSaldoQuantConsig3(2), .adSaldoValorConsig3(2), .adSaldoQuantDemo3(2), .adSaldoValorDemo3(2), .adSaldoQuantConserto3(2), .adSaldoValorConserto3(2), .adSaldoQuantOutros3(2), .adSaldoValorOutros3(2), .adSaldoQuantBenef3(2), .adSaldoValorBenef3(2), _
            .adSaldoQuantConsig3(3), .adSaldoValorConsig3(3), .adSaldoQuantDemo3(3), .adSaldoValorDemo3(3), .adSaldoQuantConserto3(3), .adSaldoValorConserto3(3), .adSaldoQuantOutros3(3), .adSaldoValorOutros3(3), .adSaldoQuantBenef3(3), .adSaldoValorBenef3(3), _
            .adSaldoQuantConsig3(4), .adSaldoValorConsig3(4), .adSaldoQuantDemo3(4), .adSaldoValorDemo3(4), .adSaldoQuantConserto3(4), .adSaldoValorConserto3(4), .adSaldoQuantOutros3(4), .adSaldoValorOutros3(4), .adSaldoQuantBenef3(4), .adSaldoValorBenef3(4), _
            .adSaldoQuantConsig3(5), .adSaldoValorConsig3(5), .adSaldoQuantDemo3(5), .adSaldoValorDemo3(5), .adSaldoQuantConserto3(5), .adSaldoValorConserto3(5), .adSaldoQuantOutros3(5), .adSaldoValorOutros3(5), .adSaldoQuantBenef3(5), .adSaldoValorBenef3(5), _
            .adSaldoQuantConsig3(6), .adSaldoValorConsig3(6), .adSaldoQuantDemo3(6), .adSaldoValorDemo3(6), .adSaldoQuantConserto3(6), .adSaldoValorConserto3(6), .adSaldoQuantOutros3(6), .adSaldoValorOutros3(6), .adSaldoQuantBenef3(6), .adSaldoValorBenef3(6), _
            .adSaldoQuantConsig3(7), .adSaldoValorConsig3(7), .adSaldoQuantDemo3(7), .adSaldoValorDemo3(7), .adSaldoQuantConserto3(7), .adSaldoValorConserto3(7), .adSaldoQuantOutros3(7), .adSaldoValorOutros3(7), .adSaldoQuantBenef3(7), .adSaldoValorBenef3(7), _
            .adSaldoQuantConsig3(8), .adSaldoValorConsig3(8), .adSaldoQuantDemo3(8), .adSaldoValorDemo3(8), .adSaldoQuantConserto3(8), .adSaldoValorConserto3(8), .adSaldoQuantOutros3(8), .adSaldoValorOutros3(8), .adSaldoQuantBenef3(8), .adSaldoValorBenef3(8), _
            .adSaldoQuantConsig3(9), .adSaldoValorConsig3(9), .adSaldoQuantDemo3(9), .adSaldoValorDemo3(9), .adSaldoQuantConserto3(9), .adSaldoValorConserto3(9), .adSaldoQuantOutros3(9), .adSaldoValorOutros3(9), .adSaldoQuantBenef3(9), .adSaldoValorBenef3(9), _
            .adSaldoQuantConsig3(10), .adSaldoValorConsig3(10), .adSaldoQuantDemo3(10), .adSaldoValorDemo3(10), .adSaldoQuantConserto3(10), .adSaldoValorConserto3(10), .adSaldoQuantOutros3(10), .adSaldoValorOutros3(10), .adSaldoQuantBenef3(10), .adSaldoValorBenef3(10), _
            .adSaldoQuantConsig3(11), .adSaldoValorConsig3(11), .adSaldoQuantDemo3(11), .adSaldoValorDemo3(11), .adSaldoQuantConserto3(11), .adSaldoValorConserto3(11), .adSaldoQuantOutros3(11), .adSaldoValorOutros3(11), .adSaldoQuantBenef3(11), .adSaldoValorBenef3(11), _
            .adSaldoQuantConsig3(12), .adSaldoValorConsig3(12), .adSaldoQuantDemo3(12), .adSaldoValorDemo3(12), .adSaldoQuantConserto3(12), .adSaldoValorConserto3(12), .adSaldoQuantOutros3(12), .adSaldoValorOutros3(12), .adSaldoQuantBenef3(12), .adSaldoValorBenef3(12), _
            iFilialEmpresa, sProduto)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 81913
    
    lErro = Comando_BuscarProximo(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81914
    
    If lErro = AD_SQL_SUCESSO Then
    
        'le 2002 do produto filial
        With tSldMesEst1_2002
            lErro = Comando_Executar(alComando(7), "SELECT QuantInicialConsig3, ValorInicialConsig3, QuantInicialDemo3, ValorInicialDemo3, QuantInicialConserto3, ValorInicialConserto3, QuantInicialOutros3, ValorInicialOutros3, QuantInicialBenef3, ValorInicialBenef3 FROM SldMesEst1 WHERE FilialEmpresa = ? AND Produto = ? AND Ano = 2002", _
                .dQuantInicialConsig3, .dValorInicialConsig3, .dQuantInicialDemo3, .dValorInicialDemo3, .dQuantInicialConserto3, .dValorInicialConserto3, .dQuantInicialOutros3, .dValorInicialOutros3, .dQuantInicialBenef3, .dValorInicialBenef3, iFilialEmpresa, sProduto)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 81911
        
        lErro = Comando_BuscarProximo(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81912
        
        If lErro = AD_SQL_SUCESSO Then
    
            With tSldMesEst1_2001
                
                dQuantConsig3 = .dQuantInicialConsig3 + .adSaldoQuantConsig3(1) + .adSaldoQuantConsig3(2) + .adSaldoQuantConsig3(3) + .adSaldoQuantConsig3(4) + .adSaldoQuantConsig3(5) + .adSaldoQuantConsig3(6) + .adSaldoQuantConsig3(7) + .adSaldoQuantConsig3(8) + .adSaldoQuantConsig3(9) + .adSaldoQuantConsig3(10) + .adSaldoQuantConsig3(11) + .adSaldoQuantConsig3(12)
                dValorConsig3 = .dValorInicialConsig3 + .adSaldoValorConsig3(1) + .adSaldoValorConsig3(2) + .adSaldoValorConsig3(3) + .adSaldoValorConsig3(4) + .adSaldoValorConsig3(5) + .adSaldoValorConsig3(6) + .adSaldoValorConsig3(7) + .adSaldoValorConsig3(8) + .adSaldoValorConsig3(9) + .adSaldoValorConsig3(10) + .adSaldoValorConsig3(11) + .adSaldoValorConsig3(12)
                dQuantDemo3 = .dQuantInicialDemo3 + .adSaldoQuantDemo3(1) + .adSaldoQuantDemo3(2) + .adSaldoQuantDemo3(3) + .adSaldoQuantDemo3(4) + .adSaldoQuantDemo3(5) + .adSaldoQuantDemo3(6) + .adSaldoQuantDemo3(7) + .adSaldoQuantDemo3(8) + .adSaldoQuantDemo3(9) + .adSaldoQuantDemo3(10) + .adSaldoQuantDemo3(11) + .adSaldoQuantDemo3(12)
                dValorDemo3 = .dValorInicialDemo3 + .adSaldoValorDemo3(1) + .adSaldoValorDemo3(2) + .adSaldoValorDemo3(3) + .adSaldoValorDemo3(4) + .adSaldoValorDemo3(5) + .adSaldoValorDemo3(6) + .adSaldoValorDemo3(7) + .adSaldoValorDemo3(8) + .adSaldoValorDemo3(9) + .adSaldoValorDemo3(10) + .adSaldoValorDemo3(11) + .adSaldoValorDemo3(12)
                dQuantConserto3 = .dQuantInicialConserto3 + .adSaldoQuantConserto3(1) + .adSaldoQuantConserto3(2) + .adSaldoQuantConserto3(3) + .adSaldoQuantConserto3(4) + .adSaldoQuantConserto3(5) + .adSaldoQuantConserto3(6) + .adSaldoQuantConserto3(7) + .adSaldoQuantConserto3(8) + .adSaldoQuantConserto3(9) + .adSaldoQuantConserto3(10) + .adSaldoQuantConserto3(11) + .adSaldoQuantConserto3(12)
                dValorConserto3 = .dValorInicialConserto3 + .adSaldoValorConserto3(1) + .adSaldoValorConserto3(2) + .adSaldoValorConserto3(3) + .adSaldoValorConserto3(4) + .adSaldoValorConserto3(5) + .adSaldoValorConserto3(6) + .adSaldoValorConserto3(7) + .adSaldoValorConserto3(8) + .adSaldoValorConserto3(9) + .adSaldoValorConserto3(10) + .adSaldoValorConserto3(11) + .adSaldoValorConserto3(12)
                dQuantOutros3 = .dQuantInicialOutros3 + .adSaldoQuantOutros3(1) + .adSaldoQuantOutros3(2) + .adSaldoQuantOutros3(3) + .adSaldoQuantOutros3(4) + .adSaldoQuantOutros3(5) + .adSaldoQuantOutros3(6) + .adSaldoQuantOutros3(7) + .adSaldoQuantOutros3(8) + .adSaldoQuantOutros3(9) + .adSaldoQuantOutros3(10) + .adSaldoQuantOutros3(11) + .adSaldoQuantOutros3(12)
                dValorOutros3 = .dValorInicialOutros3 + .adSaldoValorOutros3(1) + .adSaldoValorOutros3(2) + .adSaldoValorOutros3(3) + .adSaldoValorOutros3(4) + .adSaldoValorOutros3(5) + .adSaldoValorOutros3(6) + .adSaldoValorOutros3(7) + .adSaldoValorOutros3(8) + .adSaldoValorOutros3(9) + .adSaldoValorOutros3(10) + .adSaldoValorOutros3(11) + .adSaldoValorOutros3(12)
                dQuantBenef3 = .dQuantInicialBenef3 + .adSaldoQuantBenef3(1) + .adSaldoQuantBenef3(2) + .adSaldoQuantBenef3(3) + .adSaldoQuantBenef3(4) + .adSaldoQuantBenef3(5) + .adSaldoQuantBenef3(6) + .adSaldoQuantBenef3(7) + .adSaldoQuantBenef3(8) + .adSaldoQuantBenef3(9) + .adSaldoQuantBenef3(10) + .adSaldoQuantBenef3(11) + .adSaldoQuantBenef3(12)
                dValorBenef3 = .dValorInicialBenef3 + .adSaldoValorBenef3(1) + .adSaldoValorBenef3(2) + .adSaldoValorBenef3(3) + .adSaldoValorBenef3(4) + .adSaldoValorBenef3(5) + .adSaldoValorBenef3(6) + .adSaldoValorBenef3(7) + .adSaldoValorBenef3(8) + .adSaldoValorBenef3(9) + .adSaldoValorBenef3(10) + .adSaldoValorBenef3(11) + .adSaldoValorBenef3(12)
            
            End With
            
            With tSldMesEst1_2002
            
                dDifQuantConsig3 = .dQuantInicialConsig3 - dQuantConsig3
                dDifValorConsig3 = .dValorInicialConsig3 - dValorConsig3
                dDifQuantDemo3 = .dQuantInicialDemo3 - dQuantDemo3
                dDifValorDemo3 = .dValorInicialDemo3 - dValorDemo3
                dDifQuantConserto3 = .dQuantInicialConserto3 - dQuantConserto3
                dDifValorConserto3 = .dValorInicialConserto3 - dValorConserto3
                dDifQuantOutros3 = .dQuantInicialOutros3 - dQuantOutros3
                dDifValorOutros3 = .dValorInicialOutros3 - dValorOutros3
                dDifQuantBenef3 = .dQuantInicialBenef3 - dQuantBenef3
                dDifValorBenef3 = .dValorInicialBenef3 - dValorBenef3
            
            End With
            
            'se houver diferenca
            If Abs(dDifQuantConsig3) > QTDE_ESTOQUE_DELTA Or _
                Abs(dDifValorConsig3) > DELTA_VALORMONETARIO Or _
                Abs(dDifQuantDemo3) > QTDE_ESTOQUE_DELTA Or _
                Abs(dDifValorDemo3) > DELTA_VALORMONETARIO Or _
                Abs(dDifQuantConserto3) > QTDE_ESTOQUE_DELTA Or _
                Abs(dDifValorConserto3) > DELTA_VALORMONETARIO Or _
                Abs(dDifQuantOutros3) > QTDE_ESTOQUE_DELTA Or _
                Abs(dDifValorOutros3) > DELTA_VALORMONETARIO Or _
                Abs(dDifQuantBenef3) > QTDE_ESTOQUE_DELTA Or _
                Abs(dDifValorBenef3) > DELTA_VALORMONETARIO Then

                'modificar quant/valor inicial em sldMesEst1
                lErro = Comando_Executar(alComando(8), "UPDATE SldMesEst1 SET QuantInicialConsig3 = QuantInicialConsig3 + ?, ValorInicialConsig3 = ValorInicialConsig3 + ?, QuantInicialDemo3 = QuantInicialDemo3 + ?, ValorInicialDemo3 = ValorInicialDemo3 + ?, QuantInicialConserto3 = QuantInicialConserto3 + ?, ValorInicialConserto3 = ValorInicialConserto3 + ?, QuantInicialOutros3 = QuantInicialOutros3 + ?, ValorInicialOutros3 = ValorInicialOutros3 + ?, QuantInicialBenef3 = QuantInicialBenef3 + ?, ValorInicialBenef3 = ValorInicialBenef3 + ? WHERE FilialEmpresa = ? AND Produto = ? AND Ano = 2001", _
                    dDifQuantConsig3, dDifValorConsig3, dDifQuantDemo3, dDifValorDemo3, dDifQuantConserto3, dDifValorConserto3, dDifQuantOutros3, dDifValorOutros3, dDifQuantBenef3, dDifValorBenef3, iFilialEmpresa, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81915
                
                'modificar quant/valor inicial em sldmesestalm1 no almoxarifado padrao
                lErro = Comando_Executar(alComando(9), "UPDATE SldMesEstAlm1 SET QuantInicialConsig3 = QuantInicialConsig3 + ?, ValorInicialConsig3 = ValorInicialConsig3 + ?, QuantInicialDemo3 = QuantInicialDemo3 + ?, ValorInicialDemo3 = ValorInicialDemo3 + ?, QuantInicialConserto3 = QuantInicialConserto3 + ?, ValorInicialConserto3 = ValorInicialConserto3 + ?, QuantInicialOutros3 = QuantInicialOutros3 + ?, ValorInicialOutros3 = ValorInicialOutros3 + ?, QuantInicialBenef3 = QuantInicialBenef3 + ?, ValorInicialBenef3 = ValorInicialBenef3 + ? WHERE Almoxarifado = ? AND Produto = ? AND Ano = 2001", _
                    dDifQuantConsig3, dDifValorConsig3, dDifQuantDemo3, dDifValorDemo3, dDifQuantConserto3, dDifValorConserto3, dDifQuantOutros3, dDifValorOutros3, dDifQuantBenef3, dDifValorBenef3, iAlmoxarifado, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81916
                
                dDifQuant = dDifQuantConsig3 + dDifQuantDemo3 + dDifQuantConserto3 + dDifQuantOutros3 + dDifQuantBenef3
                dDifValor = dDifValorConsig3 + dDifValorDemo3 + dDifValorConserto3 + dDifValorOutros3 + dDifValorBenef3
                
                lErro = Comando_Executar(alComando(10), "UPDATE SldMesEst SET QuantInicial = QuantInicial + ?, QuantInicialCusto = QuantInicialCusto + ?, ValorInicial = ValorInicial + ?, ValorInicialCusto = ValorInicialCusto + ? WHERE FilialEmpresa = ? AND Produto = ? AND Ano = 2001", _
                    dDifQuant, dDifQuant, dDifValor, dDifValor, iFilialEmpresa, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81916

                lErro = Comando_Executar(alComando(11), "UPDATE SldMesEstAlm SET QuantInicial = QuantInicial + ?, QuantInicialCusto = QuantInicialCusto + ?, ValorInicial = ValorInicial + ?, ValorInicialCusto = ValorInicialCusto + ? WHERE Almoxarifado = ? AND Produto = ? AND Ano = 2001", _
                    dDifQuant, dDifQuant, dDifValor, dDifValor, iAlmoxarifado, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81916

                lErro = Comando_Executar(alComando(17), "UPDATE EstoqueProduto SET QuantInicialConsig3 = QuantInicialConsig3 + ?, ValorInicialConsig3 = ValorInicialConsig3 + ?, QuantInicialDemo3 = QuantInicialDemo3 + ?, ValorInicialDemo3 = ValorInicialDemo3 + ?, QuantInicialConserto3 = QuantInicialConserto3 + ?, ValorInicialConserto3 = ValorInicialConserto3 + ?, QuantInicialOutras3 = QuantInicialOutras3 + ?, ValorInicialOutras3 = ValorInicialOutras3 + ?, QuantInicialBenef3 = QuantInicialBenef3 + ?, ValorInicialBenef3 = ValorInicialBenef3 + ?, DataUltAtualizacao = ?, Exportado = ?  WHERE Almoxarifado = ? AND Produto = ?", _
                    dDifQuantConsig3, dDifValorConsig3, dDifQuantDemo3, dDifValorDemo3, dDifQuantConserto3, dDifValorConserto3, dDifQuantOutros3, dDifValorOutros3, dDifQuantBenef3, dDifValorBenef3, Date, 0, iAlmoxarifado, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81916
            End If
        
        End If
        
    End If

    Jones_Acerta_SaldosEst1 = SUCESSO
     
    Exit Function
    
Erro_Jones_Acerta_SaldosEst1:

    Jones_Acerta_SaldosEst1 = gErr
     
    Select Case gErr
          
        Case 81911 To 81916
            Call Rotina_Erro(vbOKOnly, "ERRO_JONES", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151665)
     
    End Select
     
    Exit Function

End Function
        
Private Function Jones_Acerta_SaldosEst(alComando() As Long, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal iAlmoxarifado As Integer) As Long
'Acerta Sldmesest e Sldmesestalm em 2001 para chegar ao saldo inicial desejado em 2002

Dim lErro As Long, dQuant2001 As Double, dValor2001 As Double, dQuant2002 As Double, dValor2002 As Double
Dim dDifQuant As Double, dDifValor As Double

On Error GoTo Erro_Jones_Acerta_SaldosEst

    'le sldMesEst 2001
    lErro = Comando_Executar(alComando(12), "SELECT QuantInicial+QuantEnt1-QuantSai1+QuantEnt2-QuantSai2+QuantEnt3-QuantSai3+QuantEnt4-QuantSai4+QuantEnt5-QuantSai5+QuantEnt6-QuantSai6+QuantEnt7-QuantSai7+QuantEnt8-QuantSai8+QuantEnt9-QuantSai9+QuantEnt10-QuantSai10+QuantEnt11-QuantSai11+QuantEnt12-QuantSai12, ValorInicial+ValorEnt1-ValorSai1+ValorEnt2-ValorSai2+ValorEnt3-ValorSai3+ValorEnt4-ValorSai4+ValorEnt5-ValorSai5+ValorEnt6-ValorSai6+ValorEnt7-ValorSai7+ValorEnt8-ValorSai8+ValorEnt9-ValorSai9+ValorEnt10-ValorSai10+ValorEnt11-ValorSai11+ValorEnt12-ValorSai12 FROM SldMesEst WHERE FilialEmpresa = ? AND Produto = ? AND Ano = 2001", _
        dQuant2001, dValor2001, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 81917

    lErro = Comando_BuscarProximo(alComando(12))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81918
    
    If lErro = AD_SQL_SUCESSO Then
    
        'le 2002 do produto filial
        lErro = Comando_Executar(alComando(13), "SELECT QuantInicial, ValorInicial FROM SldMesEst WHERE FilialEmpresa = ? AND Produto = ? AND Ano = 2002", _
            dQuant2002, dValor2002, iFilialEmpresa, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 81919
    
        lErro = Comando_BuscarProximo(alComando(13))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81920
    
        If lErro = AD_SQL_SUCESSO Then
        
            dDifQuant = dQuant2002 - dQuant2001
            dDifValor = dValor2002 - dValor2001
    
            If Abs(dDifQuant) > QTDE_ESTOQUE_DELTA Or Abs(dDifValor) > DELTA_VALORMONETARIO Then
                
                lErro = Comando_Executar(alComando(14), "UPDATE SldMesEst SET QuantInicial = QuantInicial + ?, ValorInicial = ValorInicial + ? WHERE FilialEmpresa = ? AND Produto = ? AND Ano = 2001", _
                    dDifQuant, dDifValor, iFilialEmpresa, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81921

                lErro = Comando_Executar(alComando(15), "UPDATE SldMesEstAlm SET QuantInicial = QuantInicial + ?, ValorInicial = ValorInicial + ? WHERE Almoxarifado = ? AND Produto = ? AND Ano = 2001", _
                    dDifQuant, dDifValor, iAlmoxarifado, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81922

                lErro = Comando_Executar(alComando(18), "UPDATE EstoqueProduto SET QuantidadeInicial = QuantidadeInicial + ?, SaldoInicial = SaldoInicial + ?, DataUltAtualizacao = ?, Exportado = ?  WHERE Almoxarifado = ? AND Produto = ?", _
                    dDifQuant, dDifValor, Date, 0, iAlmoxarifado, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81922

            End If
            
        End If
    
    End If
    
    Jones_Acerta_SaldosEst = SUCESSO
     
    Exit Function
    
Erro_Jones_Acerta_SaldosEst:

    Jones_Acerta_SaldosEst = gErr
     
    Select Case gErr
          
        Case 81917 To 81922
            Call Rotina_Erro(vbOKOnly, "ERRO_JONES", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151666)
     
    End Select
     
    Exit Function

End Function

Private Function Jones_Acerta_SaldosEstCusto(alComando() As Long, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal iAlmoxarifado As Integer) As Long
'Acerta Sldmesest e Sldmesestalm em 2001 para chegar ao saldo inicial desejado em 2002

Dim lErro As Long, dQuant2001 As Double, dValor2001 As Double, dQuant2002 As Double, dValor2002 As Double
Dim dDifQuant As Double, dDifValor As Double, sAuxQ As String, sAuxV As String, sSelect As String
Dim dQuant2001Atual As Double, dValor2001Atual As Double

On Error GoTo Erro_Jones_Acerta_SaldosEstCusto

    sAuxQ = "-(SaldoQuantConsig31 + SaldoQuantConsig32 + SaldoQuantConsig33 + SaldoQuantConsig34 + SaldoQuantConsig35 + SaldoQuantConsig36 + SaldoQuantConsig37 + SaldoQuantConsig38 + SaldoQuantConsig39 + SaldoQuantConsig310 + SaldoQuantConsig311 + SaldoQuantConsig312)"
    sAuxQ = sAuxQ & "-(SaldoQuantBenef31 + SaldoQuantBenef32 + SaldoQuantBenef33 + SaldoQuantBenef34 + SaldoQuantBenef35 + SaldoQuantBenef36 + SaldoQuantBenef37 + SaldoQuantBenef38 + SaldoQuantBenef39 + SaldoQuantBenef310 + SaldoQuantBenef311 + SaldoQuantBenef312)"
    sAuxQ = sAuxQ & "-(SaldoQuantOutros31 + SaldoQuantOutros32 + SaldoQuantOutros33 + SaldoQuantOutros34 + SaldoQuantOutros35 + SaldoQuantOutros36 + SaldoQuantOutros37 + SaldoQuantOutros38 + SaldoQuantOutros39 + SaldoQuantOutros310 + SaldoQuantOutros311 + SaldoQuantOutros312)"
    sAuxQ = sAuxQ & "-(SaldoQuantDemo31 + SaldoQuantDemo32 + SaldoQuantDemo33 + SaldoQuantDemo34 + SaldoQuantDemo35 + SaldoQuantDemo36 + SaldoQuantDemo37 + SaldoQuantDemo38 + SaldoQuantDemo39 + SaldoQuantDemo310 + SaldoQuantDemo311 + SaldoQuantDemo312)"

    sAuxV = "-(SaldoValorConsig31 + SaldoValorConsig32 + SaldoValorConsig33 + SaldoValorConsig34 + SaldoValorConsig35 + SaldoValorConsig36 + SaldoValorConsig37 + SaldoValorConsig38 + SaldoValorConsig39 + SaldoValorConsig310 + SaldoValorConsig311 + SaldoValorConsig312)"
    sAuxV = sAuxV & "-(SaldoValorBenef31 + SaldoValorBenef32 + SaldoValorBenef33 + SaldoValorBenef34 + SaldoValorBenef35 + SaldoValorBenef36 + SaldoValorBenef37 + SaldoValorBenef38 + SaldoValorBenef39 + SaldoValorBenef310 + SaldoValorBenef311 + SaldoValorBenef312)"
    sAuxV = sAuxV & "-(SaldoValorOutros31 + SaldoValorOutros32 + SaldoValorOutros33 + SaldoValorOutros34 + SaldoValorOutros35 + SaldoValorOutros36 + SaldoValorOutros37 + SaldoValorOutros38 + SaldoValorOutros39 + SaldoValorOutros310 + SaldoValorOutros311 + SaldoValorOutros312)"
    sAuxV = sAuxV & "-(SaldoValorDemo31 + SaldoValorDemo32 + SaldoValorDemo33 + SaldoValorDemo34 + SaldoValorDemo35 + SaldoValorDemo36 + SaldoValorDemo37 + SaldoValorDemo38 + SaldoValorDemo39 + SaldoValorDemo310 + SaldoValorDemo311 + SaldoValorDemo312)"

    sSelect = "SELECT QuantInicialCusto, ValorInicialCusto, QuantEnt1-QuantSai1+QuantEnt2-QuantSai2+QuantEnt3-QuantSai3+QuantEnt4-QuantSai4+QuantEnt5-QuantSai5+QuantEnt6-QuantSai6+QuantEnt7-QuantSai7+QuantEnt8-QuantSai8+QuantEnt9-QuantSai9+QuantEnt10-QuantSai10+QuantEnt11-QuantSai11+QuantEnt12-QuantSai12" & sAuxQ & ", " & _
        "ValorEnt1-ValorSai1+ValorEnt2-ValorSai2+ValorEnt3-ValorSai3+ValorEnt4-ValorSai4+ValorEnt5-ValorSai5+ValorEnt6-ValorSai6+ValorEnt7-ValorSai7+ValorEnt8-ValorSai8+ValorEnt9-ValorSai9+ValorEnt10-ValorSai10+ValorEnt11-ValorSai11+ValorEnt12-ValorSai12" & sAuxV & _
        " FROM SldMesEst s, sldmesest1 s1 WHERE s.FilialEmpresa = s1.filialempresa AND s.Produto = s1.produto AND s.Ano = s1.ano and s.FilialEmpresa = ? AND s.Produto = ? AND s.Ano = 2001"

    'le sldMesEst 2001
    lErro = Comando_Executar(alComando(12), sSelect, dQuant2001Atual, dValor2001Atual, dQuant2001, dValor2001, iFilialEmpresa, sProduto)
    
    If lErro <> AD_SQL_SUCESSO Then gError 81917

    lErro = Comando_BuscarProximo(alComando(12))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81918
    
    If lErro = AD_SQL_SUCESSO Then
    
        'le 2002 do produto filial
        lErro = Comando_Executar(alComando(13), "SELECT QuantInicial - (QuantInicialConsig3+QuantInicialDemo3+QuantInicialConserto3+QuantInicialOutros3+QuantInicialBenef3), ValorInicial - (ValorInicialConsig3+ValorInicialDemo3+ValorInicialConserto3+ValorInicialOutros3+ValorInicialBenef3) FROM SldMesEst s, SldMesest1 s1 WHERE s.FilialEmpresa = ? AND s.Produto = ? AND s.Ano = 2002 AND s.FilialEmpresa = s1.FilialEmpresa AND s.Produto = s1.Produto AND s.Ano = s1.ano", _
            dQuant2002, dValor2002, iFilialEmpresa, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 81919
    
        lErro = Comando_BuscarProximo(alComando(13))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81920
    
        If lErro = AD_SQL_SUCESSO Then
        
            dDifQuant = dQuant2002 - (dQuant2001 + dQuant2001Atual)
            dDifValor = dValor2002 - (dValor2001 + dValor2001Atual)
    
            If Abs(dDifQuant) > QTDE_ESTOQUE_DELTA Or Abs(dDifValor) > DELTA_VALORMONETARIO Then
                
                lErro = Comando_Executar(alComando(14), "UPDATE SldMesEst SET QuantInicialCusto = QuantInicialCusto + ?, ValorInicialCusto = ValorInicialCusto + ? WHERE FilialEmpresa = ? AND Produto = ? AND Ano = 2001", _
                    dDifQuant, dDifValor, iFilialEmpresa, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81921

                lErro = Comando_Executar(alComando(15), "UPDATE SldMesEstAlm SET QuantInicialCusto = QuantInicialCusto + ?, ValorInicialCusto = ValorInicialCusto + ? WHERE Almoxarifado = ? AND Produto = ? AND Ano = 2001", _
                    dDifQuant, dDifValor, iAlmoxarifado, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 81922

'                lErro = Comando_Executar(alComando(18), "UPDATE EstoqueProduto SET QuantidadeInicial = QuantidadeInicial + ?, SaldoInicial = SaldoInicial + ? WHERE Almoxarifado = ? AND Produto = ?", _
'                    dDifQuant, dDifValor, iAlmoxarifado, sProduto)
'                If lErro <> AD_SQL_SUCESSO Then gError 81922

            End If
            
        End If
    
    End If
    
    Jones_Acerta_SaldosEstCusto = SUCESSO
     
    Exit Function
    
Erro_Jones_Acerta_SaldosEstCusto:

    Jones_Acerta_SaldosEstCusto = gErr
     
    Select Case gErr
          
        Case 81917 To 81922
            Call Rotina_Erro(vbOKOnly, "ERRO_JONES", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151667)
     
    End Select
     
    Exit Function

End Function

Function MovEstoque_Trata_Grade(ByVal objMovEst As ClassMovEstoque) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim objItemMovEst As ClassItemMovEstoque
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemMovEstNovo As ClassItemMovEstoque
Dim colItemMovEstTemp As New Collection
Dim iIndice As Integer

On Error GoTo Erro_MovEstoque_Trata_Grade

    For Each objItemMovEst In objMovEst.colItens

        iIndice = iIndice + 1

        If objItemMovEst.colItensRomaneioGrade.Count > 0 Then
            
            For Each objItemRomaneioGrade In objItemMovEst.colItensRomaneioGrade
            
                If objItemRomaneioGrade.dQuantidade > 0 Then
                
                    Set objItemMovEstNovo = New ClassItemMovEstoque
                    
                    objItemMovEstNovo.lNumIntDoc = objItemRomaneioGrade.lNumIntDoc
                    objItemMovEstNovo.dQuantidade = objItemRomaneioGrade.dQuantidade
                    objItemMovEstNovo.dtData = objItemMovEst.dtData
                    objItemMovEstNovo.dtDataInicioProducao = objItemMovEst.dtDataInicioProducao
                    objItemMovEstNovo.dtDataRegistro = objItemMovEst.dtDataRegistro
                    objItemMovEstNovo.dtHora = objItemMovEst.dtHora
                    objItemMovEstNovo.iAlmoxarifado = objItemRomaneioGrade.iAlmoxarifado
                    objItemMovEstNovo.iAtualizaSoLote = objItemMovEst.iAtualizaSoLote
                    objItemMovEstNovo.iBenef = objItemMovEst.iBenef
                    objItemMovEstNovo.iClasseUM = objItemMovEst.iClasseUM
                    objItemMovEstNovo.iFilialEmpresa = objItemMovEst.iFilialEmpresa
                    objItemMovEstNovo.iTipoMov = objItemMovEst.iTipoMov
                    objItemMovEstNovo.lCliente = objItemMovEst.lCliente
                    objItemMovEstNovo.lCodigo = objItemMovEst.lCodigo
                    objItemMovEstNovo.lFornecedor = objItemMovEst.lFornecedor
                    objItemMovEstNovo.lHorasMaquina = objItemMovEst.lHorasMaquina
                    objItemMovEstNovo.sAlmoxarifadoNomeRed = objItemRomaneioGrade.sAlmoxarifado
                    objItemMovEstNovo.sCcl = objItemMovEst.sCcl
                    objItemMovEstNovo.sContaContabilAplic = objItemMovEst.sContaContabilAplic
                    objItemMovEstNovo.sContaContabilEst = objItemMovEst.sContaContabilEst
                    objItemMovEstNovo.sDocOrigem = objItemMovEst.sDocOrigem
                    objItemMovEstNovo.sOPCodigo = objItemRomaneioGrade.sCodOP
                    objItemMovEstNovo.sProduto = objItemRomaneioGrade.sProduto
                    objItemMovEstNovo.sProdutoOP = objItemRomaneioGrade.sProdOP
                    objItemMovEstNovo.sSiglaUM = objItemRomaneioGrade.sUMEstoque
                    objItemMovEstNovo.sSiglaUMEst = objItemRomaneioGrade.sUMEstoque
                    objItemMovEstNovo.iIndicePaiGrade = iIndice
                    
                    colItemMovEstTemp.Add objItemMovEstNovo
                    
                End If
            
            Next

        End If

    Next
    
    For Each objItemMovEst In colItemMovEstTemp
    
        objMovEst.colItens.Add1 objItemMovEst
    
    Next
    
    MovEstoque_Trata_Grade = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Trata_Grade:

    MovEstoque_Trata_Grade = gErr
    
    Select Case gErr
    
        Case 20416, 20417, 126665
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151668)
    
    End Select
    
    Exit Function

End Function

Function MovEstoque_Trata_Grade1(ByVal objMovEst As ClassMovEstoque, objItemMovEstPai As ClassItemMovEstoque, iIndicePai As Integer) As Long
'coloca o lNumIntDoc do Pai de Grade nos filhos

Dim objItemMovEst As ClassItemMovEstoque

On Error GoTo Erro_MovEstoque_Trata_Grade1

    For Each objItemMovEst In objMovEst.colItens

        If objItemMovEst.iIndicePaiGrade = iIndicePai Then
        
            objItemMovEst.lNumIntDocGrade = objItemMovEstPai.lNumIntDoc
            
        End If
        
    Next

    MovEstoque_Trata_Grade1 = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Trata_Grade1:

    MovEstoque_Trata_Grade1 = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151669)
    
    End Select
    
    Exit Function

End Function

Function MovEst_ObtemCustoMovOrig1(ByVal objItemMovEst As ClassItemMovEstoque, dCustoUnitarioMovOrig As Double) As Long
'Retorna em dCustoUnitarioMovOrig o custo unitario (correspondente a unidades de estoque) do movimento de estoque de saida associado ao que origina objItemMovEst
'Esta funcao é especifica para o tratamento de notas fiscais de transferencia de material da propria empresa entre filiais onde o custo de entrada deve ser igual ao de saida em vez do valor da nfiscal

Dim lErro As Long, alComando(1 To 2) As Long, iIndice As Integer
Dim dtDataEmissao As Date, sSerie As String, lNumNotaFiscal As Long
Dim dCusto As Double
Dim dQuantidade As Double
Dim sSiglaUM As String
Dim sSiglaUMEstoque As String
Dim iClasseUM As Integer
Dim dFator As Double
Dim lNumIntNotaOriginal As Long

On Error GoTo Erro_MovEst_ObtemCustoMovOrig1
        
    dCustoUnitarioMovOrig = 0
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 133363
    Next
    
    If objItemMovEst.lNumIntDocOrigem <> 0 And objItemMovEst.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL Then
    
        'obtem dados da nf de entrada
        lErro = Comando_Executar(alComando(1), "SELECT NumIntNotaOriginal FROM NFiscal, ItensNFiscal WHERE ItensNFiscal.NumIntDoc = ? AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF", lNumIntNotaOriginal, objItemMovEst.lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 133364
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133365
        
        If lErro <> AD_SQL_SUCESSO Then gError 133366
    
        sSiglaUM = String(STRING_UM_SIGLA, 0)
        sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
        
        'busca a nf de saida da outra filialempresa
        lErro = Comando_Executar(alComando(2), "SELECT MovimentoEstoque.Custo, MovimentoEstoque.Quantidade, MovimentoEstoque.SiglaUM, ClasseUM, SiglaUMEstoque FROM NFiscal, ItensNFiscal, MovimentoEstoque, Produtos WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.NumIntDoc = ? AND MovimentoEstoque.TipoNumIntDocOrigem = ? AND MovimentoEstoque.NumIntDocOrigem = ItensNFiscal.NumIntDoc AND ItensNFiscal.Produto = Produtos.Codigo AND Produtos.Codigo = ? AND MovimentoEstoque.FilialEmpresa = ?", _
            dCusto, dQuantidade, sSiglaUM, iClasseUM, sSiglaUMEstoque, lNumIntNotaOriginal, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL, objItemMovEst.sProduto, objItemMovEst.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 133367

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133368
        
        If lErro <> AD_SQL_SUCESSO Then gError 133369

        'Converte a unidade de medida da quantidade do Movimento de Estoque para a UM de Estoque
        lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUM, sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 133370
        
        dQuantidade = dQuantidade * dFator
        
        If Abs(dQuantidade) > QTDE_ESTOQUE_DELTA Then
            dCustoUnitarioMovOrig = dCusto / dQuantidade
        Else
            dCustoUnitarioMovOrig = 0
        End If
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    MovEst_ObtemCustoMovOrig1 = SUCESSO
     
    Exit Function
    
Erro_MovEst_ObtemCustoMovOrig1:

    MovEst_ObtemCustoMovOrig1 = gErr
     
    Select Case gErr
          
        Case 133363
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 133364, 133365
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFO_MOVORIG", gErr)
        
        Case 133367, 133368
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFO_MOVORIG3", gErr)
        
        Case 133366, 133369, 133370
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151670)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function Rastreamento_Serie_Gera(ByVal objItemMovEst As ClassItemMovEstoque, ByVal objProduto As ClassProduto, ByVal sSerieIni As String, colRastreamentoSerie As Collection) As Long

Dim lErro As Long
Dim bSaida As Boolean
Dim iEscaninho As Integer
Dim objTipoMovEstoque As New ClassTipoMovEst
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dFator As Double

On Error GoTo Erro_Rastreamento_Serie_Gera

    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemMovEst.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 177271
    
    objItemMovEst.dQuantidadeEst = objItemMovEst.dQuantidade * dFator
    
    lErro = Valor_Inteiro_Critica(Formata_Estoque(objItemMovEst.dQuantidadeEst))
    If lErro <> SUCESSO Then gError 177289
    
    If objItemMovEst.iTipoMov <> 0 Then
        
        lErro = RastreamentoSerie_Obtem_Escaninho(objItemMovEst, objTipoMovEstoque, objEstoqueProduto, iEscaninho, bSaida)
        If lErro <> SUCESSO Then gError 177269
        
        lErro = RastreamentoSerie_Valida_Serie(objItemMovEst, objEstoqueProduto, sSerieIni)
        If lErro <> SUCESSO Then gError 177272
    
        If bSaida Then
            lErro = Rastreamento_Serie_Gera_Saida(objItemMovEst, objProduto, iEscaninho, sSerieIni, colRastreamentoSerie)
        Else
            lErro = Rastreamento_Serie_Gera_Entrada(objItemMovEst, objProduto, iEscaninho, sSerieIni, colRastreamentoSerie)
        End If
        If lErro <> SUCESSO Then gError 177235
        
    End If

    Rastreamento_Serie_Gera = SUCESSO
        
    Exit Function
    
Erro_Rastreamento_Serie_Gera:

    Rastreamento_Serie_Gera = gErr
    
    Select Case gErr
    
        Case 177235, 177269, 177271, 177272, 177289

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177236)

    End Select
    
    Exit Function
    
End Function

Public Function Rastreamento_Serie_Gera_Saida(ByVal objItemMovEst As ClassItemMovEstoque, ByVal objProduto As ClassProduto, ByVal iEscaninho As Integer, ByVal sSerieIni As String, colRastreamentoSerie As Collection) As Long
'Gera na ordem de Data de Entrada descrecente os Movimentos dos Lotes

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim tRastreamentoLoteSaldo As typeRastreamentoLoteSaldo
Dim objRastreamentoMovto As New ClassRastreamentoMovto
Dim dQuantidadeEscaninho As Double
Dim dQuantidadeTotal As Double

On Error GoTo Erro_Rastreamento_Serie_Gera_Saida
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 177229
    Next
    
    'se é um produto com rastreamento por número de série
    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
    
        Set colRastreamentoSerie = New Collection
    
        tRastreamentoLoteSaldo.sProduto = String(STRING_PRODUTO, 0)
        tRastreamentoLoteSaldo.sLote = String(STRING_LOTE_RASTREAMENTO, 0) 'Alterado por Wagner 09/03/2006
        
        'Lê o RastreamentoLoteSaldo, RastreamentoLote em ordem decrescente de Data
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDocLote, QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantEmpenhada, QuantPedida, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantOP, QuantBenef, QuantBenef3, Lote, FilialOP FROM RastreamentoLoteSaldo, RastreamentoLote WHERE RastreamentoLote.NumIntDoc = RastreamentoLoteSaldo.NumIntDocLote AND RastreamentoLoteSaldo.Produto = ? AND Almoxarifado = ? AND Status = ? AND RastreamentoLote.Lote >= ? ORDER BY RastreamentoLote.Lote", tRastreamentoLoteSaldo.lNumIntDocLote, tRastreamentoLoteSaldo.dQuantDispNossa, tRastreamentoLoteSaldo.dQuantReservada, tRastreamentoLoteSaldo.dQuantReservadaConsig, tRastreamentoLoteSaldo.dQuantEmpenhada, _
        tRastreamentoLoteSaldo.dQuantPedida, tRastreamentoLoteSaldo.dQuantRecIndl, tRastreamentoLoteSaldo.dQuantIndOutras, tRastreamentoLoteSaldo.dQuantDefeituosa, tRastreamentoLoteSaldo.dQuantConsig3, tRastreamentoLoteSaldo.dQuantConsig, tRastreamentoLoteSaldo.dQuantDemo3, tRastreamentoLoteSaldo.dQuantDemo, tRastreamentoLoteSaldo.dQuantConserto3, tRastreamentoLoteSaldo.dQuantConserto, tRastreamentoLoteSaldo.dQuantOutras3, tRastreamentoLoteSaldo.dQuantOutras, tRastreamentoLoteSaldo.dQuantOP, tRastreamentoLoteSaldo.dQuantBenef, tRastreamentoLoteSaldo.dQuantBenef3, tRastreamentoLoteSaldo.sLote, tRastreamentoLoteSaldo.iFilialOP, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, RATREAMENTOLOTE_STATUS_ABERTO, sSerieIni)
        If lErro <> AD_SQL_SUCESSO Then gError 177230
               
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 177231
                
        'Se ainda não alcançou a quantidade e ainda existe lote a ser acumulado
        Do While objItemMovEst.dQuantidadeEst > dQuantidadeTotal And lErro = AD_SQL_SUCESSO
        
            Set objRastreamentoMovto = New ClassRastreamentoMovto
            
            'Depende do escaninho
            Select Case iEscaninho
                
                'Acumula uma quantidade Total e a quantidade no Lote-escaninho atual
                Case ESCANINHO_NOSSO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDispNossa
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDispNossa
                Case ESCANINHO_3_EM_CONSIGNACAO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConsig3
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConsig3
                Case ESCANINHO_NOSSO_EM_CONSIGNACAO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConsig
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConsig
                Case ESCANINHO_3_EM_DEMO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDemo3
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDemo3
                Case ESCANINHO_NOSSO_EM_DEMO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDemo
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDemo
                Case ESCANINHO_3_EM_CONSERTO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConserto3
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConserto3
                Case ESCANINHO_NOSSO_EM_CONSERTO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConserto
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConserto
                Case ESCANINHO_3_EM_OUTROS
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantOutras3
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantOutras3
                Case ESCANINHO_NOSSO_EM_OUTROS
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantOutras
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantOutras
                Case ESCANINHO_3_EM_BENEF
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantBenef3
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantBenef3
                Case ESCANINHO_NOSSO_EM_BENEF
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantBenef
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantBenef
                Case ESCANINHO_DEFEITUOSO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDefeituosa
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDefeituosa
                Case ESCANINHO_INDISPONIVEL
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantIndOutras
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantIndOutras
                Case ESCANINHO_RECEBIDO_INDISPONIVEL
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantRecIndl
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantRecIndl
            
            End Select
                    
            'Se a quantidade for maior
            If dQuantidadeTotal > objItemMovEst.dQuantidadeEst Then
                'Pega somente o necessário
                objRastreamentoMovto.dQuantidade = objItemMovEst.dQuantidadeEst - (dQuantidadeTotal - dQuantidadeEscaninho)
            Else
                'Se for menor ou igual --> Pega toda a quantidade do lote
                objRastreamentoMovto.dQuantidade = dQuantidadeEscaninho
            End If
                
            If objRastreamentoMovto.dQuantidade <> 0 Then
            
                'Armazena os dados do Lote
                objRastreamentoMovto.sSiglaUM = objItemMovEst.sSiglaUM
                objRastreamentoMovto.dQuantidade = objRastreamentoMovto.dQuantidade * (objItemMovEst.dQuantidadeEst / objItemMovEst.dQuantidade)
                objRastreamentoMovto.iFilialOP = tRastreamentoLoteSaldo.iFilialOP
                objRastreamentoMovto.sLote = tRastreamentoLoteSaldo.sLote
                objRastreamentoMovto.lNumIntDocLote = tRastreamentoLoteSaldo.lNumIntDocLote
                objRastreamentoMovto.sProduto = objItemMovEst.sProduto
                                       
                'coloca na coleção
                colRastreamentoSerie.Add objRastreamentoMovto
            
            End If
                
            'Busca o próximo Lote
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 177232
                    
        Loop
            
        'Se não tem quantidade suficiente ---> Erro
        If objItemMovEst.dQuantidadeEst > dQuantidadeTotal Then gError 177233

    End If
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Rastreamento_Serie_Gera_Saida = SUCESSO
        
    Exit Function
    
Erro_Rastreamento_Serie_Gera_Saida:

    Rastreamento_Serie_Gera_Saida = gErr
    
    Select Case gErr
            
        Case 177229
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
           
        Case 177230 To 177232
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTESALDO2", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
                
        Case 177233
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_INSUF_RASTREAMENTO", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, iEscaninho)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177234)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function Rastreamento_Serie_Gera_Entrada(ByVal objItemMovEst As ClassItemMovEstoque, ByVal objProduto As ClassProduto, ByVal iEscaninho As Integer, ByVal sSerieIni As String, colRastreamentoSerie As Collection) As Long
'Gera na ordem de Data de Entrada descrecente os Movimentos dos Lotes

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim tRastreamentoLoteSaldo As typeRastreamentoLoteSaldo
Dim objRastreamentoMovto As New ClassRastreamentoMovto
Dim dQuantidadeEscaninho As Double
Dim dQuantidadeTotal As Double
Dim dQtdAlocado As Double

On Error GoTo Erro_Rastreamento_Serie_Gera_Entrada
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 177229
    Next
    
    'se é um produto com rastreamento por número de série
    If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
    
        Set colRastreamentoSerie = New Collection
    
        tRastreamentoLoteSaldo.sProduto = String(STRING_PRODUTO, 0)
        tRastreamentoLoteSaldo.sLote = String(STRING_LOTE_RASTREAMENTO, 0) 'Alterado por Wagner 09/03/2006
        
        'Lê o RastreamentoLoteSaldo, RastreamentoLote em ordem decrescente de Data
        lErro = Comando_Executar(alComando(1), "SELECT S.NumIntDocLote, S.QuantDispNossa, S.QuantReservada, S.QuantReservadaConsig, S.QuantEmpenhada, S.QuantPedida, S.QuantRecIndl, S.QuantIndOutras, S.QuantDefeituosa, S.QuantConsig3, S.QuantConsig, S.QuantDemo3, S.QuantDemo, S.QuantConserto3, S.QuantConserto, S.QuantOutras3, S.QuantOutras, S.QuantOP, S.QuantBenef, S.QuantBenef3, L.Lote, L.FilialOP FROM RastreamentoLoteSaldo AS S RIGHT OUTER JOIN RastreamentoLote AS L ON S.NumIntDocLote = L.NumIntDoc WHERE L.Produto = ? AND L.Status = ? AND L.Lote >= ? ORDER BY L.Lote", tRastreamentoLoteSaldo.lNumIntDocLote, tRastreamentoLoteSaldo.dQuantDispNossa, tRastreamentoLoteSaldo.dQuantReservada, tRastreamentoLoteSaldo.dQuantReservadaConsig, tRastreamentoLoteSaldo.dQuantEmpenhada, _
        tRastreamentoLoteSaldo.dQuantPedida, tRastreamentoLoteSaldo.dQuantRecIndl, tRastreamentoLoteSaldo.dQuantIndOutras, tRastreamentoLoteSaldo.dQuantDefeituosa, tRastreamentoLoteSaldo.dQuantConsig3, tRastreamentoLoteSaldo.dQuantConsig, tRastreamentoLoteSaldo.dQuantDemo3, tRastreamentoLoteSaldo.dQuantDemo, tRastreamentoLoteSaldo.dQuantConserto3, tRastreamentoLoteSaldo.dQuantConserto, tRastreamentoLoteSaldo.dQuantOutras3, tRastreamentoLoteSaldo.dQuantOutras, tRastreamentoLoteSaldo.dQuantOP, tRastreamentoLoteSaldo.dQuantBenef, tRastreamentoLoteSaldo.dQuantBenef3, tRastreamentoLoteSaldo.sLote, tRastreamentoLoteSaldo.iFilialOP, objItemMovEst.sProduto, RATREAMENTOLOTE_STATUS_ABERTO, sSerieIni)
        If lErro <> AD_SQL_SUCESSO Then gError 177230
               
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 177231
                
        'Se ainda não alcançou a quantidade e ainda existe lote a ser acumulado
        Do While objItemMovEst.dQuantidadeEst > dQtdAlocado And lErro = AD_SQL_SUCESSO
        
            Set objRastreamentoMovto = New ClassRastreamentoMovto
            
            'Depende do escaninho
            Select Case iEscaninho
                
                'Acumula uma quantidade Total e a quantidade no Lote-escaninho atual
                Case ESCANINHO_NOSSO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDispNossa
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDispNossa
                Case ESCANINHO_3_EM_CONSIGNACAO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConsig3
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConsig3
                Case ESCANINHO_NOSSO_EM_CONSIGNACAO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConsig
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConsig
                Case ESCANINHO_3_EM_DEMO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDemo3
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDemo3
                Case ESCANINHO_NOSSO_EM_DEMO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDemo
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDemo
                Case ESCANINHO_3_EM_CONSERTO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConserto3
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConserto3
                Case ESCANINHO_NOSSO_EM_CONSERTO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantConserto
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantConserto
                Case ESCANINHO_3_EM_OUTROS
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantOutras3
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantOutras3
                Case ESCANINHO_NOSSO_EM_OUTROS
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantOutras
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantOutras
                Case ESCANINHO_3_EM_BENEF
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantBenef3
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantBenef3
                Case ESCANINHO_NOSSO_EM_BENEF
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantBenef
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantBenef
                Case ESCANINHO_DEFEITUOSO
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDefeituosa
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDefeituosa
                Case ESCANINHO_INDISPONIVEL
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantIndOutras
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantIndOutras
                Case ESCANINHO_RECEBIDO_INDISPONIVEL
                    dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantRecIndl
                    dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantRecIndl
            
            End Select
                
            If dQuantidadeTotal = 0 Then
            
                'Armazena os dados do Lote
                objRastreamentoMovto.sSiglaUM = objItemMovEst.sSiglaUM
                objRastreamentoMovto.dQuantidade = objRastreamentoMovto.dQuantidade * (objItemMovEst.dQuantidadeEst / objItemMovEst.dQuantidade)
                objRastreamentoMovto.iFilialOP = tRastreamentoLoteSaldo.iFilialOP
                objRastreamentoMovto.sLote = tRastreamentoLoteSaldo.sLote
                objRastreamentoMovto.lNumIntDocLote = tRastreamentoLoteSaldo.lNumIntDocLote
                objRastreamentoMovto.sProduto = objItemMovEst.sProduto
                objRastreamentoMovto.dQuantidade = 1
                
                dQtdAlocado = dQtdAlocado + objRastreamentoMovto.dQuantidade
                                       
                'coloca na coleção
                colRastreamentoSerie.Add objRastreamentoMovto
            
            End If
                
            'Busca o próximo Lote
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 177232
                    
        Loop
            
        'Se não tem quantidade suficiente ---> Erro
        If objItemMovEst.dQuantidadeEst > dQtdAlocado Then gError 177233

    End If
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Rastreamento_Serie_Gera_Entrada = SUCESSO
        
    Exit Function
    
Erro_Rastreamento_Serie_Gera_Entrada:

    Rastreamento_Serie_Gera_Entrada = gErr
    
    Select Case gErr
            
        Case 177229
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
           
        Case 177230 To 177232
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTESALDO2", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
                
        Case 177233
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_INSUF_RASTREAMENTO", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, iEscaninho)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177234)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

'#############################################################################################################
'Inserido por Wagner 03/04/2006
Function RastreamentoSerie_Obtem_Escaninho(ByVal objItemMovEst As ClassItemMovEstoque, ByVal objTipoMovEstoque As ClassTipoMovEst, ByVal objEstoqueProduto As ClassEstoqueProduto, iEscaninho As Integer, bSaida As Boolean) As Long
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long

On Error GoTo Erro_RastreamentoSerie_Obtem_Escaninho
    
    objTipoMovEstoque.iCodigo = objItemMovEst.iTipoMov
    
    'ler os dados referentes ao tipo de movimento
    lErro = CF("TipoMovEstoque_Le", objTipoMovEstoque)
    If lErro <> SUCESSO Then gError 177270
            
    bSaida = False
   
    'se a movimentação é referente a outras movimentações de material indisponivel
    If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
        objEstoqueProduto.dQuantInd = 1
        If Not bSaida Then
            iEscaninho = ESCANINHO_INDISPONIVEL
            bSaida = False
        End If
    End If
    
    'se a movimentação é referente a outras movimentações de material indisponivel
    If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
        objEstoqueProduto.dQuantInd = -1
        iEscaninho = ESCANINHO_INDISPONIVEL
        bSaida = True
    End If
    
    'se a movimentação é referente a nosso material disponivel
    If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
        objEstoqueProduto.dQuantDispNossa = 1
        If Not bSaida Then
            iEscaninho = ESCANINHO_NOSSO
            bSaida = False
        End If
    End If
    
    'se a movimentação é referente a nosso material disponivel
    If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
        objEstoqueProduto.dQuantDispNossa = -1
        iEscaninho = ESCANINHO_NOSSO
        bSaida = True
    End If
    
    'se a movimentação é referente a material defeituoso
    If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
        objEstoqueProduto.dQuantDefeituosa = 1
        If Not bSaida Then
            iEscaninho = ESCANINHO_DEFEITUOSO
            bSaida = False
        End If
    End If
    
    'se a movimentação é referente a material defeituoso
    If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
        objEstoqueProduto.dQuantDefeituosa = -1
        iEscaninho = ESCANINHO_DEFEITUOSO
        bSaida = True
    End If
    
    'se a movimentação é referente a material recebido e indisponível
    If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP Then
        objEstoqueProduto.dQuantRecIndl = 1
        If Not bSaida Then
            iEscaninho = ESCANINHO_RECEBIDO_INDISPONIVEL
            bSaida = True
        End If
    End If
    
    'se a movimentação é referente a material recebido e indisponível
    If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP Then
        objEstoqueProduto.dQuantRecIndl = -1
        iEscaninho = ESCANINHO_RECEBIDO_INDISPONIVEL
        bSaida = True
    End If
    
    'se a movimentação é referente a um conserto
    If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConserto = 1
            If Not bSaida Then
                iEscaninho = ESCANINHO_NOSSO_EM_CONSERTO
                bSaida = False
            End If
        Else
            objEstoqueProduto.dQuantConserto3 = 1
            If Not bSaida Then
                iEscaninho = ESCANINHO_3_EM_CONSERTO
                bSaida = False
            End If
        End If
    End If
    
    'se é uma movimentação referente a um conserto
    If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConserto = -1
            iEscaninho = ESCANINHO_NOSSO_EM_CONSERTO
            bSaida = True
        Else
            objEstoqueProduto.dQuantConserto3 = -1
            iEscaninho = ESCANINHO_3_EM_CONSERTO
            bSaida = True
        End If
    End If
    
    'se a movimentação é referente a uma demonstração
    If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantDemo = 1
            If Not bSaida Then
                iEscaninho = ESCANINHO_NOSSO_EM_DEMO
                bSaida = False
            End If
        Else
            objEstoqueProduto.dQuantDemo3 = 1
            If Not bSaida Then
                iEscaninho = ESCANINHO_3_EM_DEMO
                bSaida = False
            End If
        End If
    End If
    
    'se a movimentação é referente a uma demonstração
    If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantDemo = -1
            iEscaninho = ESCANINHO_NOSSO_EM_DEMO
            bSaida = True
        Else
            objEstoqueProduto.dQuantDemo3 = -1
            iEscaninho = ESCANINHO_3_EM_DEMO
            bSaida = True
        End If
    End If
    
    'se a movimentação é referente a material em consignação
    If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConsig = 1
            If Not bSaida Then
                iEscaninho = ESCANINHO_NOSSO_EM_CONSIGNACAO
                bSaida = False
            End If
        Else
            objEstoqueProduto.dQuantConsig3 = 1
            If Not bSaida Then
                iEscaninho = ESCANINHO_3_EM_CONSIGNACAO
                bSaida = False
            End If
        End If
    End If
    
    'se a movimentação é referente a material em consignação
    If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantConsig = -1
            iEscaninho = ESCANINHO_NOSSO_EM_CONSIGNACAO
            bSaida = True
        Else
            objEstoqueProduto.dQuantConsig3 = -1
            iEscaninho = ESCANINHO_3_EM_CONSIGNACAO
            bSaida = True
        End If
    End If
    
    'se a movimentação é referente a outras movimentações de material
    If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantOutras = 1
            If Not bSaida Then
                iEscaninho = ESCANINHO_NOSSO_EM_OUTROS
                bSaida = False
            End If
        Else
            objEstoqueProduto.dQuantOutras3 = 1
            If Not bSaida Then
                iEscaninho = ESCANINHO_3_EM_OUTROS
                bSaida = False
            End If
        End If
    End If
    
    'se a movimentação é referente a outras movimentações de material
    If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantOutras = -1
            iEscaninho = ESCANINHO_NOSSO_EM_OUTROS
            bSaida = True
        Else
            objEstoqueProduto.dQuantOutras3 = -1
            iEscaninho = ESCANINHO_3_EM_OUTROS
            bSaida = True
        End If
    End If
    
    'se a movimentação é referente a material em beneficiamento
    If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantBenef = 1
            If Not bSaida Then
                iEscaninho = ESCANINHO_NOSSO_EM_BENEF
                bSaida = False
            End If
        Else
            objEstoqueProduto.dQuantBenef3 = 1
            If Not bSaida Then
                iEscaninho = ESCANINHO_3_EM_BENEF
                bSaida = False
            End If
        End If
    End If
    
    'se é uma movimentação referente a um conserto
    If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
        If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
            objEstoqueProduto.dQuantBenef = -1
            iEscaninho = ESCANINHO_NOSSO_EM_BENEF
            bSaida = True
        Else
            objEstoqueProduto.dQuantBenef3 = -1
            iEscaninho = ESCANINHO_3_EM_BENEF
            bSaida = True
        End If
    End If
    
    RastreamentoSerie_Obtem_Escaninho = SUCESSO
    
    Exit Function
    
Erro_RastreamentoSerie_Obtem_Escaninho:

    RastreamentoSerie_Obtem_Escaninho = gErr
    
    Select Case gErr
    
        Case 177270
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177268)
        
    End Select
        
    Exit Function
    
End Function

Function RastreamentoSerie_Valida_Serie(ByVal objItemMovEst As ClassItemMovEstoque, ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal sSerie As String) As Long
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim objRastroSaldo As New ClassRastreamentoLoteSaldo

On Error GoTo Erro_RastreamentoSerie_Valida_Serie

    objRastroSaldo.sProduto = objItemMovEst.sProduto
    objRastroSaldo.sLote = sSerie
    objRastroSaldo.iAlmoxarifado = objItemMovEst.iAlmoxarifado
    
    lErro = CF("RastreamentoLoteSaldo_Le", objRastroSaldo)
    If lErro <> SUCESSO And lErro <> 78633 Then gError 177273
    
    'Verifica se tem saldo
    objRastroSaldo.dQuantConserto = objRastroSaldo.dQuantConserto + objEstoqueProduto.dQuantConserto
    objRastroSaldo.dQuantConserto3 = objRastroSaldo.dQuantConserto3 + objEstoqueProduto.dQuantConserto3
    objRastroSaldo.dQuantConsig = objRastroSaldo.dQuantConsig + objEstoqueProduto.dQuantConsig
    objRastroSaldo.dQuantConsig3 = objRastroSaldo.dQuantConsig3 + objEstoqueProduto.dQuantConsig3
    objRastroSaldo.dQuantDefeituosa = objRastroSaldo.dQuantDefeituosa + objEstoqueProduto.dQuantDefeituosa
    objRastroSaldo.dQuantDemo = objRastroSaldo.dQuantDemo + objEstoqueProduto.dQuantDemo
    objRastroSaldo.dQuantDemo3 = objRastroSaldo.dQuantDemo3 + objEstoqueProduto.dQuantDemo3
    objRastroSaldo.dQuantDispNossa = objRastroSaldo.dQuantDispNossa + objEstoqueProduto.dQuantDispNossa
    objRastroSaldo.dQuantIndOutras = objRastroSaldo.dQuantIndOutras + objEstoqueProduto.dQuantInd
    objRastroSaldo.dQuantOutras = objRastroSaldo.dQuantOutras + objEstoqueProduto.dQuantOutras
    objRastroSaldo.dQuantOutras3 = objRastroSaldo.dQuantOutras3 + objEstoqueProduto.dQuantOutras3
    objRastroSaldo.dQuantBenef = objRastroSaldo.dQuantBenef + objEstoqueProduto.dQuantBenef
    objRastroSaldo.dQuantBenef3 = objRastroSaldo.dQuantBenef3 + objEstoqueProduto.dQuantBenef3
    
    If objRastroSaldo.dQuantConserto < -QTDE_ESTOQUE_DELTA Then gError 177274
    If objRastroSaldo.dQuantConserto3 < -QTDE_ESTOQUE_DELTA Then gError 177275
    If objRastroSaldo.dQuantConsig < -QTDE_ESTOQUE_DELTA Then gError 177276
    If objRastroSaldo.dQuantConsig3 < -QTDE_ESTOQUE_DELTA Then gError 177277
    If objRastroSaldo.dQuantDefeituosa < -QTDE_ESTOQUE_DELTA Then gError 177278
    If objRastroSaldo.dQuantDemo < -QTDE_ESTOQUE_DELTA Then gError 177279
    If objRastroSaldo.dQuantDemo3 < -QTDE_ESTOQUE_DELTA Then gError 177280
    If objRastroSaldo.dQuantDispNossa < -QTDE_ESTOQUE_DELTA Then gError 177281
    If objRastroSaldo.dQuantIndOutras < -QTDE_ESTOQUE_DELTA Then gError 177282
    If objRastroSaldo.dQuantOutras < -QTDE_ESTOQUE_DELTA Then gError 177283
    If objRastroSaldo.dQuantOutras3 < -QTDE_ESTOQUE_DELTA Then gError 177284
    If objRastroSaldo.dQuantBenef < -QTDE_ESTOQUE_DELTA Then gError 177285
    If objRastroSaldo.dQuantBenef3 < -QTDE_ESTOQUE_DELTA Then gError 177286
    
    'Verifica se ficará com mais do que 1 da mesma série
    If objRastroSaldo.dQuantConserto + objRastroSaldo.dQuantConserto3 + objRastroSaldo.dQuantConsig + objRastroSaldo.dQuantConsig3 _
    + objRastroSaldo.dQuantDefeituosa + objRastroSaldo.dQuantDemo + objRastroSaldo.dQuantDemo3 _
    + objRastroSaldo.dQuantDispNossa + objRastroSaldo.dQuantIndOutras + objRastroSaldo.dQuantOutras _
    + objRastroSaldo.dQuantOutras3 + objRastroSaldo.dQuantBenef + objRastroSaldo.dQuantBenef3 > 1 Then gError 177287
    
    RastreamentoSerie_Valida_Serie = SUCESSO
    
    Exit Function
    
Erro_RastreamentoSerie_Valida_Serie:

    RastreamentoSerie_Valida_Serie = gErr
    
    Select Case gErr
    
        Case 177274
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSERTO_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantConserto)
            
        Case 177275
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSERTO3_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantConserto3)
            
        Case 177276
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSIG_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantConsig)
            
        Case 177277
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_CONSIG3_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantConsig3)
            
        Case 177278
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEFEITUOSO_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantDefeituosa)
            
        Case 177279
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEMO_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantDemo)
            
        Case 177280
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DEMO3_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantDemo3)
            
        Case 177281
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_DISPONIVEL_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantDispNossa)
            
        Case 177282
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_INDISPONIVEL_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantIndOutras)
            
        Case 177283
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_OUTRAS_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantOutras)
            
        Case 177284
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_OUTRAS3_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantOutras3)
            
        Case 177285
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_BENEF_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantBenef)
            
        Case 177286
            Call Rotina_Erro(vbOKOnly, "ERRO_SALDO_MAT_BENEF3_RASTRO", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie, objRastroSaldo.dQuantBenef3)
        
         Case 177287
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_RASTRO_REPETIDA", gErr, objRastroSaldo.sProduto, objRastroSaldo.iAlmoxarifado, sSerie)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177268)
        
    End Select
        
    Exit Function
    
End Function

Function FIFO_Testa(ByVal iFilialEmpresa As Integer, ByVal objMATConfig As ClassMATConfig) As Long

Dim lErro As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_FIFO_Testa

    If iFilialEmpresa <= 50 And Not (gcolFiliais Is Nothing) Then
        For Each objFiliais In gcolFiliais
            If objFiliais.iCodFilial > 50 Then
                objMATConfig.sConteudo = CStr(NAO_HABILITA_FIFO_NF)
                Exit For
            End If
        Next
    End If
    
    FIFO_Testa = SUCESSO
     
    Exit Function
    
Erro_FIFO_Testa:

    FIFO_Testa = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130531)
     
    End Select
     
    Exit Function

End Function

Public Function Rastreamento_FIFO_Gera_TEC(ByVal objItemMovEst As ClassItemMovEstoque, ByVal objProduto As ClassProduto, ByVal objProdutoMP As ClassProduto, dQuantEstTotal) As Long
'Gera na ordem de Data de Entrada descrecente os Movimentos dos Lotes

Dim lErro As Long
Dim tRastreamentoLoteSaldo As typeRastreamentoLoteSaldo
Dim objRastreamentoMovto As New ClassRastreamentoMovto
Dim dQuantidadeEscaninho As Double
Dim objMATConfig As New ClassMATConfig
Dim lComando As Long
Dim dFator As Double
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim sUM As String
Dim dQuantidade As Double
Dim dQuantidadeEst As Double
Dim dQuantReqEstTotal As Double


'##################################
'Inserido por Wagner 09/03/2006
Dim lNumIntDocLote As Long
Dim sSerieIni As String
'##################################

On Error GoTo Erro_Rastreamento_FIFO_Gera_TEC

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 199102
    Next

    objMATConfig.iFilialEmpresa = EMPRESA_TODA
    objMATConfig.sCodigo = HABILITA_FIFO_NFISCAIS

    lErro = CF("MATConfig_Le", objMATConfig)
    If lErro <> SUCESSO And lErro <> 89118 Then gError 199103

    If lErro = 89118 Then gError 199104
    
    lErro = FIFO_Testa(objItemMovEst.iFilialEmpresa, objMATConfig)
    If lErro <> SUCESSO Then gError 199105

    If objMATConfig.sConteudo = CStr(HABILITA_FIFO_NF) Then
    
        'se é um produto com rastreamento
        If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
    
            'se a colecao de rastreamentos não tiver sido inicializada ==> inicializa
            If objItemMovEst.colRastreamentoMovto Is Nothing Then Set objItemMovEst.colRastreamentoMovto = New Collection
    
            'se a colecao de rastreamentos estiver zerada ==> tem que gerar os rastreamentos.
            If objItemMovEst.colRastreamentoMovto.Count = 0 Then
    
                tRastreamentoLoteSaldo.sLote = String(STRING_LOTE_RASTREAMENTO, 0) 'Alterado por Wagner 09/03/2006
                
                'Lê o RastreamentoLoteSaldo, RastreamentoLote em ordem decrescente de Data
                lErro = Comando_Executar(alComando(1), "SELECT NumIntDocLote, Lote FROM RastreamentoLoteSaldo, RastreamentoLote WHERE RastreamentoLote.NumIntDoc = RastreamentoLoteSaldo.NumIntDocLote AND RastreamentoLoteSaldo.Produto = ? AND Almoxarifado = ? AND Status = ? AND QuantOutras3 > 0 ORDER BY DataEntrada, RastreamentoLote.NumIntDoc", tRastreamentoLoteSaldo.lNumIntDocLote, _
                tRastreamentoLoteSaldo.sLote, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, RATREAMENTOLOTE_STATUS_ABERTO)
                If lErro <> AD_SQL_SUCESSO Then gError 199106
                       
                lErro = Comando_BuscarPrimeiro(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199107
                    
                'Se ainda não alcançou a quantidade e ainda existe lote a ser acumulado
                Do While objItemMovEst.dQuantidadeEst > dQuantEstTotal And lErro = AD_SQL_SUCESSO
                
                    sUM = String(STRING_UM_SIGLA, 0)
                    
                    lErro = Comando_Executar(alComando(2), "SELECT MREQ.SiglaUM, MREQ.Quantidade FROM RastreamentoMovto AS R, RequisicaoProdTec AS MP, MovimentoEstoque AS MREQ " & _
                    "WHERE R.TipoDocOrigem = ? AND R.NumIntDocLote = ? AND R.NumIntDocOrigem = MP.NumINtDocMovEstProd AND MP.FilialEmpresa = ? AND  MREQ.TipoMov = ? AND MREQ.Codigo = MP.CodigoMovEstReq AND MREQ.Produto = ?", _
                    sUM, dQuantidade, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, tRastreamentoLoteSaldo.lNumIntDocLote, objItemMovEst.iFilialEmpresa, MOV_EST_REQ_PRODUCAO_OUTROS, objProdutoMP.sCodigo)
                    If lErro <> AD_SQL_SUCESSO Then gError 199108
                
                    lErro = Comando_BuscarPrimeiro(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199109
                
                    dQuantReqEstTotal = 0
                                    
                    Do While lErro = AD_SQL_SUCESSO
                
                        'converter qtde a faturar p/UMEstoque do produto
                        lErro = CF("UM_Conversao", objProdutoMP.iClasseUM, sUM, objProdutoMP.sSiglaUMEstoque, dFator)
                        If lErro <> SUCESSO Then gError 199110
        
                        dQuantidadeEst = dQuantidade * dFator
                    
                        dQuantEstTotal = dQuantEstTotal + dQuantidadeEst
                        
                        dQuantReqEstTotal = dQuantReqEstTotal + dQuantidadeEst
                
                        lErro = Comando_BuscarProximo(alComando(2))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199111
                
                    Loop
                
                    Set objRastreamentoMovto = New ClassRastreamentoMovto
                    
'                    'Depende do escaninho
'                    Select Case iEscaninho
'
'                        'Acumula uma quantidade Total e a quantidade no Lote-escaninho atual
'                        Case ESCANINHO_NOSSO
'                            dQuantidadeTotal = dQuantidadeTotal + tRastreamentoLoteSaldo.dQuantDispNossa
'                            dQuantidadeEscaninho = tRastreamentoLoteSaldo.dQuantDispNossa
'
'                    End Select
                            
'                    'Se a quantidade for maior
'                    If dQuantidadeTotal > objItemMovEst.dQuantidadeEst Then
'                        'Pega somente o necessário
'                        objRastreamentoMovto.dQuantidadeEst = objItemMovEst.dQuantidadeEst - (dQuantidadeTotal - dQuantidadeEscaninho)
'                    Else
'                        'Se for menor ou igual --> Pega toda a quantidade do lote
'                        objRastreamentoMovto.dQuantidadeEst = dQuantidadeEscaninho
'                    End If
                    
                    If dQuantReqEstTotal <> 0 Then
                    
                        'Armazena os dados do Lote
                        objRastreamentoMovto.sSiglaUM = objProduto.sSiglaUMEstoque
                        objRastreamentoMovto.dQuantidade = 1
                        objRastreamentoMovto.sLote = tRastreamentoLoteSaldo.sLote
                        objRastreamentoMovto.lNumIntDocLote = tRastreamentoLoteSaldo.lNumIntDocLote
                        objRastreamentoMovto.sProduto = objItemMovEst.sProduto
                                               
                        'coloca na coleção
                        objItemMovEst.colRastreamentoMovto.Add objRastreamentoMovto
                    
                    End If
                    
                    'Busca o próximo Lote
                    lErro = Comando_BuscarProximo(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199112
                            
                Loop
                
                
'                '############################################
'                'Inserido por Wagner 09/03/2006
'                'Se for Rastreado por Série tem que gravar a Série Inicial e a Qtd de Séries
'                If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
'                    For Each objRastreamentoMovto In objItemMovEst.colRastreamentoMovto
'                        If Len(Trim(sSerieIni)) = 0 Or sSerieIni > objRastreamentoMovto.sLote Then
'                            lNumIntDocLote = objRastreamentoMovto.lNumIntDocLote
'                            sSerieIni = objRastreamentoMovto.sLote
'                        End If
'                    Next
'                    For Each objRastreamentoMovto In objItemMovEst.colRastreamentoMovto
'                        objRastreamentoMovto.lNumIntDocLoteSerieIni = lNumIntDocLote
'                        objRastreamentoMovto.dQuantidadeSerie = objItemMovEst.dQuantidadeEst
'                    Next
'                End If
'                '############################################
        
            End If
            
        End If
        
    End If
        
        
        
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Rastreamento_FIFO_Gera_TEC = SUCESSO
        
    Exit Function
    
Erro_Rastreamento_FIFO_Gera_TEC:

    Rastreamento_FIFO_Gera_TEC = gErr
    
    Select Case gErr
            
        Case 199102
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 199103, 199105, 199110, 199160
        
        Case 199104
            Call Rotina_Erro(vbOKOnly, "ERRO_MATCONFIG_NAO_CADASTRADO", gErr, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
            
        Case 199106, 199107, 199112
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTESALDO2", gErr, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
                
        Case 199108, 199109, 199111
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 199113)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Public Function Producao_CalculaFatores(ByVal objItemMovEst As ClassItemMovEstoque, dQuantFator1 As Double, dQuantFator2 As Double, dQuantFator3 As Double, dQuantFator4 As Double, dQuantFator5 As Double, dQuantFator6 As Double) As Long

Dim lErro As Long
Dim alComando(1 To 5) As Long
Dim iIndice As Integer
Dim dFator As Double
Dim iClasseUM As Integer, dPesoFator1 As Double, dPesoFator2 As Double, dPesoFator3 As Double, dPesoFator4 As Double, dPesoFator5 As Double, dPesoFator6 As Double, dQuantidade As Double, sUnidadeMed As String
Dim dQuantFator As Double, dQuantFatorBD As Double, sVersao As String

On Error GoTo Erro_Producao_CalculaFatores

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201475
    Next

    dQuantFator1 = 0
    dQuantFator2 = 0
    dQuantFator3 = 0
    dQuantFator4 = 0
    dQuantFator5 = 0
    dQuantFator6 = 0
    
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    sVersao = String(STRING_KIT_VERSAO, 0)
    lErro = Comando_Executar(alComando(1), "SELECT Kit.Versao, P.ClasseUM, Kit.PesoFator1, Kit.PesoFator2, Kit.PesoFator3, Kit.PesoFator4, Kit.PesoFator5, Kit.PesoFator6, PK.Quantidade, PK.UnidadeMed FROM ItensOrdemProducaoBaixadas I, Produtos P, Kit, ProdutoKit PK WHERE Kit.Produtoraiz = PK.ProdutoRaiz AND Kit.Versao = PK.Versao AND PK.Nivel = 0 AND I.Produto = Kit.ProdutoRaiz AND (Kit.Versao = I.Versao OR I.Versao='' AND Kit.Situacao = 1) AND P.Codigo = I.Produto AND I.NumIntDoc = ?", _
        sVersao, iClasseUM, dPesoFator1, dPesoFator2, dPesoFator3, dPesoFator4, dPesoFator5, dPesoFator6, dQuantidade, sUnidadeMed, objItemMovEst.lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 201476

    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201477

    If lErro <> AD_SQL_SUCESSO Then

        sUnidadeMed = String(STRING_UM_SIGLA, 0)
        sVersao = String(STRING_KIT_VERSAO, 0)
        lErro = Comando_Executar(alComando(2), "SELECT Kit.Versao, P.ClasseUM, Kit.PesoFator1, Kit.PesoFator2, Kit.PesoFator3, Kit.PesoFator4, Kit.PesoFator5, Kit.PesoFator6, PK.Quantidade, PK.UnidadeMed FROM ItensOrdemProducao I, Produtos P, Kit, ProdutoKit PK WHERE Kit.Produtoraiz = PK.ProdutoRaiz AND Kit.Versao = PK.Versao AND PK.Nivel = 0 AND I.Produto = Kit.ProdutoRaiz AND (Kit.Versao = I.Versao OR I.Versao='' AND Kit.Situacao = 1) AND P.Codigo = I.Produto AND I.NumIntDoc = ?", _
            sVersao, iClasseUM, dPesoFator1, dPesoFator2, dPesoFator3, dPesoFator4, dPesoFator5, dPesoFator6, dQuantidade, sUnidadeMed, objItemMovEst.lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 201478
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201479

    End If
    
    If lErro = AD_SQL_SUCESSO Then
    
        If UCase(sUnidadeMed) <> UCase(objItemMovEst.sSiglaUM) Then
        
            lErro = CF("UM_Conversao_Trans", iClasseUM, objItemMovEst.sSiglaUM, sUnidadeMed, dFator)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Else
        
            dFator = 1
        
        End If
        
        dQuantFator = objItemMovEst.dQuantidade * dFator / IIf(dQuantidade = 0, 1, dQuantidade)
        
        dQuantFator1 = Arredonda_Estoque(dPesoFator1 * dQuantFator)
        dQuantFator2 = Arredonda_Estoque(dPesoFator2 * dQuantFator)
        dQuantFator3 = Arredonda_Estoque(dPesoFator3 * dQuantFator)
        dQuantFator4 = Arredonda_Estoque(dPesoFator4 * dQuantFator)
        dQuantFator5 = Arredonda_Estoque(dPesoFator5 * dQuantFator)
        dQuantFator6 = Arredonda_Estoque(dPesoFator6 * dQuantFator)
    
    End If
    
    lErro = Comando_ExecutarPos(alComando(3), "SELECT QuantFator FROM CustoFabrMovto WHERE NumIntMovEst = ?", 0, dQuantFatorBD, objItemMovEst.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 201480
    
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201481
    
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_ExecutarPos(alComando(5), "UPDATE CustoFabrMovto SET QuantFator1 = ?, QuantFator2 = ?, QuantFator3 = ?, QuantFator4 = ?, QuantFator5 = ?, QuantFator6 = ?, Versao = ?, QuantFator = ?, UnidadeMed = ?", alComando(3), _
            dQuantFator1, dQuantFator2, dQuantFator3, dQuantFator4, dQuantFator5, dQuantFator6, sVersao, dQuantFator, sUnidadeMed)
        If lErro <> AD_SQL_SUCESSO Then gError 201482
    
    Else
    
        lErro = Comando_Executar(alComando(4), "INSERT INTO CustoFabrMovto (NumIntMovEst, QuantFator1, QuantFator2, QuantFator3, QuantFator4, QuantFator5, QuantFator6, Versao, QuantFator, UnidadeMed) VALUES (?,?,?,?,?,?,?,?,?,?)", _
            objItemMovEst.lNumIntDoc, dQuantFator1, dQuantFator2, dQuantFator3, dQuantFator4, dQuantFator5, dQuantFator6, sVersao, dQuantFator, sUnidadeMed)
        If lErro <> AD_SQL_SUCESSO Then gError 201483
    
    End If
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Producao_CalculaFatores = SUCESSO
    
    Exit Function
    
Erro_Producao_CalculaFatores:

    Producao_CalculaFatores = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201474)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function Estoque_TestaIntegridade(ByVal iFilialEmpresa As Integer, Optional ByVal bForcaTestes As Boolean = False, Optional ByVal bForcaLog As Boolean = False, Optional ByVal bExibeMsgTermino As Boolean = False) As Long
'Essa rotina tem por finalidade encontrar situações que estejam gerando problemas de integridade no estoque

Dim lErro As Long
Dim alComando(1 To 10) As Long
Dim iIndice As Integer
Dim sProduto As String, iAlmoxarifado As Integer, dQuantIni As Double
Dim dQuantTotalAtual  As Double, dQuantTotalCalc  As Double, dQuantDispNossaAtual As Double
Dim dQuantDispNossaCalc As Double, dQuantReservadaAtual As Double, dQuantReservadaCalc As Double
Dim dSaldoCalcPeriodo As Double, dtData As Date, dQtdeEnt As Double, dQtdeSai As Double
Dim dQtdeEntCalc As Double, dQtdeSaiCalc As Double, iAno As Integer, iMes As Integer
Dim adQuantEnt(1 To 12) As Double, adQuantSai(1 To 12) As Double
Dim dQuantAno As Double, dQuantAnoCusto As Double, dValorAno As Double, dValorAnoCusto As Double
Dim dQuantAnoCalc As Double, dQuantAnoCustoCalc As Double, dValorAnoCalc As Double, dValorAnoCustoCalc As Double
Dim sDir As String, sArqLog As String, bArqAberto As Boolean, lNumErros As Long, iAux As Integer

Const ACAO_NAO_TESTA = 0
Const ACAO_NAO_ERRO = 1
Const ACAO_NAO_LOG = 2

On Error GoTo Erro_Estoque_TestaIntegridade

    bArqAberto = False

    If gobjMAT.iESTTestaIntAcao <> ACAO_NAO_TESTA Or bForcaTestes Then
    
        If gobjMAT.iESTTestaIntAcao = ACAO_NAO_LOG Or bForcaLog Then
            sDir = gobjMAT.sESTTestaIntArqLog
            If sDir = "" Then sDir = CurDir$
            If right(sDir, 1) <> "/" And right(sDir, 1) <> "\" Then sDir = sDir & "\"
            sArqLog = "EST_INT_LOG_" & CStr(glEmpresa) & "_" & CStr(giFilialEmpresa) & "_" & UCase(gsUsuario) & "_" & Format(Date, "YYYYMMDD") & "_" & Format(Now, "HHMMSS") & ".txt"
        
            Open sDir & sArqLog For Output As #1
            bArqAberto = True
            Print #1, CStr(Now)
        End If
    
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 213080
        Next
        
        'Procura erros na quantidade disponível e reservada em estoque produto
        If gobjMAT.iESTTestaEstoqueProduto = MARCADO Or bForcaTestes Then
        
            sProduto = String(STRING_PRODUTO, 0)
        
            lErro = Comando_Executar(alComando(1), "SELECT Produto,Almoxarifado,QuantIni,QuantTotalAtual,QuantTotalCalc,QuantDispNossaAtual,QuantDispNossaCalc,QuantReservadaAtual,QuantReservadaCalc,SaldoCalcPeriodo FROM TesteSaldoNossaDisp WHERE (ABS(QuantDispNossaAtual-QuantDispNossaCalc) > ? OR ABS(QuantReservadaAtual-QuantReservadaCalc) > ?) AND Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa = ?) ORDER BY Produto, Almoxarifado", _
            sProduto, iAlmoxarifado, dQuantIni, dQuantTotalAtual, dQuantTotalCalc, dQuantDispNossaAtual, dQuantDispNossaCalc, dQuantReservadaAtual, dQuantReservadaCalc, dSaldoCalcPeriodo, DELTA_VALORMONETARIO, DELTA_VALORMONETARIO, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 213081
            
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213082
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                If Abs(dQuantDispNossaAtual - dQuantDispNossaCalc) > DELTA_VALORMONETARIO Then
                    If bArqAberto Then
                        lNumErros = lNumErros + 1
                        Print #1, CStr(lNumErros) & SEPARADOR & "EstoqueProduto:213083: A quantidade disponível em estoque do produto " & Trim(sProduto) & " no almoxarifado " & CStr(iAlmoxarifado) & " está diferente da calculada. Atual: " & Format(dQuantDispNossaAtual, "STANDARD") & " Calculada: " & Format(dQuantDispNossaCalc, "STANDARD")
                    Else
                        gError 213083
                    End If
                End If
                
                If Abs(dQuantReservadaAtual - dQuantReservadaCalc) > DELTA_VALORMONETARIO Then
                    If bArqAberto Then
                        lNumErros = lNumErros + 1
                        Print #1, CStr(lNumErros) & SEPARADOR & "EstoqueProduto:213084: A quantidade reservada em estoque do produto " & Trim(sProduto) & " no almoxarifado " & CStr(iAlmoxarifado) & " está diferente da calculada. Atual: " & Format(dQuantReservadaAtual, "STANDARD") & " Calculada: " & Format(dQuantReservadaCalc, "STANDARD")
                    Else
                        gError 213084
                    End If
                End If
            
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213082
            
            Loop
            
        End If
        
        'Testa a integrigade das tabelas de saldos (sldmesest e slddiaest)
        If gobjMAT.iESTTestaSaldos = MARCADO Or bForcaTestes Then
        
            sProduto = String(STRING_PRODUTO, 0)
        
            'Compara SldDiaEst - Calc 1
            lErro = Comando_Executar(alComando(2), "SELECT Produto, Data, QuantEntrada, QuantSaida FROM SldDiaEst WHERE FilialEmpresa = ? ORDER BY Produto, Data", _
            sProduto, dtData, dQtdeEnt, dQtdeSai, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 213085
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213086
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                lErro = Comando_Executar(alComando(3), "SELECT FilialEmpresa, SUM(QuantEnt) AS QE, SUM(QuantSai) AS QS FROM KardexNossaDispDia WHERE Filialempresa = ? AND Produto = ? AND Data = ? GROUP BY FilialEmpresa", _
                iAux, dQtdeEntCalc, dQtdeSaiCalc, iFilialEmpresa, sProduto, dtData)
                If lErro <> AD_SQL_SUCESSO Then gError 213087
                
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213088
                
                If lErro <> AD_SQL_SUCESSO Then
                    dQtdeEntCalc = 0
                    dQtdeSaiCalc = 0
                End If
            
                'Incluindo para evitar diferenças onde de um lado não calcula nem entrada e nem saída enquanto do outro calcula entrada e saida de mesmo valor (lançamento/estorno)
                If Abs(dQtdeEntCalc - dQtdeSaiCalc - (dQtdeEnt - dQtdeSai)) > DELTA_VALORMONETARIO Then
                    If Abs(dQtdeEntCalc - dQtdeEnt) > DELTA_VALORMONETARIO Then
                        If bArqAberto Then
                            lNumErros = lNumErros + 1
                            Print #1, CStr(lNumErros) & SEPARADOR & "SldDiaEst:213089: A quantidade que entrou no estoque do produto " & Trim(sProduto) & " na data " & Format(dtData, "dd/mm/yyyy") & " está diferente da calculada. Atual: " & Format(dQtdeEnt, "STANDARD") & " Calculada: " & Format(dQtdeEntCalc, "STANDARD")
                        Else
                            gError 213089
                        End If
                    End If
                    
                    If Abs(dQtdeSaiCalc - dQtdeSai) > DELTA_VALORMONETARIO Then
                        If bArqAberto Then
                            lNumErros = lNumErros + 1
                            Print #1, CStr(lNumErros) & SEPARADOR & "SldDiaEst:213090: A quantidade que saiu do estoque do produto " & Trim(sProduto) & " na data " & Format(dtData, "dd/mm/yyyy") & " está diferente da calculada. Atual: " & Format(dQtdeSai, "STANDARD") & " Calculada: " & Format(dQtdeSaiCalc, "STANDARD")
                        Else
                            gError 213090
                        End If
                    End If
                End If
                
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213091
            
            Loop
            
            sProduto = String(STRING_PRODUTO, 0)
        
            'Compara SldDiaEst - Calc 2
            lErro = Comando_Executar(alComando(4), "SELECT Produto, Data, SUM(QuantEnt), SUM(QuantSai) FROM KardexNossaDispDia WHERE FilialEmpresa = ? GROUP BY Produto, Data ORDER BY Produto, Data ", _
            sProduto, dtData, dQtdeEntCalc, dQtdeSaiCalc, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 213092
            
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213093
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                lErro = Comando_Executar(alComando(5), "SELECT QuantEntrada, QuantSaida FROM SldDiaEst WHERE Filialempresa = ? AND Produto = ? AND Data = ?", _
                dQtdeEnt, dQtdeSai, iFilialEmpresa, sProduto, dtData)
                If lErro <> AD_SQL_SUCESSO Then gError 213094
                
                lErro = Comando_BuscarPrimeiro(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213095
                
                'Só testa o que não tem em SldDiaEst. pois quando tem já foi testado antes
                If lErro <> AD_SQL_SUCESSO Then
                    dQtdeEnt = 0
                    dQtdeSai = 0
                'End If
            
                    'Incluindo para evitar diferenças onde de um lado não calcula nem entrada e nem saída enquanto do outro calcula entrada e saida de mesmo valor (lançamento/estorno)
                    If Abs(dQtdeEntCalc - dQtdeSaiCalc - (dQtdeEnt - dQtdeSai)) > DELTA_VALORMONETARIO Then
                        If Abs(dQtdeEntCalc - dQtdeEnt) > DELTA_VALORMONETARIO Then
                            If bArqAberto Then
                                lNumErros = lNumErros + 1
                                Print #1, CStr(lNumErros) & SEPARADOR & "SldDiaEst:213096: A quantidade que entrou no estoque do produto " & Trim(sProduto) & " na data " & Format(dtData, "dd/mm/yyyy") & " está diferente da calculada. Atual: " & Format(dQtdeEnt, "STANDARD") & " Calculada: " & Format(dQtdeEntCalc, "STANDARD")
                            Else
                                gError 213096
                            End If
                        End If
                        
                        If Abs(dQtdeSaiCalc - dQtdeSai) > DELTA_VALORMONETARIO Then
                            If bArqAberto Then
                                lNumErros = lNumErros + 1
                                Print #1, CStr(lNumErros) & SEPARADOR & "SldDiaEst:213097: A quantidade que saiu do estoque do produto " & Trim(sProduto) & " na data " & Format(dtData, "dd/mm/yyyy") & " está diferente da calculada. Atual: " & Format(dQtdeSai, "STANDARD") & " Calculada: " & Format(dQtdeSaiCalc, "STANDARD")
                            Else
                                gError 213097
                            End If
                        End If
                    End If
                End If
                
                lErro = Comando_BuscarProximo(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213098
            
            Loop
            
            sProduto = String(STRING_PRODUTO, 0)
            
            'Compara SldMesEst (Saldos)
            lErro = Comando_Executar(alComando(6), "SELECT Produto, Month(Data), Year(Data), SUM(QuantEnt), SUM(QuantSai) FROM KardexNossaDispDia WHERE FilialEmpresa = ? GROUP BY Produto, Year(Data), Month(Data) ORDER BY Produto, Year(Data), Month(Data)", _
                sProduto, iMes, iAno, dQtdeEntCalc, dQtdeSaiCalc, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 213099
            
            lErro = Comando_BuscarProximo(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213100
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                lErro = Comando_Executar(alComando(7), "SELECT QuantEnt1, QuantSai1, QuantEnt2, QuantSai2, QuantEnt3, QuantSai3, QuantEnt4, QuantSai4, QuantEnt5, QuantSai5, QuantEnt6, QuantSai6, QuantEnt7, QuantSai7, QuantEnt8, QuantSai8, QuantEnt9, QuantSai9, QuantEnt10, QuantSai10, QuantEnt11, QuantSai11, QuantEnt12, QuantSai12 FROM SldMesEst WHERE Filialempresa = ? AND Produto = ? AND Ano = ? ", _
                adQuantEnt(1), adQuantSai(1), adQuantEnt(2), adQuantSai(2), adQuantEnt(3), adQuantSai(3), adQuantEnt(4), adQuantSai(4), _
                adQuantEnt(5), adQuantSai(5), adQuantEnt(6), adQuantSai(6), adQuantEnt(7), adQuantSai(7), adQuantEnt(8), adQuantSai(8), _
                adQuantEnt(9), adQuantSai(9), adQuantEnt(10), adQuantSai(10), adQuantEnt(11), adQuantSai(11), adQuantEnt(12), adQuantSai(12), _
                iFilialEmpresa, sProduto, iAno)
                If lErro <> AD_SQL_SUCESSO Then gError 213101
                
                lErro = Comando_BuscarPrimeiro(alComando(7))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213102
                
                If lErro <> AD_SQL_SUCESSO Then
                    For iIndice = 1 To 12
                        adQuantEnt(iIndice) = 0
                        adQuantSai(iIndice) = 0
                    Next
                End If
            
                'Incluindo para evitar diferenças onde de um lado não calcula nem entrada e nem saída enquanto do outro calcula entrada e saida de mesmo valor (lançamento/estorno)
                If Abs(dQtdeEntCalc - dQtdeSaiCalc - (adQuantEnt(iMes) - adQuantSai(iMes))) > DELTA_VALORMONETARIO Then
                    
                    If Abs(adQuantEnt(iMes) - dQtdeEntCalc) > DELTA_VALORMONETARIO Then
                        If bArqAberto Then
                            lNumErros = lNumErros + 1
                            Print #1, CStr(lNumErros) & SEPARADOR & "SldMesEst:213103: A quantidade que entrou no estoque do produto " & Trim(sProduto) & " no mês/ano " & Format(iMes, "00") & "/" & CStr(iAno) & " está diferente da calculada. Atual: " & Format(adQuantEnt(iMes), "STANDARD") & " Calculada: " & Format(dQtdeEntCalc, "STANDARD")
                        Else
                            gError 213103
                        End If
                    End If
                    
                    If Abs(adQuantSai(iMes) - dQtdeSaiCalc) > DELTA_VALORMONETARIO Then
                        If bArqAberto Then
                            lNumErros = lNumErros + 1
                            Print #1, CStr(lNumErros) & SEPARADOR & "SldMesEst:213104: A quantidade que saiu do estoque do produto " & Trim(sProduto) & " no mês/ano " & Format(iMes, "00") & "/" & CStr(iAno) & " está diferente da calculada. Atual: " & Format(adQuantSai(iMes), "STANDARD") & " Calculada: " & Format(dQtdeSaiCalc, "STANDARD")
                        Else
                            gError 213104
                        End If
                    End If
                End If
                
                lErro = Comando_BuscarProximo(alComando(6))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213105
            
            Loop
        
            sProduto = String(STRING_PRODUTO, 0)
        
            'Compara SldMesEst (QuantidadeInicial)
            lErro = Comando_Executar(alComando(8), "SELECT Produto, Ano, (QuantInicial+QuantEnt1-QuantSai1+QuantEnt2-QuantSai2+QuantEnt3-QuantSai3+QuantEnt4-QuantSai4+QuantEnt5-QuantSai5+QuantEnt6-QuantSai6+QuantEnt7-QuantSai7+QuantEnt8-QuantSai8+QuantEnt9-QuantSai9+QuantEnt10-QuantSai10+QuantEnt11-QuantSai11+QuantEnt12-QuantSai12), (QuantInicialCusto+SaldoQuantCusto1+SaldoQuantCusto2+SaldoQuantCusto3+SaldoQuantCusto4+SaldoQuantCusto5+SaldoQuantCusto6+SaldoQuantCusto7+SaldoQuantCusto8+SaldoQuantCusto9+SaldoQuantCusto10+SaldoQuantCusto11+SaldoQuantCusto12),  " & _
            "(ValorInicial+ValorEnt1-ValorSai1+ValorEnt2-ValorSai2+ValorEnt3-ValorSai3+ValorEnt4-ValorSai4+ValorEnt5-ValorSai5+ValorEnt6-ValorSai6+ValorEnt7-ValorSai7+ValorEnt8-ValorSai8+ValorEnt9-ValorSai9+ValorEnt10-ValorSai10+ValorEnt11-ValorSai11+ValorEnt12-ValorSai12), (ValorInicialCusto+SaldoValorCusto1+SaldoValorCusto2+SaldoValorCusto3+SaldoValorCusto4+SaldoValorCusto5+SaldoValorCusto6+SaldoValorCusto7+SaldoValorCusto8+SaldoValorCusto9+SaldoValorCusto10+SaldoValorCusto11+SaldoValorCusto12) FROM SldMesEst WHERE Filialempresa = ? ORDER BY Produto, Ano ", _
             sProduto, iAno, dQuantAnoCalc, dQuantAnoCustoCalc, dValorAnoCalc, dValorAnoCustoCalc, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 213106
            
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213107
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                lErro = Comando_Executar(alComando(7), "SELECT QuantInicial, QuantInicialCusto, ValorInicial, ValorInicialCusto FROM SldMesEst WHERE Produto = ? AND Ano = ? AND Filialempresa = ? ", dQuantAno, dQuantAnoCusto, dValorAno, dValorAnoCusto, sProduto, iAno + 1, iFilialEmpresa)
                If lErro <> AD_SQL_SUCESSO Then gError 213108
                
                lErro = Comando_BuscarProximo(alComando(7))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213109
                
                If lErro = AD_SQL_SUCESSO Then
            
                    If Abs(dQuantAno - dQuantAnoCalc) > DELTA_VALORMONETARIO Then
                        If bArqAberto Then
                            lNumErros = lNumErros + 1
                            Print #1, CStr(lNumErros) & SEPARADOR & "SldMesEst:213110: A quantidade inicial do produto " & Trim(sProduto) & " no ano " & CStr(iAno + 1) & " está diferente da calculada. Atual: " & Format(dQuantAno, "STANDARD") & " Calculada: " & Format(dQuantAnoCalc, "STANDARD")
                        Else
                            gError 213110
                        End If
                    End If
                
                    If Abs(dQuantAnoCusto - dQuantAnoCustoCalc) > DELTA_VALORMONETARIO Then
                        If bArqAberto Then
                            lNumErros = lNumErros + 1
                            Print #1, CStr(lNumErros) & SEPARADOR & "SldMesEst:213111: A quantidade inicial custo do produto " & Trim(sProduto) & " no ano " & CStr(iAno + 1) & " está diferente da calculada. Atual: " & Format(dQuantAnoCusto, "STANDARD") & " Calculada: " & Format(dQuantAnoCustoCalc, "STANDARD")
                        Else
                            gError 213111
                        End If
                    End If
                    
                    If Abs(dValorAno - dValorAnoCalc) > DELTA_VALORMONETARIO Then
                        If bArqAberto Then
                            lNumErros = lNumErros + 1
                            Print #1, CStr(lNumErros) & SEPARADOR & "SldMesEst:213112: O valor inicial do produto " & Trim(sProduto) & " no ano " & CStr(iAno + 1) & " está diferente da calculada. Atual: " & Format(dValorAno, "STANDARD") & " Calculada: " & Format(dValorAnoCalc, "STANDARD")
                        Else
                            gError 213112
                        End If
                    End If
                    
                    If Abs(dValorAnoCusto - dValorAnoCustoCalc) > DELTA_VALORMONETARIO Then
                        If bArqAberto Then
                            lNumErros = lNumErros + 1
                            Print #1, CStr(lNumErros) & SEPARADOR & "SldMesEst:213113: O valor inicial custo do produto " & Trim(sProduto) & " no ano " & CStr(iAno + 1) & " está diferente da calculada. Atual: " & Format(dValorAnoCusto, "STANDARD") & " Calculada: " & Format(dValorAnoCustoCalc, "STANDARD")
                        Else
                            gError 213113
                        End If
                    End If
            
                End If
            
                lErro = Comando_BuscarProximo(alComando(6))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213114
            
            Loop
        End If
    
        'libera comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
           
        If bArqAberto Then
            Print #1, CStr(Now)
            Close #1
            If bExibeMsgTermino Then
                If lNumErros > 0 Then
                    Call Rotina_Aviso(vbOKOnly, "AVISO_EST_TESTA_INT_FINALIZADO_LOG", sDir & sArqLog)
                Else
                    Call Rotina_Aviso(vbOKOnly, "AVISO_EST_TESTA_INT_FINALIZADO")
                End If
            End If
        Else
            If bExibeMsgTermino Then Call Rotina_Aviso(vbOKOnly, "AVISO_EST_TESTA_INT_FINALIZADO")
        End If
        
    End If
    
    Estoque_TestaIntegridade = SUCESSO
    
    Exit Function
    
Erro_Estoque_TestaIntegridade:

    Estoque_TestaIntegridade = gErr

    Select Case gErr
    
         Case 213080
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
   
        Case 213081, 213082, 213085 To 213088, 213091 To 213095, 213098 To 213102, 213105 To 213109, 213114
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUESPRODUTO", gErr)

        Case 213083
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTA_INT_EST_ESTPROD_QTDEDISP", gErr, sProduto, iAlmoxarifado, Format(dQuantDispNossaAtual, "STANDARD"), Format(dQuantDispNossaCalc, "STANDARD"))

        Case 213084
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTA_INT_EST_ESTPROD_QTDERESERV", gErr, sProduto, iAlmoxarifado, Format(dQuantReservadaAtual, "STANDARD"), Format(dQuantReservadaCalc, "STANDARD"))

        Case 213089, 213096
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTA_INT_EST_SLDDIAEST_QTDEENT", gErr, sProduto, dtData, Format(dQtdeEnt, "STANDARD"), Format(dQtdeEntCalc, "STANDARD"))

        Case 213090, 213098
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTA_INT_EST_SLDDIAEST_QTDESAI", gErr, sProduto, dtData, Format(dQtdeSai, "STANDARD"), Format(dQtdeSaiCalc, "STANDARD"))
        
        Case 213103
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTA_INT_EST_SLDMESEST_QTDEENT", gErr, sProduto, Format(iMes, "00") & "/" & CStr(iAno), Format(dQtdeEnt, "STANDARD"), Format(adQuantEnt(iMes), "STANDARD"))

        Case 213104
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTA_INT_EST_SLDMESEST_QTDESAI", gErr, sProduto, Format(iMes, "00") & "/" & CStr(iAno), Format(dQtdeSai, "STANDARD"), Format(adQuantSai(iMes), "STANDARD"))

        Case 213110
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTA_INT_EST_SLDMESEST_QTDEINI", gErr, sProduto, iAno + 1, Format(dQuantAno, "STANDARD"), Format(dQuantAnoCalc, "STANDARD"))

        Case 213111
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTA_INT_EST_SLDMESEST_QTDEINICUST", gErr, sProduto, iAno + 1, Format(dQuantAno, "STANDARD"), Format(dQuantAnoCalc, "STANDARD"))

        Case 213112
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTA_INT_EST_SLDMESEST_VLRINI", gErr, sProduto, iAno + 1, Format(dQuantAno, "STANDARD"), Format(dQuantAnoCalc, "STANDARD"))

        Case 213113
            Call Rotina_Erro(vbOKOnly, "ERRO_TESTA_INT_EST_SLDMESEST_VLRINICUST", gErr, sProduto, iAno + 1, Format(dQuantAno, "STANDARD"), Format(dQuantAnoCalc, "STANDARD"))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213115)
    
    End Select
    
    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    If bArqAberto Then Close #1
    
    Exit Function

End Function

Private Function Estoque_Custo_Trata_TipoMov205(ByVal objItemMovEst As ClassItemMovEstoque) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 3) As Long
Dim iItemNF As Integer, dCustoItemNF As Double, dQuantItemNF As Double
Dim dQuantInsumo As Double, lNumIntItemNF As Long

On Error GoTo Erro_Estoque_Custo_Trata_TipoMov205
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213357
    Next
    
    If objItemMovEst.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_NFISCAL Then
    
        lErro = Comando_Executar(alComando(1), "SELECT QuantInsumos, ItemNF FROM MovimentoEstoque WHERE NumIntDoc = ?", dQuantInsumo, iItemNF, objItemMovEst.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 213358
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213359
                
        lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc FROM ItensNFiscal WHERE NumIntNF = ? AND Item = ?", lNumIntItemNF, objItemMovEst.lNumIntDocOrigem, iItemNF)
        If lErro <> AD_SQL_SUCESSO Then gError 213360
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213361
        
        'If lErro = AD_SQL_SEM_DADOS Then gError 213362
        If lErro <> AD_SQL_SEM_DADOS Then
        
            dQuantItemNF = 0
            lErro = Comando_Executar(alComando(3), "SELECT Quantidade, Custo FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? ", dQuantItemNF, dCustoItemNF, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL, lNumIntItemNF)
            If lErro <> AD_SQL_SUCESSO Then gError 213363
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213364
            
            If dQuantItemNF > 0 Then
                objItemMovEst.dCusto = Arredonda_Moeda((dCustoItemNF / dQuantItemNF) * dQuantInsumo)
            End If
            
        End If
        
    End If
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Estoque_Custo_Trata_TipoMov205 = SUCESSO
        
    Exit Function
    
Erro_Estoque_Custo_Trata_TipoMov205:

    Estoque_Custo_Trata_TipoMov205 = gErr
    
    Select Case gErr
            
        Case 213357
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 213358 To 213361, 213363 To 213364
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
          
'        Case 213362
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 213365)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function
