VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function ESTConfig_Grava_Configs(ByVal colConfigs As ColESTConfig) As Long
'grava na tabela ESTConfig uma coleção de registros de configuração

Dim lErro As Long, iIndice As Integer
Dim lTransacao As Long

On Error GoTo Erro_ESTConfig_Grava_Configs

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 33654

    lErro = CF("ESTConfig_Grava_Configs_Trans", colConfigs)
    If lErro <> SUCESSO Then Error 33655

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 33656

    ESTConfig_Grava_Configs = SUCESSO

    Exit Function

Erro_ESTConfig_Grava_Configs:

    ESTConfig_Grava_Configs = Err

    Select Case Err

        Case 33654
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 33655

        Case 33656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149482)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function ESTConfig_Grava_Configs_Trans(ByVal colConfigs As ColESTConfig) As Long
'grava na tabela ESTConfig uma coleção de registros de configuração
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim objESTConfig As ClassESTConfig

On Error GoTo Erro_ESTConfig_Grava_Configs_Trans

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 33653
    Next

   'Para cada elemento da coleção
    For Each objESTConfig In colConfigs

        lErro = ESTConfig_Grava_Trans(objESTConfig, alComando())
        If lErro <> SUCESSO Then Error 33655

    Next

   'Fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ESTConfig_Grava_Configs_Trans = SUCESSO

    Exit Function

Erro_ESTConfig_Grava_Configs_Trans:

    ESTConfig_Grava_Configs_Trans = Err

    Select Case Err

        Case 33653
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33655

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149483)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function ESTConfig_Grava_Trans(ByVal objESTConfig As ClassESTConfig, alComando() As Long) As Long
'grava registro em ESTConfig. Apenas altera o conteúdo, se não existir dá erro.
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, tESTConfig As typeESTConfig

On Error GoTo Erro_ESTConfig_Grava_Trans

    'Inicializar buffer
    tESTConfig.sConteudo = String(STRING_ESTCONFIG_CONTEUDO, 0)

    'Lê o registo
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Conteudo FROM ESTConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, tESTConfig.sConteudo, objESTConfig.sCodigo, objESTConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 33657

    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33658

    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then Error 33659

    'Se o conteúdo foi modificado
    If objESTConfig.sConteudo <> tESTConfig.sConteudo Then
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ESTConfig SET Conteudo = ?", alComando(1), objESTConfig.sConteudo)
        If lErro <> AD_SQL_SUCESSO Then Error 33660
    End If

    Exit Function

Erro_ESTConfig_Grava_Trans:

    ESTConfig_Grava_Trans = Err

    Select Case Err

        Case 33657, 33658
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTCONFIG2", Err, objESTConfig.sCodigo, objESTConfig.iFilialEmpresa)

        Case 33659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTCONFIG_INEXISTENTE", Err, objESTConfig.sCodigo, objESTConfig.iFilialEmpresa)

        Case 33660
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ESTCONFIG", Err, objESTConfig.sCodigo, objESTConfig.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149484)

    End Select

    Exit Function

End Function

Function Almoxarifado_Automatico(iCodigo As Integer) As Long
'Obtem o próximo número da Almoxarifado

Dim lErro As Long, lNumAuto As Long

On Error GoTo Erro_Almoxarifado_Automatico
    
    'seleciona o codigo no bd e verifica se já existe
    lErro = CF("Config_ObterAutomatico", "CPRConfig", "NUM_PROX_ALMOXARIFADO", "Almoxarifado", "Codigo", lNumAuto)
    If lErro <> SUCESSO And lErro <> 25191 Then Error 22239
   
    iCodigo = CInt(lNumAuto)

    Almoxarifado_Automatico = SUCESSO

    Exit Function

Erro_Almoxarifado_Automatico:

    Almoxarifado_Automatico = Err

   Select Case Err

        'Erro já tratado
        Case 22239

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149485)

    End Select

    Exit Function

End Function

''Function NFiscalEntradaDevolucao_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'''Faz a gravação de Uma Nota Fiscal de Entrada-Devolução
''
''Dim lErro As Long
''Dim objItemNFOrig As New ClassItemNF
''Dim dQuantDevolvida As Double
''Dim lTransacao As Long
''Dim alComando(0 To 2) As Long
''Dim iIndice As Integer
''Dim lNumIntDoc As Long
''Dim vbMsgRes As VbMsgBoxResult
''Dim objTipoDocInfo As New ClassTipoDocInfo
''Dim iTipoNFOriginal As Integer
''
''On Error GoTo Erro_NFiscalEntradaDevolucao_Grava
''
''    'Abre a Transacao
''    lTransacao = Transacao_Abrir()
''    If lTransacao = 0 Then Error 35378
''
''    'Abre os Comandos
''    For iIndice = LBound(alComando) To UBound(alComando)
''        alComando(iIndice) = Comando_Abrir()
''        If alComando(iIndice) = 0 Then Error 35379
''    Next
''
''    'Guarda o Tipo Original de NFiscal para usar em estatística
''    iTipoNFOriginal = objNFiscal.iTipoNFiscal
''
''    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
''    'Lê o Tipo de Documento
''    lErro = CF("TipoDocInfo_Le_Codigo",objTipoDocInfo)
''    If lErro <> SUCESSO And lErro <> 31415 Then Error 42070
''
''    'Se não encontrou o Tipo de Documento --> erro
''    If lErro <> SUCESSO Then Error 42072
''
''    'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
''    lErro = CF("NFiscalEntrada_Verifica_Existencia2",objNFiscal, objTipoDocInfo)
''    If lErro <> SUCESSO And lErro <> 61414 Then Error 39498
''
''    'Se for uma alteração só altera os dados da contabilidade
''    If lErro = SUCESSO Then
''
''        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
''
''            'gera o Número para a Nota Fiscal Interna
''            lErro = CF("NFiscalNumAuto",objNFiscal)
''            If lErro <> SUCESSO Then Error 60448
''
''        End If
''
''        If objNFiscal.lFornecedor > 0 Then
''            'Lê a Filial Fornecedor
''            lErro = CF("FilialFornecedor_Lock",objNFiscal.lFornecedor, objNFiscal.iFilialForn)
''            If lErro <> SUCESSO And lErro <> 42870 Then Error 42082
''
''            'Se não encontrou a Filial Fornecedor --> erro
''            If lErro <> SUCESSO Then Error 35390
''
''        Else
''            'Lê a Filial Cliente
''            lErro = CF("FilialCliente_Lock",objNFiscal.lCliente, objNFiscal.iFilialCli)
''            If lErro <> SUCESSO And lErro <> 43050 Then Error 35389
''
''            'Se não encontrou a Filial Cliente
''            If lErro <> SUCESSO Then Error 35442
''
''        End If
''
''        'Se a Transportadora estiver preenchida
''        If objNFiscal.iCodTransportadora > 0 Then
''
''            'Lê a Transportadora
''            lErro = CF("Transportadora_Lock1",objNFiscal.iCodTransportadora)
''            If lErro <> SUCESSO And lErro <> 43038 Then Error 35392
''
''            'Se não encontrou a Transportadora --> erro
''            If lErro <> SUCESSO Then Error 35394
''
''        End If
''
''        'Se Possui Nota Fiscal original
''        If objNFiscal.lNumIntNotaOriginal > 0 Then
''
''            'Busca a Nota Fiscal Original
''            lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM NFiscal WHERE NumIntDoc = ? AND Status <> ? ", 0, lNumIntDoc, objNFiscal.lNumIntNotaOriginal, STATUS_EXCLUIDO)
''            If lErro <> AD_SQL_SUCESSO Then Error 35396
''
''            lErro = Comando_BuscarPrimeiro(alComando(2))
''            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35397
''            If lErro = AD_SQL_SEM_DADOS Then Error 35398 'Não encontrou
''
''            'Faz o "lock" na Nota Fiscal Original
''            lErro = Comando_LockShared(alComando(2))
''            If lErro <> AD_SQL_SUCESSO Then Error 35399
''
''        End If
''
''        'Faz os "locks" dos dados ligados aos Ítens da Nota Fiscal
''        lErro = CF("ItensNFiscalEntDev_Lock_Gravacao",objNFiscal)
''        If lErro <> SUCESSO Then Error 35414
''
''        'Faz a Inserção da Nova Nota Fiscal Entrada-Devolução no BD
''        lErro = CF("NFiscalEntradaDevolucao_Grava_BD",objNFiscal)
''        If lErro <> SUCESSO Then Error 35415
''
''        objNFiscal.iTipoNFiscal = iTipoNFOriginal
''
''        'Grava a Estatística do Produto com Relação a devolução
''        lErro = CF("ProdutoVendido_Grava_EstatisticasDev",objNFiscal)
''        If lErro <> SUCESSO Then Error 52961
''
''    End If
''
''    If objNFiscal.lFornecedor > 0 Then
''
''        'Grava os dados contábeis (contabilidade)
''        lErro = objContabil.Contabil_Gravar_Registro(objNFiscal.lNumIntDoc, objNFiscal.lFornecedor, objNFiscal.iFilialForn, NAO_AVALIA_PELO_CUSTO_REAL_PRODUCAO, AVISA_LANCAMENTOS_CONTABILIZADOS, objNFiscal.lNumNotaFiscal)
''        If lErro <> SUCESSO Then Error 36497
''
''    Else
''
''        'Grava os dados contábeis (contabilidade)
''        lErro = objContabil.Contabil_Gravar_Registro(objNFiscal.lNumIntDoc, objNFiscal.lCliente, objNFiscal.iFilialCli, NAO_AVALIA_PELO_CUSTO_REAL_PRODUCAO, AVISA_LANCAMENTOS_CONTABILIZADOS, objNFiscal.lNumNotaFiscal)
''        If lErro <> SUCESSO Then Error 36498
''
''    End If
''
''    'Faz o Commit na Transacao
''    lErro = Transacao_Commit()
''    If lErro <> AD_SQL_SUCESSO Then Error 35443
''
''    'Fecha os Comandos
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    NFiscalEntradaDevolucao_Grava = SUCESSO
''
''    Exit Function
''
''Erro_NFiscalEntradaDevolucao_Grava:
''
''    NFiscalEntradaDevolucao_Grava = Err
''
''    Select Case Err
''
''        Case 35090
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn)
''
''        Case 35378
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
''
''        Case 35379
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 35414, 35415, 36497, 36498, 39498, 42070, 42082, 52961
''
''        Case 42072
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)
''
''        Case 35389
''
''        Case 35392
''
''        Case 35394
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objNFiscal.iCodTransportadora)
''
''        Case 35396, 35397
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)
''
''        Case 35398
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", Err, objNFiscal.lNumNotaFiscal)
''
''        Case 35399
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", Err)
''
''        Case 35442
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objNFiscal.lCliente, objNFiscal.iFilialCli)
''
''        Case 35443
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
''
''        Case 60448 'Tratado na rotina chamada
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149486)
''
''    End Select
''
''    Call Transacao_Rollback
''    For iIndice = LBound(alComando) To UBound(alComando)
''        Call Comando_Fechar(alComando(iIndice))
''    Next
''
''    Exit Function
''
''End Function

Function NFiscalEntDevRem_Baixa_Recebimento(ByVal objNFiscal As ClassNFiscal) As Long
'Faz a baixa de um Recebimento

Dim alComando(0 To 4) As Long
Dim tNFiscal As typeNFiscal
Dim objMovEstoqueBaixa As ClassMovEstoque
Dim iItem As Integer
Dim iIndice As Integer
Dim lErro As Long
Dim tItemNF As typeItemNF
Dim lCodigo As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objResevaItem As ClassReservaItem
Dim objItemNF As ClassItemNF
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemMovEstoqueINF As ClassItemMovEstoque
Dim colMovEstoqueINF As Collection

On Error GoTo Erro_NFiscalEntDevRem_Baixa_Recebimento

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 35426
    Next

    With tNFiscal

        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
               
        'Produra no BD um Recebimento com os dados compatíveis aos da Nota Fiscal
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Status, Cliente, FilialCli, FilialEntrega, DataEmissao, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, NumIntDoc FROM NFiscal WHERE FilialEmpresa = ? AND NumRecebimento = ? AND (TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ?)", 0, .iStatus, .lCliente, _
        .iFilialCli, .iFilialEntrega, .dtDataEmissao, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.lNumRecebimento, DOCINFO_NRFF, DOCINFO_NRFP, DOCINFO_NRCC, DOCINFO_NRCP)
        If lErro <> AD_SQL_SUCESSO Then Error 35427

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35428
        
        'Se encontrar
        If lErro = AD_SQL_SUCESSO Then
        
            'Faz um "lock" no Recebimento
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> SUCESSO Then Error 35429
            
            'Insere o Recebimento em NFiscalBaixadas
            objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
            'Lê o Tipo da Nota Fiscal
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then Error 44992
    
            'Se não achou o Tipo de Documento --> erro
            If lErro = 31415 Then Error 44993
    
            'Se o TipoMovtoEstoque for positivo
            If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
                Set objMovEstoqueBaixa = New ClassMovEstoque
    
                'Obtém o Codigo Automático do Movimento de Estoque
                lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
                If lErro <> SUCESSO Then Error 35437
    
                objMovEstoqueBaixa.dtData = objNFiscal.dtDataEntrada
                objMovEstoqueBaixa.iFilialEmpresa = objNFiscal.iFilialEmpresa
                objMovEstoqueBaixa.iTipoMov = MOV_EST_BAIXA_RECEBIMENTO_MATERIAL
                objMovEstoqueBaixa.lCodigo = lCodigo
                objMovEstoqueBaixa.lCliente = objNFiscal.lCliente
                objMovEstoqueBaixa.lFornecedor = objNFiscal.lFornecedor
                objMovEstoqueBaixa.sDocOrigem = objTipoDocInfo.sSigla

                tItemNF.sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
                tItemNF.sProduto = String(STRING_PRODUTO, 0)
                tItemNF.sUnidadeMed = String(STRING_UM_SIGLA, 0)
    
                'Pesquisa os Ítens do Recebimento
                lErro = Comando_ExecutarPos(alComando(2), "SELECT Item,Status,Produto,UnidadeMed,Quantidade,Almoxarifado,PrecoUnitario,PercDesc,ValorDesconto,DataEntrega,DescricaoItem,ValorAbatComissao,NumIntPedVenda,NumIntItemPedVenda,NumIntDoc,NumIntTrib,NumIntDocOrig FROM ItensNFiscal WHERE NumIntNF = ? ORDER BY Item ", 0, tItemNF.iItem, tItemNF.iStatus, tItemNF.sProduto, tItemNF.sUnidadeMed, tItemNF.dQuantidade, tItemNF.iAlmoxarifado, tItemNF.dPrecoUnitario, tItemNF.dPercDesc, tItemNF.dValorDesconto, tItemNF.dtDataEntrega, tItemNF.sDescricaoItem, tItemNF.dValorAbatComissao, tItemNF.lNumIntPedVenda, tItemNF.lNumIntItemPedVenda, tItemNF.lNumIntDoc, tItemNF.lNumIntTrib, tItemNF.lNumIntDocOrig, tNFiscal.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then Error 35431
    
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35432
                'Se encontrar
                Do While lErro <> AD_SQL_SEM_DADOS
                    'Faz o "lock"
                    lErro = Comando_LockExclusive(alComando(2))
                    If lErro <> AD_SQL_SUCESSO Then Error 35434
    
                    Set objItemNF = New ClassItemNF
                    
                    objItemNF.lNumIntDoc = tItemNF.lNumIntDoc
    
                    lErro = CF("ItensNFGrade_Le", objItemNF)
                    If lErro <> SUCESSO Then Error 35139
    
                    If objItemNF.iPossuiGrade <> MARCADO Then
                       
                        '#########################################
                        'Inserido por Wagner - 08/07/2008
                        Set colMovEstoqueINF = New Collection
                        Set objItemMovEstoqueINF = New ClassItemMovEstoque
                        
                        objItemMovEstoqueINF.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                        objItemMovEstoqueINF.lNumIntDocOrigem = tItemNF.lNumIntDoc
                        objItemMovEstoqueINF.iFilialEmpresa = objNFiscal.iFilialEmpresa
                        
                        lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoqueINF, colMovEstoqueINF)
                        If lErro <> SUCESSO Then Error 35139
                        
                        For Each objItemMovEstoqueINF In colMovEstoqueINF
                            Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (objItemMovEstoqueINF.dQuantidade * tItemNF.dPrecoUnitario) - tItemNF.dValorDesconto, 0, tItemNF.sProduto, tItemNF.sDescricaoItem, objItemMovEstoqueINF.sSiglaUM, objItemMovEstoqueINF.dQuantidade, objItemMovEstoqueINF.iAlmoxarifado, "", tItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                        Next
                        '#########################################
                            
'                        'Atribui os ítens ao Movimento
'                        With tItemNF
'
'                            Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (.dQuantidade * .dPrecoUnitario) - .dValorDesconto, 0, .sProduto, .sDescricaoItem, .sUnidadeMed, .dQuantidade, .iAlmoxarifado, "", .lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
'                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'
'                        End With
                    Else
                        For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                            Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (((objItemRomaneio.dQuantidade * tItemNF.dPrecoUnitario) - (tItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / tItemNF.dQuantidade))) * objItemRomaneio.dQuantidade), 0, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, tItemNF.sUnidadeMed, objItemRomaneio.dQuantidade, objItemRomaneio.iAlmoxarifado, "", objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                        Next
                    End If

                    'Atualiza o status do ítem para baixado
                    lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensNFiscal SET Status = ?", alComando(2), STATUS_BAIXADO)
                    If lErro <> AD_SQL_SUCESSO Then Error 35438
    
                    lErro = Comando_BuscarProximo(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35433
    
                Loop
    
                lErro = Comando_ExecutarPos(alComando(4), "UPDATE NFiscal SET Status = ?, Serie = ?, NumNotaFiscal = ?, DataAtualizacao=?", alComando(0), STATUS_BAIXADO, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, gdtDataHoje)
                If lErro <> AD_SQL_SUCESSO Then Error 35439
    
                'Gera o Movimento de Baixa de Recebimento
                lErro = CF("MovEstoque_Grava", objMovEstoqueBaixa)
                If lErro <> SUCESSO Then Error 35436

            End If

        End If

    End With

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalEntDevRem_Baixa_Recebimento = SUCESSO

    Exit Function

Erro_NFiscalEntDevRem_Baixa_Recebimento:

    NFiscalEntDevRem_Baixa_Recebimento = Err

    Select Case Err

        Case 35139, 35436, 35437, 44992

        Case 35426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 35427, 35428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", Err, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 35429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", Err)

        Case 35430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)

        Case 35431, 35432, 35433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", Err)

        Case 35434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSNFISCAL", Err)

        Case 35435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSNFISCAL", objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 35438
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSNFISCAL", Err, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, tItemNF.iItem)

        Case 35439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)

        Case 44993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149487)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalEntrada_Verifica_Existencia2(ByVal objNFiscal As ClassNFiscal, ByVal objTipoDocInfo As ClassTipoDocInfo, Optional ByVal bSemAviso As Boolean = False) As Long
'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
'OBS: Da aviso que vai gravar os dados contabeis

Dim lErro As Long, lErro1 As Long
Dim lNumIntDoc As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lComando As Long
Dim objNFiscalBD As New ClassNFiscal
Dim dtDataEmissao As Date

On Error GoTo Erro_NFiscalEntrada_Verifica_Existencia2

    If bSemAviso Then vbMsgRes = vbYes
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 61412
    
    lErro = CF("NFiscal_Testa_Numero", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    
    'Verifica se tem Nota Fiscal Igual (Série, Número, DataEmissão)
    lErro = CF("NFiscalEntrada_Verifica_Igual", objNFiscal, objTipoDocInfo.iTipo, lComando, objNFiscalBD)
    If lErro <> SUCESSO And lErro <> 35384 And lErro <> 42083 Then gError 61413
        
'    'se a nota já existir e for de importação e estiver com a data de entrada gravada nula ==> erro
'    If lErro <> SUCESSO And objTipoDocInfo.iCodigo = DOCINFO_NFIEIMP And objNFiscalBD.dtDataEntrada = DATA_NULA Then gError 89149
'
'    lErro1 = CF("NFiscalEntrada_Verifica_Existencia2_Cust", lErro, objNFiscal, vbMsgRes, bSemAviso)
'    If lErro1 <> SUCESSO Then gError 61413
'
'    If lErro <> SUCESSO Then
'        If vbMsgRes = vbYes Then
'            gError 61414
'        ElseIf vbMsgRes = vbNo Then
'            gError 61415
'        End If
'    End If
        
    'se a nota ja existir
    If lErro <> SUCESSO Then
    
        'verifica se a nota fiscal foi processada pela sefaz ou se foi enviada e se confirma a alteracao
        lErro = CF("NFiscal_Testa_Existencia1", objNFiscal, bSemAviso)
        If lErro <> SUCESSO And lErro <> 207565 And lErro <> 207567 And lErro <> 207624 And lErro <> 207626 Then gError 207638
        
        'precisa excluir a nota e inclui-la com o mesmo numero
        gError 61414
    
    End If
        
    'Verifica se tem Nota na data Proxima. Se tiver pergunta se deseja realizar a alteração da data de emissão.
    lErro = NFiscalEntrada_Verifica_DataProxima(objNFiscal, objTipoDocInfo.iTipo, lComando, dtDataEmissao)
    If lErro <> SUCESSO And lErro <> 89810 And lErro <> 89811 Then gError 61416
    
    'se a nota ja existir
    If lErro <> SUCESSO Then
    
        'verifica se a nota fiscal foi processada pela sefaz ou se foi enviada e se confirma a alteracao
        lErro = CF("NFiscal_Testa_Existencia1", objNFiscal, bSemAviso)
        If lErro <> SUCESSO And lErro <> 207565 And lErro <> 207567 And lErro <> 207624 And lErro <> 207626 Then gError 207639
        
        'precisa excluir a nota e inclui-la com o mesmo numero
        gError 89723
    
    End If
    
'    'se encontrou uma nota em data proxima ==> pergunta se deseja prosseguir com a alteracao
'    If lErro = 89810 Or lErro = 89811 Then
'
'        If lErro = 89810 Then
'
'            'Pergunta se deseja alterar uma nota com dados parecidos aos da já existentes
'            If Not bSemAviso Then vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NF_EXTERNA_DATA_PROXIMA2", objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, dtDataEmissao)
'
'        Else
'
'            'Pergunta se deseja alterar uma nota com dados parecidos aos da já existentes
'            If Not bSemAviso Then vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NF_INTERNA_DATA_PROXIMA2", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, dtDataEmissao)
'
'        End If
'
'        'se desejar alterar ==> retorna erro
'        If vbMsgRes = vbYes Then gError 89723
'
'        'Se não --> Erro
'        If vbMsgRes = vbNo Then gError 89812
'
'    End If

    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
        'Verifica se a numeração não foi inutilizada
        lErro = CF("NFeInutFaixa_Valida2", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    Call Comando_Fechar(lComando)

    NFiscalEntrada_Verifica_Existencia2 = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Verifica_Existencia2:

    NFiscalEntrada_Verifica_Existencia2 = gErr

    Select Case gErr
    
        Case 35382, 35383, 35385, 35386, 42079, 42081, 42085, 42092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)

        Case 61414 'Tratar a Gravação da Contabilidade
            
        Case 61411, 61412, 61413, 61415, 61416, 89149, 89723, 89812, 207638, 207639
        
        Case 55642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149488)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function NFiscalEntrada_Verifica_Existencia(ByVal objNFiscal As ClassNFiscal, ByVal iTipo As Integer) As Long
'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima

Dim lErro As Long
Dim lNumIntDoc As Long
Dim vbMsgRes As VbMsgBoxResult
Dim lComando As Long
Dim dtDataEmissao As Date
Dim objNFiscalBD As New ClassNFiscal

On Error GoTo Erro_NFiscalEntrada_Verifica_Existencia

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 55642
    
    'Verifica se tem Nota Fiscal Igual (Série, Número, DataEmissão)
    lErro = CF("NFiscalEntrada_Verifica_Igual", objNFiscal, iTipo, lComando, objNFiscalBD)
    If lErro <> SUCESSO And lErro <> 35384 And lErro <> 42083 Then gError 61411
        
    'Nota Fiscal Externa Igual
    If lErro = 35384 Then gError 61409
    
    'Nota Fiscal Interna Igual
    If lErro = 42083 Then gError 61410
    
    'Verifica se tem Nota na data Proxima
    lErro = NFiscalEntrada_Verifica_DataProxima(objNFiscal, iTipo, lComando, dtDataEmissao)
    If lErro <> SUCESSO And lErro <> 89810 And lErro <> 89811 Then gError 61412
    
    'Se tiver uma nota fiscal externa com mesmo numero e data proxima ==> erro
    If lErro = 89810 Then gError 89813
    
    'Se tiver uma nota fiscal interno com mesmo numero e data proxima ==> erro
    If lErro = 89811 Then gError 89814
    
    Call Comando_Fechar(lComando)

    NFiscalEntrada_Verifica_Existencia = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Verifica_Existencia:

    NFiscalEntrada_Verifica_Existencia = gErr

    Select Case gErr
    
        Case 35382, 35383, 35385, 35386, 42079, 42081, 42085, 42092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)

        Case 61409
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_NFISCAL_EXTERNA2", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
            
        Case 61410
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_NFISCAL_INTERNA", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
    
        Case 61411, 61412
        
        Case 89813
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_NFISCAL_EXTERNA3", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, dtDataEmissao)
            
        Case 89814
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_NFISCAL_INTERNA1", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, dtDataEmissao)
        
        Case 55642
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149489)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function NFiscalEntrada_Verifica_Igual(ByVal objNFiscal As ClassNFiscal, ByVal iTipo As Integer, ByVal lComando As Long, ByVal objNFiscalBD As ClassNFiscal) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim vbMsgRes
Dim dtDataEntrada As Date

On Error GoTo Erro_NFiscalEntrada_Verifica_Igual

    If iTipo = DOCINFO_NF_EXTERNA Then

        'Busca Nota Fiscal com os mesmos dados contidos em objNFiscal
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEntrada FROM TipoDocInfoNFiscal WHERE Fornecedor = ? AND FilialForn = ? AND Cliente = ? AND FilialCli = ? AND DataEmissao = ? AND Serie = ? AND NumNotaFiscal = ? AND Status <> ? AND Tipo = ? ", lNumIntDoc, dtDataEntrada, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, STATUS_EXCLUIDO, DOCINFO_NF_EXTERNA)
        If lErro <> AD_SQL_SUCESSO Then Error 35382
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35383
        
        If lErro = AD_SQL_SUCESSO Then
        
            objNFiscal.lNumIntDoc = lNumIntDoc
            objNFiscalBD.dtDataEntrada = dtDataEntrada
        
             Error 35384 'Se encontrar --> Erro

        End If

    ElseIf iTipo = DOCINFO_NF_INT_ENTRADA Then

        'Busca Nota Fiscal com os mesmos dados contidos em objNFiscal
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEntrada FROM TipoDocInfoNFiscal WHERE FilialEmpresa = ? AND Serie = ? AND NumNotaFiscal = ? AND DataEmissao=? AND Status <> ? AND (Tipo = ? OR Tipo =?) ", lNumIntDoc, dtDataEntrada, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, STATUS_EXCLUIDO, DOCINFO_NF_INT_ENTRADA, DOCINFO_NF_INT_SAIDA)
        If lErro <> AD_SQL_SUCESSO Then Error 42079
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42081
        
        If lErro = AD_SQL_SUCESSO Then
        
            objNFiscal.lNumIntDoc = lNumIntDoc
            
            objNFiscalBD.dtDataEntrada = dtDataEntrada
            
            Error 42083 'Se encontrar --> Erro

        End If

    End If

    NFiscalEntrada_Verifica_Igual = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Verifica_Igual:

    NFiscalEntrada_Verifica_Igual = Err

    Select Case Err
    
        Case 35382, 35383, 42079, 42081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", Err)

        Case 35384 'Externa Tratar erro de Nota Fiscal Igual
            
        Case 42083 'Interna Tratar erro de Nota Fiscal Igual
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149490)
    
    End Select
    
    Exit Function

End Function

Private Function NFiscalEntrada_Verifica_DataProxima(ByVal objNFiscal As ClassNFiscal, ByVal iTipo As Integer, ByVal lComando As Long, ByVal dtDataEmissao As Date) As Long
'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima

Dim lErro As Long
Dim vbMsgRes
Dim lNumIntDoc As Long

On Error GoTo Erro_NFiscalEntrada_Verifica_DataProxima

    If iTipo = DOCINFO_NF_EXTERNA Then

        'Procura Nota Fiscal com os mesmos dados gerada em um periodo de emissao
        lErro = Comando_Executar(lComando, "SELECT DataEmissao, NumIntDoc FROM TipoDocInfoNFiscal WHERE FilialEmpresa = ? AND Fornecedor = ? AND FilialForn = ? AND Cliente = ? AND FilialCli = ? AND DataEmissao > ? AND DataEmissao < ? AND Serie = ? AND NumNotaFiscal = ? AND Status <> ? AND Tipo = ? ", dtDataEmissao, lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao - PERIODO_EMISSAO, objNFiscal.dtDataEmissao + PERIODO_EMISSAO, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, STATUS_EXCLUIDO, DOCINFO_NF_EXTERNA)
        If lErro <> AD_SQL_SUCESSO Then gError 35385
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 35386
        
        'Se encontrar
        If lErro = AD_SQL_SUCESSO Then
            objNFiscal.lNumIntDoc = lNumIntDoc
            gError 89810
        End If
        
    ElseIf iTipo = DOCINFO_NF_INT_ENTRADA Then

        'Procura Nota Fiscal com os mesmos dados gerada em um periodo de emissao
        lErro = Comando_Executar(lComando, "SELECT DataEmissao, NumIntDoc FROM TipoDocInfoNFiscal WHERE FilialEmpresa = ? AND DataEmissao > ? AND DataEmissao < ? AND Serie = ? AND NumNotaFiscal = ? AND Status <> ? AND (Tipo = ? OR Tipo = ?)", dtDataEmissao, lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.dtDataEmissao - PERIODO_EMISSAO, objNFiscal.dtDataEmissao + PERIODO_EMISSAO, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, STATUS_EXCLUIDO, DOCINFO_NF_INT_ENTRADA, DOCINFO_NF_INT_SAIDA)
        If lErro <> AD_SQL_SUCESSO Then gError 42085
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 42092
        
        'Se encontrar
        If lErro = AD_SQL_SUCESSO Then
            objNFiscal.lNumIntDoc = lNumIntDoc
            gError 89811
        End If
        
    End If

    NFiscalEntrada_Verifica_DataProxima = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Verifica_DataProxima:

    NFiscalEntrada_Verifica_DataProxima = gErr

    Select Case gErr
    
        Case 35385, 35386, 42085, 42092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)

        Case 89810, 89811

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149491)
    
    End Select
    
    Exit Function
    
End Function

Public Sub TipoNFiscal_Correspondente_Recebimento(ByVal objTipoDocInfo As ClassTipoDocInfo, iTipoRecebimento As Integer)

    If objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA And objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        iTipoRecebimento = DOCINFO_NRCP
    ElseIf objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA And objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        iTipoRecebimento = DOCINFO_NRFP
    ElseIf objTipoDocInfo.iTipo = DOCINFO_NF_EXTERNA And objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        iTipoRecebimento = DOCINFO_NRCC
    ElseIf objTipoDocInfo.iTipo = DOCINFO_NF_EXTERNA And objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        iTipoRecebimento = DOCINFO_NRFF
    End If

End Sub

''Function NFiscalFatEntrada_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
''
''Dim lErro As Long
''Dim objTipoDocInfo As New ClassTipoDocInfo
''Dim lTransacao As Long
''Dim lComando As Long
''Dim lNumIntDoc As Long
''Dim vbMsgRes As VbMsgBoxResult
''Dim dtDataEmissao As Date
''Dim iIndice As Integer
''Dim sDescricao As String
''
''On Error GoTo Erro_NFiscalFatEntrada_Grava
''
''    lTransacao = Transacao_Abrir()
''    If lTransacao = 0 Then Error 42665
''
''    lComando = Comando_Abrir()
''    If lComando = 0 Then Error 42666
''
''    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
''
''    lErro = CF("TipoDocInfo_Le_Codigo",objTipoDocInfo)
''    If lErro <> SUCESSO And lErro <> 31415 Then Error 42667
''
''    If lErro <> SUCESSO Then Error 42668
''
''    'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
''    lErro = CF("NFiscalEntrada_Verifica_Existencia2",objNFiscal, objTipoDocInfo.iTipo)
''    If lErro <> SUCESSO And lErro <> 61414 Then Error 55645
''
''    'Se for uma alteração só altera os dados da contabilidade
''    If lErro = SUCESSO Then
''
''        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
''
''            'gera o Número para a Nota Fiscal Interna
''            lErro = CF("NFiscalNumAuto",objNFiscal)
''            If lErro <> SUCESSO Then Error 60446
''
''        End If
''
''        lErro = CF("NFiscalFatEntrada_Lock_Gravacao",objNFiscal)
''        If lErro <> SUCESSO Then Error 42682
''
''        lErro = CF("NFiscalFatEntrada_Grava_BD",objNFiscal)
''        If lErro <> SUCESSO Then Error 42683
''
''        'Grava a Estatística da Filial Fornecedor
''        lErro = CF("FilialFornecedor_Grava_Estatisticas",objNFiscal)
''        If lErro <> SUCESSO Then Error 52959
''
''    End If
''
''    'Grava os dados contábeis (contabilidade)
''    lErro = objContabil.Contabil_Gravar_Registro(objNFiscal.lNumIntDoc, objNFiscal.lFornecedor, objNFiscal.iFilialForn, NAO_AVALIA_PELO_CUSTO_REAL_PRODUCAO, AVISA_LANCAMENTOS_CONTABILIZADOS, objNFiscal.lNumNotaFiscal)
''    If lErro <> SUCESSO Then Error 36501
''
''    lErro = Transacao_Commit()
''    If lErro <> AD_SQL_SUCESSO Then Error 42684
''
''    Call Comando_Fechar(lComando)
''
''    NFiscalFatEntrada_Grava = SUCESSO
''
''    Exit Function
''
''Erro_NFiscalFatEntrada_Grava:
''
''    NFiscalFatEntrada_Grava = Err
''
''    Select Case Err
''
''        Case 42665
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
''
''        Case 42666
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 42668
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err)
''
''        Case 42671, 42674
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", Err)
''
''        Case 36501, 42667, 42682, 42683, 52959, 55645
''
''        Case 42684
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
''
''        Case 60446 'Tratado na rotina chamada
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149492)
''
''    End Select
''
''    Call Transacao_Rollback
''    Call Comando_Fechar(lComando)
''
''    Exit Function
''
''End Function
    
Function NFiscalFatEntrada_Baixa_Recebimento(ByVal objNFiscal As ClassNFiscal) As Long
'Faz a baixa de um Recebimento

Dim alComando(0 To 4) As Long
Dim tNFiscal As typeNFiscal
Dim objMovEstoqueBaixa As ClassMovEstoque
Dim iItem As Integer
Dim iIndice As Integer
Dim lErro As Long
Dim tItemNF As typeItemNF
Dim lCodigo As Long
Dim objProduto As New ClassProduto
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iTipoNFiscal As Integer
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objResevaItem As ClassReservaItem
Dim objItemNF As ClassItemNF
Dim objItemMovEstoqueINF As ClassItemMovEstoque
Dim colMovEstoqueINF As Collection
Dim dFator As Double, dCusto As Double

On Error GoTo Erro_NFiscalFatEntrada_Baixa_Recebimento
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEFN Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEFN Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEFBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEFBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEFBEN Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEFBEN Then
    
        'Abre os Comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then Error 42723
        Next
    
        Select Case objNFiscal.iTipoNFiscal
        
            Case DOCINFO_NFIEFN, DOCINFO_NFIEFBF, DOCINFO_NFIEFBEN
                iTipoNFiscal = DOCINFO_NRFP
            Case DOCINFO_NFEEFN, DOCINFO_NFEEFBF, DOCINFO_NFEEFBEN
                iTipoNFiscal = DOCINFO_NRFF
                
        End Select
    
        With tNFiscal
    
            .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
            .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
            .sPlaca = String(STRING_NFISCAL_PLACA, 0)
            .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
            .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
        
            'Procura no BD um Recebimento com os dados compatíveis aos da Nota Fiscal
            lErro = Comando_ExecutarPos(alComando(0), "SELECT DataEntrada, Status, Cliente, FilialCli, FilialEntrega, DataEmissao, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, NumIntDoc FROM NFiscal WHERE FilialEmpresa = ? AND NumRecebimento = ? AND (TipoNFiscal = ? OR TipoNFiscal = ?)", 0, .dtDataEntrada, .iStatus, .lCliente, _
            .iFilialCli, .iFilialEntrega, .dtDataEmissao, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.lNumRecebimento, DOCINFO_NRFF, DOCINFO_NRFP)
            If lErro <> AD_SQL_SUCESSO Then Error 42724
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42725
            'Se encontrar
            If lErro = AD_SQL_SUCESSO Then
                'Faz um "lock" no Recebimento
                lErro = Comando_LockExclusive(alComando(0))
                If lErro <> SUCESSO Then Error 42726

                objTipoDocInfo.iCodigo = iTipoNFiscal
        
                'Lê o Tipo da Nota Fiscal
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then Error 44994
        
                'Se não achou o Tipo de Documento --> erro
                If lErro = 31415 Then Error 44995
        
                'Se o TipoMovtoEstoque for positivo
                If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
                    Set objMovEstoqueBaixa = New ClassMovEstoque
        
                    'Obtém o Codigo Automático do Movimento de Estoque
                    lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
                    If lErro <> SUCESSO Then Error 42728
        
                    objMovEstoqueBaixa.dtData = objNFiscal.dtDataEntrada
                    objMovEstoqueBaixa.iFilialEmpresa = objNFiscal.iFilialEmpresa
                    objMovEstoqueBaixa.iTipoMov = MOV_EST_BAIXA_RECEBIMENTO_MATERIAL
                    objMovEstoqueBaixa.lCodigo = lCodigo
                    objMovEstoqueBaixa.lCliente = objNFiscal.lCliente
                    objMovEstoqueBaixa.lFornecedor = objNFiscal.lFornecedor
                    objMovEstoqueBaixa.sDocOrigem = objTipoDocInfo.sSigla
        
                    tItemNF.sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
                    tItemNF.sProduto = String(STRING_PRODUTO, 0)
                    tItemNF.sUnidadeMed = String(STRING_UM_SIGLA, 0)
        
                    'Pesquisa os Ítens do Recebimento
                    lErro = Comando_ExecutarPos(alComando(2), "SELECT Item,Status,Produto,UnidadeMed,Quantidade,Almoxarifado,PrecoUnitario,PercDesc,ValorDesconto,DataEntrega,DescricaoItem,ValorAbatComissao,NumIntPedVenda,NumIntItemPedVenda,NumIntDoc,NumIntTrib,NumIntDocOrig FROM ItensNFiscal WHERE NumIntNF = ? ORDER BY Item ", 0, tItemNF.iItem, tItemNF.iStatus, tItemNF.sProduto, tItemNF.sUnidadeMed, tItemNF.dQuantidade, tItemNF.iAlmoxarifado, tItemNF.dPrecoUnitario, tItemNF.dPercDesc, tItemNF.dValorDesconto, tItemNF.dtDataEntrega, tItemNF.sDescricaoItem, tItemNF.dValorAbatComissao, tItemNF.lNumIntPedVenda, tItemNF.lNumIntItemPedVenda, tItemNF.lNumIntDoc, tItemNF.lNumIntTrib, tItemNF.lNumIntDocOrig, tNFiscal.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then Error 42729
        
                    lErro = Comando_BuscarPrimeiro(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42730
                    'Se encontrar
                    Do While lErro <> AD_SQL_SEM_DADOS
                        'Faz o "lock"
                        lErro = Comando_LockExclusive(alComando(2))
                        If lErro <> AD_SQL_SUCESSO Then Error 42731
        
                        Set objItemNF = New ClassItemNF
                        
                        objItemNF.lNumIntDoc = tItemNF.lNumIntDoc
        
                        lErro = CF("ItensNFGrade_Le", objItemNF)
                        If lErro <> SUCESSO Then Error 35139
        
                        If objItemNF.iPossuiGrade <> MARCADO Then
            
                            objProduto.sCodigo = tItemNF.sProduto
                        
                            lErro = CF("Produto_Le", objProduto)
                            If lErro <> SUCESSO And lErro <> 28030 Then Error 42732
                            If lErro <> SUCESSO Then Error 42733
                            
                            tItemNF.iControleEstoque = objProduto.iControleEstoque
                            
                            '#########################################
                            'Inserido por Wagner - 08/07/2008
                            Set colMovEstoqueINF = New Collection
                            Set objItemMovEstoqueINF = New ClassItemMovEstoque
                            
                            objItemMovEstoqueINF.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                            objItemMovEstoqueINF.lNumIntDocOrigem = tItemNF.lNumIntDoc
                            objItemMovEstoqueINF.iFilialEmpresa = objNFiscal.iFilialEmpresa
                            
                            lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoqueINF, colMovEstoqueINF)
                            If lErro <> SUCESSO Then Error 35139
                            
                            For Each objItemMovEstoqueINF In colMovEstoqueINF
                                If UCase(objItemMovEstoqueINF.sSiglaUM) <> UCase(tItemNF.sUnidadeMed) Then
                                
                                    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemMovEstoqueINF.sSiglaUM, tItemNF.sUnidadeMed, dFator)
                                    If lErro <> SUCESSO Then gError 35139
                                
                                    dCusto = (objItemMovEstoqueINF.dQuantidade * (tItemNF.dPrecoUnitario * dFator)) - ((tItemNF.dValorDesconto / tItemNF.dQuantidade) * (objItemMovEstoqueINF.dQuantidade * dFator))
                                
                                Else
                                    dCusto = (objItemMovEstoqueINF.dQuantidade * tItemNF.dPrecoUnitario) - ((tItemNF.dValorDesconto / tItemNF.dQuantidade) * objItemMovEstoqueINF.dQuantidade)
                                End If
                            
                                'Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (objItemMovEstoqueINF.dQuantidade * tItemNF.dPrecoUnitario) - tItemNF.dValorDesconto, 0, tItemNF.sProduto, tItemNF.sDescricaoItem, objItemMovEstoqueINF.sSiglaUM, objItemMovEstoqueINF.dQuantidade, objItemMovEstoqueINF.iAlmoxarifado, "", tItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                                Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, dCusto, 0, tItemNF.sProduto, tItemNF.sDescricaoItem, objItemMovEstoqueINF.sSiglaUM, objItemMovEstoqueINF.dQuantidade, objItemMovEstoqueINF.iAlmoxarifado, "", tItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                            Next
                            '#########################################
        
'                            'Atribui os ítens ao Movimento
'                            With tItemNF
'
'                                Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (.dQuantidade * .dPrecoUnitario) - .dValorDesconto, 0, .sProduto, .sDescricaoItem, .sUnidadeMed, .dQuantidade, .iAlmoxarifado, "", .lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
'                                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'
'                            End With
                        Else
                            For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                                
                                Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (((objItemRomaneio.dQuantidade * tItemNF.dPrecoUnitario) - (tItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / tItemNF.dQuantidade))) * objItemRomaneio.dQuantidade), 0, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, tItemNF.sUnidadeMed, objItemRomaneio.dQuantidade, objItemRomaneio.iAlmoxarifado, "", objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            
                            Next
                        End If
                        
                        'Atualiza o status dos itens da Nota para baixado
                        lErro = Comando_ExecutarPos(alComando(4), "UPDATE ItensNFiscal SET Status = ?", alComando(2), STATUS_BAIXADO)
                        If lErro <> AD_SQL_SUCESSO Then Error 42735
        
                        lErro = Comando_BuscarProximo(alComando(2))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42736
        
                    Loop
        
                    'Atualiza o status da nota para baixado
                    lErro = Comando_ExecutarPos(alComando(4), "UPDATE NFiscal SET Status = ? , Serie = ?, NumNotaFiscal = ?, DataAtualizacao=?", alComando(0), STATUS_BAIXADO, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, gdtDataHoje)
                    If lErro <> AD_SQL_SUCESSO Then Error 42737
                    
                    'Gera o Movimento de Baixa de Recebimento
                    lErro = CF("MovEstoque_Grava", objMovEstoqueBaixa)
                    If lErro <> SUCESSO Then Error 42738
        
                End If
                
            End If
            
        End With
    
        'Fecha os comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    NFiscalFatEntrada_Baixa_Recebimento = SUCESSO

    Exit Function

Erro_NFiscalFatEntrada_Baixa_Recebimento:

    NFiscalFatEntrada_Baixa_Recebimento = Err

    Select Case Err

        Case 35139, 42728, 42732, 42738, 44994

        Case 42723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 42724, 42725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL1", Err, iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 42726
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", Err)

        Case 42727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCALBAIXADAS", Err, tNFiscal.lNumIntDoc)

        Case 42729, 42730, 42736
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", Err)

        Case 42731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSNFISCAL", Err)

        Case 42733
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 42734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCALBAIXADAS", tNFiscal.lNumIntDoc)

        Case 42735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSNFISCAL", Err, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, tItemNF.iItem)

        Case 42737
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)

        Case 44995
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, iTipoNFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149493)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function MovEstoque_ReqProducao_Grava(ByVal objMovEstoque As ClassMovEstoque, ByVal objContabil As ClassContabil) As Long
'Grava o Movimento de Estoque de Requisição de Produção

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_MovEstoque_ReqProducao_Grava

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 34882

    lErro = CF("MovEstoque_ReqProducao_Grava0", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 105144
    
    'Confirmacao da transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 34888

    MovEstoque_ReqProducao_Grava = SUCESSO

    Exit Function

Erro_MovEstoque_ReqProducao_Grava:

    MovEstoque_ReqProducao_Grava = gErr

    Select Case gErr

        Case 34882
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 105144

        Case 34888
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149494)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    Exit Function

End Function

Function RecebMaterialC_Grava(ByVal objNFiscal As ClassNFiscal) As Long
'Grava a Nota Fiscal de Recebimento de Cliente

Dim lErro As Long
Dim lNumIntNFItem As Long
Dim lNumIntDocOrigem As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim objItemNF As ClassItemNF
Dim alComando(0 To 12) As Long
Dim dtData As Date
Dim iTipo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
'#################################
'Inserido por Wagner 26/01/2006
Dim dFator As Double
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dQuantItemNFUMEst As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim objProduto As ClassProduto
'#################################

On Error GoTo Erro_RecebMaterialC_Grava

    'Abre os comandos de Transação
    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 30716

    Next

    'Abre Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 30717

    'Faz lock dos objetos associados
    lErro = RecebMaterialC_Lock_ObjAssociados(objNFiscal, alComando(0), alComando(1))
    If lErro <> SUCESSO Then gError 30718

    If objNFiscal.iTipoNFiscal = DOCINFO_NRCP Then
        iTipo = DOCINFO_NF_INT_ENTRADA
    Else
        iTipo = DOCINFO_NF_EXTERNA
        
        'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
        lErro = CF("NFiscalEntrada_Verifica_Existencia", objNFiscal, iTipo)
        If lErro <> SUCESSO Then gError 55643
    
        If objNFiscal.lNumRecebimento = 0 Then
            'Verifica a Existencia de um Recebimento  com os Valores de Nota igual
            lErro = CF("RecebimentoC_Verifica_Existencia", objNFiscal)
            If lErro <> SUCESSO Then gError 61110
        End If
    
    End If
    
    dtData = objNFiscal.dtDataEmissao
    
    objNFiscal.dtDataEmissao = objNFiscal.dtDataEntrada
        
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo de documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 27623 Then gError 61041
        
    'Se não encontrar --> erro
    If lErro = 27623 Then gError 61042
        
    objNFiscal.dtDataEmissao = dtData
    
   '###################################################
    'Inserido por Wagner 25/01/2006
    iIndice = 0
    
    For Each objItemNF In objNFiscal.colItensNF
        
        iIndice = iIndice + 1
        
        If objItemNF.colItensRomaneioGrade.Count > 0 Then
            
            For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                Set objProduto = New ClassProduto
            
                objProduto.sCodigo = objItemRomaneio.sProduto
                
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 23080 Then gError 141600
                
                If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    
                    lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemRomaneio.sUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 141600
    
                    dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
    
                    dSomaAlocacao = 0
                                                                    
                    For Each objReservaItem In objItemRomaneio.colLocalizacao
                                            
                        dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                                                    
                    Next
                    
                    dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)
                    
                    If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > QTDE_ESTOQUE_DELTA Then
                    
                        If dSomaAlocacao < dQuantItemNFUMEst Then gError 141601
                        
                        If dSomaAlocacao > dQuantItemNFUMEst Then gError 141602
                    
                    End If
                    
                End If
                
            Next
            
        Else
                    
            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                If objItemNF.ColAlocacoes.Count = 0 Then gError 141603

                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 141604

                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
    
                dSomaAlocacao = 0
                                                                
                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
                    
                Next

                dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)
                
                If Abs(dSomaAlocacao - dQuantItemNFUMEst) > QTDE_ESTOQUE_DELTA Then
                 
                    If dSomaAlocacao < dQuantItemNFUMEst Then gError 141605
                    
                    If dSomaAlocacao > dQuantItemNFUMEst Then gError 141606
            
                End If
                
            End If
            
        End If
        
    Next
    '###################################################
    
    'Se não encontrar o Recebimento ==>> INSERÇÃO
    If objNFiscal.lNumRecebimento = 0 Then

        lErro = RecebMaterialC_Grava_Insercao(objNFiscal, alComando(4), alComando(5), alComando(2), alComando(11))
        If lErro <> SUCESSO Then gError 30728

    Else 'Se encontrar o Recebimento ==>> ALTERAÇÃO

        lErro = RecebMaterialC_Grava_Atualiza(objNFiscal, alComando(3), alComando(6), alComando(7), alComando(8), alComando(9), alComando(10), alComando(12))
        If lErro <> SUCESSO Then gError 30730

    End If

    'Confirma a Transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 30731

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    RecebMaterialC_Grava = SUCESSO

    Exit Function

Erro_RecebMaterialC_Grava:

    RecebMaterialC_Grava = gErr

    Select Case gErr

        Case 30716
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 30717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 30731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 30718, 30728, 30730, 55643, 61041, 61042, 61110

        Case 141601, 141605
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INCOMPLETA", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case 141602, 141606
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)
        
        Case 141603
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INEXISTENTE", gErr, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149495)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function RecebMaterialC_Lock_ObjAssociados(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Faz o Lock dos Objetos associados ao RecebMaterialC
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim iControleEstoque As Integer
Dim lCodCliente As Long
Dim iCodigo As Integer
Dim lNumIntDoc As Long
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim objItemRomaneioGrade As ClassItemRomaneioGrade

On Error GoTo Erro_Lock_ObjAssociados

    'Faz Lock Shared no FilialCliente
    lErro = CF("FilialCliente_Lock", objNFiscal.lCliente, objNFiscal.iFilialCli)
    If lErro <> SUCESSO And lErro <> 43050 Then Error 30719
    
    'Não achou a Filial do Cliente --> erro
    If lErro <> SUCESSO Then Error 30720

    'Verifica se o código da Transportadora foi preenchido
    If objNFiscal.iCodTransportadora <> 0 Then

        'Faz lock da Transportadora
        lErro = CF("Transportadora_Lock1", objNFiscal.iCodTransportadora)
        If lErro <> SUCESSO And lErro <> 43038 Then Error 30722
        
        'Não achou a Transportadora --> erro
        If lErro <> SUCESSO Then Error 30724

    End If

    'Se NumIntNotaOriginal estiver preenchido
    If objNFiscal.lNumIntNotaOriginal <> 0 Then

        'Seleciona no BD NFiscal a partir do NumIntDoc
        lErro = Comando_ExecutarLockado(lComando1, "SELECT NumIntDoc FROM NFiscal WHERE NumIntDoc = ?", lNumIntDoc, objNFiscal.lNumIntNotaOriginal)
        If lErro <> AD_SQL_SUCESSO Then Error 30726

        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30727

        If lErro = AD_SQL_SEM_DADOS Then Error 30728

        'Faz Lock shared na NFiscal
        lErro = Comando_LockShared(lComando1)
        If lErro <> AD_SQL_SUCESSO Then Error 30729

    End If

    'Para cada Item de ObjNFiscal.colItens
    For Each objItemNF In objNFiscal.colItensNF

        objProduto.sCodigo = objItemNF.sProduto

        'Faz lock do Produto
        lErro = CF("Produto_Lock", objProduto)
        If lErro <> SUCESSO Then Error 30899
        

        For Each objItemRomaneioGrade In objItemNF.colItensRomaneioGrade
            
            objProduto.sCodigo = objItemRomaneioGrade.sProduto
    
            'Faz lock do Produto
            lErro = CF("Produto_Lock", objProduto)
            If lErro <> SUCESSO Then Error 30899
        
            objItemRomaneioGrade.iControleEstoque = objProduto.iControleEstoque
        Next


        'Armazena o ControleEstoque do Produto
        objItemNF.iControleEstoque = objProduto.iControleEstoque
       
        If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
            
            '??? otimizar como em RecebeMaterialF
            
            'Faz o Lock Shared no Almoxarifado
''''            lErro = CF("Almoxarifado_Lock", objItemNF.iAlmoxarifado, lComando2)
''''            If lErro <> SUCESSO Then Error 25180
    
        End If
        
    Next

    RecebMaterialC_Lock_ObjAssociados = SUCESSO

    Exit Function

Erro_Lock_ObjAssociados:

    RecebMaterialC_Lock_ObjAssociados = Err

    Select Case Err

        Case 25180, 30722, 30525, 30899

        Case 30720
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objNFiscal.iFilialCli, objNFiscal.lCliente)

        Case 30721
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISCLIENTES", Err, objNFiscal.lCliente, objNFiscal.iFilialCli)

        Case 30724
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objNFiscal.iCodTransportadora)

        Case 30726, 30727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 30728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA", Err, objNFiscal.lNumIntNotaOriginal)

        Case 30729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", Err, objNFiscal.lNumIntNotaOriginal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149496)

    End Select

    Exit Function

End Function

'Subir para RotinasEST - classGrava (Incluso VolumeNumero e FreteRespons no INSERT de NFiscal)

Private Function RecebMaterialC_Grava_Insercao(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long) As Long
'Faz a Inserção do RecebMaterialC na Tabela NFiscal
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lCodigo As Long
Dim dtDataEntrada As Date
Dim vbMsg As VbMsgBoxResult
Dim objItemNF As ClassItemNF
Dim objMovEstoque As New ClassMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
'distribuicao
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dFator As Double
Dim iIndice As Integer
Dim dQuantItemNFUMEst As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim objProduto As ClassProduto

On Error GoTo Erro_RecebMaterialC_Grava_Insercao

        
    'Le o Número do Proximo Código
    lErro = CF("Recebimento_Automatico", objNFiscal)
    If lErro <> SUCESSO Then Error 61043
    
    'Lê o Número do Próximo Código Automático Interno.
    lErro = CF("NFiscal_Automatico", lNumIntDoc)
    If lErro <> SUCESSO Then Error 30735
    
    'Preenche o objNFiscal com Código
    objNFiscal.lNumIntDoc = lNumIntDoc

    'Insere no BD objNFiscal em NFiscal
    With objNFiscal
        lErro = Comando_Executar(lComando2, "INSERT INTO NFiscal(DataAtualizacao, NumRecebimento, NumIntDoc, FilialEmpresa, Serie, NumNotaFiscal, Status, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn,DataEntrada, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, VolumeNumero, FreteRespons, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, HoraEntrada, Moeda, TaxaMoeda, ChvNFe, HoraEmissao, DataCadastro, HoraCadastro, UsuCadastro) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", gdtDataHoje, objNFiscal.lNumRecebimento, objNFiscal.lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iStatus, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.iFilialEntrega, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada, objNFiscal.lNumPedidoVenda, _
            .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sVolumeNumero, .iFreteRespons, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, _
            .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, CDbl(.dtHoraEntrada), .iMoeda, .dTaxaMoeda, .sChvNFe, CDbl(IIf(.dtDataSaida <> DATA_NULA, IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraSaida, .dtHoraEmissao, .dtHoraSaida), IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraEntrada, .dtHoraEmissao, .dtHoraEntrada))), gdtDataHoje, CDbl(Time), gsUsuario)
    End With
    If lErro <> AD_SQL_SUCESSO Then Error 30736

    'Inserção de Itens da Nota Fiscal
    lErro = RecebimentoMatC_Insere_ItensNFiscal(objNFiscal, lComando3, lComando4)
    If lErro <> SUCESSO Then Error 30737

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 20744
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then Error 20745

    'Se o TipoMovtoEstoque for positivo
    If objTipoDocInfo.iTipoMovtoEstoque > 0 Then

        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then Error 30888
    
        'Preenche todos os Atributos exceto Código e colItens
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque
        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = objTipoDocInfo.sSigla & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
    
'mario distribuicao substituicao do codigo original acrescentando o tab de distribuicao
    
        'Para cada elemento de colItens de objNFiscal, adiciona elemento em colItens de objMovEstoque
        For Each objItemNF In objNFiscal.colItensNF
    
            iIndice = iIndice + 1
    
            If objItemNF.colItensRomaneioGrade.Count > 0 Then
    
                For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                    Set objProduto = New ClassProduto
                
                    objProduto.sCodigo = objItemRomaneio.sProduto
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 89606
                    
                    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    
                        lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemRomaneio.sUMEstoque, dFator)
                        If lErro <> SUCESSO Then gError 89606
                   
                        dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
                   
                        dSomaAlocacao = 0
                                                                        
                        For Each objReservaItem In objItemRomaneio.colLocalizacao
                                                
                            dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                   
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, objItemNF.sCcl, 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                            objItemMovEstoque.iControleEstoque = objItemRomaneio.iControleEstoque
                            objItemMovEstoque.sSiglaUMEst = objItemRomaneio.sUMEstoque
            
                        Next
                        
                        If dSomaAlocacao > dQuantItemNFUMEst Then gError 89607
                    End If
                Next
    
            ElseIf objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
           
                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 89606
           
                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
           
                dSomaAlocacao = 0
                                                                
                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
           
                    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemNF.dQuantidade * objItemNF.dPrecoUnitario) - objItemNF.dValorDesconto) * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, objItemNF.sCcl, 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                    objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                    objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                    objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
    
                Next
                
                If dSomaAlocacao > dQuantItemNFUMEst Then gError 89607
    
            End If
        Next
    
        'Chama função de Gravação de MovEstoque
        lErro = CF("MovEstoque_Grava", objMovEstoque) 'Chamada de dentro de Transação
        If lErro <> SUCESSO Then Error 30738

    End If
    
    RecebMaterialC_Grava_Insercao = SUCESSO

    Exit Function

Erro_RecebMaterialC_Grava_Insercao:

    RecebMaterialC_Grava_Insercao = Err

    Select Case Err

        Case 20744, 30735, 30738, 30737, 30734, 30888, 30748, 61043, 89606

        Case 20745
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)

        Case 30732, 30733
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 30736
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL", Err, objNFiscal.lNumNotaFiscal)

        Case 30749
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RECEB_MAT_CLI_NF", Err, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEntrada, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
        
        Case 89607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149497)

    End Select

    Exit Function

End Function

'Subir para RotinasEST - classGrava (Incluso VolumeNumero e FreteRespons no UPDATE de NFiscal)

Private Function RecebMaterialC_Grava_Atualiza(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long) As Long
'Faz a Atualização do RecebMaterialC na Tabela NFiscal
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lCodigo As Long
Dim iTipoDoc As Integer
Dim objItemNF As ClassItemNF
Dim objMovEstoque As ClassMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
'distribuicao
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dFator As Double
Dim iIndice As Integer
Dim dQuantItemNFUMEst As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim objProduto As ClassProduto

On Error GoTo Erro_RecebMaterialC_Grava_Atualiza

    'Testa se o Recebimento já foi baixado
    lErro = CF("Recebimento_Testa_Baixado", objNFiscal)
    If lErro <> SUCESSO Then gError 61044
    
    'Tenta ler o Recebimento na Tabela NFiscal
    lErro = Comando_ExecutarPos(lComando3, "SELECT NumIntDoc FROM NFiscal WHERE FilialEmpresa = ? AND NumRecebimento = ?", 0, lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.lNumRecebimento)
    If lErro <> AD_SQL_SUCESSO Then gError 30726

    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 30727

    'Se não encontrar o Recebimento --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 61045
    
    'Faz Lock Exclusive no Recebimento
    lErro = Comando_LockExclusive(lComando3)
    If lErro <> SUCESSO Then gError 30729
   
    objNFiscal.lNumIntDoc = lNumIntDoc
   
    'Faz Atualização do Recebimento
    lErro = Comando_ExecutarPos(lComando2, "UPDATE NFiscal SET ValorTotal = ?, ValorProdutos = ?, ValorSeguro = ?, ValorDesconto = ?, ValorFrete = ?, ValorOutrasDespesas = ?, CodTransportadora = ?, Placa = ?, PlacaUF = ?, VolumeQuant = ?, VolumeEspecie = ?, VolumeMarca = ?, VolumeNumero = ?, FreteRespons = ?, MensagemNota = ?, PesoLiq = ?, PesoBruto = ?, Serie = ? , NumNotaFiscal= ?, TipoNFiscal = ?, Cliente = ?, FilialCli = ?, DataEntrada = ?, HoraEntrada = ?, DataAtualizacao=?", lComando3, _
    objNFiscal.dValorTotal, objNFiscal.dValorProdutos, objNFiscal.dValorSeguro, objNFiscal.dValorDesconto, objNFiscal.dValorFrete, objNFiscal.dValorOutrasDespesas, objNFiscal.iCodTransportadora, objNFiscal.sPlaca, objNFiscal.sPlacaUF, objNFiscal.lVolumeQuant, objNFiscal.lVolumeEspecie, objNFiscal.lVolumeMarca, objNFiscal.sVolumeNumero, objNFiscal.iFreteRespons, objNFiscal.sMensagemNota, objNFiscal.dPesoLiq, objNFiscal.dPesoBruto, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iTipoNFiscal, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEntrada, CDbl(objNFiscal.dtHoraEntrada), gdtDataHoje)
    If lErro <> AD_SQL_SUCESSO Then gError 30744
    
    '==========================
    'EXCLUSÃO ITENS NOTA FISCAL
    '==========================
    
    lErro = RecebimentoMatC_Exclui_ItensNFiscal(objNFiscal, lComando4, lComando5)
    If lErro <> SUCESSO Then gError 30745

    'Inserção de Itens da Nota Fiscal
    lErro = RecebimentoMatC_Insere_ItensNFiscal(objNFiscal, lComando6, lComando7)
    If lErro <> SUCESSO Then gError 30746

    '====================
    'INSERÇÃO MOV ESTOQUE
    '====================
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 20742
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 20743
    
    'Se o TipoMovtoEstoque for positivo
    If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
        'Cria objMovEstoque
        Set objMovEstoque = New ClassMovEstoque
    
        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then gError 30889
    
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque
        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = objTipoDocInfo.sSigla & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
    
'inicio distribuicao
        'Para cada elemento de colItens de objNFiscal, adiciona elemento em colItens de objMovEstoque
        For Each objItemNF In objNFiscal.colItensNF
    
            iIndice = iIndice + 1
    
            If objItemNF.colItensRomaneioGrade.Count > 0 Then
            
                For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                    Set objProduto = New ClassProduto
                
                    objProduto.sCodigo = objItemRomaneio.sProduto
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 89606
                    
                    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    
                        lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemRomaneio.sUMEstoque, dFator)
                        If lErro <> SUCESSO Then gError 89606
                   
                        dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
                   
                        dSomaAlocacao = 0
                                                                        
                        For Each objReservaItem In objItemRomaneio.colLocalizacao
                                                
                            dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                   
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, objItemNF.sCcl, 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                            objItemMovEstoque.iControleEstoque = objItemRomaneio.iControleEstoque
                            objItemMovEstoque.sSiglaUMEst = objItemRomaneio.sUMEstoque
            
                        Next
                        
                        If dSomaAlocacao > dQuantItemNFUMEst Then gError 89607
                    End If
                Next
    
            ElseIf objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
               
                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 89608

                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
           
                dSomaAlocacao = 0
                                                                
                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
           
                    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemNF.dQuantidade * objItemNF.dPrecoUnitario) - objItemNF.dValorDesconto) * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, objItemNF.sCcl, 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                    objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                    objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                    objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
    
                Next
                
                If dSomaAlocacao > dQuantItemNFUMEst Then gError 89609
    
            End If
        Next
'fim distribuicao
    
        'Chama função de Gravação de MovEstoque
        lErro = CF("MovEstoque_Grava", objMovEstoque) 'Chamada de dentro de Transação
        If lErro <> SUCESSO Then gError 30550

    End If
    
    RecebMaterialC_Grava_Atualiza = SUCESSO

    Exit Function

Erro_RecebMaterialC_Grava_Atualiza:

    RecebMaterialC_Grava_Atualiza = gErr

    Select Case gErr

        Case 20742, 30550, 30746, 30745, 30889, 61044, 89608

        Case 20743
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 30725
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RECEB_MAT_CLI_NF", gErr, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEntrada, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
        
        Case 30726, 30727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 30729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 30744
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", gErr, objNFiscal.lNumNotaFiscal)
        
        Case 61045
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NAO_CADASTRADO", gErr, objNFiscal.lNumNotaFiscal)
        
        Case 89609
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149498)

    End Select

    Exit Function

End Function

Private Function RecebimentoMatF_Insere_ItensNFiscal(ByVal objNFiscal As ClassNFiscal, ByVal lComando As Long, ByVal lComando2 As Long) As Long
'Insere os Itens de Nota Fiscal de RecebMaterialF
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lNumIntNFItem As Long
Dim objItemNF As ClassItemNF

On Error GoTo Erro_RecebimentoMatF_Insere_ItensNFiscal

    'Para cada Item da Nota Fiscal
    For Each objItemNF In objNFiscal.colItensNF

        'Preenche lNumIntNF com objNFiscal.lNumIntDoc
        objItemNF.lNumIntNF = objNFiscal.lNumIntDoc

        'Lê o número do próximo Código Automático interno do Ítem
        lErro = CF("NFiscalItem_Automatico", lNumIntNFItem)
        If lErro <> SUCESSO Then Error 30536

        objItemNF.lNumIntDoc = lNumIntNFItem

        'Insere no BD em ItensNFiscal
        lErro = Comando_Executar(lComando, "INSERT INTO ItensNFiscal(NumIntNF, Item, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, Almoxarifado, NumIntDocOrig, Status,Ccl) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objItemNF.lNumIntNF, objItemNF.iItem, objItemNF.sProduto, objItemNF.sUnidadeMed, objItemNF.dQuantidade, objItemNF.dPrecoUnitario, objItemNF.dPercDesc, objItemNF.dValorDesconto, objItemNF.dtDataEntrega, objItemNF.sDescricaoItem, objItemNF.dValorAbatComissao, objItemNF.lNumIntPedVenda, objItemNF.lNumIntItemPedVenda, objItemNF.lNumIntDoc, objItemNF.lNumIntTrib, objItemNF.iAlmoxarifado, objItemNF.lNumIntDocOrig, objItemNF.iStatus, objItemNF.sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 30537
        
        lErro = CF("ItensNFGrade_Inclui", objItemNF)
        If lErro <> SUCESSO Then Error 30739
        
        objItemNF.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
        objItemNF.objInfoAdicDocItem.lNumIntDocItem = objItemNF.lNumIntDoc
        lErro = CF("InfoAdicDocItem_Insere", objItemNF.objInfoAdicDocItem, lComando2)
        If lErro <> SUCESSO Then Error 30739

    Next

    RecebimentoMatF_Insere_ItensNFiscal = SUCESSO

    Exit Function

Erro_RecebimentoMatF_Insere_ItensNFiscal:

    RecebimentoMatF_Insere_ItensNFiscal = Err

    Select Case Err

        Case 30536, 30739

        Case 30537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", Err, objItemNF.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149499)

    End Select

    Exit Function

End Function

Private Function RecebimentoMatC_Insere_ItensNFiscal(ByVal objNFiscal As ClassNFiscal, ByVal lComando As Long, ByVal lComando2 As Long) As Long
'Insere os Itens de Nota Fiscal de RecebMaterialF
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lNumIntNFItem As Long
Dim objItemNF As ClassItemNF

On Error GoTo Erro_RecebimentoMatC_Insere_ItensNFiscal

    'Para cada Item da Nota Fiscal
    For Each objItemNF In objNFiscal.colItensNF

        'Preenche lNumIntNF com objNFiscal.lNumIntDoc
        objItemNF.lNumIntNF = objNFiscal.lNumIntDoc

        'Lê o número do próximo Código Automático interno do Ítem
        lErro = CF("NFiscalItem_Automatico", lNumIntNFItem)
        If lErro <> SUCESSO Then Error 30739

        objItemNF.lNumIntDoc = lNumIntNFItem

        'Insere no BD em ItensNFiscal
        lErro = Comando_Executar(lComando, "INSERT INTO ItensNFiscal(NumIntNF, Item, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, Almoxarifado, NumIntDocOrig, Status,Ccl) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objItemNF.lNumIntNF, objItemNF.iItem, objItemNF.sProduto, objItemNF.sUnidadeMed, objItemNF.dQuantidade, objItemNF.dPrecoUnitario, objItemNF.dPercDesc, objItemNF.dValorDesconto, objItemNF.dtDataEntrega, objItemNF.sDescricaoItem, objItemNF.dValorAbatComissao, objItemNF.lNumIntPedVenda, objItemNF.lNumIntItemPedVenda, objItemNF.lNumIntDoc, objItemNF.lNumIntTrib, objItemNF.iAlmoxarifado, objItemNF.lNumIntDocOrig, objItemNF.iStatus, objItemNF.sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 30740

        lErro = CF("ItensNFGrade_Inclui", objItemNF)
        If lErro <> SUCESSO Then Error 30739
        
        objItemNF.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
        objItemNF.objInfoAdicDocItem.lNumIntDocItem = objItemNF.lNumIntDoc
        lErro = CF("InfoAdicDocItem_Insere", objItemNF.objInfoAdicDocItem, lComando2)
        If lErro <> SUCESSO Then Error 30739
    
    Next

    RecebimentoMatC_Insere_ItensNFiscal = SUCESSO

    Exit Function

Erro_RecebimentoMatC_Insere_ItensNFiscal:

    RecebimentoMatC_Insere_ItensNFiscal = Err

    Select Case Err

        Case 30739

        Case 30740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", Err, objItemNF.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149500)

    End Select

    Exit Function

End Function

Private Function RecebimentoMatC_Exclui_ItensNFiscal(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Exclui os Itens da Nota Fiscal que estão ==>>
'==>>relacionados ao RecebMaterial Selecionado
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando3 As Long
Dim objMovEstoqueExcluir As ClassMovEstoque
Dim dQuantidade As Double
Dim lCodigo As Long
Dim sProduto As String
Dim sSiglaUM As String
Dim iAlmoxarifado As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocOrigem As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim dCusto As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemNF As New ClassItemNF
Dim objProduto As New ClassProduto

On Error GoTo Erro_RecebimentoMatC_Exclui_ItensNFiscal

    'Inicializa comando
    lComando3 = Comando_Abrir
    If lComando3 = 0 Then Error 55540

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 20756
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then Error 20757
    
    'Se o TipoMovtoEstoque for positivo
    If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
        'Cria objMovEstoque
        Set objMovEstoqueExcluir = New ClassMovEstoque
    
        'Preenche os atributos exceto código e colItens
        objMovEstoqueExcluir.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque2
        objMovEstoqueExcluir.dtData = objNFiscal.dtDataEntrada
        objMovEstoqueExcluir.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoqueExcluir.lCliente = objNFiscal.lCliente
        objMovEstoqueExcluir.lFornecedor = objNFiscal.lFornecedor
        objMovEstoqueExcluir.sDocOrigem = objTipoDocInfo.sSigla
    
        sProduto = String(STRING_PRODUTO, 0)
        sSiglaUM = String(STRING_UM_SIGLA, 0)
    
        'Para cada Item na Tabela ItensNFiscal que ==>>
        'tem NumIntNF igual ao NumIntDoc do Recebimento selecionado
        lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc, Produto, Quantidade, UnidadeMed, Almoxarifado FROM ItensNFiscal WHERE NumIntNF = ?", 0, lNumIntDocOrigem, sProduto, dQuantidade, sSiglaUM, iAlmoxarifado, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 55536
    
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55537
    
        While lErro <> AD_SQL_SEM_DADOS
            
            Set objItemNF.colItensRomaneioGrade = New Collection
            objItemNF.lNumIntDoc = lNumIntDocOrigem
            
            lErro = CF("ItensNFGrade_Le", objItemNF)
            If lErro <> SUCESSO Then Error 20756
        
            If objItemNF.iPossuiGrade = DESMARCADO Then
    
                objProduto.sCodigo = sProduto
                
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then Error 25188
            
                'Se não achou o Produto --> erro
                If lErro <> SUCESSO Then Error 25189
                                
                If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                    lErro = Comando_Executar(lComando3, "SELECT NumIntDoc, Custo FROM MovimentoEstoque WHERE NumIntDocOrigem = ? AND TipoNumIntDocOrigem=?", lNumIntDoc, dCusto, lNumIntDocOrigem, TIPO_ORIGEM_ITEMNF)
                    If lErro <> AD_SQL_SUCESSO Then Error 55541
            
                    lErro = Comando_BuscarPrimeiro(lComando3)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55542
            
                    If lErro = SUCESSO Then
            
                        'Adiciona ítem em colItens de objMovEstoqueExcluir
                        objMovEstoqueExcluir.colItens.Add lNumIntDoc, 0, dCusto, 0, sProduto, "", sSiglaUM, dQuantidade, iAlmoxarifado, "", lNumIntDocOrigem, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                        
                    End If
                End If
            Else
                For Each objItemRomaneioGrade In objItemNF.colItensRomaneioGrade
                
                    If objItemRomaneioGrade.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                        lErro = Comando_Executar(lComando3, "SELECT NumIntDoc, Custo FROM MovimentoEstoque WHERE NumIntDocOrigem = ? AND TipoNumIntDocOrigem=?", lNumIntDoc, dCusto, objItemRomaneioGrade.lNumIntDoc, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE)
                        If lErro <> AD_SQL_SUCESSO Then Error 55541
                
                        lErro = Comando_BuscarPrimeiro(lComando3)
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55542
                
                        Do While lErro = AD_SQL_SUCESSO
                            objMovEstoqueExcluir.colItens.Add lNumIntDoc, 0, dCusto, 0, objItemRomaneioGrade.sProduto, objItemNF.sDescricaoItem, sSiglaUM, objItemRomaneioGrade.dQuantidade, 0, "", objItemRomaneioGrade.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                            objMovEstoqueExcluir.colItens(objMovEstoqueExcluir.colItens.Count).iFilialEmpresa = objNFiscal.iFilialEmpresa
                            objMovEstoqueExcluir.colItens(objMovEstoqueExcluir.colItens.Count).iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            
                            lErro = Comando_BuscarProximo(lComando3)
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55542
                        Loop
                    End If
                Next
    
            End If
            
            'Exclui ItenNFiscal do BD
            lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ItensNFiscal", lComando1)
            If lErro <> AD_SQL_SUCESSO Then Error 55539
            
            lErro = CF("ItensNFGrade_Exclui", lNumIntDoc)
            If lErro <> SUCESSO Then Error 20756
                        
            
            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55538
    
        Wend
    
        'Faz exclusão de MovEstoque Associados
        lErro = CF("MovEstoque_Grava", objMovEstoqueExcluir)
        If lErro <> SUCESSO Then Error 55535

    End If
    
    'Libera comando
    Call Comando_Fechar(lComando3)

    RecebimentoMatC_Exclui_ItensNFiscal = SUCESSO

    Exit Function

Erro_RecebimentoMatC_Exclui_ItensNFiscal:

    RecebimentoMatC_Exclui_ItensNFiscal = Err

    Select Case Err

        Case 20756, 25188, 55535, 55536

        Case 20757
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)

        Case 25189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)

        Case 55536, 55537, 55538
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", Err)

        Case 55539
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSNFISCAL", Err, objNFiscal.lNumIntDoc)

        Case 55540
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55541, 55542
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE3", Err)

        Case 55543
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO", Err, lNumIntDocOrigem)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149501)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function RecebMaterialC_Exclui(ByVal objNFiscal As ClassNFiscal) As Long
'Exclui a Nota Fiscal de Recebimento de material de Cliente

Dim lErro As Long
Dim iIndice As Integer
Dim iTipoDoc As Integer
Dim lNumIntDocOrigem As Long
Dim lNumIntDoc As Long
Dim lComando(5) As Long
Dim lTransacao As Long
Dim iStatus As Integer

On Error GoTo Erro_RecebMaterialC_Exclui

    'Abre os Comandos
    For iIndice = 0 To 4
        lComando(iIndice) = Comando_Abrir
        If lComando(iIndice) = 0 Then Error 30750
    Next

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 30751

    'Tenta selecionar o Recebimento
    lErro = Comando_ExecutarPos(lComando(0), "SELECT NumIntDoc, Status FROM NFiscal WHERE FilialEmpresa = ? AND NumRecebimento = ? ", 0, lNumIntDoc, iStatus, objNFiscal.iFilialEmpresa, objNFiscal.lNumRecebimento)
    If lErro <> AD_SQL_SUCESSO Then Error 30752

    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30753

    'Se não achou o Recebimento --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30754

    'se o recebimento existe e está abaixado ==> nao pode ser excluido
    If iStatus = STATUS_BAIXADO Then Error 55647

    'Faz o Lock exclusivo no Recebimento
    lErro = Comando_LockExclusive(lComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 30755

    objNFiscal.lNumIntDoc = lNumIntDoc

'    'Tenta ler a Nota Fiscal correspondente ao Recebimento
'    lErro = CF("RecebMaterialC_Pesquisa_NotaFiscal",objNFiscal, lComando(1))
'    If lErro <> SUCESSO And lErro <> 30743 Then Error 30756
'
'    'Se encontrou a NFiscal correspondente ao Recebimento de Material de Cliente --> erro
'    If lErro <> SUCESSO Then Error 30758

    'EXCLUSÃO DE ITENS DE RECEBIMENTO
    lErro = RecebimentoMatC_Exclui_ItensNFiscal(objNFiscal, lComando(2), lComando(3))
    If lErro <> SUCESSO Then Error 30759

    'EXCLUI RECEBIMENTO
    lErro = Comando_ExecutarPos(lComando(4), "DELETE FROM NFiscal", lComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 30760

    'Confirma transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then Error 30761

    'Libera comando
    For iIndice = 0 To 4
        lErro = Comando_Fechar(lComando(iIndice))

    Next

    RecebMaterialC_Exclui = SUCESSO

    Exit Function

Erro_RecebMaterialC_Exclui:

    RecebMaterialC_Exclui = Err

    Select Case Err

        Case 30750
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 30752, 30753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 30754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEB_MAT_CLI_NAO_CADASTRADO", Err, objNFiscal.lNumRecebimento)

        Case 30755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", Err, objNFiscal.lNumNotaFiscal)

        Case 30756, 30759
        
        Case 30758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RECEB_MAT_CLI_NF", Err, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEntrada, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 30760
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NFISCAL", Err, objNFiscal.lNumNotaFiscal)

        Case 30761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 55647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RECEB_BAIXADO_CLI", Err, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEntrada, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149502)

    End Select

    Call Transacao_Rollback

    'Libera comandos
    For iIndice = 0 To 4
        lErro = Comando_Fechar(lComando(iIndice))

    Next

    Exit Function

End Function

Function RecebMaterialF_Grava(ByVal objNFiscal As ClassNFiscal) As Long
'Grava a Nota Fiscal de Recebimento de Material de Fornecedor

Dim lErro As Long
Dim lNumIntNFItem As Long
Dim lNumIntDoc As Long
Dim lNumIntDocOrigem As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim objItemNF As ClassItemNF
Dim alComando(15) As Long
Dim iTipo As Integer
Dim dtData As Date
Dim objTipoDocInfo As New ClassTipoDocInfo
'#################################
'Inserido por Wagner 26/01/2006
Dim dFator As Double
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dQuantItemNFUMEst As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
'#################################
Dim objProduto As ClassProduto

On Error GoTo Erro_RecebMaterialF_Grava

    'Abre os comandos de Transação
    For iIndice = LBound(alComando) To UBound(alComando) 'iIndice = 0 To 12

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 30520

    Next

    'Abre Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 30521

    'Faz lock dos objetos associados
    lErro = RecebMaterialF_Lock_ObjAssociados(objNFiscal, alComando(0), alComando(1), alComando(11), alComando(12))
    If lErro <> SUCESSO Then gError 30552

    dtData = objNFiscal.dtDataEmissao
    objNFiscal.dtDataEmissao = objNFiscal.dtDataEntrada

    If objNFiscal.iTipoNFiscal = DOCINFO_NRFP Then
        iTipo = DOCINFO_NF_INT_ENTRADA
    Else
        iTipo = DOCINFO_NF_EXTERNA
        
        'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
        lErro = CF("NFiscalEntrada_Verifica_Existencia", objNFiscal, iTipo)
        If lErro <> SUCESSO Then gError 55641
        
        If objNFiscal.lNumRecebimento = 0 Then
            'Verifica a Existencia de um Recebimento  com os Valores de Nota igual
            lErro = CF("RecebimentoF_Verifica_Existencia", objNFiscal)
            If lErro <> SUCESSO Then gError 61089
        End If
        
    End If
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo de documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 27623 Then gError 61072
        
    'Se não encontrar --> erro
    If lErro = 27623 Then gError 61073
    
    objNFiscal.dtDataEmissao = dtData
    
    '###################################################
    'Inserido por Wagner 25/01/2006
    iIndice = 0
    
    For Each objItemNF In objNFiscal.colItensNF
        
        iIndice = iIndice + 1
        
        If objItemNF.colItensRomaneioGrade.Count > 0 Then
            
            For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                Set objProduto = New ClassProduto
            
                objProduto.sCodigo = objItemRomaneio.sProduto
                
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 23080 Then gError 141600
                
                If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    
                    lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemRomaneio.sUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 141600
    
                    dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
    
                    dSomaAlocacao = 0
                                                                    
                    For Each objReservaItem In objItemRomaneio.colLocalizacao
                                            
                        dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                                                    
                    Next
                    
                    dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)
                    
                    If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > QTDE_ESTOQUE_DELTA Then
                    
                        If dSomaAlocacao < dQuantItemNFUMEst Then gError 141601
                        
                        If dSomaAlocacao > dQuantItemNFUMEst Then gError 141602
                    
                    End If
                    
                End If
                
            Next
            
        Else
                    
            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                If objItemNF.ColAlocacoes.Count = 0 Then gError 141603

                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 141604

                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
    
                dSomaAlocacao = 0
                                                                
                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
                    
                Next

                dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)
                
                If Abs(dSomaAlocacao - dQuantItemNFUMEst) > QTDE_ESTOQUE_DELTA Then
                 
                    If dSomaAlocacao < dQuantItemNFUMEst Then gError 141605
                    
                    If dSomaAlocacao > dQuantItemNFUMEst Then gError 141606
            
                End If
                
            End If
            
        End If
        
    Next
    '###################################################
    'Se o código do Recebimento estiver = 0 ==>> INSERÇÃO
    If objNFiscal.lNumRecebimento = 0 Then

        lErro = RecebMaterialF_Grava_Insercao(objNFiscal, alComando(3), alComando(4), alComando(10), alComando(14))
        If lErro <> SUCESSO Then gError 30553

    Else 'Se não ==>> ALTERAÇÃO

        lErro = RecebMaterialF_Grava_Atualiza(objNFiscal, alComando(2), alComando(5), alComando(6), alComando(7), alComando(8), alComando(9), alComando(15))
        If lErro <> SUCESSO Then gError 30557

    End If

    'Confirma a Transação
    lErro = Transacao_Commit
    If lErro <> SUCESSO Then gError 30551

    'Fecha os comandos
    For iIndice = 0 To 12
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    RecebMaterialF_Grava = SUCESSO

    Exit Function

Erro_RecebMaterialF_Grava:

    RecebMaterialF_Grava = gErr

    Select Case gErr

        Case 30520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 30521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 30551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 30552, 30553, 30557, 55641, 61089 'Tratados na Rotinas chamadas
        
        Case 61072
    
        Case 61073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO", gErr, objTipoDocInfo)

        Case 141601, 141605
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INCOMPLETA", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case 141602, 141606
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)
        
        Case 141603
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INEXISTENTE", gErr, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149503)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = 0 To 12
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function RecebMaterialF_Lock_ObjAssociados(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long) As Long
'Faz o Lock dos Objetos associados ao RecebMaterialF
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim iControleEstoque As Integer
Dim lErro As Long
Dim lCodFornecedor As Long
Dim iCodigo As Integer
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim objCodigoNome As AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome
Dim iLocado As Integer
Dim objItemRomaneioGrade As ClassItemRomaneioGrade

On Error GoTo Erro_Lock_ObjAssociados

    lErro = CF("FilialFornecedor_Lock", objNFiscal.lFornecedor, objNFiscal.iFilialForn)
    If lErro <> SUCESSO And lErro <> 42870 Then Error 30524
        
    'Se não achou a Filial do Fornecedor --> erro
    If lErro <> SUCESSO Then Error 30526

    'Verifica se o código da Transportadora foi preenchido
    If objNFiscal.iCodTransportadora <> 0 Then

        'Faz Lock shared na Transportadora
        lErro = CF("Transportadora_Lock1", objNFiscal.iCodTransportadora)
        If lErro <> SUCESSO And lErro <> 43038 Then Error 30528
        
        'Se não achou a Transportadora --> erro
        If lErro <> SUCESSO Then Error 30530

    End If

    'Para cada Item de ObjNFiscal.colItens
    For Each objItemNF In objNFiscal.colItensNF
                
        objProduto.sCodigo = objItemNF.sProduto
        
        lErro = CF("Produto_Lock", objProduto)
        If lErro <> SUCESSO Then Error 30903

        'Armazena o ControleEstoque do Produto
        objItemNF.iControleEstoque = objProduto.iControleEstoque
       
        For Each objItemRomaneioGrade In objItemNF.colItensRomaneioGrade
            
            objProduto.sCodigo = objItemRomaneioGrade.sProduto
    
            'Faz lock do Produto
            lErro = CF("Produto_Lock", objProduto)
            If lErro <> SUCESSO Then Error 30899
        
            objItemRomaneioGrade.iControleEstoque = objProduto.iControleEstoque
        Next

    Next

    RecebMaterialF_Lock_ObjAssociados = SUCESSO

    Exit Function

Erro_Lock_ObjAssociados:

    RecebMaterialF_Lock_ObjAssociados = Err

    Select Case Err

        Case 25186, 30524, 30528, 30903
        
        Case 30526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 30530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objNFiscal.iCodTransportadora)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149504)

    End Select

    Exit Function

End Function

Private Function RecebMaterialF_Grava_Insercao(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long) As Long
'Faz a Inserção do RecebMaterialF na Tabela NFiscal
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lCodigo As Long
Dim lNumIntDoc As Long
Dim dtDataEntrada As Date
Dim vbMsg As VbMsgBoxResult
Dim objItemNF As ClassItemNF
Dim objMovEstoque As ClassMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
'distribuicao
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dFator As Double
Dim iIndice As Integer
Dim dQuantItemNFUMEst As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim objProduto As ClassProduto

On Error GoTo Erro_RecebMaterialF_Grava_Insercao

    'Gera o código do Recebimento
    lErro = CF("Recebimento_Automatico", objNFiscal)
    If lErro <> SUCESSO Then gError 61084
    
    'Lê o Número do Próximo Código Automático Interno.
    lErro = CF("NFiscal_Automatico", lNumIntDoc)
    If lErro <> SUCESSO Then gError 30534
    
    'Preenche o objNFiscal com Código
    objNFiscal.lNumIntDoc = lNumIntDoc

    'Insere no BD objNFiscal em NFiscal
    With objNFiscal
        lErro = Comando_Executar(lComando2, "INSERT INTO NFiscal(DataAtualizacao, NumRecebimento , NumIntDoc, FilialEmpresa, Serie, NumNotaFiscal, Status, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEntrada, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, VolumeNumero, FreteRespons, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,Observacao, HoraEntrada, Moeda, TaxaMoeda, ChvNFe, HoraEmissao, DataCadastro, HoraCadastro, UsuCadastro) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", gdtDataHoje, _
            .lNumRecebimento, .lNumIntDoc, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .iStatus, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEntrada, .lNumPedidoVenda, _
            .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sVolumeNumero, .iFreteRespons, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sObservacao, CDbl(.dtHoraEntrada), .iMoeda, .dTaxaMoeda, .sChvNFe, CDbl(IIf(.dtDataSaida <> DATA_NULA, IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraSaida, .dtHoraEmissao, .dtHoraSaida), IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraEntrada, .dtHoraEmissao, .dtHoraEntrada))), gdtDataHoje, CDbl(Time), gsUsuario)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 30535

    'Inserção de Itens da Nota Fiscal
    lErro = RecebimentoMatF_Insere_ItensNFiscal(objNFiscal, lComando3, lComando4)
    If lErro <> SUCESSO Then gError 30555

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 20748
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 20749
    
    'Se o TipoMovtoEstoque for positivo
    If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
        'Cria MovEstoque
        Set objMovEstoque = New ClassMovEstoque
    
        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then gError 30884
    
        'Preenche o objMovEstoque
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque
        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = objTipoDocInfo.sSigla & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
    
'mario distribuicao substituicao do codigo original acrescentando o tab de distribuicao
    
        'Para cada elemento de colItens de objNFiscal, adiciona elemento em colItens de objMovEstoque
        For Each objItemNF In objNFiscal.colItensNF
    
            iIndice = iIndice + 1
    
            If objItemNF.colItensRomaneioGrade.Count > 0 Then
    
                For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                    Set objProduto = New ClassProduto
                
                    objProduto.sCodigo = objItemRomaneio.sProduto
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 89606
                    
                    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    
                        lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemRomaneio.sUMEstoque, dFator)
                        If lErro <> SUCESSO Then gError 89606
                   
                        dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
                   
                        dSomaAlocacao = 0
                                                                        
                        For Each objReservaItem In objItemRomaneio.colLocalizacao
                                                
                            dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                   
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, objItemNF.sCcl, 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                            objItemMovEstoque.iControleEstoque = objItemRomaneio.iControleEstoque
                            objItemMovEstoque.sSiglaUMEst = objItemRomaneio.sUMEstoque
            
                        Next
                        
                        If dSomaAlocacao > dQuantItemNFUMEst Then gError 89607
                    End If
                Next
    
            ElseIf objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
           
                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 89593

                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
           
                dSomaAlocacao = 0
                                                                
                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
           
                    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemNF.dQuantidade * objItemNF.dPrecoUnitario) - objItemNF.dValorDesconto) * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                    objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                    objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                    objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
    
                Next
                
                If dSomaAlocacao > dQuantItemNFUMEst Then gError 89594
    
            End If
        Next
    
        'Chama função de Gravação de MovEstoque
        lErro = CF("MovEstoque_Grava", objMovEstoque) 'Chamada de dentro de Transação
        If lErro <> SUCESSO Then gError 30537

    End If
    
    RecebMaterialF_Grava_Insercao = SUCESSO

    Exit Function

Erro_RecebMaterialF_Grava_Insercao:

    RecebMaterialF_Grava_Insercao = gErr

    Select Case gErr

        Case 20749
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 30534, 30537, 30555, 30602, 30884, 20748, 61084, 89593

        Case 30600, 30601
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 30535
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 89594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149505)

    End Select

    Exit Function

End Function

Private Function RecebimentoMatF_Exclui_ItensNFiscal(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Exclui os Itens da Nota Fiscal que estão ==>>
'==>>relacionados ao RecebMaterial Selecionado
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando3 As Long
Dim objMovEstoqueExcluir As ClassMovEstoque
Dim dQuantidade As Double
Dim lCodigo As Long
Dim sProduto As String
Dim sSiglaUM As String
Dim iAlmoxarifado As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocOrigem As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim dCusto As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemNF As New ClassItemNF
Dim objProduto As New ClassProduto

On Error GoTo Erro_RecebimentoMatF_Exclui_ItensNFiscal

    'Inicializa comando
    lComando3 = Comando_Abrir
    If lComando3 = 0 Then Error 30914

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 20754
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then Error 20755
    
    'Se o TipoMovtoEstoque for positivo
    If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
        'Cria objMovEstoque
        Set objMovEstoqueExcluir = New ClassMovEstoque
    
        'Preenche os atributos exceto código e colItens
        objMovEstoqueExcluir.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque2
        objMovEstoqueExcluir.dtData = objNFiscal.dtDataEntrada
        objMovEstoqueExcluir.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoqueExcluir.lCliente = objNFiscal.lCliente
        objMovEstoqueExcluir.lFornecedor = objNFiscal.lFornecedor
        objMovEstoqueExcluir.sDocOrigem = objTipoDocInfo.sSigla
    
        sProduto = String(STRING_PRODUTO, 0)
        sSiglaUM = String(STRING_UM_SIGLA, 0)
    
        'Para cada Item na Tabela ItensNFiscal que ==>>
        'tem NumIntNF igual ao NumIntDoc do Recebimento selecionado
        lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc, Produto, Quantidade, UnidadeMed, Almoxarifado FROM ItensNFiscal WHERE NumIntNF = ?", 0, lNumIntDocOrigem, sProduto, dQuantidade, sSiglaUM, iAlmoxarifado, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 30543
    
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30544
    
        While lErro <> AD_SQL_SEM_DADOS
    
            Set objItemNF.colItensRomaneioGrade = New Collection
            objItemNF.lNumIntDoc = lNumIntDocOrigem
            
            lErro = CF("ItensNFGrade_Le", objItemNF)
            If lErro <> SUCESSO Then Error 20756
        
            If objItemNF.colItensRomaneioGrade.Count = 0 Then
                
                objProduto.sCodigo = sProduto
                
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then Error 25188
            
                'Se não achou o Produto --> erro
                If lErro <> SUCESSO Then Error 25189
                                
                If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                    lErro = Comando_Executar(lComando3, "SELECT NumIntDoc, Custo FROM MovimentoEstoque WHERE NumIntDocOrigem = ? And TipoNumIntDocOrigem=?", lNumIntDoc, dCusto, lNumIntDocOrigem, TIPO_ORIGEM_ITEMNF)
                    If lErro <> AD_SQL_SUCESSO Then Error 30915
            
                    lErro = Comando_BuscarPrimeiro(lComando3)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30916
            
                    If lErro = SUCESSO Then
            
                        'Adiciona ítem em colItens de objMovEstoqueExcluir
                        objMovEstoqueExcluir.colItens.Add lNumIntDoc, 0, dCusto, 0, sProduto, "", sSiglaUM, dQuantidade, iAlmoxarifado, "", lNumIntDocOrigem, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                
                    End If
                
                End If
            
            Else
            
                For Each objItemRomaneioGrade In objItemNF.colItensRomaneioGrade
                
                    If objItemRomaneioGrade.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    
                        lErro = Comando_Executar(lComando3, "SELECT NumIntDoc, Custo FROM MovimentoEstoque WHERE NumIntDocOrigem = ? AND TipoNumIntDocOrigem=?", lNumIntDoc, dCusto, objItemRomaneioGrade.lNumIntDoc, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE)
                        If lErro <> AD_SQL_SUCESSO Then Error 55541
                
                        lErro = Comando_BuscarPrimeiro(lComando3)
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55542
                
                        Do While lErro = AD_SQL_SUCESSO
                            objMovEstoqueExcluir.colItens.Add lNumIntDoc, 0, dCusto, 0, objItemRomaneioGrade.sProduto, objItemNF.sDescricaoItem, sSiglaUM, objItemRomaneioGrade.dQuantidade, 0, "", objItemRomaneioGrade.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                            objMovEstoqueExcluir.colItens(objMovEstoqueExcluir.colItens.Count).iFilialEmpresa = objNFiscal.iFilialEmpresa
                            objMovEstoqueExcluir.colItens(objMovEstoqueExcluir.colItens.Count).iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            
                            lErro = Comando_BuscarProximo(lComando3)
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55542
                        Loop
                    End If
                Next
    
            End If
                        
            'Exclui ItenNFiscal do BD
            lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ItensNFiscal", lComando1)
            If lErro <> AD_SQL_SUCESSO Then Error 30545
       
            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30546
    
        Wend
    
        'Faz exclusão de MovEstoque Associados
        lErro = CF("MovEstoque_Grava", objMovEstoqueExcluir)
        If lErro <> SUCESSO Then Error 30547

    End If

    'Libera comando
    Call Comando_Fechar(lComando3)

    RecebimentoMatF_Exclui_ItensNFiscal = SUCESSO

    Exit Function

Erro_RecebimentoMatF_Exclui_ItensNFiscal:

    RecebimentoMatF_Exclui_ItensNFiscal = Err

    Select Case Err

        Case 20754, 25188, 30547

        Case 20755
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)

        Case 25189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)

        Case 30543, 30544, 30546
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", Err)

        Case 30545
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSNFISCAL", Err, objNFiscal.lNumIntDoc)

        Case 30914
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30915, 30916
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE3", Err)

        Case 30917
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO", Err, lNumIntDocOrigem)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149506)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Private Function RecebMaterialF_Grava_Atualiza(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long) As Long
'Faz a Atualização do RecebMaterialF na Tabela NFiscal
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lCodigo As Long
Dim lNumIntDoc As Long
Dim iTipoDoc As Integer
Dim objItemNF As ClassItemNF
Dim objMovEstoque As ClassMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
'distribuicao
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dFator As Double
Dim iIndice As Integer
Dim dQuantItemNFUMEst As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim objProduto As ClassProduto

On Error GoTo Erro_RecebMaterialF_Grava_Atualiza

    'Testa se o Recebimento já foi baixado
    lErro = CF("Recebimento_Testa_Baixado", objNFiscal)
    If lErro <> SUCESSO Then gError 61080
    
    'Tenta ler o Recebimento na Tabela NFiscal
    lErro = Comando_ExecutarPos(lComando3, "SELECT NumIntDoc FROM NFiscal WHERE FilialEmpresa = ? AND NumRecebimento = ?", 0, lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.lNumRecebimento)
    If lErro <> AD_SQL_SUCESSO Then gError 30532

    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 30533
    
    'Faz Lock Exclusive no Recebimento
    lErro = Comando_LockExclusive(lComando3)
    If lErro <> SUCESSO Then gError 30538

    objNFiscal.lNumIntDoc = lNumIntDoc

    'Atualiza o Recebimento
    lErro = Comando_ExecutarPos(lComando2, "UPDATE NFiscal SET ValorTotal = ?, ValorProdutos = ?, ValorSeguro = ?, ValorDesconto = ?, ValorFrete = ?, ValorOutrasDespesas = ?, CodTransportadora = ?, Placa = ?, PlacaUF = ?, VolumeQuant = ?, VolumeEspecie = ?, VolumeMarca = ?, VolumeNumero = ?, FreteRespons = ?, MensagemNota = ?, PesoLiq = ?, PesoBruto = ?, Observacao = ?, Serie = ?, NumNotaFiscal = ?, DataEntrada = ?, TipoNFiscal = ?, Fornecedor = ? , FilialForn = ?, HoraEntrada = ?, DataAtualizacao=?", lComando3, _
    objNFiscal.dValorTotal, objNFiscal.dValorProdutos, objNFiscal.dValorSeguro, objNFiscal.dValorDesconto, objNFiscal.dValorFrete, objNFiscal.dValorOutrasDespesas, objNFiscal.iCodTransportadora, objNFiscal.sPlaca, objNFiscal.sPlacaUF, objNFiscal.lVolumeQuant, objNFiscal.lVolumeEspecie, objNFiscal.lVolumeMarca, objNFiscal.sVolumeNumero, objNFiscal.iFreteRespons, objNFiscal.sMensagemNota, objNFiscal.dPesoLiq, objNFiscal.dPesoBruto, objNFiscal.sObservacao, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEntrada, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, CDbl(objNFiscal.dtHoraEntrada), gdtDataHoje)
    If lErro <> AD_SQL_SUCESSO Then gError 30542

    '==========================
    'EXCLUSÃO ITENS NOTA FISCAL
    '==========================
    lErro = RecebimentoMatF_Exclui_ItensNFiscal(objNFiscal, lComando4, lComando5)
    If lErro <> SUCESSO Then gError 30556

    'Inserção de Itens da Nota Fiscal
    lErro = RecebimentoMatF_Insere_ItensNFiscal(objNFiscal, lComando6, lComando7)
    If lErro <> SUCESSO Then gError 30554

    '====================
    'INSERÇÃO MOV ESTOQUE
    '====================
    'Cria objMovEstoque
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 20746
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 20747
    
    'Se o TipoMovtoEstoque for positivo
    If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
        Set objMovEstoque = New ClassMovEstoque
    
        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then gError 30886
    
        'Preenche o objMovEstoque
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque
        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = objTipoDocInfo.sSigla & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
    
'inicio distribuicao
        'Para cada elemento de colItens de objNFiscal, adiciona elemento em colItens de objMovEstoque
        For Each objItemNF In objNFiscal.colItensNF
    
            iIndice = iIndice + 1
    
            If objItemNF.colItensRomaneioGrade.Count > 0 Then
            
                For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                    Set objProduto = New ClassProduto
                
                    objProduto.sCodigo = objItemRomaneio.sProduto
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 89606
                    
                    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    
                        lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemRomaneio.sUMEstoque, dFator)
                        If lErro <> SUCESSO Then gError 89606
                   
                        dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
                   
                        dSomaAlocacao = 0
                                                                        
                        For Each objReservaItem In objItemRomaneio.colLocalizacao
                                                
                            dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                   
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, objItemNF.sCcl, 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                            objItemMovEstoque.iControleEstoque = objItemRomaneio.iControleEstoque
                            objItemMovEstoque.sSiglaUMEst = objItemRomaneio.sUMEstoque
            
                        Next
                        
                        If dSomaAlocacao > dQuantItemNFUMEst Then gError 89607
                    End If
                Next
    
            ElseIf objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
           
                dSomaAlocacao = 0
                                                                
                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 89595
                                                                
                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))

                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
           
                    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemNF.dQuantidade * objItemNF.dPrecoUnitario) - objItemNF.dValorDesconto) * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                    objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                    objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                    objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
    
                Next
                
                If dSomaAlocacao > dQuantItemNFUMEst Then gError 89596
    
            End If
        Next
'fim distribuicao
    
        'Chama função de Gravação de MovEstoque
        lErro = CF("MovEstoque_Grava", objMovEstoque) 'Chamada de dentro de Transação
        If lErro <> SUCESSO Then gError 30550

    End If
    
    RecebMaterialF_Grava_Atualiza = SUCESSO

    Exit Function

Erro_RecebMaterialF_Grava_Atualiza:

    RecebMaterialF_Grava_Atualiza = gErr

    Select Case gErr

        Case 20746, 30550, 30554, 30556, 30886, 30907, 61080, 89595

        Case 20747
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 30532, 30533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 30538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 30539, 30540
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 30757
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RECEB_MAT_FORN_NF", objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 30542
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 89596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149507)

    End Select

    Exit Function

End Function

Function RecebMaterialF_Exclui(ByVal objNFiscal As ClassNFiscal) As Long
'Exclui o Recebimento de Material de Fornecedor

Dim lErro As Long
Dim iIndice As Integer
Dim iTipoDoc As Integer
Dim lNumIntDocOrigem As Long
Dim lNumIntDoc As Long
Dim lComando(5) As Long
Dim lTransacao As Long
Dim iStatus As Integer

On Error GoTo Erro_RecebMaterialF_Exclui

    'Abre os Comandos
    For iIndice = 0 To 4
        lComando(iIndice) = Comando_Abrir
        If lComando(iIndice) = 0 Then Error 30588
    Next

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 30589

    'Tenta selecionar o Recebimento
    lErro = Comando_ExecutarPos(lComando(0), "SELECT NumIntDoc, Status FROM NFiscal WHERE FilialEmpresa = ? AND NumRecebimento = ?", 0, lNumIntDoc, iStatus, objNFiscal.iFilialEmpresa, objNFiscal.lNumRecebimento)
    If lErro <> AD_SQL_SUCESSO Then Error 30590

    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30591

    'Se não achou o Recebimneto --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30592

    'se o recebimento existe e já está baixado ==> não pode excluir
    If iStatus = STATUS_BAIXADO Then Error 55646
    
    'Faz o Lock exclusivo no Recebimento
    lErro = Comando_LockExclusive(lComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 30593

    objNFiscal.lNumIntDoc = lNumIntDoc

'    lErro = CF("RecebMaterialF_Pesquisa_NotaFiscal",objNFiscal, lComando(1))
'    If lErro <> SUCESSO And lErro <> 30541 Then Error 30908
'
'    If lErro <> SUCESSO Then Error 30596

    'EXCLUSÃO DE ITENS DE RECEBIMENTO
    lErro = RecebimentoMatF_Exclui_ItensNFiscal(objNFiscal, lComando(2), lComando(3))
    If lErro <> SUCESSO Then Error 30597

    'EXCLUI RECEBIMENTO
    lErro = Comando_ExecutarPos(lComando(4), "DELETE FROM NFiscal", lComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 30598

    'Confirma Transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then Error 30599

    For iIndice = 0 To 4
    
        lErro = Comando_Fechar(lComando(iIndice))

    Next

    RecebMaterialF_Exclui = SUCESSO

    Exit Function

Erro_RecebMaterialF_Exclui:

    RecebMaterialF_Exclui = Err

    Select Case Err

        Case 30588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 30590, 30591, 30594, 30595
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 30592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEB_MAT_FORN_NAO_CADASTRADO", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.dtDataEntrada)

        Case 30593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", Err, objNFiscal.lNumNotaFiscal)

        Case 30596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RECEB_MAT_FORN_NF", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 30597, 30908

        Case 30598
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NFISCAL", Err, objNFiscal.lNumNotaFiscal)

        Case 30599
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 55646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RECEB_BAIXADO_FORN", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149508)

    End Select

    Call Transacao_Rollback

    For iIndice = 0 To 4
        
        lErro = Comando_Fechar(lComando(iIndice))

    Next

    Exit Function

End Function

Function Custos_Grava(ByVal iMes As Integer, ByVal dCustoReposicao As Double, ByVal objSldMesEst As ClassSldMesEst, ByVal objTela As Object) As Long

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objProduto As New ClassProduto
Dim objProdutoFilial As New ClassProdutoFilial
Dim objEstoqueMes As New ClassEstoqueMes
Dim lTransacao As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim dCustoStd As Double
Dim objMovEstoque As New ClassMovEstoque
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As ClassAlmoxarifado
Dim lCodigo As Long
Dim objContabAutomatica As ClassContabAutomatica
Dim objEstoqueMesAberto As New ClassEstoqueMes

On Error GoTo Erro_Custos_Grava

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 41756
    Next

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 41757

'    Verificar se o produto está cadastrado. Se estiver locar. Senão, dar aviso perguntando se quer cadastrar
    objProduto.sCodigo = objSldMesEst.sProduto

    lErro = CF("Produto_Lock1", alComando(2), objProduto)
    If lErro <> SUCESSO Then gError 41758

'    Verificar se o produto filial está cadastrado. Se estiver locar. Senão, dar aviso perguntando se quer cadastrar.
    objProdutoFilial.iFilialEmpresa = objSldMesEst.iFilialEmpresa
    objProdutoFilial.sProduto = objSldMesEst.sProduto

    lErro = CF("ProdutoFilial_Le_Lock", alComando(0), objProdutoFilial)
    If lErro <> SUCESSO And lErro <> 41770 Then gError 41760

    If lErro = 41770 Then gError 41761

'    Ler o EstoqueMes e locar.
    objEstoqueMes.iFilialEmpresa = objSldMesEst.iFilialEmpresa
    objEstoqueMes.iAno = objSldMesEst.iAno
    objEstoqueMes.iMes = iMes
    
    lErro = CF("EstoqueMes_Le_Lock", alComando(1), objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 41774 Then gError 41762

    If lErro = 41774 Then gError 41763

    'Se o atributo CustoProdApurado da tabela EstoqueMes estiver ligado para o mes selecionado, avisar que não pode alterar o custo pois o custo médio de produção já foi apurado para o mes em questão.
    If objProduto.iApropriacaoCusto = APROPR_CUSTO_REAL Then
        If objEstoqueMes.iCustoProdApurado = CUSTO_APURADO Then gError 41764
    End If
    
'    If objProduto.iApropriacaoCusto = APROPR_CUSTO_STANDARD Then
'
'        'Move Produto para o objeto
'        objItemMovEstoque.sProduto = objProduto.sCodigo
'
'        'verifica se tem algum Movimento relacioado com aquele Ano-Mes
'        lErro = CF("MovEstoque_Le_ProdutoMes",iMes, objSldMesEst.iAno, objItemMovEstoque)
'        If lErro <> SUCESSO And lErro <> 49971 Then gError 49972
'
'        If lErro = SUCESSO Then gError 49973
'
'    Else
'
'        If objSldMesEst.dCustoStandard(iMes) <> 0 Then gError 49974
'
'    End If
    
    'Grava o Custo reposição na tabela de Produtos
    lErro = Comando_ExecutarPos(alComando(3), "UPDATE Produtos SET CustoReposicao=?", alComando(2), dCustoReposicao)
    If lErro <> AD_SQL_SUCESSO Then gError 52862
    
    'Gravar o Custo Standard e o Custo Real de produção em SldMesEst
    lErro = SldMesEst_Custos_Grava(iMes, objSldMesEst, dCustoStd)
    If lErro <> SUCESSO And lErro <> 41778 Then gError 41765

    If lErro = 41778 Then gError 41766

'Parei aqui
'Vamos criar uma função envolvendo o codigo a partir de dCustoStd <> objSldMesEst....
'Parametros dCustoStd Leitura do SldMesEst do mes/Ano em questão
'a data do movimento de estoque
'tem que verificar se tem SldMesAlm para o almoxarifado em questão e como o ClassMatEstoque resolve
'o classmatestoque coloca zero na quantidade se não existe e cria o sldmesestalm se necessário.

    If objProduto.iApropriacaoCusto = APROPR_CUSTO_STANDARD Then

        objEstoqueMesAberto.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMesAberto)
        If lErro <> SUCESSO Then gError 83599
        
        'se o custo standard estiver sendo alterado do mes aberto ==> gera reajuste do estoque segundo o novo custo
        If objEstoqueMesAberto.iAno = objSldMesEst.iAno And objEstoqueMesAberto.iMes = iMes Then

            If dCustoStd <> objSldMesEst.dCustoStandard(iMes) Then
    
                lErro = CF("MovEstoque_Automatico_EmTransacao", giFilialEmpresa, lCodigo)
                If lErro <> SUCESSO Then gError 30901
            
                objMovEstoque.dtData = gdtDataHoje
                objMovEstoque.iFilialEmpresa = giFilialEmpresa
                objMovEstoque.lCodigo = lCodigo
                
                'Lê todos os Almoxarifados do Produto para todas as filiais
                lErro = CF("EstoqueProduto_Le_Almoxarifados", objProduto.sCodigo, colAlmoxarifados)
                If lErro <> SUCESSO Then gError 20789
                
                For Each objAlmoxarifado In colAlmoxarifados
                
                        objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_NOSSO, objSldMesEst.dCustoStandard(iMes) - dCustoStd, APROPR_CUSTO_INFORMADO, objProduto.sCodigo, objProduto.sDescricao, objProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                        objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_CONSIG_NOSSO, objSldMesEst.dCustoStandard(iMes) - dCustoStd, APROPR_CUSTO_INFORMADO, objProduto.sCodigo, objProduto.sDescricao, objProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                        objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_DEMO_NOSSO, objSldMesEst.dCustoStandard(iMes) - dCustoStd, APROPR_CUSTO_INFORMADO, objProduto.sCodigo, objProduto.sDescricao, objProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                        objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_CONSERTO_NOSSO, objSldMesEst.dCustoStandard(iMes) - dCustoStd, APROPR_CUSTO_INFORMADO, objProduto.sCodigo, objProduto.sDescricao, objProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                        objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_OUTROS_NOSSO, objSldMesEst.dCustoStandard(iMes) - dCustoStd, APROPR_CUSTO_INFORMADO, objProduto.sCodigo, objProduto.sDescricao, objProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                        objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_BENEF_NOSSO, objSldMesEst.dCustoStandard(iMes) - dCustoStd, APROPR_CUSTO_INFORMADO, objProduto.sCodigo, objProduto.sDescricao, objProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                        
                Next
    
                lErro = CF("MovEstoque_Grava", objMovEstoque)
                If lErro <> SUCESSO Then gError 20790
    
                If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
                
                    Set objContabAutomatica = New ClassContabAutomatica
                
                    lErro = objContabAutomatica.Inicializa_Contab(objTela, MODULO_ESTOQUE, "CES", objMovEstoque.dtData, objMovEstoque.dtData)
                    If lErro <> SUCESSO Then gError 83577
                    
                    lErro = objContabAutomatica.GeraContabilizacao(objMovEstoque)
                    If lErro <> SUCESSO Then gError 83579
                
                    lErro = objContabAutomatica.Finaliza_Contab()
                    If lErro <> SUCESSO Then gError 83578
                    
                End If

            End If

        End If

    End If

    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 41767

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Custos_Grava = SUCESSO

    Exit Function

Erro_Custos_Grava:

    Custos_Grava = gErr

    Select Case gErr

        Case 20789, 20790, 83577, 83578, 83579

        Case 41756
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 41757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 41758, 41760, 41762, 41765, 49972
                        
        Case 41761
            'Não encontrou Produto no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTOFILIAL", objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)

            If vbMsgRes = vbYes Then
                'Chama a tela de Estoque
                Call Chama_Tela("Estoque", objProduto)
            End If
            
        Case 41763
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
            
        Case 41764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_CUSTO_APURADO", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
            
        Case 41766
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTOS_INEXISTENTES", gErr, objSldMesEst.iFilialEmpresa, objSldMesEst.iAno, objSldMesEst.sProduto)
        
        Case 41767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 49973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTO_STANDARD_MOVIMENTO", gErr, objEstoqueMes.iMes, objSldMesEst.iAno)
        
        Case 49974
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTO_STANDARD_PREENCHIDO", gErr)
            
        Case 52862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PRODUTOS", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149509)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function SldMesEst_Custos_Grava(ByVal iMes As Integer, ByVal objSldMesEst As ClassSldMesEst, dCustoStd As Double) As Long
'gravar o custo standard e o custo real de produção em SldMesEst
'retorna o custo std do bd

Dim lErro As Long
Dim sComando_SQL As String
Dim sMes As String
Dim dCustoProducao As Double
Dim lComando As Long, lComando1 As Long

On Error GoTo Erro_SldMesEst_Custos_Grava

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41704

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 41705
    
    sMes = CStr(iMes)

    sComando_SQL = "SELECT CustoProducao" & sMes & ", CustoStandard" & sMes & " FROM SldMesEst WHERE Ano = ? AND FilialEmpresa = ? AND Produto = ?"

    lErro = Comando_ExecutarPos(lComando, sComando_SQL, 0, dCustoProducao, dCustoStd, objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 41776

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41777

    If lErro = AD_SQL_SEM_DADOS Then Error 41778

    sComando_SQL = "UPDATE SldMesEst SET CustoProducao" & sMes & " = ? , CustoStandard" & sMes & " = ?, CustoReposicao" & sMes & " = ?"

    lErro = Comando_ExecutarPos(lComando1, sComando_SQL, lComando, objSldMesEst.dCustoProducao(iMes), objSldMesEst.dCustoStandard(iMes), objSldMesEst.dCustoReposicao(iMes))
    If lErro <> AD_SQL_SUCESSO Then Error 41779

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    SldMesEst_Custos_Grava = SUCESSO

    Exit Function

Erro_SldMesEst_Custos_Grava:

    SldMesEst_Custos_Grava = Err

    Select Case Err

        Case 41704, 41705
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 41776, 41777
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", Err, objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)

        Case 41778
        
        Case 41779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDMESEST", Err, objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149510)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function Reserva_Grava(ByVal objReserva As ClassReserva) As Long
'Grava uma Reserva

Dim lErro As Long
Dim alComando(0 To 11) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim dQuantidadeReservar As Double
Dim dQuantDisponivel As Double, dDifQtdeReserva As Double
Dim tReserva As typeReserva
Dim objMovEstoque As New ClassMovEstoque
Dim lCodigo As Long
Dim tItemPedido As typeItemPedido
Dim objProduto As New ClassProduto
Dim objReservaAux As ClassReserva

On Error GoTo Erro_Reserva_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 23867
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 23868

    tReserva.sProduto = String(STRING_PRODUTO, 0)

    'Pesquisa no Banco de Dados a Reserva em questão
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Produto, Almoxarifado, TipoDoc, DocOrigem, NumIntOrigem, Quantidade FROM Reserva WHERE Codigo=? AND FilialEmpresa=? ", 0, tReserva.sProduto, tReserva.iAlmoxarifado, tReserva.iTipoDoc, tReserva.lDocOrigem, tReserva.lNumIntOrigem, tReserva.dQuantidade, objReserva.lCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 23890

    'Lê a Reserva, se estiver no Banco de Dados
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23891

    'Encontrou a Reserva ==> isto é uma alteração
    If lErro = AD_SQL_SUCESSO Then

        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 23903

        'Compara os Campos Lidos no Banco de Dados com os Campos no Objeto
        If objReserva.sProduto <> tReserva.sProduto Then Error 23904
        If objReserva.iAlmoxarifado <> tReserva.iAlmoxarifado Then Error 23905
        If objReserva.iTipoDoc <> tReserva.iTipoDoc Then Error 23906
        If objReserva.lDocOrigem <> tReserva.lDocOrigem Then Error 23907
        If objReserva.lNumIntOrigem <> tReserva.lNumIntOrigem Then Error 23908

        'Atualiza a Reserva
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Reserva SET Quantidade = ?, DataValidade = ?, DataReserva = ?, CodUsuario = ?, Responsavel = ? ", alComando(0), objReserva.dQuantidade, objReserva.dtDataValidade, objReserva.dtDataReserva, objReserva.sCodUsuario, objReserva.sResponsavel)
        If lErro <> AD_SQL_SUCESSO Then Error 23892

    Else 'não encontrou a Reserva ==>  é uma inserção

        lErro = Reserva_Grava_BD(objReserva, alComando(2), alComando(3), alComando(4), tReserva)
        If lErro <> SUCESSO Then Error 30909

    End If

    dDifQtdeReserva = objReserva.dQuantidade - tReserva.dQuantidade

    If dDifQtdeReserva <> 0 Then

        lErro = CF("Reserva_Grava_MovtoEstoque", objReserva, dDifQtdeReserva)
        If lErro <> SUCESSO Then Error 23933

    End If

    objReserva.dQuantidade = dDifQtdeReserva

    If objReserva.iTipoDoc = TIPO_PEDIDO_COD Then
       
        lErro = CF("Reserva_Grava_TipoPedidoCod", objReserva, alComando(7), alComando(8))
        If lErro <> SUCESSO Then Error 30052

    ElseIf objReserva.iTipoDoc = TIPO_PEDIDO_GRADE Then
    
        lErro = CF("Reserva_Grava_TipoPedidoGrade", objReserva, alComando(7), alComando(8), alComando(9), alComando(10))
        If lErro <> SUCESSO Then Error 30052
       
    ElseIf objReserva.iTipoDoc = TIPO_PEDIDO_SRV_COD Then
    
        lErro = Reserva_Grava_TipoPedidoSRVCod(objReserva, alComando(7), alComando(8))
        If lErro <> SUCESSO Then Error 30052
              
    End If
    
    lErro = CF("Estoque_TestaIntegridade", objReserva.iFilialEmpresa)
    If lErro <> SUCESSO Then Error 30052

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 30043

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Reserva_Grava = SUCESSO

    Exit Function

Erro_Reserva_Grava:

    Reserva_Grava = Err

    Select Case Err

        Case 23883, 23980, 23990, 23933, 30027, 30052, 30909

        Case 23867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 23890, 23891
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA", Err)

        Case 23892
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RESERVA", Err, objReserva.lCodigo)

        Case 23894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_INSUFICIENTE", Err, objReserva.sProduto, objReserva.iAlmoxarifado, dQuantDisponivel, objReserva.dQuantidade)

        Case 23903
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RESERVA", Err)

        Case 23904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COINCIDE", Err, tReserva.sProduto, objReserva.sProduto)

        Case 23905
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_COINCIDE", Err, tReserva.iAlmoxarifado, objReserva.iAlmoxarifado)

        Case 23906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_NAO_COINCIDE", Err, tReserva.iTipoDoc, objReserva.iTipoDoc)

        Case 23907
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DOCORIGEM_NAO_COINCIDE", Err, tReserva.lDocOrigem, objReserva.lDocOrigem)

        Case 23908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_COINCIDE", Err)

        Case 23909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_INSUFICIENTE", Err, objReserva.sProduto, objReserva.iAlmoxarifado, dQuantDisponivel, objReserva.dQuantidade)

        Case 30028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", Err, objReserva.lNumIntOrigem)

        Case 30042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ESTOQUEPRODUTO", Err, objReserva.iAlmoxarifado)

        Case 30043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149511)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Bloqueio_ReservaGrava(ByVal objReserva As ClassReserva) As Long
'Grava o bloqueio com os dados passados em objReserva

Dim lErro As Long
Dim lAchou As Long
Dim dQuantidadeFaltaReservar As Double
Dim dPedidoFaltaReservar As Double
Dim objBloqueioPV As New ClassBloqueioPV
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objItemPedido As ClassItemPedido
Dim iIndice As Integer
Dim iNumItens As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim bPossui As Boolean

On Error GoTo Erro_Bloqueio_ReservaGrava

    objPedidoVenda.iFilialEmpresa = objReserva.iFilialEmpresa
    objPedidoVenda.lCodigo = objReserva.lDocOrigem
    
    'Lê os ítens de PedidoVenda associados
    lErro = CF("PedidoDeVenda_Le_Itens", objPedidoVenda)
    If lErro <> SUCESSO Then Error 23991

    'Retira da coleção os ítens que correspondem a Produtos que tem ControleEstoque<>PRODUTO_CONTROLE_RESERVA
    iIndice = 1
    iNumItens = objPedidoVenda.colItensPedido.Count
    Do While True
        
        If objPedidoVenda.colItensPedido(iIndice).iPossuiGrade = DESMARCADO Then
            If objPedidoVenda.colItensPedido(iIndice).iControleEstoque <> PRODUTO_CONTROLE_RESERVA Then
                objPedidoVenda.colItensPedido.Remove (iIndice)
                iIndice = iIndice - 1
                iNumItens = iNumItens - 1
            End If
            If iIndice = iNumItens Then Exit Do
            iIndice = iIndice + 1
        Else
            bPossui = False
            For Each objItemRomaneio In objPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade
                If Not (objItemRomaneio.iControleEstoque <> PRODUTO_CONTROLE_RESERVA) Then
                    bPossui = True
                    Exit For
                End If
            Next
            
            If Not bPossui Then
                objPedidoVenda.colItensPedido.Remove (iIndice)
                iIndice = iIndice - 1
                iNumItens = iNumItens - 1
            End If
            If iIndice = iNumItens Then Exit Do
            iIndice = iIndice + 1
            
        End If
    Loop

    'Acumula quantidade que falta reservar do Pedido
    For Each objItemPedido In objPedidoVenda.colItensPedido

        dQuantidadeFaltaReservar = 0
        'Verifica a quantidade que falta reservar deste ítem do pedido
        lErro = CF("ItemPedido_QuantFaltaReservar", objItemPedido, dQuantidadeFaltaReservar)
        If lErro <> SUCESSO Then Error 23992

        'Calcula quantidade
        dPedidoFaltaReservar = dPedidoFaltaReservar + dQuantidadeFaltaReservar

    Next

    objBloqueioPV.lPedidoDeVendas = objReserva.lDocOrigem
    
    'Lê bloqueio de estoque (não liberado) correspondente ao Pedido
    'Se for Bloqueio_Total, dá erro
    lErro = CF("BloqueioEstoquePV_Le", objBloqueioPV)
    If lErro <> SUCESSO And lErro <> 46032 Then Error 23993
                                      '|--> Não encontrou Bloqueio
    lAchou = lErro

    'Se encontrou BloqueioTotal erro (impede gravação de Reserva)
    If lAchou = SUCESSO And objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_TOTAL Then Error 25663

    'Se não falta nada a reservar no Pedido e encontrou Bloqueio de Estoque
    If dPedidoFaltaReservar = 0 And lAchou = SUCESSO Then

        objBloqueioPV.sCodUsuarioLib = objReserva.sCodUsuario
        objBloqueioPV.sResponsavelLib = objReserva.sResponsavel
        objBloqueioPV.dtDataLib = gdtDataHoje

        'Libera o bloqueio de Estoque (que não é total)
        lErro = CF("BloqueioEstoquePV_Libera", objBloqueioPV)
        If lErro <> SUCESSO And lErro <> 30037 Then Error 23994
                                           '|--> SEM DADOS
    End If

    If dPedidoFaltaReservar <> 0 And objBloqueioPV.iTipoDeBloqueio <> BLOQUEIO_PARCIAL And lAchou = SUCESSO Then

        objBloqueioPV.sCodUsuarioLib = objReserva.sCodUsuario
        objBloqueioPV.sResponsavelLib = objReserva.sResponsavel
        objBloqueioPV.dtDataLib = gdtDataHoje

        'Libera o bloqueio do pedido de venda
        lErro = CF("BloqueioEstoquePV_Libera", objBloqueioPV)
        If lErro <> SUCESSO And lErro <> 30037 Then Error 23995

        objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_PARCIAL
        objBloqueioPV.lPedidoDeVendas = objReserva.lDocOrigem
        objBloqueioPV.sCodUsuario = objReserva.sCodUsuario
        objBloqueioPV.sResponsavel = objReserva.sResponsavel
        objBloqueioPV.dtData = gdtDataHoje
        objBloqueioPV.iSequencial = 0

        'Grava o blqueio
        lErro = CF("BloqueioPV_Grava", objBloqueioPV)
        If lErro <> SUCESSO Then Error 23996

    End If

    If dPedidoFaltaReservar <> 0 And lAchou = 46032 Then
                                        'Erro na leitura(AD_SQL_SEM_DADOS)

        objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_PARCIAL
        objBloqueioPV.lPedidoDeVendas = objReserva.lDocOrigem
        objBloqueioPV.sCodUsuario = objReserva.sCodUsuario
        objBloqueioPV.sResponsavel = objReserva.sResponsavel
        objBloqueioPV.dtData = gdtDataHoje
        objBloqueioPV.iSequencial = 0

        'Grava o bloqueio
        lErro = CF("BloqueioPV_Grava", objBloqueioPV)
        If lErro <> SUCESSO Then Error 23997

    End If

    Bloqueio_ReservaGrava = SUCESSO

    Exit Function

Erro_Bloqueio_ReservaGrava:

    Bloqueio_ReservaGrava = Err

    Select Case Err

        Case 23991, 23992, 23993, 23994, 23995, 23996, 23997

        Case 25663
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RESERVA_BLOQUEIO_TOTAL_EXISTENTE", Err, objBloqueioPV.lPedidoDeVendas)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149512)

    End Select

    Exit Function

End Function

Function Reserva_Exclui(ByVal objReserva As ClassReserva) As Long
'Exclui do BD a Reserva passada em objReserva

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Reserva_Exclui

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 30004

    lErro = CF("Reserva_Exclui_EmTrans", objReserva)
    If lErro <> SUCESSO And lErro <> 30007 Then Error 30023
    If lErro = 30007 Then Error 30007
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 30023

    Reserva_Exclui = SUCESSO

    Exit Function

Erro_Reserva_Exclui:

    Reserva_Exclui = Err

    Select Case Err

        Case 30004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 30007
            'Reserva inexistente
        
        Case 30023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149513)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Reserva_Exclui_EmTrans(ByVal objReserva As ClassReserva) As Long
'Exclui do BD a Reserva passada em objReserva

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 8) As Long
Dim lCodigo As Long
Dim tReserva As typeReserva
Dim tItemPedido As typeItemPedido
Dim objItemPedido As New ClassItemPedido
Dim objMovEstoque As New ClassMovEstoque
Dim objProduto As New ClassProduto
Dim dQuantReservada As Double
Dim dDifQtdeReserva As Double

On Error GoTo Erro_Reserva_Exclui_EmTrans

    'Inicializa comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 30003
    Next

    With tReserva

        .sProduto = String(STRING_PRODUTO, 0)
        .sCodUsuario = String(STRING_SIGLA_USUARIO, 0)
        .sResponsavel = String(STRING_RESPONSAVEL_RESERVA, 0)

        'Seleciona a Reserva a ser excluída
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Produto, Almoxarifado, TipoDoc, DocOrigem, NumIntOrigem, Quantidade, DataReserva, DataValidade, CodUsuario, Responsavel FROM Reserva WHERE FilialEmpresa = ? AND Codigo = ?", 0, .sProduto, .iAlmoxarifado, .iTipoDoc, .lDocOrigem, .lNumIntOrigem, .dQuantidade, .dtDataReserva, .dtDataValidade, .sCodUsuario, .sResponsavel, giFilialEmpresa, objReserva.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 30005

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30006

        'Se não encontrou a Reserva --> erro
        If lErro = AD_SQL_SEM_DADOS Then Error 30007

        'Lock da Reserva
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 30008

        'Preenche objeto com os Dados lidos no Banco de Dados
        objReserva.sProduto = .sProduto
        objReserva.iAlmoxarifado = .iAlmoxarifado
        objReserva.iTipoDoc = .iTipoDoc
        objReserva.lDocOrigem = .lDocOrigem
        objReserva.lNumIntOrigem = .lNumIntOrigem
        objReserva.dQuantidade = .dQuantidade
        objReserva.dtDataValidade = .dtDataValidade
        objReserva.dtDataReserva = .dtDataReserva
        objReserva.sCodUsuario = .sCodUsuario
        objReserva.sResponsavel = .sResponsavel
        'objReserva.iFilialEmpresa = giFilialEmpresa
        
    End With

    lErro = CF("Reserva_Grava_MovtoEstoque", objReserva, -objReserva.dQuantidade)
    If lErro <> SUCESSO Then Error 30527

    'Exclui Reserva do Banco de Dados
    lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM Reserva", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 30019

    If objReserva.iTipoDoc = TIPO_PEDIDO_COD Then

        lErro = CF("Reserva_Exclui_TipoPedidoCod", objReserva, alComando(2), alComando(3), alComando(5), alComando(6))
        If lErro <> SUCESSO Then Error 25187

    ElseIf objReserva.iTipoDoc = TIPO_PEDIDO_GRADE Then
        
        lErro = CF("Reserva_Exclui_TipoPedidoGrade", objReserva, alComando(2), alComando(3), alComando(5), alComando(6), alComando(7), alComando(8))
        If lErro <> SUCESSO Then Error 25187
    
    End If

    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    lErro = CF("Estoque_TestaIntegridade", objReserva.iFilialEmpresa)
    If lErro <> SUCESSO Then Error 30594

    Reserva_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_Reserva_Exclui_EmTrans:

    Reserva_Exclui_EmTrans = Err

    Select Case Err

        Case 30003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 30005, 30006
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA", Err)

        Case 25187, 30007, 30527, 30594
            'Reserva inexistente

        Case 30008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_RESERVA", Err)

        Case 30009, 30010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", Err)

        Case 30011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", Err, objReserva.lCodigo)

        Case 30019
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RESERVA", Err, objReserva.lCodigo)

        Case 30023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 30057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149513)

    End Select

    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Bloqueio_ReservaExclui(ByVal objReserva As ClassReserva) As Long
'CHAMADA DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lAchou As Long
Dim dPedidoQuantReservada As Double
Dim objItemPedido As ClassItemPedido
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objBloqueioPV As New ClassBloqueioPV
Dim iIndice As Integer
Dim iNumItens As Integer

On Error GoTo Erro_Bloqueio_ReservaExclui

    'obtem identificacao do PedidoVenda associado à reserva
    objPedidoVenda.iFilialEmpresa = objReserva.iFilialEmpresa
    objPedidoVenda.lCodigo = objReserva.lDocOrigem

    'Lê os Itens do Pedido de Venda
    lErro = CF("PedidoDeVenda_Le_Itens", objPedidoVenda)
    If lErro <> SUCESSO Then Error 30046

    'Retira da coleção os ítens que correspondem a Produtos que tem ControleEstoque<>PRODUTO_CONTROLE_RESERVA
    iIndice = 1
    iNumItens = objPedidoVenda.colItensPedido.Count
    Do While True
        If objPedidoVenda.colItensPedido(iIndice).iControleEstoque = PRODUTO_CONTROLE_RESERVA Then
            objPedidoVenda.colItensPedido.Remove (iIndice)
            iIndice = iIndice - 1
            iNumItens = iNumItens - 1
        End If
        If iIndice = iNumItens Then Exit Do
        iIndice = iIndice + 1

    Loop

    'Calcula a QuantReservadaTotal
    For Each objItemPedido In objPedidoVenda.colItensPedido
        dPedidoQuantReservada = dPedidoQuantReservada + objItemPedido.dQuantReservada
    Next

    objBloqueioPV.lPedidoDeVendas = objReserva.lDocOrigem

    lErro = CF("BloqueioEstoquePV_Le", objBloqueioPV)
    If lErro <> SUCESSO And lErro <> 46032 Then Error 30022
                                    'Não encontrou Bloqueio

    lAchou = lErro

    If lAchou <> 46032 And objBloqueioPV.iTipoDeBloqueio <> BLOQUEIO_PARCIAL Then Error 30045

    If dPedidoQuantReservada = 0 Then

        objBloqueioPV.sCodUsuarioLib = objReserva.sCodUsuario
        objBloqueioPV.sResponsavelLib = objReserva.sResponsavel
        objBloqueioPV.dtDataLib = gdtDataHoje

        lErro = CF("BloqueioEstoquePV_Libera", objBloqueioPV)
        If lErro <> SUCESSO And lErro <> 30037 Then Error 30023

        objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_NAO_RESERVA
        objBloqueioPV.lPedidoDeVendas = objReserva.lDocOrigem
        objBloqueioPV.sCodUsuario = objReserva.sCodUsuario
        objBloqueioPV.sResponsavel = objReserva.sResponsavel
        objBloqueioPV.dtData = gdtDataHoje
        objBloqueioPV.iSequencial = 0

        lErro = CF("BloqueioPV_Grava", objBloqueioPV)
        If lErro <> SUCESSO Then Error 30024

    End If

    If dPedidoQuantReservada > 0 And lAchou = 46032 Then
                                     'Não encontrou bloqueio
        objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_PARCIAL
        objBloqueioPV.lPedidoDeVendas = objReserva.lDocOrigem
        objBloqueioPV.sCodUsuario = objReserva.sCodUsuario
        objBloqueioPV.sResponsavel = objReserva.sResponsavel
        objBloqueioPV.dtData = gdtDataHoje
        objBloqueioPV.iSequencial = 0

        lErro = CF("BloqueioPV_Grava", objBloqueioPV)
        If lErro <> SUCESSO Then Error 30025

    End If

    Bloqueio_ReservaExclui = SUCESSO

    Exit Function

Erro_Bloqueio_ReservaExclui:

    Bloqueio_ReservaExclui = Err

    Select Case Err

        Case 30046, 30022, 30023, 30024, 30025

        Case 30045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIO_INVALIDO", Err, objBloqueioPV.iSequencial, objBloqueioPV.lPedidoDeVendas, objBloqueioPV.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149514)

    End Select

    Exit Function

End Function

Function BloqueioPV_Grava(ByVal objBloqueioPV As ClassBloqueioPV) As Long
'Grava o bloqueio do pedido de venda

Dim lErro As Long
Dim iIndice As Integer
Dim iSequencial As Integer
Dim iFilialEmpresa As Integer
Dim lComando(3) As Long

On Error GoTo Erro_BloqueioPV_Grava

    'Abre os comandos
    For iIndice = 0 To 2
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 30040
    Next

    lErro = Comando_ExecutarPos(lComando(0), "SELECT FilialEmpresa FROM BloqueiosPV WHERE FilialEmpresa=? AND PedidoDeVenda=? AND Sequencial = ?", 0, iFilialEmpresa, giFilialEmpresa, objBloqueioPV.lPedidoDeVendas, objBloqueioPV.iSequencial)
    If lErro <> AD_SQL_SUCESSO Then Error 30047

    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30048

    'Se encontrou --> Alteração
    If lErro <> AD_SQL_SEM_DADOS Then

        lErro = Comando_ExecutarPos(lComando(1), "UPDATE BloqueiosPV SET TipoDeBloqueio=?, CodUsuario=? , Responsavel=? , Data=? , DataLib = ?, Observacao = ? ", lComando(0), objBloqueioPV.iTipoDeBloqueio, objBloqueioPV.sCodUsuario, objBloqueioPV.sResponsavel, objBloqueioPV.dtData, objBloqueioPV.dtDataLib, objBloqueioPV.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then Error 30049

    Else 'Se não encontrou --> Inserção
        lErro = CF("BloqueioPV_Automatico", objBloqueioPV.lPedidoDeVendas, giFilialEmpresa, iSequencial)
        If lErro <> SUCESSO Then Error 30050

        lErro = Comando_Executar(lComando(2), "INSERT INTO BloqueiosPV(FilialEmpresa,PedidoDeVenda,Sequencial,TipoDeBloqueio,CodUsuario,Responsavel,Data,DataLib,Observacao) VALUES (?,?,?,?,?,?,?,?,?)", giFilialEmpresa, objBloqueioPV.lPedidoDeVendas, iSequencial, objBloqueioPV.iTipoDeBloqueio, objBloqueioPV.sCodUsuario, objBloqueioPV.sResponsavel, objBloqueioPV.dtData, objBloqueioPV.dtDataLib, objBloqueioPV.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then Error 30051

    End If

    'Fecha os comandos
    For iIndice = 0 To 2
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    BloqueioPV_Grava = SUCESSO

    Exit Function

Erro_BloqueioPV_Grava:

    BloqueioPV_Grava = Err

    Select Case Err

        Case 30040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 30047, 30048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPV", Err)

        Case 30049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_BLOQUEIOSPV", Err, objBloqueioPV.lPedidoDeVendas)

        Case 30050

        Case 30051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_BLOQUEIOSPV", Err, objBloqueioPV.lPedidoDeVendas)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149515)

    End Select

    'Fecha os comandos
    For iIndice = 0 To 2
        lErro = Comando_Fechar(lComando(iIndice))

    Next

    Exit Function

End Function

Private Function Reserva_Grava_BD(ByVal objReserva As ClassReserva, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, tReserva As typeReserva) As Long
'Grava a Reserva no Banco de Dados

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dQuantReservadaBD As Double
Dim vbMsg As VbMsgBoxResult
Dim lNumIntDoc As Long

On Error GoTo Erro_Reserva_Grava_BD

    lErro = CF("Config_ObterNumInt", "MATConfig", NUM_PROX_INT_RESERVA, lNumIntDoc)
    If lErro <> SUCESSO Then Error 23934

    objReserva.lNumIntDoc = lNumIntDoc

    'se a origem é um pedido
    If objReserva.iTipoDoc = TIPO_PEDIDO_COD Or objReserva.iTipoDoc = TIPO_PEDIDO_GRADE Then

        tReserva.sProduto = String(STRING_PRODUTO, 0)

        'Procura reserva p/o mesmo pedido,almoxarifado e produto
        lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo, Quantidade FROM Reserva WHERE FilialEmpresa = ? AND DocOrigem = ? AND NumIntOrigem = ? AND Almoxarifado = ? AND Produto =?", 0, tReserva.lCodigo, tReserva.dQuantidade, giFilialEmpresa, objReserva.lDocOrigem, objReserva.lNumIntOrigem, objReserva.iAlmoxarifado, objReserva.sProduto)
        If lErro <> AD_SQL_SUCESSO Then Error 23895

        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23896

        'Se encontrar reserva p/o mesmo item de pedido
        If lErro = SUCESSO Then Error 55505

    End If

    objProduto.sCodigo = objReserva.sProduto

    lErro = CF("Produto_Lock", objProduto)
    If lErro <> SUCESSO Then Error 23845

    If objProduto.iGerencial = GERENCIAL Then Error 23846
    
    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_RESERVA Then Error 30091

    'Insere a Reserva
    lErro = Comando_Executar(lComando3, "INSERT INTO Reserva(NumIntDoc, FilialEmpresa, Codigo, Produto, Almoxarifado, TipoDoc, DocOrigem, NumIntOrigem, Quantidade, DataValidade, DataReserva, CodUsuario, Responsavel) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", objReserva.lNumIntDoc, objReserva.iFilialEmpresa, objReserva.lCodigo, objReserva.sProduto, objReserva.iAlmoxarifado, objReserva.iTipoDoc, objReserva.lDocOrigem, objReserva.lNumIntOrigem, objReserva.dQuantidade, objReserva.dtDataValidade, objReserva.dtDataReserva, objReserva.sCodUsuario, objReserva.sResponsavel)
    If lErro <> AD_SQL_SUCESSO Then Error 23897

    Reserva_Grava_BD = SUCESSO

    Exit Function

Erro_Reserva_Grava_BD:

    Reserva_Grava_BD = Err

    Select Case Err

        Case 23845, 23934

        Case 23846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", Err, objProduto.sCodigo)

        Case 23895, 23896
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA", Err)

        Case 23897
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RESERVA", Err)

        Case 30091
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_CONTROLE_NAO_RESERVA", Err, objReserva.sProduto)

        Case 55505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RESERVA_CODIGO_DIFERENTE", Err, tReserva.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149516)

    End Select

    Exit Function

End Function

Function Reserva_Exclui_TipoPedidoCod(ByVal objReserva As ClassReserva, ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long) As Long
'Exclui o Tipo de Pedido associado à Reserva

Dim lErro As Long
Dim lCodigo As Long
Dim tItemPedido As typeItemPedido
Dim objItemPedido As New ClassItemPedido
Dim dFator As Double
Dim dQuantReservada As Double

On Error GoTo Erro_Reserva_Exclui_TipoPedidoCod

    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo FROM PedidosDeVenda WHERE Codigo = ? AND FilialEmpresa = ?", 0, lCodigo, objReserva.lDocOrigem, objReserva.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 30012

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30013

    'Se não achou o Pedido de Venda --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30014

    lErro = Comando_LockExclusive(lComando)
    If lErro <> SUCESSO Then Error 30015

    tItemPedido.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)
    tItemPedido.sUMEstoque = String(STRING_UM_SIGLA_UM_BASE, 0)

    lErro = Comando_ExecutarPos(lComando1, "SELECT ClasseUM, UnidadeMed, QuantReservada FROM ItensPedidoDeVenda WHERE NumIntDoc = ? AND FilialEmpresa = ?", 0, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dQuantReservada, objReserva.lNumIntOrigem, objReserva.iFilialEmpresa)
    If lErro <> SUCESSO Then Error 30016

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30017

    'Se não achou o Item do Pedido --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30018

    objItemPedido.iClasseUM = tItemPedido.iClasseUM
    objItemPedido.sUnidadeMed = tItemPedido.sUnidadeMed
    objItemPedido.dQuantReservada = tItemPedido.dQuantReservada

    lErro = Comando_Executar(lComando2, "SELECT SiglaUMEstoque FROM Produtos WHERE Codigo = ?", tItemPedido.sUMEstoque, objReserva.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 30071

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30072

    'Se não achou o Produto --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30073

    objItemPedido.sUMEstoque = tItemPedido.sUMEstoque

    lErro = CF("UM_Conversao", objItemPedido.iClasseUM, objItemPedido.sUnidadeMed, objItemPedido.sUMEstoque, dFator)
    If lErro <> SUCESSO Then Error 30044

    dQuantReservada = (objItemPedido.dQuantReservada - objReserva.dQuantidade) * dFator

    lErro = Comando_ExecutarPos(lComando3, "UPDATE ItensPedidoDeVenda SET QuantReservada=?", lComando1, dQuantReservada)
    If lErro <> AD_SQL_SUCESSO Then Error 30020

    lErro = CF("Bloqueio_ReservaExclui", objReserva)
    If lErro <> SUCESSO Then Error 30021

    Reserva_Exclui_TipoPedidoCod = SUCESSO

    Exit Function

Erro_Reserva_Exclui_TipoPedidoCod:

    Reserva_Exclui_TipoPedidoCod = Err

    Select Case Err

        Case 30012, 30013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDAS1", Err)

        Case 30014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_INEXISTENTE", Err, objReserva.lCodigo)

        Case 30015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOS_DE_VENDA", Err, objReserva.lDocOrigem)

        Case 30016, 30017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", Err)

        Case 30018
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_NAO_CADASTRADO", Err, objReserva.lCodigo)

        Case 30020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", Err, objReserva.lNumIntOrigem)

        Case 30021

        Case 30044

        Case 30071, 30072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case 30073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objReserva.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149517)

    End Select

    Exit Function

End Function

Function Reserva_Grava_TipoPedidoCod(ByVal objReserva As ClassReserva, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Faz as atualizações à Nível de Item de Pedido

Dim lErro As Long
Dim tItemPedido As typeItemPedido

On Error GoTo Erro_Reserva_Grava_TipoPedidoCod

    tItemPedido.sProduto = String(STRING_PRODUTO, 0)

    lErro = Comando_ExecutarPos(lComando1, "SELECT Produto, Quantidade, QuantCancelada, QuantFaturada FROM ItensPedidoDeVenda WHERE NumIntDoc = ?", 0, tItemPedido.sProduto, tItemPedido.dQuantidade, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, objReserva.lNumIntOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 30529

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30531

    'Se não achou o Item de Pedido --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30595

    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 30904

    tItemPedido.dQuantAFaturar = tItemPedido.dQuantidade - tItemPedido.dQuantFaturada - tItemPedido.dQuantCancelada
    
    If objReserva.dQuantidade > tItemPedido.dQuantAFaturar Then Error 30906

    If objReserva.sProduto <> tItemPedido.sProduto Then Error 55506

    lErro = Comando_ExecutarPos(lComando2, "UPDATE ItensPedidoDeVenda SET QuantReservada = QuantReservada + ?", lComando1, objReserva.dQuantidade)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30805

    lErro = CF("Bloqueio_ReservaGrava", objReserva)
    If lErro <> SUCESSO Then Error 56971

    Reserva_Grava_TipoPedidoCod = SUCESSO

    Exit Function

Erro_Reserva_Grava_TipoPedidoCod:

    Reserva_Grava_TipoPedidoCod = Err

    Select Case Err

        Case 30529, 30531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", Err)

        Case 30595
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", Err, objReserva.lDocOrigem, objReserva.lNumIntOrigem)

        Case 30904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDIDODEVENDA", Err, objReserva.lDocOrigem, objReserva.lNumIntOrigem)

        Case 30906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_RESERVADA_MAIOR_FATURAR", Err, objReserva.dQuantidade, tItemPedido.dQuantAFaturar)

        Case 55506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SEM_PRODUTO", Err, objReserva.sProduto)

        Case 56971
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149518)

    End Select

    Exit Function

End Function

Function Reserva_Grava_MovtoEstoque(ByVal objReserva As ClassReserva, ByVal dDifQtdeReserva As Double) As Long
'chamar movestoque_grava pela diferenca entre o que estava no bd
'e a nova qtde de reserva, se diferentes

Dim objMovEstoque As New ClassMovEstoque
Dim lErro As Long
Dim lCodigo As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Reserva_Grava_MovtoEstoque

    objProduto.sCodigo = objReserva.sProduto
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 25188

    'Se não achou o Produto --> erro
    If lErro <> SUCESSO Then Error 25189

    lErro = CF("MovEstoque_Automatico_EmTransacao", giFilialEmpresa, lCodigo)
    If lErro <> SUCESSO Then Error 30901

    objMovEstoque.dtData = objReserva.dtDataReserva
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    objMovEstoque.lCodigo = lCodigo

    If dDifQtdeReserva > 0 Then
        objMovEstoque.iTipoMov = MOV_EST_INCLUI_RESERVA
        objMovEstoque.colItens.Add 0, MOV_EST_INCLUI_RESERVA, 0, 0, objProduto.sCodigo, objProduto.sDescricao, objProduto.sSiglaUMEstoque, dDifQtdeReserva, objReserva.iAlmoxarifado, "", objReserva.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
    Else
        objMovEstoque.iTipoMov = MOV_EST_CANCELA_RESERVA
        objMovEstoque.colItens.Add 0, MOV_EST_CANCELA_RESERVA, 0, 0, objProduto.sCodigo, objProduto.sDescricao, objProduto.sSiglaUMEstoque, -dDifQtdeReserva, objReserva.iAlmoxarifado, "", objReserva.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
    End If

    lErro = CF("MovEstoque_Grava", objMovEstoque)
    If lErro <> SUCESSO Then Error 30902

    Reserva_Grava_MovtoEstoque = SUCESSO

    Exit Function

Erro_Reserva_Grava_MovtoEstoque:

    Reserva_Grava_MovtoEstoque = Err

    Select Case Err

        Case 25188, 30901, 30902

        Case 25189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objReserva.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149519)

    End Select

    Exit Function

End Function

Function EST_Instalacao_Filial(ByVal iFilialEmpresa As Integer, ByVal sIntervaloProducao As String) As Long
'faz as inicializacoes necessarias à criacao de uma nova filial especificas do modulo

Dim lErro As Long, lComando As Long
Dim lComando2 As Long
Dim sCodigo As String
Dim sDescricao As String
Dim sConteudo As String
Dim iTipo As Integer

On Error GoTo Erro_EST_Instalacao_Filial

    'a matriz já vem pré-inicializada
    If iFilialEmpresa <> FILIAL_MATRIZ Then
        
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 41918
            
        lErro = Comando_Executar(lComando, "INSERT INTO MATConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", NUM_PROX_RESERVA, iFilialEmpresa, "Código automático da próxima Reserva", 2, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 41919
        
        lErro = Comando_Executar(lComando, "INSERT INTO MATConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", NUM_PROX_MOV_ESTOQUE, iFilialEmpresa, "Codigo do proximo Movimento de Estoque p/ esta FilialEmpresa", 2, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 41920
        
        lErro = Comando_Executar(lComando, "INSERT INTO MATConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", NUM_PROX_INVENTARIO, iFilialEmpresa, "Código do próximo Inventário p/ esta FilialEmpresa", 2, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 41921
        
        lErro = Comando_Executar(lComando, "INSERT INTO MATConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", NUM_PROX_EMPENHO, iFilialEmpresa, "Codigo do prox. Empenho p/Ordem de Producao", 2, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 41922
        
        lErro = Comando_Executar(lComando, "INSERT INTO MATConfig(Codigo,FilialEmpresa,Descricao,Tipo,Conteudo)VALUES (?,?,?,?,?)", INTERVALO_INSUMO_INIC_PRODUCAO, iFilialEmpresa, "Intervalo entre a produção dos insumos e a produção da mercadoria que utiliza os insumos produzidos.", 1, sIntervaloProducao)
        If lErro <> AD_SQL_SUCESSO Then gError 74991
        
        lErro = Comando_Executar(lComando, "INSERT INTO ESTConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", NUM_PROX_LOTE_INVENTARIO, iFilialEmpresa, "numero do proximo lote a ser gerado automaticamente", 2, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 41923
        
        lErro = Comando_Executar(lComando, "INSERT INTO ESTConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", NUM_PROX_RECEBIMENTO, iFilialEmpresa, "numero do próximo Recebimento (automático, depende de FilialEmpresa)", 2, "1")
        If lErro <> AD_SQL_SUCESSO Then gError 59292
        
        lErro = Comando_Executar(lComando, "INSERT INTO ESTConfig (Codigo,FilialEmpresa,Descricao,Tipo,Conteudo) VALUES (?,?,?,?,?)", ESTCFG_ACEITA_ESTOQUE_NEGATIVO, iFilialEmpresa, "0-> Não aceita que a quantidade em estoque seja negativa 1-> Aceita", 1, "0")
        If lErro <> AD_SQL_SUCESSO Then gError 126847
        
        Call Comando_Fechar(lComando)
    
    'Se a filial é matriz
    Else
        
        'Abre os comandos
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 76000

        lComando2 = Comando_Abrir()
        If lComando = 0 Then gError 76001

        sCodigo = String(STRING_COMCONFIG_CODIGO, 0)
        sDescricao = String(STRING_COMCONFIG_DESCRICAO, 0)
        sConteudo = String(STRING_CONTEUDO, 0)

        'Lê em MatConfig o Intervalo de Produção para a matriz
        lErro = Comando_ExecutarPos(lComando, "SELECT Codigo, FilialEmpresa,Descricao,Tipo,Conteudo FROM MATConfig WHERE Codigo=? AND FilialEmpresa=?", 0, sCodigo, iFilialEmpresa, sDescricao, iTipo, sConteudo, "INTERVALO_INSUMO_INIC_PRODUCAO", FILIAL_MATRIZ)
        If lErro <> AD_SQL_SUCESSO Then gError 76002

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76003

        'Atualiza MatConfig
        lErro = Comando_ExecutarPos(lComando2, "UPDATE MATConfig SET Codigo=?,FilialEmpresa=?,Descricao=?,Tipo=?,Conteudo=?", lComando, "INTERVALO_INSUMO_INIC_PRODUCAO", FILIAL_MATRIZ, "Intervalo entre a produção dos insumos e a produção da mercadoria que utiliza os insumos produzidos.", 1, sIntervaloProducao)
        If lErro <> AD_SQL_SUCESSO Then gError 76004

        Call Comando_Fechar(lComando)
        Call Comando_Fechar(lComando2)
        
    End If
    
    EST_Instalacao_Filial = SUCESSO
     
    Exit Function
    
Erro_EST_Instalacao_Filial:

    EST_Instalacao_Filial = gErr
     
    Select Case gErr
          
        Case 41918, 76000, 76001
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 41919 To 41923, 74991, 126847
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ARQCONFIG", gErr)
        
        Case 76002, 76003
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONFIGURACAO", gErr)
            
        Case 76004
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONFIGURACAO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149520)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


'Function NFiscalEntrada_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
''Faz a Gravação de Uma Nota Fiscal de Entrada
'
'Dim lErro As Long
'Dim objTipoDocInfo As New ClassTipoDocInfo
'Dim lTransacao As Long
'Dim lComando As Long
'Dim sNome As String
'Dim iCodigo As Integer
'Dim lNumIntDoc As Long
'Dim vbMsgRes As VbMsgBoxResult
'Dim dtDataEmissao As Date
'
'On Error GoTo Erro_NFiscalEntrada_Grava
'
'    'Abre a Transação
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then Error 35043
'
'    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
'
'    lErro = CF("TipoDocInfo_Le_Codigo",objTipoDocInfo)
'    If lErro <> SUCESSO And lErro <> 31415 Then Error 55599
'
'    If lErro <> SUCESSO Then Error 55600
'
'    'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
'    lErro = CF("NFiscalEntrada_Verifica_Existencia2",objNFiscal, objTipoDocInfo)
'    If lErro <> SUCESSO And lErro <> 61414 Then Error 55640
'
'    'Se for uma alteração só altera os dados da contabilidade
'    If lErro = SUCESSO Then
'
'        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
'
'            'gera o Número para a Nota Fiscal Interna
'            lErro = CF("NFiscalNumAuto",objNFiscal)
'            If lErro <> SUCESSO Then Error 60445
'
'        End If
'
'        lErro = CF("NFiscalEntrada_Lock_Gravacao",objNFiscal)
'        If lErro <> SUCESSO Then Error 35663
'
'        'Dispara a Gravação da Nota Fiscal
'        lErro = CF("NFiscalEntrada_Grava_BD",objNFiscal)
'        If lErro <> SUCESSO Then Error 35078
'
'        'Grava a Estatística da Filial Fornecedor
'        lErro = CF("FilialFornecedor_Grava_Estatisticas",objNFiscal)
'        If lErro <> SUCESSO Then Error 52958
'
'    End If
'
'
'    If objNFiscal.lCliente <> 0 Then
'
'        'Grava os dados contábeis (contabilidade)
'        lErro = objContabil.Contabil_Gravar_Registro(objNFiscal.lNumIntDoc, objNFiscal.lCliente, objNFiscal.iFilialCli, NAO_AVALIA_PELO_CUSTO_REAL_PRODUCAO, AVISA_LANCAMENTOS_CONTABILIZADOS, objNFiscal.lNumNotaFiscal)
'        If lErro <> SUCESSO Then Error 36728
'
'    Else
'
'        'Grava os dados contábeis (contabilidade)
'        lErro = objContabil.Contabil_Gravar_Registro(objNFiscal.lNumIntDoc, objNFiscal.lFornecedor, objNFiscal.iFilialForn, NAO_AVALIA_PELO_CUSTO_REAL_PRODUCAO, AVISA_LANCAMENTOS_CONTABILIZADOS, objNFiscal.lNumNotaFiscal)
'        If lErro <> SUCESSO Then Error 36729
'
'    End If
'
'    'Faz o Commit da Transação
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then Error 35079
'
'    NFiscalEntrada_Grava = SUCESSO
'
'    Exit Function
'
'Erro_NFiscalEntrada_Grava:
'
'    NFiscalEntrada_Grava = Err
'
'    Select Case Err
'
'        Case 35043
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
'
'        Case 35078, 35663, 36728, 36729, 52958, 55606, 55612
'
'        Case 35079
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
'
'        Case 55599, 55640
'
'        Case 55600
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err)
'
'        Case 60445 'Tratado na Rotina chamada
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149521)
'
'    End Select
'
'    Call Transacao_Rollback
'
'    Exit Function
'
'End Function
'
'Private Function NFiscalEntrada_Pesquisa_Existencia(objNFiscal As ClassNFiscal, iTipo) As Long
''pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
'
'Dim lErro As Long
'Dim objTipoDocInfo As New ClassTipoDocInfo
'Dim lTransacao As Long
'Dim lComando As Long
'Dim sNome As String
'Dim iCodigo As Integer
'Dim lNumIntDoc As Long
'Dim vbMsgRes As VbMsgBoxResult
'Dim dtDataEmissao As Date
'
'On Error GoTo Erro_NFiscalEntrada_Pesquisa_Existencia
'
'    'Abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then Error 35044
'
'    If iTipo = DOCINFO_NF_EXTERNA Then
'
'        'pesquisa se nota fiscal com os mesmos identificadores está cadastrada
'        lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM NFiscalTipoDocInfo WHERE Fornecedor = ? AND FilialForn = ? AND Serie = ? AND NumNotaFiscal = ? AND DataEmissao = ? AND Status <> ? AND Tipo = ? AND FilialEmpresa = ?", lNumIntDoc, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, STATUS_EXCLUIDO, DOCINFO_NF_EXTERNA, giFilialEmpresa)
'        If lErro <> AD_SQL_SUCESSO Then Error 55601
'
'        lErro = Comando_BuscarPrimeiro(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55602
'
'        'se já estiver cadastrada a nota fiscal ==> erro
'        If lErro = AD_SQL_SUCESSO Then Error 55603
'
'        'pesquisa se já está cadastrada uma nota fiscal com os mesmos identificadores em data proxima
'        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM NFiscalTipoDocInfo WHERE Fornecedor = ? AND FilialForn = ? AND Serie = ? AND NumNotaFiscal = ? AND DataEmissao > ? AND DataEmissao < ? AND Status <> ? AND Tipo = ? AND FilialEmpresa = ?", dtDataEmissao, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, (objNFiscal.dtDataEmissao - PERIODO_EMISSAO), (objNFiscal.dtDataEmissao + PERIODO_EMISSAO), STATUS_EXCLUIDO, DOCINFO_NF_EXTERNA, giFilialEmpresa)
'        If lErro <> AD_SQL_SUCESSO Then Error 55604
'
'        lErro = Comando_BuscarPrimeiro(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55605
'
'        If lErro = AD_SQL_SUCESSO Then
'
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NF_EXTERNA_DATA_PROXIMA", objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, dtDataEmissao)
'            If vbMsgRes = vbNo Then Error 55606
'
'        End If
'
'    ElseIf iTipo = DOCINFO_NF_INT_ENTRADA Then
'
'        'pesquisa se nota fiscal com os mesmos identificadores está cadastrada
'        lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM NFiscalTipoDocInfo WHERE FilialEmpresa =? AND Serie = ? AND NumNotaFiscal = ?  AND DataEmissao = ? AND Status <> ? AND(Tipo = ? OR Tipo = ?) ", lNumIntDoc, giFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, STATUS_EXCLUIDO, DOCINFO_NF_INT_ENTRADA, DOCINFO_NF_INT_SAIDA)
'        If lErro <> AD_SQL_SUCESSO Then Error 55607
'
'        lErro = Comando_BuscarPrimeiro(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55608
'
'        'se já estiver cadastrada a nota fiscal ==> erro
'        If lErro = AD_SQL_SUCESSO Then Error 55609
'
'        'pesquisa se já está cadastrada uma nota fiscal com os mesmos identificadores em data proxima
'        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM NFiscalTipoDocInfo WHERE FilialEmpresa =? AND Serie = ? AND NumNotaFiscal = ?  AND DataEmissao > ? AND DataEmissao < ? AND Status <> ? AND (Tipo = ? OR Tipo = ?) ", dtDataEmissao, giFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao - PERIODO_EMISSAO, objNFiscal.dtDataEmissao + PERIODO_EMISSAO, STATUS_EXCLUIDO, DOCINFO_NF_INT_ENTRADA, DOCINFO_NF_INT_SAIDA)
'        If lErro <> AD_SQL_SUCESSO Then Error 55610
'
'        lErro = Comando_BuscarPrimeiro(lComando)
'        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 55611
'
'        If lErro = AD_SQL_SUCESSO Then
'
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NF_INTERNA_DATA_PROXIMA", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, dtDataEmissao)
'            If vbMsgRes = vbNo Then Error 55612
'
'        End If
'
'    End If
'
'    'Fecha os Comandos
'    Call Comando_Fechar(lComando)
'
'    NFiscalEntrada_Pesquisa_Existencia = SUCESSO
'
'    Exit Function
'
'Erro_NFiscalEntrada_Pesquisa_Existencia:
'
'    NFiscalEntrada_Pesquisa_Existencia = Err
'
'    Select Case Err
'
'        Case 35044
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 55601, 55602, 55604, 55605, 55607, 55608, 55610, 55611
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", Err)
'
'        Case 55603
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_NFISCAL_EXTERNA", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
'
'        Case 55609
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_NFISCAL_INTERNA", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149522)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Function NFiscalEntrada_Grava_BD(ByVal objNFiscal As ClassNFiscal) As Long
'Insere uma Nova Nota Fiscal de Entrada no Banco de Dados

Dim alComando(0 To 7) As Long
Dim lNumIntDoc As Long
Dim lErro As Long
Dim iClasseDocCPR As Integer
Dim lNumIntDocCPR As Long
Dim objItemNF As ClassItemNF
Dim objMovEstoque As ClassMovEstoque
Dim objMovEstoqueBaixa As ClassMovEstoque
Dim iIndice As Integer
Dim iTipoMovtoEstoque As Integer
Dim lCodigo As Long
Dim sDocOrigem As String
Dim iTipoMovtoEstoqueBenef As Integer, objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objRastroItemNF As ClassRastroItemNF
Dim objRastroMovto As ClassRastreamentoMovto
Dim colInsumos As Collection
Dim colRastroMovto As Collection
Dim dFator As Double
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dQuantItemNFUMEst As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim objProduto As ClassProduto
Dim objProduto2 As ClassProduto, dCusto As Double
Dim bTrataRateioMovEst As Boolean, bExisteCtrlEst As Boolean

On Error GoTo Erro_NFiscalEntrada_Grava_BD

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 35054
    Next

    sDocOrigem = String(STRING_TIPO_DOC_INFO_SIGLA, 0)

    'Lê TipoMovtoEstoque em TiposDocInfo
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then gError 59259
    
    iTipoMovtoEstoque = objTipoDocInfo.iTipoMovtoEstoque
    sDocOrigem = objTipoDocInfo.sSigla
    iTipoMovtoEstoqueBenef = objTipoDocInfo.iTipoMovtoEstoqueBenef

    'Gera Conta a Pagar e devolve iClasseDocCPR e lNumIntDocCPR
    lErro = CF("CPR_Gera", objNFiscal, iClasseDocCPR, lNumIntDocCPR)
    If lErro <> SUCESSO Then gError 35061

    objNFiscal.iClasseDocCPR = iClasseDocCPR
    objNFiscal.lNumIntDocCPR = lNumIntDocCPR

    'Obtem o Número Interno da nova Nota Fiscal
    lErro = CF("NFiscal_Automatico", lNumIntDoc)
    If lErro <> SUCESSO Then gError 35082

    objNFiscal.lNumIntDoc = lNumIntDoc
    
    lErro = CF("NFiscal_Grava_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError 35082

    With objNFiscal

        .iNFe = SerieEletronica(.sSerie)

        'Insere a nova Nota Fiscal de Entrada no BD
        lErro = Comando_Executar(alComando(0), "INSERT INTO NFiscal (DataAtualizacao, NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,FornecedorBenef,FilialFornBenef,VolumeNumero,FreteRespons,DataReferencia,Observacao, NumRecebimento, HoraEntrada, Moeda, TaxaMoeda, NFe, ValorDescontoTit, ValorItens, ValorDescontoItens, ChvNFe, HoraEmissao, DataCadastro, HoraCadastro, UsuCadastro) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            gdtDataHoje, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lFornecedorBenef, .iFilialFornBenef, .sVolumeNumero, .iFreteRespons, .dtDataReferencia, .sObservacao, .lNumRecebimento, CDbl(.dtHoraEntrada), .iMoeda, .dTaxaMoeda, .iNFe, .dValorDescontoTit, .dValorItens, .dValorDescontoItens, objNFiscal.sChvNFe, _
            CDbl(IIf(.dtDataSaida <> DATA_NULA, IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraSaida, .dtHoraEmissao, .dtHoraSaida), IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraEntrada, .dtHoraEmissao, .dtHoraEntrada))), _
            gdtDataHoje, CDbl(Time), gsUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 35062
    End With

    lErro = CF("NFiscal_GravaTributacao", objNFiscal, alComando(4), alComando(5))
    If lErro <> SUCESSO Then gError 22855

    'Para cada Ítem da Nota
    For Each objItemNF In objNFiscal.colItensNF
    
        objItemNF.lNumIntNF = objNFiscal.lNumIntDoc

        'Obtém o Número Interno do Ítem
        lErro = CF("NFiscalItem_Automatico", lNumIntDoc)
        If lErro <> SUCESSO Then gError 35063

        objItemNF.lNumIntDoc = lNumIntDoc

        With objItemNF
            'Insere o Ítem no Banco de dados
            lErro = Comando_Executar(alComando(1), "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, NumIntDocOrig,Almoxarifado, Ccl) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lNumIntNF, .iItem, STATUS_LANCADO, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, lNumIntDoc, .lNumIntTrib, 0, .iAlmoxarifado, .sCcl)
            If lErro <> AD_SQL_SUCESSO Then gError 35064
        End With
        
        lErro = CF("ItensNFGrade_Inclui", objItemNF)
        If lErro <> SUCESSO Then Error 30739
    
        lErro = CF("NFiscal_GravaTributacaoItem", objItemNF, alComando(3), alComando(6))
        If lErro <> SUCESSO Then gError 22856
        
        objItemNF.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
        objItemNF.objInfoAdicDocItem.lNumIntDocItem = objItemNF.lNumIntDoc
        lErro = CF("InfoAdicDocItem_Insere", objItemNF.objInfoAdicDocItem, alComando(7))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Next

    lErro = CF("NFe_Grava_Vinculo_Xml", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("NFiscalItens_Grava_Customizado", objNFiscal)
    If lErro <> SUCESSO Then gError 177049
    
    'Faz a baixa de um recebimento
    'a baixa do recebimento deve ocorrer antes da gravacao da nota fiscal para que os valores dos saldos estejam corretos e permitam calcular o novo custo do produto, se for o caso.
    'Só ocorre a Baixa de Recebimento se o Número do Recebimento estiver preenchido
    If objNFiscal.lNumRecebimento > 0 Then
        
        'Lê o Recebimento e Verifica os Dados
        lErro = NFiscal_Valida_Recebimento(objNFiscal)
        If lErro <> SUCESSO Then gError 64246
        
        lErro = CF("NFiscalEntrada_Baixa_Recebimento", objNFiscal)
        If lErro <> SUCESSO Then gError 35080
    
    End If
                
    If iTipoMovtoEstoque > 0 Then
    
        Set objMovEstoque = New ClassMovEstoque
    
        'Obtém o Código do Movimento
        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then gError 35137
    
        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.iTipoMov = iTipoMovtoEstoque
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = sDocOrigem & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)

        'Se for para incluir o valor do frete e outras depesas no custo do ItemNF
        If gobjEST.iIncluiFreteNoCusto = INCLUI_FRETE_CUSTO Then
            Call Inclui_Frete_CustoItemNF(objNFiscal)
        End If
        
        bTrataRateioMovEst = False
        If objNFiscal.lNumIntNotaOriginal <> 0 Then
            'Se é uma despesa aduaneira e tem NF original rateia de qualquer jeito
            If objNFiscal.iTipoNFiscal = DOCINFO_DDAI Or objNFiscal.iTipoNFiscal = DOCINFO_FDDAI Then
                bTrataRateioMovEst = True
            End If
            'Se for um conhecimento de transporte só vai ratear se todos itens não controlarem estoque
            If objNFiscal.iTipoNFiscal = DOCINFO_CFECT Or objNFiscal.iTipoNFiscal = DOCINFO_NFSET Or objNFiscal.iTipoNFiscal = DOCINFO_NFSEFT Then
                bExisteCtrlEst = False
                For Each objItemNF In objNFiscal.colItensNF
                    Set objProduto = New ClassProduto
                    objProduto.sCodigo = objItemNF.sProduto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 89464
                    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                        bExisteCtrlEst = True
                        Exit For
                    End If
                Next
                If Not bExisteCtrlEst Then
                    bTrataRateioMovEst = True
                    If objNFiscal.colItensNF.Count > 2 Then gError 211177 'Só pode ter 1 ou dois itens a serem rateados (com e sem crédito)
                    If objNFiscal.colItensNF.Count = 2 Then
                        If objNFiscal.colItensNF.Item(1).objTributacao.dICMSCredito = 0 And objNFiscal.colItensNF.Item(2).objTributacao.dICMSCredito = 0 Then gError 211178  'Não poder ter crédito nos dois
                        If objNFiscal.colItensNF.Item(1).objTributacao.dICMSCredito <> 0 And objNFiscal.colItensNF.Item(2).objTributacao.dICMSCredito <> 0 Then gError 211179  'Não poder não ter crédito nos dois
                    End If
                End If
            End If
        End If
        
'mario distribuicao substituicao do codigo original acrescentando o tab de distribuicao
               
        iIndice = 0
        'Adiciona itens ao Movimento
        For Each objItemNF In objNFiscal.colItensNF
            
            iIndice = iIndice + 1
            
            If objItemNF.colItensRomaneioGrade.Count > 0 Then
                
                objItemNF.dCusto = 0
                
                'calcula o custo do item da NF em função do credito do ICMS e do IPI
                lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
                If lErro <> SUCESSO Then gError 60787
                
                Set objProduto2 = New ClassProduto
            
                objProduto2.sCodigo = objItemNF.sProduto
                
                lErro = CF("Produto_Le", objProduto2)
                If lErro <> SUCESSO And lErro <> 23080 Then gError 89464
                
                For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                    Set objProduto = New ClassProduto
                
                    objProduto.sCodigo = objItemRomaneio.sProduto
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 89464
                    
                    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                        
                        If objProduto2.iKitVendaComp = MARCADO Then
                            dFator = 1 'A UM do componente já está na UM de Estoque e a classe UM pode ser diferente da Classe UM do item da NF, logo a conversão feita para grade é desnecessária e pode dar erro
                        Else
                            lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemRomaneio.sUMEstoque, dFator)
                            If lErro <> SUCESSO Then gError 89464
                        End If
        
                        dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
        
'                        'calcula o custo do item da NF em função do credito do ICMS e do IPI
'                        lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
'                        If lErro <> SUCESSO Then gError 60787
        
                        dSomaAlocacao = 0
                                                                        
                        For Each objReservaItem In objItemRomaneio.colLocalizacao
                                                
                            dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                            
                            Set colInsumos = New Collection
                            Set colRastroMovto = New Collection
                            
                            If objProduto2.iKitVendaComp = MARCADO Then
                                'dCusto = 0
                                'Não existem parâmetros de rateio para entrada de kit de venda então o valor vai dividido igualmente entre os componentes
                                dCusto = ((objItemNF.dCusto / objItemNF.colItensRomaneioGrade.Count) / objItemRomaneio.dQuantidade) * objReservaItem.dQuantidade
                            Else
                                'O cálculo anterior dava um resultado errado e não considerava os impostos
                                'dCusto = ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst
                                dCusto = ((objItemNF.dCusto / objItemNF.dQuantidade) * (objReservaItem.dQuantidade / dFator))
                            End If
                                                
                            'Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, dCusto, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                            
                            '################################################################
                            'Inserido por Wagner 02/06/2006
                            'Guarda o Rastreamento dos ItensNF
                            If Not (objNFiscal.objRastreamento Is Nothing) Then
                                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
                                If lErro <> SUCESSO Then gError 180072
                            End If
                            '################################################################
                             
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                            objItemMovEstoque.iControleEstoque = objItemRomaneio.iControleEstoque
                            objItemMovEstoque.sSiglaUMEst = objItemRomaneio.sUMEstoque
''''                            objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
                            
                            If TipoDocInfo_ZeraQtdeMovto(objNFiscal.iTipoNFiscal) Then
                            
                                objItemMovEstoque.dQuantidade = 0
                                
                            End If
                            
                        Next
                        
                        '###########################################################
                        'Alterado por Wagner 16/11/04
                        'Calcula o número de casas decimais do Formato de Estoque
'                        iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                        dAcrescimo = 0
'                        If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                            'alteracao por tulio070303...
'                        If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > dAcrescimo + QTDE_ESTOQUE_DELTA Then

                        dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)

                        If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > QTDE_ESTOQUE_DELTA Then
                        '###########################################################
                        
                            If dSomaAlocacao < dQuantItemNFUMEst Then gError 89465
                            
                            If dSomaAlocacao > dQuantItemNFUMEst Then gError 89467
                    
                        End If
                        
                    End If
                Next
            Else
            
                'Se for uma NF de despesa aduaneira com NF orginal (ou outro tipo que precise ratear o custo)
                'If (objNFiscal.iTipoNFiscal = DOCINFO_DDAI Or objNFiscal.iTipoNFiscal = DOCINFO_FDDAI Or (objNFiscal.iTipoNFiscal = DOCINFO_CFECT And objItemNF.ColAlocacoes.Count = 0)) And objNFiscal.lNumIntNotaOriginal <> 0 Then
                If bTrataRateioMovEst Then
                            
                    'rateia o custo entre os itens da NF original
                    lErro = NF_Trata_MovEst_DDAI(objNFiscal, objItemNF, objMovEstoque)
                    If lErro <> SUCESSO Then gError 89463
 
                Else
                
                    If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                        
                        If objItemNF.ColAlocacoes.Count = 0 Then gError 89462
        
                        lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                        If lErro <> SUCESSO Then gError 89464
        
                        dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
        
                        'calcula o custo do item da NF em função do credito do ICMS e do IPI
                        lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
                        If lErro <> SUCESSO Then gError 60787
        
                        dSomaAlocacao = 0
                                                                        
                        For Each objItemNFAloc In objItemNF.ColAlocacoes
                                                
                            dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
                            
                            Set colInsumos = New Collection
                            Set colRastroMovto = New Collection
                                                
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                            
                            'Guarda o Rastreamento dos ItensNF
                            If Not (objNFiscal.objRastreamento Is Nothing) Then
                                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
                                If lErro <> SUCESSO Then gError 89463
                            End If
                                                    
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                            objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                            objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                            objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                            objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
                            
                            'se for uma nota de complemento de preço ou outra que nao deve afetar qtde ==> zera a quantidade pois somente o custo será utilizado
                            'If (objNFiscal.iTipoNFiscal = DOCINFO_NFIEICM) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEICM) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEECP) Or (objNFiscal.iTipoNFiscal = DOCINFO_CFECT) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEIPI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIECP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEIPI) Or (objNFiscal.iTipoNFiscal = DOCINFO_CFEV) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFETEL) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFES) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIECS) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEECS) Or (objNFiscal.iTipoNFiscal = DOCINFO_DDAI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEESAC) Then
                            If TipoDocInfo_ZeraQtdeMovto(objNFiscal.iTipoNFiscal) Then
                            
                                objItemMovEstoque.dQuantidade = 0
                                
                            End If
                            
                        Next
                        
                        '############################################################
                        'Alterado por Wagner 16/11/04
                        'Calcula o número de casas decimais do Formato de Estoque
    '                    iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
    '
    '                    dAcrescimo = 0
    '                    If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
    '
    '                        'alteracao por tulio070303...
    '                    If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > dAcrescimo + QTDE_ESTOQUE_DELTA Then
                        
                        dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)
    
                        
                        If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > QTDE_ESTOQUE_DELTA Then
                        '############################################################
                        
                            If dSomaAlocacao < dQuantItemNFUMEst Then gError 89465
                            
                            If dSomaAlocacao > dQuantItemNFUMEst Then gError 89467
                    
                        End If
                        
                    End If
                    
                End If
            End If
        Next
                        
                        
'mario distribuicao fim da alteração para incluir o tab de distribuicao
                        
                        
                        
'mario parte substituida
'        'Adiciona os ítem ao Movimento
'        For Each objItemNF In objNFiscal.colItensNF
'
'            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
'
'                'calcula o custo do item da NF em função do credito do ICMS e do IPI
'                lErro = CF("NFiscalEntrada_Calcula_Custo",objItemNF)
'                If lErro <> SUCESSO Then gError 60787
'
'                Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNF.sUnidadeMed, objItemNF.dQuantidade, objItemNF.iAlmoxarifado, objItemNF.sAlmoxarifadoNomeRed, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
'
'                Set objItemMovEstoque.colRastreamentoMovto = New Collection
'                Set objItemMovEstoque.colApropriacaoInsumo = New Collection
'
'                'Guarda o Rastreamento dos ItensNF
'                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
'                If lErro <> SUCESSO Then gError 83235
'
'                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'                objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
'                objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
'                objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
'                objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
'
'
'                'se for uma nota de complemento de preço ou outra que nao deve afetar qtde ==> zera a quantidade pois somente o custo será utilizado
'                If (objNFiscal.iTipoNFiscal = DOCINFO_NFEECP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIECP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEECPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIECPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_CFECT) Then
'                    objItemMovEstoque.dQuantidade = 0
'                End If
'
'            End If
'
'        Next
'
'mario fim da continuacao do codigo original


        'Gera um Movimento de Estoque
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then gError 35065

    End If
    
    If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEBEN) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEEBEN) Then
        
        'Gera um Movimento de Estoque
        lErro = CF("NFiscalEntrada_Grava_MovimentosEst", objNFiscal, iTipoMovtoEstoqueBenef, sDocOrigem)
        If lErro <> SUCESSO Then gError 52530
        
    End If
        
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalEntrada_Grava_BD = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Grava_BD:

    NFiscalEntrada_Grava_BD = gErr

    Select Case gErr

        Case 35054
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 35061, 35063, 35082, 35065, 35080, 35137, 22855, 22856, 52530, 59259, _
             60787, 64246, 89463, 89464, 30739, 180072, 177049

        Case 35062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 35064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 35092, 35093
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO", gErr, objNFiscal.iTipoNFiscal)

        Case 35094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 89462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INEXISTENTE", gErr, iIndice)

        Case 89465
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INCOMPLETA", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case 89467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case 211177
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_CFECT_QTDE_ITENS_INVALIDA", gErr)

        Case 211178, 211179
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_CFECT_ITENS_TRIB_INVALIDA", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149523)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Sub Inclui_Frete_CustoItemNF(ByVal objNFiscal As ClassNFiscal)
'Inclui o valor do Frete, ValorSeguro e outras despesas ao Custo dos ItensNF
'proporcional ao Valor do ItemNF

Dim objItemNF As ClassItemNF
    
    'Para cada item da nota fiscal
    'Adiciona o valor rateado de Frete\Seguro e Outras Despesas (O Valor Creditado já está na tributação do Item)
    For Each objItemNF In objNFiscal.colItensNF

        'Se é importação ignora as outras despesas
        If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC Then
            If gobjCRFAT.iNFImportacaoTribFlag12 Then
                objItemNF.dCusto = Arredonda_Moeda(objItemNF.dCusto + objItemNF.objTributacaoItemNF.dICMSValor + objItemNF.objTributacao.dValorFreteItem + objItemNF.objTributacao.dValorSeguroItem + objItemNF.dDespImpValorRateado + objItemNF.objTributacao.dIIValor)
            Else
                objItemNF.dCusto = Arredonda_Moeda(objItemNF.dCusto + objItemNF.objTributacaoItemNF.dICMSValor + objItemNF.objTributacao.dValorFreteItem + objItemNF.objTributacao.dValorSeguroItem + objItemNF.dDespImpValorRateado)
            End If
        Else
            objItemNF.dCusto = Arredonda_Moeda(objItemNF.dCusto + objItemNF.objTributacao.dValorFreteItem + objItemNF.objTributacao.dValorOutrasDespesasItem + objItemNF.objTributacao.dValorSeguroItem)
        End If
        
    Next
    
End Sub

Function NFiscalEntrada_Calcula_Custo(ByVal objItemNF As ClassItemNF) As Long
'calcula o custo do item da NF em função do credito do ICMS, do IPI e do PIS

On Error GoTo Erro_NFiscalEntrada_Calcula_Custo
    
    objItemNF.dCusto = objItemNF.dCusto + objItemNF.dValorTotal
    
    If objItemNF.objTributacaoItemNF.dICMSCredito <> 0 Then
        objItemNF.dCusto = objItemNF.dCusto - objItemNF.objTributacaoItemNF.dICMSCredito
    Else
        objItemNF.dCusto = objItemNF.dCusto + objItemNF.objTributacaoItemNF.dICMSSubstValor
    End If
    
    If objItemNF.objTributacaoItemNF.dIPICredito = 0 Then
        objItemNF.dCusto = objItemNF.dCusto + objItemNF.objTributacaoItemNF.dIPIValor
    End If
        
    If objItemNF.objTributacaoItemNF.dPISCredito <> 0 Then
        objItemNF.dCusto = objItemNF.dCusto - objItemNF.objTributacaoItemNF.dPISCredito
    End If
        
    If objItemNF.objTributacaoItemNF.dCOFINSCredito <> 0 Then
        objItemNF.dCusto = objItemNF.dCusto - objItemNF.objTributacaoItemNF.dCOFINSCredito
    End If
        
    NFiscalEntrada_Calcula_Custo = SUCESSO
    
    Exit Function
    
Erro_NFiscalEntrada_Calcula_Custo:

    NFiscalEntrada_Calcula_Custo = Err
    
    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149524)

    End Select

    Exit Function
            
End Function

Private Function NFiscalEntrada_Calcula_Custo_ImpCC(ByVal objItemNF As ClassItemNF) As Long
'calcula o custo do item da NF em função do credito do ICMS, do IPI e do PIS

On Error GoTo Erro_NFiscalEntrada_Calcula_Custo_ImpCC
    
    objItemNF.dCusto = objItemNF.dCusto + objItemNF.dPrecoUnitario
    
    If objItemNF.objTributacaoItemNF.dICMSCredito <> 0 Then
        objItemNF.dCusto = objItemNF.dCusto - objItemNF.objTributacaoItemNF.dICMSCredito
    Else
        objItemNF.dCusto = objItemNF.dCusto + objItemNF.objTributacaoItemNF.dICMSSubstValor
    End If
    
    If objItemNF.objTributacaoItemNF.dIPICredito = 0 Then
        objItemNF.dCusto = objItemNF.dCusto + objItemNF.objTributacaoItemNF.dIPIValor
    End If
        
    If objItemNF.objTributacaoItemNF.dPISCredito <> 0 Then
        objItemNF.dCusto = objItemNF.dCusto - objItemNF.objTributacaoItemNF.dPISCredito
    End If
        
    If objItemNF.objTributacaoItemNF.dCOFINSCredito <> 0 Then
        objItemNF.dCusto = objItemNF.dCusto - objItemNF.objTributacaoItemNF.dCOFINSCredito
    End If
        
    NFiscalEntrada_Calcula_Custo_ImpCC = SUCESSO
    
    Exit Function
    
Erro_NFiscalEntrada_Calcula_Custo_ImpCC:

    NFiscalEntrada_Calcula_Custo_ImpCC = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149524)

    End Select

    Exit Function
            
End Function

Function NFiscalEntradaImpParc_Calcula_Custo(ByVal objNF As ClassNFiscal, ByVal objItemNF As ClassItemNF) As Long
'calcula o custo do item da NF em função do credito do ICMS, do IPI e do PIS

Dim lErro As Long
Dim objNFOrig As New ClassNFiscal
Dim objItemNFOrig As ClassItemNF
Dim dFator As Double, bAchou As Boolean

On Error GoTo Erro_NFiscalEntradaImpParc_Calcula_Custo

    objNFOrig.lNumIntDoc = objNF.lNumIntNotaOriginal
    
    lErro = CF("NFiscal_Le", objNFOrig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFOrig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Lê a parte de Tributação
    lErro = CF("NFiscal_Le_Tributacao", objNFOrig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê a parte de tributacao dos Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le_Trib", objNFOrig)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    bAchou = False
    For Each objItemNFOrig In objNFOrig.colItensNF
        If objItemNFOrig.sProduto = objItemNF.sProduto Then
            bAchou = True
            Exit For
        End If
    Next
    If bAchou = False Then gError 211205
    dFator = objItemNF.dQuantidade / objItemNFOrig.dQuantidade
        
    objItemNF.dCusto = objItemNF.dCusto + objItemNF.dValorTotal + Arredonda_Moeda(objItemNFOrig.objTributacao.dICMSValor * dFator, 4)
    
    If objItemNFOrig.objTributacaoItemNF.dICMSCredito <> 0 Then
        objItemNF.dCusto = objItemNF.dCusto - Arredonda_Moeda(objItemNFOrig.objTributacaoItemNF.dICMSCredito * dFator, 4)
    Else
        objItemNF.dCusto = objItemNF.dCusto + Arredonda_Moeda(objItemNFOrig.objTributacaoItemNF.dICMSSubstValor * dFator, 4)
    End If
    
    If objItemNFOrig.objTributacaoItemNF.dIPICredito = 0 Then
        objItemNF.dCusto = objItemNF.dCusto + Arredonda_Moeda(objItemNFOrig.objTributacaoItemNF.dIPIValor * dFator, 4)
    End If
        
    If objItemNFOrig.objTributacaoItemNF.dPISCredito <> 0 Then
        objItemNF.dCusto = objItemNF.dCusto - Arredonda_Moeda(objItemNFOrig.objTributacaoItemNF.dPISCredito * dFator, 4)
    End If
        
    If objItemNFOrig.objTributacaoItemNF.dCOFINSCredito <> 0 Then
        objItemNF.dCusto = objItemNF.dCusto - Arredonda_Moeda(objItemNFOrig.objTributacaoItemNF.dCOFINSCredito * dFator, 4)
    End If
        
    NFiscalEntradaImpParc_Calcula_Custo = SUCESSO
    
    Exit Function
    
Erro_NFiscalEntradaImpParc_Calcula_Custo:

    NFiscalEntradaImpParc_Calcula_Custo = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 211205
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_ORIG_PROD_NAO_ENCONTRADO", gErr, objItemNF.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211206)

    End Select

    Exit Function
            
End Function

Function NFiscalEntrada_Exclui(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Exclui a Nota Fiscal do BD
'Ao tratar a exclusão de notas de compra não esquecer de atualizar os saldos.

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Long
Dim lNumIntDoc As Long
Dim lNumIntDocItem As Long
Dim lNumIntDocCPR As Long
Dim lNumIntTitPag As Long
Dim lNumNFPag As Long
Dim objMovEstoqueExcluir As ClassMovEstoque
Dim dQuantidade As Double
Dim sUnidadeMed As String
Dim sProduto As String
Dim iAlmoxarifado As Integer
Dim iItem As Integer
Dim lCodigo As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemNF As New ClassItemNF

On Error GoTo Erro_NFiscalEntrada_Exclui

    'Abre a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 35112

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 35113
    Next

    'Procura a Nota Fiscal com os dados passados
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, NumIntDocCPR FROM NFiscal WHERE FilialEmpresa = ? AND Fornecedor = ? AND FilialForn = ? AND DataEmissao = ? AND TipoNFiscal = ? AND Serie = ? AND NumNotaFiscal = ? AND Status <> ?", 0, lNumIntDoc, lNumIntDocCPR, objNFiscal.iFilialEmpresa, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 35114

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35115
    
    If lErro = AD_SQL_SEM_DADOS Then Error 35116 'Não encontrou --> Erro
    
    'Faz "lock" na NF
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 35117

    'exclui os lançamentos contábeis (se houverem) associados ao titulo em questao
    lErro = objContabil.Contabil_Exclui(objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO Then Error 36730
    
    'Procura NF a Pagar vínculada a Nota Fiscal de Entrada
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumNotaFiscal, NumIntTitPag FROM NFsPag WHERE NumIntDoc = ? AND FilialEmpresa = ? AND Status <> ?", 0, lNumNFPag, lNumIntTitPag, lNumIntDocCPR, objNFiscal.iFilialEmpresa, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 35118

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35119
    If lErro = AD_SQL_SEM_DADOS Then Error 35120 'Não encontrou --> Erro
    
    'Faz "lock" na NFPag
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 35121

    'Verifica se ela está vinculada
    If lNumIntTitPag <> 0 Then Error 35122

    'Faz a exclusão da NFPag (EXCLUSÃO LÓGICA)
    lErro = Comando_ExecutarPos(alComando(2), "UPDATE NFsPag SET Status = ? ", alComando(1), STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 35123

    Set objMovEstoqueExcluir = New ClassMovEstoque
    'Obtém o Código do Movimento
    lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
    If lErro <> SUCESSO Then Error 35138

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 44996
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then Error 44997

    objMovEstoqueExcluir.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque2
    objMovEstoqueExcluir.dtData = gdtDataAtual
    objMovEstoqueExcluir.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objMovEstoqueExcluir.lCodigo = lCodigo
    objMovEstoqueExcluir.lCliente = objNFiscal.lCliente
    objMovEstoqueExcluir.lFornecedor = objNFiscal.lFornecedor
    objMovEstoqueExcluir.sDocOrigem = objTipoDocInfo.sSigla

    sProduto = String(STRING_PRODUTO, 0)
    sUnidadeMed = String(STRING_UM_SIGLA, 0)

    'L6e os Ítens da Nota Fiscal
    lErro = Comando_ExecutarPos(alComando(3), "SELECT Item, NumIntDoc, Quantidade, Produto, UnidadeMed, Almoxarifado FROM ItensNFiscal WHERE NumIntNF = ? ORDER BY Item", 0, iItem, lNumIntDocItem, dQuantidade, sProduto, sUnidadeMed, iAlmoxarifado, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 35124

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35125

    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Faz o "lock "do ítem
        lErro = Comando_LockExclusive(alComando(3))
        If lErro <> AD_SQL_SUCESSO Then Error 35126
        
        Set objItemNF.colItensRomaneioGrade = New Collection
        objItemNF.lNumIntDoc = lNumIntDocItem
        
        lErro = CF("ItensNFGrade_Le", objItemNF)
        If lErro <> SUCESSO Then Error 20756

        If objItemNF.iPossuiGrade = DESMARCADO Then
        
            'Atribui o ítem ao Movimento
            objMovEstoqueExcluir.colItens.Add 0, 0, 0, 0, sProduto, "", sUnidadeMed, dQuantidade, iAlmoxarifado, "", lNumIntDocItem, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
        Else
        
            For Each objItemRomaneioGrade In objItemNF.colItensRomaneioGrade
                'Atribui o ítem ao Movimento
                objMovEstoqueExcluir.colItens.Add 0, 0, 0, 0, objItemRomaneioGrade.sProduto, "", sUnidadeMed, objItemRomaneioGrade.dQuantidade, objItemRomaneioGrade.iAlmoxarifado, "", objItemRomaneioGrade.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
            Next
        
        End If

        lErro = CF("ItensNFGrade_Exclui", lNumIntDocItem)
        If lErro <> SUCESSO Then Error 20756
        
        'faz a exclusão do ítem(EXCLUSÃO LÓGICA)
        lErro = Comando_ExecutarPos(alComando(4), "UPDATE ItensNFiscal SET Status = ?", alComando(3), STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then Error 35127
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35128

    Loop
    
    'Faz a Exclusão de MovEstoque associados
    lErro = CF("MovEstoque_Grava", objMovEstoqueExcluir)
    If lErro <> SUCESSO Then Error 35129

    'Exclui a Nota Fiscal (EXCLUSÃO LÓGICA)
    lErro = Comando_ExecutarPos(alComando(5), "UPDATE NFiscal SET Status = ?", alComando(0), STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 35130
    
    'Faz o Commit da Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 35131

    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalEntrada_Exclui = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Exclui:

    NFiscalEntrada_Exclui = Err

    Select Case Err

        Case 20756, 35129, 35138, 36730, 44996

        Case 35112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 35113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 35114, 35115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL1", Err, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 35116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFiscal.lNumNotaFiscal)

        Case 35117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", Err)

        Case 35118, 35119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG1", Err, lNumNFPag)

        Case 35120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VINCULO_NFENTRADA_NFPAGAR", Err, lNumIntDoc, lNumIntDocCPR)

        Case 35121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFSPAG", Err, lNumNFPag)

        Case 35122
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_VINCULADA", Err, lNumNFPag, lNumIntTitPag)

        Case 35123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFSPAG", Err, lNumNFPag)

        Case 35124, 35125, 35128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", Err)

        Case 35126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSNFISCAL", Err)

        Case 35127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSNFISCAL", Err, lNumIntDoc)

        Case 35130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NFISCAL", Err, lNumIntDoc)

        Case 35131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err, Error$)
            
        Case 44997
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149525)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalEntrada_Baixa_Recebimento(ByVal objNFiscal As ClassNFiscal) As Long
'Faz a baixa de um Recebimento

Dim alComando(0 To 4) As Long
Dim tNFiscal As typeNFiscal
Dim objMovEstoqueBaixa As ClassMovEstoque
Dim iItem As Integer
Dim iIndice As Integer
Dim lErro As Long
Dim tItemNF As typeItemNF
Dim lCodigo As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
Dim iTipoNFiscal As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objResevaItem As ClassReservaItem
Dim objItemNF As ClassItemNF
Dim objItemMovEstoqueINF As ClassItemMovEstoque
Dim colMovEstoqueINF As Collection

On Error GoTo Erro_NFiscalEntrada_Baixa_Recebimento
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEN Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEN Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEEBEN Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEBEN Then
    
        'Abre os Comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then Error 35066
        Next
    
        Select Case objNFiscal.iTipoNFiscal
        
            Case DOCINFO_NFIEN, DOCINFO_NFIEBF, DOCINFO_NFIEEBEN
                iTipoNFiscal = DOCINFO_NRFP
            Case DOCINFO_NFEEN, DOCINFO_NFEEBF, DOCINFO_NFEEBEN
                iTipoNFiscal = DOCINFO_NRFF
                
        End Select
            
        With tNFiscal
    
            .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
            .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
            .sPlaca = String(STRING_NFISCAL_PLACA, 0)
            .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
            .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
    
            'Procura no BD um Recebimento com os dados compatíveis aos da Nota Fiscal
            lErro = Comando_ExecutarPos(alComando(0), "SELECT DataEntrada, Status, Cliente, FilialCli, FilialEntrega, DataEmissao, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, NumIntDoc FROM NFiscal WHERE FilialEmpresa = ? AND NumRecebimento = ? AND (TipoNFiscal = ? OR TipoNFiscal = ?)", 0, .dtDataEntrada, .iStatus, .lCliente, _
            .iFilialCli, .iFilialEntrega, .dtDataEmissao, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.lNumRecebimento, DOCINFO_NRFF, DOCINFO_NRFP)
            If lErro <> AD_SQL_SUCESSO Then Error 35067
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35070
            
            'Se encontrar
            If lErro = AD_SQL_SUCESSO Then
            
                'Faz um "lock" no Recebimento
                lErro = Comando_LockExclusive(alComando(0))
                If lErro <> SUCESSO Then Error 35068
    
                objTipoDocInfo.iCodigo = iTipoNFiscal
            
                'Lê o Tipo da Nota Fiscal
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then Error 44998
                
                'Se não achou o Tipo de Documento --> erro
                If lErro = 31415 Then Error 44999
                
                'Se o TipoMovtoEstoque for positivo
                If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
                    Set objMovEstoqueBaixa = New ClassMovEstoque
        
                    'Obtém o Codigo Automático do Movimento de Estoque
                    lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
                    If lErro <> SUCESSO Then Error 35139
        
                    objMovEstoqueBaixa.dtData = objNFiscal.dtDataEntrada
                    objMovEstoqueBaixa.iFilialEmpresa = objNFiscal.iFilialEmpresa
                    objMovEstoqueBaixa.iTipoMov = MOV_EST_BAIXA_RECEBIMENTO_MATERIAL
                    objMovEstoqueBaixa.lCodigo = lCodigo
                    objMovEstoqueBaixa.lCliente = objNFiscal.lCliente
                    objMovEstoqueBaixa.lFornecedor = objNFiscal.lFornecedor
                    objMovEstoqueBaixa.sDocOrigem = objTipoDocInfo.sSigla & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
        
                    tItemNF.sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
                    tItemNF.sProduto = String(STRING_PRODUTO, 0)
                    tItemNF.sUnidadeMed = String(STRING_UM_SIGLA, 0)
        
                    'Pesquisa os Ítens do Recebimento
                    lErro = Comando_ExecutarPos(alComando(2), "SELECT Item,Status,Produto,UnidadeMed,Quantidade,Almoxarifado,PrecoUnitario,PercDesc,ValorDesconto,DataEntrega,DescricaoItem,ValorAbatComissao,NumIntPedVenda,NumIntItemPedVenda,NumIntDoc,NumIntTrib,NumIntDocOrig FROM ItensNFiscal WHERE NumIntNF = ? ORDER BY Item", 0, tItemNF.iItem, tItemNF.iStatus, tItemNF.sProduto, tItemNF.sUnidadeMed, tItemNF.dQuantidade, tItemNF.iAlmoxarifado, tItemNF.dPrecoUnitario, tItemNF.dPercDesc, tItemNF.dValorDesconto, tItemNF.dtDataEntrega, tItemNF.sDescricaoItem, tItemNF.dValorAbatComissao, tItemNF.lNumIntPedVenda, tItemNF.lNumIntItemPedVenda, tItemNF.lNumIntDoc, tItemNF.lNumIntTrib, tItemNF.lNumIntDocOrig, tNFiscal.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then Error 35071
        
                    lErro = Comando_BuscarPrimeiro(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35072
                    'Se encontrar
                    Do While lErro <> AD_SQL_SEM_DADOS
                        'Faz o "lock"
                        lErro = Comando_LockExclusive(alComando(2))
                        If lErro <> AD_SQL_SUCESSO Then Error 35073
        
                        Set objItemNF = New ClassItemNF
                        
                        objItemNF.lNumIntDoc = tItemNF.lNumIntDoc
        
                        lErro = CF("ItensNFGrade_Le", objItemNF)
                        If lErro <> SUCESSO Then Error 35139
        
                        If objItemNF.iPossuiGrade <> MARCADO Then
        
                            '#########################################
                            'Inserido por Wagner - 08/07/2008
                            Set colMovEstoqueINF = New Collection
                            Set objItemMovEstoqueINF = New ClassItemMovEstoque
                            
                            objItemMovEstoqueINF.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                            objItemMovEstoqueINF.lNumIntDocOrigem = tItemNF.lNumIntDoc
                            objItemMovEstoqueINF.iFilialEmpresa = objNFiscal.iFilialEmpresa
                            
                            lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoqueINF, colMovEstoqueINF)
                            If lErro <> SUCESSO Then Error 35139
                            
                            For Each objItemMovEstoqueINF In colMovEstoqueINF
                                Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (objItemMovEstoqueINF.dQuantidade * tItemNF.dPrecoUnitario) - tItemNF.dValorDesconto, 0, tItemNF.sProduto, tItemNF.sDescricaoItem, objItemMovEstoqueINF.sSiglaUM, objItemMovEstoqueINF.dQuantidade, objItemMovEstoqueINF.iAlmoxarifado, "", tItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                            Next
                            '#########################################
                            
'                            'Atribui os ítens ao Movimento
'                            With tItemNF
'
'                                Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (.dQuantidade * .dPrecoUnitario) - .dValorDesconto, 0, .sProduto, .sDescricaoItem, .sUnidadeMed, .dQuantidade, .iAlmoxarifado, "", .lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
'                                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'
'                            End With
        
                        Else
                            For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                                
                                Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (((objItemRomaneio.dQuantidade * tItemNF.dPrecoUnitario) - (tItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / tItemNF.dQuantidade))) * objItemRomaneio.dQuantidade), 0, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, tItemNF.sUnidadeMed, objItemRomaneio.dQuantidade, objItemRomaneio.iAlmoxarifado, "", objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            
                            Next
        
                        End If
       
                        lErro = Comando_ExecutarPos(alComando(4), "UPDATE ItensNFiscal SET Status = ?", alComando(2), STATUS_BAIXADO)
                        If lErro <> AD_SQL_SUCESSO Then Error 35076
        
                        lErro = Comando_BuscarProximo(alComando(2))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35055
        
                    Loop
                            
                    lErro = Comando_ExecutarPos(alComando(4), "UPDATE NFiscal SET Status = ? , Serie = ?, NumNotaFiscal = ?, DataAtualizacao=?", alComando(0), STATUS_BAIXADO, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, gdtDataHoje)
                    If lErro <> AD_SQL_SUCESSO Then Error 35077
                    
                    'Gera o Movimento de Baixa de Recebimento
                    lErro = CF("MovEstoque_Grava", objMovEstoqueBaixa)
                    If lErro <> SUCESSO Then Error 35075
    
                End If
            End If
        End With
    
        'Fecha os comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    NFiscalEntrada_Baixa_Recebimento = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Baixa_Recebimento:

    NFiscalEntrada_Baixa_Recebimento = Err

    Select Case Err

        Case 35066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 35067, 35070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL1", Err, iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 35068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", Err)

        Case 35069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCALBAIXADAS", Err, tNFiscal.lNumIntDoc)

        Case 35071, 35072, 35055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", Err)

        Case 35073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSNFISCAL", Err)

        Case 35074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCALBAIXADAS", tNFiscal.lNumIntDoc)

        Case 35075, 35139, 44998

        Case 35076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSNFISCAL", Err, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 35077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)

        Case 44999
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, iTipoNFiscal)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149526)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalEntNum_Automatico(ByVal sSerie As String, lNumAuto As Long) As Long
'Retorna o número do próximo NumNFiscal disponivel

Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim lNumNFiscal As Long
Dim lNumNFiscalUsado As Long
Dim iFim_de_Arquivo As Integer

On Error GoTo Erro_NFiscalEntNum_Automatico

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 39006

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 30577

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 30579

    'Pesquisa o número do NumNFiscal gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT ProxNumNFiscalEntrada FROM Serie WHERE Serie = ? AND FilialEmpresa = ?", 0, lNumNFiscal, sSerie, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 30582

    'le o registro com o número do NumNFiscal automático
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 30583

    'loca o registro que guarda o numero do NumNFiscal gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 30584

    'verifica se este número ainda não foi utilizado
    'Pesquisa as numerações de NumNFiscal maiores ou iguais a numeração gerada automaticamente
    lErro = Comando_Executar(lComando1, "SELECT NFiscal.NumNotaFiscal FROM NFiscal, TiposDocInfo WHERE NFiscal.NumNotaFiscal >= ? AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND TiposDocInfo.Tipo = ? ORDER BY NumNotaFiscal", lNumNFiscalUsado, lNumNFiscal, DOCINFO_NF_INT_ENTRADA)
    If lErro <> AD_SQL_SUCESSO Then Error 30585

    'le o primeiro codigo >= ao codigo automatico
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando1)

    'pesquisa o proximo número de codigo disponível
    Do While iFim_de_Arquivo <> AD_SQL_SEM_DADOS

        If lNumNFiscal < lNumNFiscalUsado Then Exit Do

        If lNumNFiscalUsado = lNumNFiscal Then
            lNumNFiscal = lNumNFiscal + 1
        End If

        iFim_de_Arquivo = Comando_BuscarProximo(lComando1)

    Loop

    'Valor de retorno da funcao
    lNumAuto = lNumNFiscal

    lNumNFiscal = lNumNFiscal + 1

    'atualiza o numero do NumNFiscal gerado automaticamente
    lErro = Comando_ExecutarPos(lComando1, "UPDATE Serie SET ProxNumNFiscalEntrada = ?", lComando, lNumNFiscal)
    If lErro <> AD_SQL_SUCESSO Then Error 30586

    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 39007
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)

    NFiscalEntNum_Automatico = SUCESSO

    Exit Function

Erro_NFiscalEntNum_Automatico:

    NFiscalEntNum_Automatico = Err

    Select Case Err
    
        Case 39006
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 39007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 30577, 30579
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 30582, 30583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE", Err)
    
        Case 30584
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERIE", Err)
    
        Case 30585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL4", Err)
    
        Case 30586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SERIE", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149527)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function NFiscalEntradaDevolucao_Grava_BD(ByVal objNFiscal As ClassNFiscal) As Long
'Faz a Inserção de uma Nota Fiscal de Entrada-Devolução no BD

Dim lErro As Long
Dim alComando(0 To 8) As Long
Dim lNumIntNF As Long
Dim lNumIntItemNF As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objMovEstoque As New ClassMovEstoque
Dim lCodigo As Long
Dim iTipoNFiscal As Integer
Dim iClasseDocCPR As Integer
Dim lNumIntDocCPR As Long
Dim objItemMovEstoque As ClassItemMovEstoque
Dim iTipoRecebimento As Integer
Dim objRastroItemNF As ClassRastroItemNF
Dim objRastroMovto As ClassRastreamentoMovto
Dim colInsumos As Collection
Dim colRastroMovto As Collection
Dim dFator As Double
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dQuantItemNFUMEst As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim objItemMovEst As ClassItemMovEstoque
Dim sDocOrigem As String
Dim iTipoMovtoEstoqueBenef As Integer
Dim objProduto As ClassProduto

On Error GoTo Erro_NFiscalEntradaDevolucao_Grava_BD

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 35416
    Next

    lErro = CF("CPR_Gera", objNFiscal, iClasseDocCPR, lNumIntDocCPR)
    If lErro <> SUCESSO Then gError 46543

    objNFiscal.iClasseDocCPR = iClasseDocCPR
    objNFiscal.lNumIntDocCPR = lNumIntDocCPR

    'Gerar o Número Interno da Nota Fiscal
    lErro = CF("NFiscal_Automatico", lNumIntNF)
    If lErro <> SUCESSO Then gError 35417

    objNFiscal.lNumIntDoc = lNumIntNF
    
    lErro = CF("NFiscal_Grava_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError 35417

    With objNFiscal

        .iNFe = SerieEletronica(.sSerie)
        
        'Insere no BD
        lErro = Comando_Executar(alComando(0), "INSERT INTO NFiscal (DataAtualizacao, NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,VolumeNumero,FreteRespons,DataReferencia, HoraEntrada, Moeda, TaxaMoeda,NumRecebimento, NFe, ValorDescontoTit, ValorItens, ValorDescontoItens, ChvNFe, HoraEmissao, DataCadastro, HoraCadastro, UsuCadastro) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            gdtDataHoje, lNumIntNF, STATUS_BAIXADO, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sVolumeNumero, .iFreteRespons, .dtDataReferencia, CDbl(.dtHoraEntrada), .iMoeda, .dTaxaMoeda, .lNumRecebimento, .iNFe, .dValorDescontoTit, .dValorItens, .dValorDescontoItens, .sChvNFe, _
            CDbl(IIf(.dtDataSaida <> DATA_NULA, IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraSaida, .dtHoraEmissao, .dtHoraSaida), IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraEntrada, .dtHoraEmissao, .dtHoraEntrada))), gdtDataHoje, CDbl(Time), gsUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 35418

    End With

    lErro = CF("NFiscal_GravaTributacao", objNFiscal, alComando(4), alComando(5))
    If lErro <> SUCESSO Then gError 35264

    'Gera o Número Automático
    lErro = CF("CRFATConfig_ObterNumInt1", NUM_PROX_ITEM_NOTA_FISCAL, lNumIntItemNF, objNFiscal.colItensNF.Count)
    If lErro <> SUCESSO Then gError 35419

    'Para cada Ítem da Nota Fiscal
    For Each objItemNF In objNFiscal.colItensNF

        With objItemNF

            .lNumIntNF = lNumIntNF
            .lNumIntDoc = lNumIntItemNF

            'Insere no BD
            lErro = Comando_Executar(alComando(3), "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, Almoxarifado, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, NumIntDocOrig,Ccl) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntNF, .iItem, STATUS_BAIXADO, .sProduto, .sUnidadeMed, .dQuantidade, .iAlmoxarifado, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, .lNumIntDoc, .lNumIntTrib, .lNumIntDocOrig, .sCcl)
            If lErro <> AD_SQL_SUCESSO Then gError 35420

        End With
        
        lErro = CF("ItemNFEnt_Devolucao_Customizado", objNFiscal, objItemNF)
        If lErro <> SUCESSO Then gError 199716

        lErro = CF("NFiscal_GravaTributacaoItem", objItemNF, alComando(6), alComando(7))
        If lErro <> SUCESSO Then gError 42075

        lErro = CF("ItensNFGrade_Inclui", objItemNF)
        If lErro <> SUCESSO Then gError 35419
        
        objItemNF.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
        objItemNF.objInfoAdicDocItem.lNumIntDocItem = objItemNF.lNumIntDoc
        lErro = CF("InfoAdicDocItem_Insere", objItemNF.objInfoAdicDocItem, alComando(8))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lNumIntItemNF = lNumIntItemNF + 1

    Next

    lErro = CF("NFe_Grava_Vinculo_Xml", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("NFiscalItens_Grava_Customizado", objNFiscal)
    If lErro <> SUCESSO Then gError 177049

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 35421

    'Se não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 35422

    sDocOrigem = objTipoDocInfo.sSigla
    iTipoMovtoEstoqueBenef = objTipoDocInfo.iTipoMovtoEstoqueBenef

    Call TipoNFiscal_Correspondente_Recebimento(objTipoDocInfo, iTipoRecebimento)

    objNFiscal.iTipoNFiscal = iTipoRecebimento

    'Faz a baixa do Recebimento
    'a baixa do recebimento deve ocorrer antes da gravacao da nota fiscal para que os valores dos saldos estejam corretos e permitam calcular o novo custo do produto, se for o caso.
    If objNFiscal.lNumRecebimento > 0 Then
        
        'Lê o Recebimento e Verifica os Dados
        lErro = NFiscal_Valida_Recebimento(objNFiscal)
        If lErro <> SUCESSO Then gError 64250
        
        lErro = CF("NFiscalEntDevRem_Baixa_Recebimento", objNFiscal)
        If lErro <> SUCESSO Then gError 35425
        
    End If

    objNFiscal.iTipoNFiscal = objTipoDocInfo.iCodigo

    'Se o TipoMovtoEstoque for positivo
    If objTipoDocInfo.iTipoMovtoEstoque > 0 Then

        Set objMovEstoque = New ClassMovEstoque

        'Gera um novo Código para o Movimento de Estoque
        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then gError 35423

        'Preenche o objMovEstoque
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque
        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = objTipoDocInfo.sSigla & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)

'mario distribuicao substituicao do codigo original acrescentando o tab de distribuicao
                        
        'Adiciona itens ao Movimento
        iIndice = 0
        For Each objItemNF In objNFiscal.colItensNF
            
            iIndice = iIndice + 1
            
            If objItemNF.colItensRomaneioGrade.Count > 0 Then
            
                For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                    Set objProduto = New ClassProduto
                
                    objProduto.sCodigo = objItemRomaneio.sProduto
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 89589
                    
                    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                        
                        dSomaAlocacao = 0
                                                                        
                        lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                        If lErro <> SUCESSO Then gError 89589
                                                                        
                        dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
                                                                        
                        For Each objReservaItem In objItemRomaneio.colLocalizacao
                                                
                            dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                            
                            Set colInsumos = New Collection
                            Set colRastroMovto = New Collection
                                                
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                            
                            objItemMovEstoque.iItemNF = objItemNF.iItem
                            
                            '###############################################################
                            'Inserido por Wagner
                            'Guarda o Rastreamento dos ItensNF
                            If Not (objNFiscal.objRastreamento Is Nothing) Then
                            
                                If objItemMovEstoque.colRastreamentoMovto Is Nothing Then
                                    Set objItemMovEstoque.colRastreamentoMovto = New Collection
                                End If
                            
                                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
                                If lErro <> SUCESSO Then gError 180065
                            End If
                            '###############################################################
                            
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                            objItemMovEstoque.iControleEstoque = objItemRomaneio.iControleEstoque
                            objItemMovEstoque.sSiglaUMEst = objItemRomaneio.sUMEstoque
                            
                        Next
                        
                        '###########################################################
                        'Alterado por Wagner 16/11/04
                        'Calcula o número de casas decimais do Formato de Estoque
'                        iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                        dAcrescimo = 0
'                        If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                        If Abs(dSomaAlocacao - dQuantItemNFUMEst) > dAcrescimo + QTDE_ESTOQUE_DELTA Then
                        
                        dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)

                        
                        If Abs(dSomaAlocacao - dQuantItemNFUMEst) > QTDE_ESTOQUE_DELTA Then
                        '###########################################################
                        
                            If dSomaAlocacao < dQuantItemNFUMEst Then gError 89590
                            
                            If dSomaAlocacao > dQuantItemNFUMEst Then gError 89591
                    
                        End If
                        
                    End If
                Next
            Else
                If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    
                    lErro = CF("MovEstoque_ItemNF_Customizado", objItemNF, objNFiscal)
                    If lErro <> SUCESSO Then gError 197727
                   
                        If Not objItemNF Is Nothing Then
                        
                        If objItemNF.ColAlocacoes.Count = 0 Then gError 89589
        
                        dSomaAlocacao = 0
                                                                        
                        lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                        If lErro <> SUCESSO Then gError 89589
                                                                        
                        dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
                                                                        
                        For Each objItemNFAloc In objItemNF.ColAlocacoes
                                                
                            dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
                            
                            Set colInsumos = New Collection
                            Set colRastroMovto = New Collection
                                                
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dValorTotal * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                            
                            objItemMovEstoque.iItemNF = objItemNF.iItem
                            
                            'Guarda o Rastreamento dos ItensNF
                            If Not (objNFiscal.objRastreamento Is Nothing) Then
                                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
                                If lErro <> SUCESSO Then gError 89588
                            End If
                            
        'mario
        '                    If Not (objItemNF.colRastreamento Is Nothing) Then
        '
        '                        'Guarda o Rastreamento dos ItensNF
        '                        For Each objRastroItemNF In objItemNF.colRastreamento
        '
        '                            If objRastroItemNF.iAlmoxCodigo = objItemNFAloc.iAlmoxarifado Then
        '
        '                                Set objRastroMovto = New ClassRastreamentoMovto
        '                                objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
        '                                objRastroMovto.sLote = objRastroItemNF.sLote
        '                                objRastroMovto.sProduto = objItemNF.sProduto
        '
        '                                'Adiciona objRastroMovto na coleção de Rastreamento
        '                                objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
        '
        '                            End If
        '
        '                        Next
        '
        '                    End If
        'fim mario
                            
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                            objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                            objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                            objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                            objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
                            
                        Next
                        
                        '####################################################
                        'Alterado por Wagner 16/11/04
                        'Calcula o número de casas decimais do Formato de Estoque
    '                    iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
    '
    '                    dAcrescimo = 0
    '                    If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
    '
    '                    If Abs(dSomaAlocacao - dQuantItemNFUMEst) > dAcrescimo + QTDE_ESTOQUE_DELTA Then
                        
                        dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)
    
                        If Abs(dSomaAlocacao - dQuantItemNFUMEst) > QTDE_ESTOQUE_DELTA Then
                        '####################################################
                        
                            If dSomaAlocacao < dQuantItemNFUMEst Then gError 89590
                            
                            If dSomaAlocacao > dQuantItemNFUMEst Then gError 89591
                    
                        End If
                    End If
                End If
            End If
        Next


'        For Each objItemNF In objNFiscal.colItensNF
'
'            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
'
'                Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, 0, 0, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNF.sUnidadeMed, objItemNF.dQuantidade, objItemNF.iAlmoxarifado, objItemNF.sAlmoxarifadoNomeRed, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
'                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'
'                Set objItemMovEstoque.colRastreamentoMovto = New Collection
'                Set objItemMovEstoque.colApropriacaoInsumo = New Collection
'
'                'Guarda o Rastreamento dos ItensNF
'                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
'                If lErro <> SUCESSO Then gError 83570
'
'
''                'Guarda o Rastreamento dos ItensNF
''                For Each objRastroItemNF In objItemNF.colRastreamento
''
''                    Set objRastroMovto = New ClassRastreamentoMovto
''                    objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
''                    objRastroMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
''                    objRastroMovto.sLote = objRastroItemNF.sLote
''                    objRastroMovto.lNumIntDocOrigem = objItemNF.lNumIntDoc
''                    objRastroMovto.sProduto = objItemNF.sProduto
''                    objRastroMovto.iFilialOP = objRastroItemNF.iLoteFilialOP
''
''                    'Adiciona objRastroMovto na coleção de Rastreamento
''                    objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
''
''                Next
'
'            End If
'        Next

        'Gera um Movimento de Estoque
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then gError 35424

    End If

    If (objNFiscal.iTipoNFiscal = DOCINFO_NFIEDSB) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEDSB) Then
    
        For Each objItemMovEstoque In objNFiscal.objMovEstoque.colItens
            'Descobre o movimento que retirou de beneficiamento para obter o custo (o custo do material consumido vai para o material beneficiado)
            For Each objItemMovEst In objMovEstoque.colItens
                If objItemMovEst.iItemNF = objItemMovEstoque.iItemNF Then
                    objItemMovEstoque.dCusto = Arredonda_Moeda((objItemMovEst.dCusto / objItemMovEst.dQuantidade) * objItemMovEstoque.dQuantInsumos)
                    Exit For
                End If
            Next
        Next
        
        'Gera um Movimento de Estoque
        lErro = CF("NFiscalEntrada_Grava_MovimentosEst", objNFiscal, iTipoMovtoEstoqueBenef, sDocOrigem)
        If lErro <> SUCESSO Then gError 126897
        
    End If

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalEntradaDevolucao_Grava_BD = SUCESSO

    Exit Function

Erro_NFiscalEntradaDevolucao_Grava_BD:

    NFiscalEntradaDevolucao_Grava_BD = gErr

    Select Case gErr

        Case 35416
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", 0)

        Case 35417, 35419, 35421, 35423, 35424, 35425, 35264, 42075, 64250, 83570, 89588, 89589, 46543, 126897, 177049, 180065, 199716

        Case 35418
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL1", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 35420
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 35422
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 89587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INEXISTENTE", gErr, iIndice)

        Case 89590
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INCOMPLETA", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case 89591
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case ERRO_SEM_MENSAGEM

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149528)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalEntradaRemessa_Grava_BD(ByVal objNFiscal As ClassNFiscal) As Long
'Faz a Inserção de uma Nota Fiscal de Entrada-Remessa no BD

Dim lErro As Long
Dim alComando(0 To 11) As Long
Dim lNumIntNF As Long
Dim lNumIntItemNF As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objMovEstoque As New ClassMovEstoque
Dim lCodigo As Long
Dim iTipoRecebimento As Integer
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objRastroItemNF As ClassRastroItemNF
Dim objRastroMovto As ClassRastreamentoMovto
Dim colItensNF As New colItensNF
Dim objItemNFBD As ClassItemNF
Dim colInsumos As Collection
Dim colRastroMovto As Collection
Dim dFator As Double
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dQuantItemNFUMEst As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim objProduto As New ClassProduto
Dim objProduto2 As ClassProduto, dCusto As Double

On Error GoTo Erro_NFiscalEntradaRemessa_Grava_BD

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 35625
    Next

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 35630

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 35631
    
    'se se tratar de uma nota de importação gravada com data nula ==> não grava novamente os dados que foram anteriormente registrados
    'senão processa normalmente
    If objNFiscal.iStatus <> STATUS_GRAVADA_IMPORTACAO_DATA_NULA Then

        'Gerar o Número Interno da Nota Fiscal
        lErro = CF("NFiscal_Automatico", lNumIntNF)
        If lErro <> SUCESSO Then gError 35626
    
        objNFiscal.lNumIntDoc = lNumIntNF
        
        lErro = CF("NFiscal_Grava_Msg", objNFiscal)
        If lErro <> SUCESSO Then gError 35626
    
        With objNFiscal

            .iNFe = SerieEletronica(.sSerie)
        
            lErro = Comando_Executar(alComando(0), "INSERT INTO NFiscal (DataAtualizacao, NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,VolumeNumero,FreteRespons,DataReferencia, HoraEntrada, Moeda, TaxaMoeda,NumRecebimento, NFe, ValorDescontoTit, ValorItens, ValorDescontoItens, ClienteBenef, FilialCliBenef, FornecedorBenef, FilialFornBenef, ChvNFe, HoraEmissao, DataCadastro, HoraCadastro, UsuCadastro) " & _
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                gdtDataHoje, lNumIntNF, STATUS_BAIXADO, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sVolumeNumero, .iFreteRespons, .dtDataReferencia, CDbl(.dtHoraEntrada), .iMoeda, .dTaxaMoeda, .lNumRecebimento, .iNFe, .dValorDescontoTit, .dValorItens, .dValorDescontoItens, .lClienteBenef, .iFilialCliBenef, .lFornecedorBenef, .iFilialFornBenef, .sChvNFe, _
                CDbl(IIf(.dtDataSaida <> DATA_NULA, IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraSaida, .dtHoraEmissao, .dtHoraSaida), IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraEntrada, .dtHoraEmissao, .dtHoraEntrada))), gdtDataHoje, CDbl(Time), gsUsuario)
            If lErro <> AD_SQL_SUCESSO Then gError 35627
    
        End With
    
        lErro = CF("NFiscal_GravaTributacao", objNFiscal, alComando(4), alComando(5))
        If lErro <> SUCESSO Then gError 35803
    
        'Gera o Número Automático
        lErro = CF("CRFATConfig_ObterNumInt1", NUM_PROX_ITEM_NOTA_FISCAL, lNumIntItemNF, objNFiscal.colItensNF.Count)
        If lErro <> SUCESSO Then gError 35628
    
        'Para cada Ítem da Nota Fiscal
        For Each objItemNF In objNFiscal.colItensNF
    
            objItemNF.lNumIntNF = lNumIntNF
            objItemNF.lNumIntDoc = lNumIntItemNF
    
            With objItemNF
    
                lErro = Comando_Executar(alComando(3), "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, Almoxarifado, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, NumIntDocOrig,Ccl) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntNF, .iItem, .iStatus, .sProduto, .sUnidadeMed, .dQuantidade, .iAlmoxarifado, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, lNumIntItemNF, .lNumIntTrib, .lNumIntDocOrig, .sCcl)
                If lErro <> AD_SQL_SUCESSO Then gError 35629
    
            End With
    
            If objNFiscal.iTipoNFiscal = DOCINFO_NFEE3BF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIE3BF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIERMB3 Or objNFiscal.iTipoNFiscal = DOCINFO_NFEERMB3 Then
            
                lErro = Comando_Executar(alComando(10), "INSERT INTO ItensNFRemBenef (NumIntItemNF, SaldoADevolver) VALUES (?,?)", objItemNF.lNumIntDoc, objItemNF.dQuantidade)
                If lErro <> AD_SQL_SUCESSO Then gError 35629
            
            End If
    
            lErro = CF("ItensNFGrade_Inclui", objItemNF)
            If lErro <> SUCESSO Then gError 30739

            lErro = CF("NFiscal_GravaTributacaoItem", objItemNF, alComando(6), alComando(7))
            If lErro <> SUCESSO Then gError 35804
            
            objItemNF.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
            objItemNF.objInfoAdicDocItem.lNumIntDocItem = objItemNF.lNumIntDoc
            lErro = CF("InfoAdicDocItem_Insere", objItemNF.objInfoAdicDocItem, alComando(11))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            lNumIntItemNF = lNumIntItemNF + 1
            
            If objNFiscal.iTipoNFiscal = DOCINFO_NFEEC Or objNFiscal.iTipoNFiscal = DOCINFO_NFEED Or _
            objNFiscal.iTipoNFiscal = DOCINFO_NFEECNT Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEOC Or _
            objNFiscal.iTipoNFiscal = DOCINFO_NFEEOF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEC Or _
            objNFiscal.iTipoNFiscal = DOCINFO_NFIECNT Or objNFiscal.iTipoNFiscal = DOCINFO_NFIED Or _
            objNFiscal.iTipoNFiscal = DOCINFO_NFIEOC Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEOF Or _
            objNFiscal.iTipoNFiscal = DOCINFO_NFIERMB3 Or objNFiscal.iTipoNFiscal = DOCINFO_NFEERMB3 Then
            
                lErro = Comando_Executar(alComando(9), "INSERT INTO ItemNFResto (NumIntItemNF, QuantResto) VALUES (?, ?)", _
                objItemNF.lNumIntDoc, objItemNF.dQuantidade)
                If lErro <> AD_SQL_SUCESSO Then gError 199072
                
            End If
    
        Next
    
        If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPCC Then
        
            lErro = NFiscal_Grava_NFImportacao(objNFiscal)
            If lErro <> SUCESSO Then gError 184704
            
        End If
        
        Call TipoNFiscal_Correspondente_Recebimento(objTipoDocInfo, iTipoRecebimento)
    
        objNFiscal.iTipoNFiscal = iTipoRecebimento
    
        'Faz a baixa do Recebimento
        'a baixa do recebimento deve ocorrer antes da gravacao da nota fiscal para que os valores dos saldos estejam corretos e permitam calcular o novo custo do produto, se for o caso.
        If objNFiscal.lNumRecebimento > 0 Then
            
            'Lê o Recebimento e Verifica os Dados
            lErro = NFiscal_Valida_Recebimento(objNFiscal)
            If lErro <> SUCESSO Then gError 64251
            
            lErro = CF("NFiscalEntDevRem_Baixa_Recebimento", objNFiscal)
            If lErro <> SUCESSO Then gError 35634
            
        End If

    Else
    
        If objNFiscal.dtDataEntrada <> DATA_NULA Then
            
            objNFiscal.iNFe = SerieEletronica(objNFiscal.sSerie)
            
            lErro = Comando_Executar(alComando(8), "UPDATE NFiscal SET DataEntrada = ?, DataAtualizacao=? WHERE NumIntDoc = ?", objNFiscal.dtDataEntrada, gdtDataHoje, objNFiscal.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 89150
            
            Set colItensNF = objNFiscal.colItensNF
            
            Set objNFiscal.colItensNF = New colItensNF
            
            lErro = CF("NFiscalItens_Le", objNFiscal)
            If lErro <> SUCESSO Then gError 89151
        
            If colItensNF.Count <> objNFiscal.colItensNF.Count Then gError 89154
                        
            'coloca o rastreamento da tela em objNFiscal
            For iIndice = 1 To objNFiscal.colItensNF.Count
                        
                Set objItemNF = colItensNF.Item(iIndice)
                Set objItemNFBD = objNFiscal.colItensNF.Item(iIndice)
                
                If objItemNF.sProduto <> objItemNFBD.sProduto Or objItemNF.dQuantidade <> objItemNFBD.dQuantidade Or objItemNF.dValorTotal <> objItemNF.dValorTotal Then gError 89155
                
                Set objItemNFBD.colRastreamento = objItemNF.colRastreamento
                Set objItemNFBD.ColAlocacoes = objItemNF.ColAlocacoes
                
                objProduto.sCodigo = objItemNF.sProduto
                'Faz o Lock do Produto
                lErro = CF("Produto_Lock", objProduto)
                If lErro <> SUCESSO Then gError 89151
                
                objItemNFBD.iControleEstoque = objProduto.iControleEstoque
                objItemNFBD.iClasseUM = objProduto.iClasseUM
                objItemNFBD.sUMEstoque = objProduto.sSiglaUMEstoque
                
            Next
    
            'Lê a parte de Tributação
            lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
            If lErro <> SUCESSO Then gError 89152

            'Lê a parte de tributacao dos Ítens da Nota Fiscal
            lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
            If lErro <> SUCESSO Then gError 89153
    
            '??? DOCINFO_NFIEIMP ler NFImportacao
            
        End If
    
    End If
    
    lErro = CF("NFe_Grava_Vinculo_Xml", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objNFiscal.iTipoNFiscal = objTipoDocInfo.iCodigo

    'se a nota for de importação e não estiver com a data de entrada preenchida ==> não processa a gravação das informações
    'caso contrário, processa
    If (objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMP And objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMPSE And objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMPPC) Or objNFiscal.dtDataEntrada <> DATA_NULA Then

        'Se o TipoMovtoEstoque for positivo
        If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
            Set objMovEstoque = New ClassMovEstoque
    
            'Gera um novo Código para o Movimento de Estoque
            lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
            If lErro <> SUCESSO Then gError 35632
    
            'Preenche o objMovEstoque
            objMovEstoque.lCodigo = lCodigo
            objMovEstoque.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque
            objMovEstoque.dtData = objNFiscal.dtDataEntrada
            objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
            objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
            objMovEstoque.lCliente = objNFiscal.lCliente
            objMovEstoque.lFornecedor = objNFiscal.lFornecedor
            objMovEstoque.sDocOrigem = objTipoDocInfo.sSigla & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
    
            'Se for para incluir o valor do frete e outras depesas no custo do ItemNF
            If gobjEST.iIncluiFreteNoCusto = INCLUI_FRETE_CUSTO Then
                Call Inclui_Frete_CustoItemNF(objNFiscal)
            End If
    
'mario distribuicao substituicao do codigo original acrescentando o tab de distribuicao
                        
            'Adiciona itens ao Movimento
            iIndice = 0
            For Each objItemNF In objNFiscal.colItensNF
                
                iIndice = iIndice + 1
                
                If objItemNF.colItensRomaneioGrade.Count > 0 Then
                
                    objItemNF.dCusto = 0
                    
                    'calcula o custo do item da NF em função do credito do ICMS e do IPI
                    lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
                    If lErro <> SUCESSO Then gError 60787
                    
                    Set objProduto2 = New ClassProduto
                
                    objProduto2.sCodigo = objItemNF.sProduto
                    
                    lErro = CF("Produto_Le", objProduto2)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 89464
                    
                    For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                        Set objProduto = New ClassProduto
                    
                        objProduto.sCodigo = objItemRomaneio.sProduto
                        
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 23080 Then gError 89464
                        
                        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                            
                            If objProduto2.iKitVendaComp = MARCADO Then
                                dFator = 1 'A UM do componente já está na UM de Estoque e a classe UM pode ser diferente da Classe UM do item da NF, logo a conversão feita para grade é desnecessária e pode dar erro
                            Else
                                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemRomaneio.sUMEstoque, dFator)
                                If lErro <> SUCESSO Then gError 89464
                            End If
            
                            dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
            
'                            'calcula o custo do item da NF em função do credito do ICMS e do IPI
'                            lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
'                            If lErro <> SUCESSO Then gError 60787
            
                            dSomaAlocacao = 0
                                                                            
                            For Each objReservaItem In objItemRomaneio.colLocalizacao
                                                    
                                dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                                
                                Set colInsumos = New Collection
                                Set colRastroMovto = New Collection
                                
                                If objProduto2.iKitVendaComp = MARCADO Then
                                    'dCusto = 0
                                    'Não existem parâmetros de rateio para entrada de kit de venda então o valor vai dividido igualmente entre os componentes
                                    dCusto = ((objItemNF.dCusto / objItemNF.colItensRomaneioGrade.Count) / objItemRomaneio.dQuantidade) * objReservaItem.dQuantidade
                                Else
                                    'O cálculo anterior dava um resultado errado e não considerava os impostos
                                    'dCusto = ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst
                                    dCusto = ((objItemNF.dCusto / objItemNF.dQuantidade) * (objReservaItem.dQuantidade / dFator))
                                End If
                                
                                'Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                                Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, dCusto, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                                
                                '################################################################
                                'Inserido por Wagner 02/06/2006
                                'Guarda o Rastreamento dos ItensNF
                                If Not (objNFiscal.objRastreamento Is Nothing) Then
                                    lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
                                    If lErro <> SUCESSO Then gError 180073
                                End If
                                '################################################################
                                 
                                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                                objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                                objItemMovEstoque.iControleEstoque = objItemRomaneio.iControleEstoque
                                objItemMovEstoque.sSiglaUMEst = objItemRomaneio.sUMEstoque
                                
                                If TipoDocInfo_ZeraQtdeMovto(objNFiscal.iTipoNFiscal) Then
                                
                                    objItemMovEstoque.dQuantidade = 0
                                    
                                End If
                                
                            Next
                            
                            '##########################################################
                            'Alterado por Wagner 16/11/04
                            'Calcula o número de casas decimais do Formato de Estoque
'                            iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                            dAcrescimo = 0
'                            If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                                'alteracao por tulio070303...
'                            If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > dAcrescimo + QTDE_ESTOQUE_DELTA Then

                            dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)

                            If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > QTDE_ESTOQUE_DELTA Then
                            '##########################################################
                            
                                If dSomaAlocacao < dQuantItemNFUMEst Then gError 89574
                                
                                If dSomaAlocacao > dQuantItemNFUMEst Then gError 89575
                        
                            End If
                            
                        End If
                    Next
                Else
                
                
                    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPCC Then
                                
                        'rateia o custo entre os itens da NF original
                        'lErro = NF_Trata_MovEst_IMPCC(objNFiscal, objItemNF, objMovEstoque)
                        lErro = NF_Trata_MovEst_DDAI(objNFiscal, objItemNF, objMovEstoque)
                        If lErro <> SUCESSO Then gError 89577
     
                    Else
                                
                        If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                            
                            If objItemNF.ColAlocacoes.Count = 0 And objItemNF.dQuantidade > 0 Then gError 89573
            
                            lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                            If lErro <> SUCESSO Then gError 89577
            
                            dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
            
                            'calcula o custo do item da NF em função do credito do ICMS e do IPI
                            If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC Then
                                lErro = CF("NFiscalEntradaImpParc_Calcula_Custo", objNFiscal, objItemNF)
                            Else
                                lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
                            End If
                            If lErro <> SUCESSO Then gError 60789
            
                            dSomaAlocacao = 0
                                                                            
                            For Each objItemNFAloc In objItemNF.ColAlocacoes
                                                    
                                dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
                                
                                Set colInsumos = New Collection
                                Set colRastroMovto = New Collection
                                                    
                                Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                                
                                'Guarda o Rastreamento dos ItensNF
                                If Not (objNFiscal.objRastreamento Is Nothing) Then
                                    lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
                                    If lErro <> SUCESSO Then gError 89576
                                End If
                                
                                
                                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                                objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                                objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                                objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                                objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
                                
                                'se for uma nota de complemento de preço ou outra que nao deve afetar qtde ==> zera a quantidade pois somente o custo será utilizado
            '                    If (objNFiscal.iTipoNFiscal = DOCINFO_NFIERCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEERCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEECCP) Then
                                If TipoDocInfo_ZeraQtdeMovto(objNFiscal.iTipoNFiscal) Then
                                   
                                    objItemMovEstoque.dQuantidade = 0
                                    
                                End If
                                
                            Next
                            
                            '###################################################################
                            'Alterado por Wagner 16/11/04
                            'Calcula o número de casas decimais do Formato de Estoque
    '                        iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
    '
    '                        dAcrescimo = 0
    '                        If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
    '
    '                        If Abs(dSomaAlocacao - dQuantItemNFUMEst) > dAcrescimo + QTDE_ESTOQUE_DELTA Then
                            
                            dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)
                            
                            If objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMPSE Then
                                If Abs(dSomaAlocacao - dQuantItemNFUMEst) > QTDE_ESTOQUE_DELTA Then
                                '###################################################################
                                
                                    If dSomaAlocacao < dQuantItemNFUMEst Then gError 89574
                                    
                                    If dSomaAlocacao > dQuantItemNFUMEst Then gError 89575
                            
                                End If
                            End If
                            If objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPPC Then
                                'Testa para ver se não recebeu mais que o total
                                '@@@@@@@@ FAZER
                            End If
                        End If
                    End If
                End If
           Next
                           
            'Gera um Movimento de Estoque
            lErro = CF("MovEstoque_Grava", objMovEstoque)
            If lErro <> SUCESSO Then gError 35633
    
            If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC And objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPCC Then
                'Gera a reserva nas importações
                lErro = CF("NFiscalEntCom_Gera_Reservas", objNFiscal, objMovEstoque)
                If lErro <> SUCESSO Then gError 35633
            End If
    
        End If

    End If

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalEntradaRemessa_Grava_BD = SUCESSO

    Exit Function

Erro_NFiscalEntradaRemessa_Grava_BD:

    NFiscalEntradaRemessa_Grava_BD = gErr

    Select Case gErr

        Case 35625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", 0)

        Case 30739, 35626, 35628, 35630, 35632, 35633, 35634, 35803, 35804, 60787, 60789, _
             64251, 83572, 89151, 89152, 89153, 89464, 89576, 89577, 180073, 184704, 199072

        Case 35627
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL1", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 35629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 35631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 89150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)

        Case 89154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMITENS_NFISCAL_DIFERE_NFISCALBD", gErr, colItensNF.Count, objNFiscal.colItensNF.Count)
        
        Case 89155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENS_NFISCAL_DIFERE_NFISCALBD", gErr)
            
        Case 89573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INEXISTENTE", gErr, iIndice)

        Case 89574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INCOMPLETA", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case 89575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149529)

    End Select

    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalFatEntrada_Grava_BD(ByVal objNFiscal As ClassNFiscal) As Long
'Insere uma Nova Nota Fiscal Fatura de Entrada no Banco de Dados

Dim alComando(0 To 7) As Long
Dim lNumIntDoc As Long
Dim lErro As Long
Dim iClasseDocCPR As Integer
Dim lNumIntDocCPR As Long
Dim objItemNF As ClassItemNF
Dim objMovEstoque As ClassMovEstoque
Dim objMovEstoqueBaixa As ClassMovEstoque
Dim iIndice As Integer
Dim iTipoMovtoEstoque As Integer
Dim lCodigo As Long
Dim sDocOrigem As String
Dim iTipoMovtoEstoqueBenef As Integer, objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objRastroItemNF As ClassRastroItemNF
Dim objRastroMovto As ClassRastreamentoMovto
Dim colInsumos As Collection
Dim colRastroMovto As Collection
Dim dFator As Double
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dQuantItemNFUMEst As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim iTipoMov As Integer, bMovEst As Boolean
Dim objProduto As ClassProduto
Dim objProduto2 As ClassProduto, dCusto As Double
Dim bTrataRateioMovEst As Boolean, bExisteCtrlEst As Boolean
Dim objItemMovEstoqueProdBenef As ClassItemMovEstoque
Dim objItemMovEstoqueAux As ClassItemMovEstoque

On Error GoTo Erro_NFiscalFatEntrada_Grava_BD

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 42709
    Next

    'Busca iClasseDocCPR e lNumIntDocCPR
    lErro = CF("CPR_Gera", objNFiscal, iClasseDocCPR, lNumIntDocCPR)
    If lErro <> SUCESSO Then Error 42710

    objNFiscal.iClasseDocCPR = iClasseDocCPR
    objNFiscal.lNumIntDocCPR = lNumIntDocCPR

    sDocOrigem = String(STRING_TIPO_DOC_INFO_SIGLA, 0)

    'Lê TipoMovtoEstoque em TiposDocInfo
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then Error 59260
    
    iTipoMovtoEstoque = objTipoDocInfo.iTipoMovtoEstoque
    sDocOrigem = objTipoDocInfo.sSigla
    iTipoMovtoEstoqueBenef = objTipoDocInfo.iTipoMovtoEstoqueBenef

    'Obtem o Número Interno da nova Nota Fiscal
    lErro = CF("NFiscal_Automatico", lNumIntDoc)
    If lErro <> SUCESSO Then Error 42714

    objNFiscal.lNumIntDoc = lNumIntDoc
    
    lErro = CF("NFiscal_Grava_Msg", objNFiscal)
    If lErro <> SUCESSO Then Error 42714

    With objNFiscal

        .iNFe = SerieEletronica(.sSerie)
        
        'Insere a nova Nota Fiscal de Entrada no BD
        lErro = Comando_Executar(alComando(0), "INSERT INTO NFiscal (DataAtualizacao, NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, FornecedorBenef, FilialFornBenef, VolumeNumero, FreteRespons,DataReferencia,Observacao, NumRecebimento, HoraEntrada, Moeda, TaxaMoeda, NFe, ValorDescontoTit, ValorItens, ValorDescontoItens, ChvNFe, HoraEmissao, DataCadastro, HoraCadastro, UsuCadastro,SerieNFPOrig,NumNFPOrig) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            gdtDataHoje, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lFornecedorBenef, .iFilialFornBenef, .sVolumeNumero, .iFreteRespons, .dtDataReferencia, .sObservacao, .lNumRecebimento, CDbl(.dtHoraEntrada), .iMoeda, .dTaxaMoeda, .iNFe, .dValorDescontoTit, .dValorItens, .dValorDescontoItens, .sChvNFe, _
            CDbl(IIf(.dtDataSaida <> DATA_NULA, IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraSaida, .dtHoraEmissao, .dtHoraSaida), IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraEntrada, .dtHoraEmissao, .dtHoraEntrada))), gdtDataHoje, CDbl(Time), gsUsuario, .sSerieNFPOrig, .lNumNFPOrig)
        If lErro <> AD_SQL_SUCESSO Then Error 42715
    End With

    lErro = CF("NFiscal_GravaTributacao", objNFiscal, alComando(4), alComando(5))
    If lErro <> SUCESSO Then Error 42716

    'Para cada Ítem da Nota
    For Each objItemNF In objNFiscal.colItensNF
    
        objItemNF.lNumIntNF = objNFiscal.lNumIntDoc

        'Obtém o Número Interno do Ítem
        lErro = CF("NFiscalItem_Automatico", lNumIntDoc)
        If lErro <> SUCESSO Then Error 42717

        objItemNF.lNumIntDoc = lNumIntDoc

        With objItemNF
            'Insere o Ítem no Banco de dados
            lErro = Comando_Executar(alComando(1), "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, NumIntDocOrig,Almoxarifado,Ccl) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lNumIntNF, .iItem, STATUS_LANCADO, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, lNumIntDoc, .lNumIntTrib, 0, .iAlmoxarifado, .sCcl)
            If lErro <> AD_SQL_SUCESSO Then Error 42718
        End With
    
        lErro = CF("ItensNFGrade_Inclui", objItemNF)
        If lErro <> SUCESSO Then Error 30739
    
        lErro = CF("NFiscal_GravaTributacaoItem", objItemNF, alComando(3), alComando(6))
        If lErro <> SUCESSO Then Error 42719
    
        objItemNF.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
        objItemNF.objInfoAdicDocItem.lNumIntDocItem = objItemNF.lNumIntDoc
        lErro = CF("InfoAdicDocItem_Insere", objItemNF.objInfoAdicDocItem, alComando(7))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objItemNF.objCobrItensContrato.lNumIntItensContrato <> 0 Then
            lErro = CF("ItensDeContrato_Insere_Fatura", objNFiscal, objItemNF)
            If lErro <> SUCESSO Then gError 136010
        End If
       
    Next
    
    lErro = CF("NFe_Grava_Vinculo_Xml", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("NFiscal_Grava_Customizado", objNFiscal)
    If lErro <> SUCESSO Then gError 177049

    lErro = CF("NFiscalItens_Grava_Customizado", objNFiscal)
    If lErro <> SUCESSO Then gError 177049
    
    'Faz a baixa de um recebimento
    'a baixa do recebimento deve ocorrer antes da gravacao da nota fiscal para que os valores dos saldos estejam corretos e permitam calcular o novo custo do produto, se for o caso.
    
    If objNFiscal.lNumRecebimento > 0 Then
        
        'Lê o Recebimento e Verifica os Dados
        lErro = NFiscal_Valida_Recebimento(objNFiscal)
        If lErro <> SUCESSO Then Error 64249
        
        lErro = CF("NFiscalFatEntrada_Baixa_Recebimento", objNFiscal)
        If lErro <> SUCESSO Then Error 42722
                
    End If

    If iTipoMovtoEstoque > 0 Then

        Set objMovEstoque = New ClassMovEstoque

        'Obtém o Código do Movimento
        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then Error 42720

        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.iTipoMov = iTipoMovtoEstoque
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = sDocOrigem & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)

        'Se for para incluir o valor do frete e outras depesas no custo do ItemNF
        If gobjEST.iIncluiFreteNoCusto = INCLUI_FRETE_CUSTO Then
            Call Inclui_Frete_CustoItemNF(objNFiscal)
        End If

        bTrataRateioMovEst = False
        If objNFiscal.lNumIntNotaOriginal <> 0 Then
            'Se é uma despesa aduaneira e tem NF original rateia de qualquer jeito
            If objNFiscal.iTipoNFiscal = DOCINFO_DDAI Or objNFiscal.iTipoNFiscal = DOCINFO_FDDAI Then
                bTrataRateioMovEst = True
            End If
            'Se for um conhecimento de transporte só vai ratear se todos itens não controlarem estoque
            If objNFiscal.iTipoNFiscal = DOCINFO_CFECT Or objNFiscal.iTipoNFiscal = DOCINFO_NFSET Or objNFiscal.iTipoNFiscal = DOCINFO_NFSEFT Then
                bExisteCtrlEst = False
                For Each objItemNF In objNFiscal.colItensNF
                    Set objProduto = New ClassProduto
                    objProduto.sCodigo = objItemNF.sProduto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 89464
                    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                        bExisteCtrlEst = True
                        Exit For
                    End If
                Next
                If Not bExisteCtrlEst Then
                    bTrataRateioMovEst = True
                    If objNFiscal.colItensNF.Count > 2 Then gError 211177 'Só pode ter 1 ou dois itens a serem rateados (com e sem crédito)
                    If objNFiscal.colItensNF.Count = 2 Then
                        If objNFiscal.colItensNF.Item(1).objTributacao.dICMSCredito = 0 And objNFiscal.colItensNF.Item(2).objTributacao.dICMSCredito = 0 Then gError 211178  'Não poder ter crédito nos dois
                        If objNFiscal.colItensNF.Item(1).objTributacao.dICMSCredito <> 0 And objNFiscal.colItensNF.Item(2).objTributacao.dICMSCredito <> 0 Then gError 211179  'Não poder não ter crédito nos dois
                    End If
                End If
            End If
        End If

'mario distribuicao substituicao do codigo original acrescentando o tab de distribuicao
        
        iIndice = 0
        'Adiciona itens ao Movimento
        For Each objItemNF In objNFiscal.colItensNF
            
            iIndice = iIndice + 1
            
            If objItemNF.colItensRomaneioGrade.Count > 0 Then
            
                objItemNF.dCusto = 0
                
                'calcula o custo do item da NF em função do credito do ICMS e do IPI
                lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
                If lErro <> SUCESSO Then gError 60787
                
                Set objProduto2 = New ClassProduto
            
                objProduto2.sCodigo = objItemNF.sProduto
                
                lErro = CF("Produto_Le", objProduto2)
                If lErro <> SUCESSO And lErro <> 23080 Then gError 89464
                
                For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                
                    Set objProduto = New ClassProduto
                
                    objProduto.sCodigo = objItemRomaneio.sProduto
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 89464
                    
                    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                        
                        If objProduto2.iKitVendaComp = MARCADO Then
                            dFator = 1 'A UM do componente já está na UM de Estoque e a classe UM pode ser diferente da Classe UM do item da NF, logo a conversão feita para grade é desnecessária e pode dar erro
                        Else
                            lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemRomaneio.sUMEstoque, dFator)
                            If lErro <> SUCESSO Then gError 89464
                        End If
        
                        dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
                        
'                        'calcula o custo do item da NF em função do credito do ICMS e do IPI
'                        lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
'                        If lErro <> SUCESSO Then gError 60787
                                
                        dSomaAlocacao = 0
                                                                        
                        For Each objReservaItem In objItemRomaneio.colLocalizacao
                                                
                            dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                            
                            Set colInsumos = New Collection
                            Set colRastroMovto = New Collection
                            
                            If objProduto2.iKitVendaComp = MARCADO Then
                                'dCusto = 0
                                'Não existem parâmetros de rateio para entrada de kit de venda então o valor vai dividido igualmente entre os componentes
                                dCusto = ((objItemNF.dCusto / objItemNF.colItensRomaneioGrade.Count) / objItemRomaneio.dQuantidade) * objReservaItem.dQuantidade
                            Else
                                'O cálculo anterior dava um resultado errado e não considerava os impostos
                                'dCusto = ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst
                                dCusto = ((objItemNF.dCusto / objItemNF.dQuantidade) * (objReservaItem.dQuantidade / dFator))
                            End If
                                                
                            'Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, dCusto, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)

                            '################################################################
                            'Inserido por Wagner 02/06/2006
                            'Guarda o Rastreamento dos ItensNF
                            If Not (objNFiscal.objRastreamento Is Nothing) Then
                                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
                                If lErro <> SUCESSO Then gError 180075
                            End If
                            '################################################################
                            
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                            objItemMovEstoque.iControleEstoque = objItemRomaneio.iControleEstoque
                            objItemMovEstoque.sSiglaUMEst = objItemRomaneio.sUMEstoque
''''                            objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
                            
                            If TipoDocInfo_ZeraQtdeMovto(objNFiscal.iTipoNFiscal) Then
                            
                                objItemMovEstoque.dQuantidade = 0
                                
                            End If
                            
                        Next
                        
                        '###################################################################
                        'Alterado por Wagner 16/11/04
                        'Calcula o número de casas decimais do Formato de Estoque
'                        iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                        dAcrescimo = 0
'                        If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                            'alteracao por tulio070303...
'                        If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > dAcrescimo + QTDE_ESTOQUE_DELTA Then
                        
                        dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)

                        
                        If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > QTDE_ESTOQUE_DELTA Then
                        '###################################################################
                        
                            If dSomaAlocacao < dQuantItemNFUMEst Then gError 89536
                            
                            If dSomaAlocacao > dQuantItemNFUMEst Then gError 89537
                        
                        End If
                        
                    End If
                Next
            Else
                        
                'Se for uma NF de despesa aduaneira com NF orginal (ou outro tipo que precise ratear o custo)
                'If (objNFiscal.iTipoNFiscal = DOCINFO_DDAI Or objNFiscal.iTipoNFiscal = DOCINFO_FDDAI Or (objNFiscal.iTipoNFiscal = DOCINFO_CFECT And objItemNF.ColAlocacoes.Count = 0)) And objNFiscal.lNumIntNotaOriginal <> 0 Then
                If bTrataRateioMovEst Then
                            
                    'rateia o custo entre os itens da NF original
                    lErro = NF_Trata_MovEst_DDAI(objNFiscal, objItemNF, objMovEstoque)
                    If lErro <> SUCESSO Then gError 89538
                    
                Else
                
                    lErro = CF("NF_Testa_Item_Mov_Est", objNFiscal, objItemNF, bMovEst, iTipoMov)
                    If lErro <> SUCESSO Then gError 197727
                
                    If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And bMovEst Then
                        
                        If objItemNF.ColAlocacoes.Count = 0 Then gError 89535
        
                        lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                        If lErro <> SUCESSO Then gError 89539
        
                        dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
        
                        'calcula o custo do item da NF em função do credito do ICMS e do IPI
                        lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
                        If lErro <> SUCESSO Then gError 60787
        
                        dSomaAlocacao = 0
                                                                        
                        For Each objItemNFAloc In objItemNF.ColAlocacoes
                                                
                            dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
                            
                            Set colInsumos = New Collection
                            Set colRastroMovto = New Collection
                                                
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                            
                            'Guarda o Rastreamento dos ItensNF
                            If Not (objNFiscal.objRastreamento Is Nothing) Then
                                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
                                If lErro <> SUCESSO Then gError 89538
                            End If
                                                    
                            If iTipoMov <> 0 Then objItemMovEstoque.iTipoMov = iTipoMov
                                                    
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                            objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                            objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                            objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                            objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
                            
                            'se for uma nota de complemento de preço ou outra que nao deve afetar qtde ==> zera a quantidade pois somente o custo será utilizado
        '                    If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEFS) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCS) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFICM) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFIPI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEFTEL) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFELUZ) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCS) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFICM) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFIPI) Or (objNFiscal.iTipoNFiscal = DOCINFO_FDDAI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFFEESAC) Then
                            If TipoDocInfo_ZeraQtdeMovto(objNFiscal.iTipoNFiscal) Then
        
                                objItemMovEstoque.dQuantidade = 0
                                
                            End If
                            
                        Next
                        
                        '#############################################################
                        'Alterado por Wagner 16/11/04
                        'Calcula o número de casas decimais do Formato de Estoque
    '                    iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
    '
    '                    dAcrescimo = 0
    '                    If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
    '
    '                    If Abs(dSomaAlocacao - dQuantItemNFUMEst) > dAcrescimo + QTDE_ESTOQUE_DELTA Then
                        
                        dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)
    
                        
                        If Abs(dSomaAlocacao - dQuantItemNFUMEst) > QTDE_ESTOQUE_DELTA Then
                        '#############################################################
                        
                            If dSomaAlocacao < dQuantItemNFUMEst Then gError 89536
                            
                            If dSomaAlocacao > dQuantItemNFUMEst Then gError 89537
                    
                        End If
                    End If
                    
                End If
            End If
        Next
                        
                        
'mario distribuicao fim da alteração para incluir o tab de distribuicao


'        'Adiciona os ítem ao Movimento
'        For Each objItemNF In objNFiscal.colItensNF
'
'            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
'
'                'calcula o custo do item da NF em função do credito do ICMS e do IPI
'                lErro = CF("NFiscalEntrada_Calcula_Custo",objItemNF)
'                If lErro <> SUCESSO Then Error 60788
'
'                Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNF.sUnidadeMed, objItemNF.dQuantidade, objItemNF.iAlmoxarifado, objItemNF.sAlmoxarifadoNomeRed, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
'                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'
'                Set objItemMovEstoque.colRastreamentoMovto = New Collection
'                Set objItemMovEstoque.colApropriacaoInsumo = New Collection
'
'                'Guarda o Rastreamento dos ItensNF
'                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
'                If lErro <> SUCESSO Then gError 83573
'
''                'Guarda o Rastreamento dos ItensNF
''                For Each objRastroItemNF In objItemNF.colRastreamento
''
''                    Set objRastroMovto = New ClassRastreamentoMovto
''                    objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
''                    objRastroMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
''                    objRastroMovto.sLote = objRastroItemNF.sLote
''                    objRastroMovto.lNumIntDocOrigem = objItemNF.lNumIntDoc
''                    objRastroMovto.sProduto = objItemNF.sProduto
''                    objRastroMovto.iFilialOP = objRastroItemNF.iLoteFilialOP
''
''                    'Adiciona objRastroMovto na coleção de Rastreamento
''                    objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
''
''                Next
'
'                objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
'                objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
'                objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
'                objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
'
'                'se for uma nota de complemento de preço ou outra que nao deve afetar qtde ==> zera a quantidade pois somente o custo será utilizado
'                If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCPCO) Then
'                    objItemMovEstoque.dQuantidade = 0
'                End If
'
'            End If
'
'        Next

        'Gera um Movimento de Estoque
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then Error 42721

    End If

    'se for entrada de material
    If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFBEN) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFBEN) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEFBFRS) Then
    
        If objNFiscal.iTipoNFiscal = DOCINFO_NFEFBFRS Then
            Set objNFiscal.objMovEstoque = New ClassMovEstoque
            
            For Each objItemMovEstoque In objMovEstoque.colItens
                If objItemMovEstoque.iTipoMov = 139 Then
                    Set objItemMovEstoqueProdBenef = objItemMovEstoque
                Else
                    Set objItemMovEstoqueAux = objNFiscal.objMovEstoque.colItens.Add(0, 0, objItemMovEstoque.dCusto, APROPR_CUSTO_INFORMADO, objItemMovEstoqueProdBenef.sProduto, objItemMovEstoqueProdBenef.sDescricao, objItemMovEstoqueProdBenef.sSiglaUMEst, 0, objItemMovEstoqueProdBenef.iAlmoxarifado, objItemMovEstoqueProdBenef.sAlmoxarifadoNomeRed, objItemMovEstoque.lNumIntDocOrigem, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                
                    objItemMovEstoqueAux.dQuantInsumos = objItemMovEstoque.dQuantidadeEst
                End If
            Next
            
        End If
        
        'Gera um Movimento de Estoque
        lErro = CF("NFiscalEntrada_Grava_MovimentosEst", objNFiscal, iTipoMovtoEstoqueBenef, sDocOrigem)
        If lErro <> SUCESSO Then Error 55453
        
    End If


    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalFatEntrada_Grava_BD = SUCESSO

    Exit Function

Erro_NFiscalFatEntrada_Grava_BD:

    NFiscalFatEntrada_Grava_BD = gErr

    Select Case gErr

        Case 30739, 42710, 42717, 42714, 42721, 42722, 42720, 42716, 42719, 55453, _
            59260, 60787, 60788, 64249, 83573, 89464, 89538, 89539, 136010, 180075, 177049

        Case 42709
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 42715
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 42718
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 42711, 42712
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO", gErr, objNFiscal.iTipoNFiscal)

        Case 42713
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 89535
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INEXISTENTE", gErr, iIndice)

        Case 89536
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INCOMPLETA", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case 89537
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case 211177
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_CFECT_QTDE_ITENS_INVALIDA", gErr)

        Case 211178, 211179
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_CFECT_ITENS_TRIB_INVALIDA", gErr)
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149530)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalEntrada_Grava_MovimentosEst(ByVal objNFiscal As ClassNFiscal, ByVal iTipoMovtoEstoque As Integer, ByVal sDocOrigem As String) As Long
'Gera um Movimento de Estoque

Dim lErro As Long
Dim objItemMovEstoque As ClassItemMovEstoque
Dim lCodigo As Long

On Error GoTo Erro_NFiscalEntrada_Grava_MovimentosEst

    'Obtém o Código do Movimento
    lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
    If lErro <> SUCESSO Then Error 52378
    
    'Preenche os dados do Movimento
    objNFiscal.objMovEstoque.dtData = objNFiscal.dtDataEntrada
    objNFiscal.objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
    objNFiscal.objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objNFiscal.objMovEstoque.iTipoMov = iTipoMovtoEstoque
    objNFiscal.objMovEstoque.lCodigo = lCodigo
    objNFiscal.objMovEstoque.sDocOrigem = sDocOrigem
    objNFiscal.objMovEstoque.lCliente = objNFiscal.lCliente
    objNFiscal.objMovEstoque.lFornecedor = objNFiscal.lFornecedor
    
    'Adiciona alguns dados a coleção de itens que ja foi preenchida
    For Each objItemMovEstoque In objNFiscal.objMovEstoque.colItens
        objItemMovEstoque.lNumIntDocOrigem = objNFiscal.lNumIntDoc
        objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_NFISCAL
    Next
        
    'Gera um Movimento de Estoque
    lErro = CF("MovEstoque_Grava", objNFiscal.objMovEstoque)
    If lErro <> SUCESSO Then Error 52379

    NFiscalEntrada_Grava_MovimentosEst = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Grava_MovimentosEst:

    NFiscalEntrada_Grava_MovimentosEst = Err

    Select Case Err

        Case 52378, 52379

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149531)

    End Select

    Exit Function

End Function


Function Recebimento_Automatico(ByVal objNFiscal As ClassNFiscal) As Long
'Preenche o Código do Recebimento de objNFiscal

Dim lCodigo As Long

    Recebimento_Automatico = CF("Config_ObterAutomatico_EmTrans", "EstConfig", "NUM_PROX_RECEBIMENTO", "NFiscal", "NumRecebimento", lCodigo)
    
    objNFiscal.lNumRecebimento = lCodigo
    
End Function

Public Function NFiscal_Valida_Recebimento(ByVal objNFiscal As ClassNFiscal) As Long
'Testa a Integridade dos Dados Básicos do Recebimento contra a NF
'Fornecedor ou Cliente, FilialCliente ou FilialFornecedor, NúmeroNF, Série, DataEntrada

Dim lErro As Long
Dim objRecebimento As New ClassNFiscal

On Error GoTo Erro_NFiscal_Valida_Recebimento
    
    objRecebimento.lNumRecebimento = objNFiscal.lNumRecebimento
    
    'Lê o Recebimento
    lErro = CF("NFiscal_Le_Recebimento", objRecebimento)
    If lErro <> SUCESSO And lErro <> 61050 Then Error 64237
    
    If lErro = 61050 Then Error 64238
    
    If objRecebimento.iStatus = STATUS_BAIXADO Then Error 64247
    
    If objNFiscal.lCliente > 0 Then
        If objRecebimento.lCliente <> objNFiscal.lCliente Then Error 64239
    End If
    
    If objNFiscal.lFornecedor > 0 Then
        If objRecebimento.lFornecedor <> objNFiscal.lFornecedor Then Error 64240
    End If
    
    If objNFiscal.iFilialCli > 0 Then
        If objRecebimento.iFilialCli <> objNFiscal.iFilialCli Then Error 64241
    End If
    
    If objNFiscal.iFilialForn > 0 Then
        If objRecebimento.iFilialForn <> objNFiscal.iFilialForn Then Error 64242
    End If
    
    If objNFiscal.dtDataEntrada <> objRecebimento.dtDataEntrada Then Error 64243
    
    If Len(Trim(objRecebimento.sSerie)) > 0 Then
        If objNFiscal.sSerie <> objRecebimento.sSerie Then Error 64244
    End If
    
    If objRecebimento.lNumNotaFiscal > 0 Then
        If objNFiscal.lNumNotaFiscal <> objRecebimento.lNumNotaFiscal Then Error 64245
    End If

    NFiscal_Valida_Recebimento = SUCESSO
    
    Exit Function
    
Erro_NFiscal_Valida_Recebimento:
    
    NFiscal_Valida_Recebimento = Err
    
    Select Case Err
        
        Case 64237
        
        Case 64238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEB_NAO_CADASTRADO", Err, objNFiscal.lNumRecebimento)
        
        Case 64239
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_RECEB_DIFERENTE_NF", Err, objRecebimento.lCliente)
        
        Case 64240
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORN_RECEB_DIFERENTE_NF", Err, objRecebimento.lFornecedor)
        
        Case 64241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILCLIENTE_RECEB_DIFERENTE_NF", Err, objRecebimento.iFilialCli)
        
        Case 64242
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILFORN_RECEB_DIFERENTE_NF", Err, objRecebimento.iFilialForn)
        
        Case 64243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_RECEB_DIFERENTE_NF", Err, objRecebimento.dtDataEntrada)
        
        Case 64244
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_RECEB_DIFERENTE_NF", Err, objRecebimento.sSerie)
        
        Case 64245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMNF_RECEB_DIFERENTE_NF", Err, objRecebimento.lNumNotaFiscal)
        
        Case 64247
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_VINCULADO_NF", Err, objRecebimento.lNumRecebimento, objRecebimento.lNumNotaFiscal)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149532)
            
    End Select
    
    Exit Function
        
End Function

Private Function FilialFornecedor_Grava_EstatisticasCom(ByVal objNFiscal As ClassNFiscal, ByVal colPedCompras As Collection) As Long
'Tem que ser chamado dentro de Transação
'Atualiza a gravação da parte de estatística do Fornecedor

Dim lErro As Long
Dim objFilFornFilEmp As New ClassFilFornFilEmp
Dim objPedidoCompra As ClassPedidoCompras

On Error GoTo Erro_FilialFornecedor_Grava_EstatisticasCom

    For Each objPedidoCompra In colPedCompras
        
        objFilFornFilEmp.lAtrasoAcumulado = objFilFornFilEmp.lAtrasoAcumulado + objPedidoCompra.lAtrasoAcumulado
        objFilFornFilEmp.lItensPedidosRecebidos = objFilFornFilEmp.lItensPedidosRecebidos + objPedidoCompra.lItensPCRecebidos
        
        'se o pedido de compra foi baixado
        If objPedidoCompra.dtDataBaixa <> 0 Then
            objFilFornFilEmp.lPedidosEmAberto = objFilFornFilEmp.lPedidosEmAberto - 1
            objFilFornFilEmp.lPedidosRecebidos = objFilFornFilEmp.lPedidosRecebidos + 1
        End If
    
    Next
                
    'Grava a Estatística da Filial Fornecedor
    lErro = CF("FilialFornecedor_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_FILIALFORNFILEMP, objFilFornFilEmp)
    If lErro <> SUCESSO Then gError 89417
    
    FilialFornecedor_Grava_EstatisticasCom = SUCESSO

    Exit Function

Erro_FilialFornecedor_Grava_EstatisticasCom:

    FilialFornecedor_Grava_EstatisticasCom = gErr

    Select Case gErr

        Case 89417

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149533)

    End Select

    Exit Function

End Function

'Funcao de Estatística do Fornecedor -- Compras
Function PedCompraAbertas_CalculaAtraso(ByVal objNFiscal As ClassNFiscal, lMaiorAtraso As Long) As Long
'Calcula Estatística do Fornecedor com relacao aos Pedidos de Compra (MaiorAtraso)

Dim lErro As Long
Dim dtDataLimite As Date
Dim lComando As Long

On Error GoTo Erro_PedCompraAbertas_CalculaAtraso

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58027

    dtDataLimite = DATA_NULA

    'Lê a menor data limite dos ItensPedCompra Abertos
    lErro = Comando_Executar(lComando, "SELECT MIN(ItensPedCompra.DataLimite) FROM PedidoCompra, ItensPedCompra WHERE PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND PedidoCompra.TipoDestino = ? AND PedidoCompra.FilialDestino = ? AND PedidoCompra.Fornecedor = ? AND PedidoCompra.Filial = ? AND ItensPedCompra.DataLimite < ? AND (ItensPedCompra.Quantidade - ItensPedCompra.QuantRecebida - ItensPedCompra.QuantRecebimento) > 0", dtDataLimite, TIPO_DESTINO_EMPRESA, objNFiscal.iFilialEmpresa, objNFiscal.lFornecedor, objNFiscal.iFilialForn, gdtDataAtual)
    If lErro <> SUCESSO Then Error 58028

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58029

    'Caso a DataLimite seja "Data_Nula" então lMaiorAtraso = 0
    If dtDataLimite <> DATA_NULA Then
        'O maior atraso será a DataAtual menos a Menor DataLimite dos Itens Abertos
        lMaiorAtraso = CLng(gdtDataAtual - dtDataLimite)
    Else
        lMaiorAtraso = 0
    End If

    Call Comando_Fechar(lComando)

    PedCompraAbertas_CalculaAtraso = SUCESSO

    Exit Function

Erro_PedCompraAbertas_CalculaAtraso:

    PedCompraAbertas_CalculaAtraso = Err

    Select Case Err

    Case 58027
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

    Case 58028, 58029
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", Err)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149534)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NotaFiscalEntrada_Cancelar(ByVal objNFiscal As ClassNFiscal, ByVal dtDataCancelamento As Date) As Long
'Faz o cancelamento de uma nota fiscal de Saida

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_NotaFiscalEntrada_Cancelar
    
    'ABre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197962

    'Grava o Livro Fiscal a partir da Nota Fiscal
    lErro = CF("NotaFiscalEntrada_Cancelar_EmTrans", objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 197963
    
    If objNFiscal.iRollBack = 1 Then
        Call Transacao_Rollback
    Else
        'Confirma a transacao
        lErro = Transacao_Commit()
        If lErro <> SUCESSO Then gError 197964
    End If
    
    NotaFiscalEntrada_Cancelar = SUCESSO

    Exit Function

Erro_NotaFiscalEntrada_Cancelar:

    NotaFiscalEntrada_Cancelar = gErr

    Select Case gErr

        Case 197962
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197963
        
        Case 197964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197965)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NotaFiscalEntrada_Cancelar_EmTrans(ByVal objNFiscal As ClassNFiscal, ByVal dtDataCancelamento As Date) As Long
'Faz o cancelamento de uma nota fiscal de Entrada

Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 12) As Long
Dim iClasseDocCPR As Integer, lNumIntDocCPR As Long
Dim lNumMovtoEstoque As Long, lErro2 As Long
Dim objMovEstoque As New ClassMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lCodigo As Long, lNumNotaFiscal As Long, iTipoNF As Integer, iStatus As Integer
Dim lNumRecebimento As Long
Dim lNumIntItemNF As Long
Dim objFilFornFilEmp As New ClassFilFornFilEmp
Dim tItemNF As typeItemNF
Dim objProduto As New ClassProduto
Dim objItemNF As ClassItemNF
Dim lNumIntOrigem As Long
Dim iTipoOrigem As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim iIndiceAux As Integer, lAux As Long

On Error GoTo Erro_NotaFiscalEntrada_Cancelar_EmTrans
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67461
    Next
    
    'Lê a nota fiscal
    lErro = Comando_ExecutarPos(alComando(0), "SELECT  ClasseDocCPR, NumIntDocCPR, NumRecebimento FROM NFiscal WHERE NumIntDoc = ?", 0, iClasseDocCPR, lNumIntDocCPR, lNumRecebimento, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 67463

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67464
    If lErro <> AD_SQL_SUCESSO Then gError 67467 'Não encontrou

    'faz o lock na Nota Fiscal
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 67468
    
    'Se for uma NF Eletrônica só deixa cancelar se estiver autorizada
    If ISSerieEletronica(objNFiscal.sSerie) And objNFiscal.iFilialEmpresa < 50 Then
        'Lê a nota fiscal
        lErro = Comando_Executar(alComando(9), "SELECT NumIntNF FROM NFeFedProtNFe WHERE NumIntNF = ? AND (cStat = '100' Or cStat = '205' Or cStat = '110' Or cStat = '301' Or cStat = '302')", lAux, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 206867
    
        lErro = Comando_BuscarPrimeiro(alComando(9))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206868
        If lErro <> AD_SQL_SUCESSO Then gError 206869 'Não encontrou
    End If

    'Lê o tipodocinfo da nota fiscal
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 67470
    
    If lErro <> SUCESSO Then gError 67479
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC And objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPCC Then
    
        lErro = NFiscal_Exclui_NFImportacao(objNFiscal)
        If lErro <> SUCESSO Then gError 67470
    
    End If
    
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 67470
    
    'Se é uma nota que possui vinculação com outra que a sucede. Ex. Uma nota que tem uma devolução vinculada ou uma nota que tem um complemento vinculado.
    'If Len(Trim(objTipoDocInfo.sSiglaNFOriginal)) > 0 Or objTipoDocInfo.iComplementar <> DOCINFO_COMPLEMENTO Then
    
    'verifica se a sua devolução já foi feita
    lErro = Comando_Executar(alComando(1), "SELECT NumNotaFiscal, TipoNFiscal FROM NFiscal WHERE NumIntNotaOriginal = ? AND Status <> ? ", lNumNotaFiscal, iTipoNF, objNFiscal.lNumIntDoc, STATUS_CANCELADO)
    If lErro <> AD_SQL_SUCESSO Then gError 67465

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67466
    If lErro = AD_SQL_SUCESSO Then gError 67489 'Se já foi, erro.
        
    lErro = Comando_Executar(alComando(12), "SELECT N.NumNotaFiscal, N.TipoNFiscal FROM NFiscal AS N, ItensNFiscal AS I WHERE N.NumIntDoc = I.NumIntNF AND I.NumIntDocOrig IN (SELECT NumIntDoc FROM ItensNFiscal WHERE NumIntNF = ?) AND N.Status <> 7", lNumNotaFiscal, iTipoNF, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 67465

    lErro = Comando_BuscarPrimeiro(alComando(12))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67466
    
    If lErro = AD_SQL_SUCESSO Then gError 67489 'Se já foi, erro.
    
    'End If

    'Atualiza o status da nota fiscal
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE NFiscal SET Status = ?, CodUsuarioCancel=?, MotivoCancel=?, DataRegCancel = ?, HoraRegCancel = ?, DataCancel = ?", alComando(0), STATUS_CANCELADO, gsUsuario, objNFiscal.sMotivoCancel, Date, CDbl(Time), dtDataCancelamento)
    If lErro <> AD_SQL_SUCESSO Then gError 67469
    
    'se a nota nao é uma remessa de importacao ou a data de entrada estiver preenchida
    If (objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMP And objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMPSE And objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMPPC) Or objNFiscal.dtDataEntrada <> DATA_NULA Then
    
        If lNumRecebimento > 0 Then
            
            objNFiscal.lNumRecebimento = lNumRecebimento
            
            lErro = CF("Recebimento_Desfaz_Baixa", objNFiscal)
            If lErro <> SUCESSO Then gError 67477
            
        End If
        
        'Se a nota gerou documentos em  CPR
        If iClasseDocCPR > 0 And lNumIntDocCPR > 0 Then
            
            'Desfaz os documentos associados
            lErro = CF("CPR_Exclui", iClasseDocCPR, lNumIntDocCPR)
            If lErro <> SUCESSO Then gError 67471
            
        End If
    
'        'Busca os itens da nota associada
'        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM ItensNFiscal WHERE NumIntNF = ? ", 0, tItemNF.lNumIntDoc, objNFiscal.lNumIntDoc)
'        If lErro <> AD_SQL_SUCESSO Then gError 67480
    
        tItemNF.sUnidadeMed = String(STRING_UM_SIGLA, 0)
        tItemNF.sProduto = String(STRING_PRODUTO, 0)
    
        'Busca os itens da nota associada
        lErro2 = Comando_ExecutarPos(alComando(2), "SELECT Produto, NumIntDoc, Quantidade, UnidadeMed, NumIntDocOrig FROM ItensNFiscal WHERE NumIntNF = ? ", 0, tItemNF.sProduto, tItemNF.lNumIntDoc, tItemNF.dQuantidade, tItemNF.sUnidadeMed, tItemNF.lNumIntDocOrig, objNFiscal.lNumIntDoc)
        If lErro2 <> AD_SQL_SUCESSO Then gError 67480
    
        lErro2 = Comando_BuscarPrimeiro(alComando(2))
        If lErro2 <> AD_SQL_SUCESSO And lErro2 <> AD_SQL_SEM_DADOS Then gError 67481
        
        If lErro2 <> AD_SQL_SUCESSO Then gError 89005
        
        'Guarda o número do Item
        lNumIntItemNF = tItemNF.lNumIntDoc
                
        iTipoOrigem = 0
        
        For iIndiceAux = 1 To objNFiscal.colItensNF.Count
        
            If objNFiscal.colItensNF(iIndiceAux).iPossuiGrade = DESMARCADO Then
                
                If objNFiscal.colItensNF(iIndiceAux).ColAlocacoes.Count > 0 Then
                    lNumIntOrigem = lNumIntItemNF
                    iTipoOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    Exit For
                End If
            Else
            
                For Each objItemRomaneio In objNFiscal.colItensNF(iIndiceAux).colItensRomaneioGrade
                    If objItemRomaneio.colLocalizacao.Count > 0 Then
                        iTipoOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                        lNumIntOrigem = objItemRomaneio.lNumIntDoc
                        Exit For
                    End If
                Next
                If iTipoOrigem > 0 Then Exit For
            End If
        Next
        
        If iTipoOrigem <> 0 Then
        
            'Verifica se há movimento de estoque ligado ao item
            lErro = Comando_ExecutarPos(alComando(4), "SELECT Codigo FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? AND FilialEmpresa = ?", 0, lNumMovtoEstoque, iTipoOrigem, lNumIntOrigem, objNFiscal.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 67484
        
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67485
            
            If lErro = AD_SQL_SUCESSO Then
        
                'Lê o movto de estoque da NOta fiscal
                objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
                objMovEstoque.lCodigo = lNumMovtoEstoque
                
                lErro = CF("MovEstoque_Le", objMovEstoque)
                If lErro <> SUCESSO And lErro <> 30128 Then gError 67472
                
                If lErro <> SUCESSO Then gError 67486
                
                'Lê os movimento dos itens de nota fiscal
                lErro = CF("MovEstoqueItens_Le_Desc", objMovEstoque)
                If lErro <> SUCESSO Then gError 67474
            
                objMovEstoque.iTipoMov = 0
                
                For iIndice = 1 To objMovEstoque.colItens.Count
                    objMovEstoque.colItens(iIndice).iTipoMov = objTipoDocInfo.iTipoMovtoEstoque2
                Next
                
                'Grava o movimento de estoque
                lErro = CF("MovEstoque_Grava", objMovEstoque)
                If lErro <> SUCESSO Then gError 67475
            
            End If
        
        End If
        
        'Enquanto encontrar ItensNF
        Do While lErro2 <> AD_SQL_SEM_DADOS
            
            For Each objItemNF In objNFiscal.colItensNF
                If objItemNF.lNumIntDoc = tItemNF.lNumIntDoc Then Exit For
            Next

            objProduto.sCodigo = objItemNF.sProduto

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 89424

            If lErro <> SUCESSO Then gError 89425

            objItemNF.iClasseUM = objProduto.iClasseUM
            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
            objItemNF.sUMVenda = objProduto.sSiglaUMVenda

            objItemNF.dQuantidade = tItemNF.dQuantidade
            objItemNF.sUnidadeMed = tItemNF.sUnidadeMed
            objItemNF.lNumIntDocOrig = tItemNF.lNumIntDocOrig
            objItemNF.sProduto = tItemNF.sProduto
            
            'Cancela o item
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensNFiscal SET Status = ?", alComando(2), STATUS_CANCELADO)
            If lErro <> AD_SQL_SUCESSO Then gError 67483
        
            'Desfaz baixas de Pedidos de Compras
            lErro = CF("PedidoCompra_Desfaz_Baixa", objNFiscal, lNumIntItemNF, objFilFornFilEmp)
            If lErro <> SUCESSO Then gError 67502
            
            lErro = CF("RequisicaoCompra_Desfaz_Baixa", objNFiscal.lNumIntDoc, lNumIntItemNF)
            If lErro <> SUCESSO Then gError 67503
            
            lErro = Comando_Executar(alComando(8), "DELETE FROM ItemNFResto WHERE NumIntItemNF = ?", _
            tItemNF.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199073
            
            If objNFiscal.iTipoNFiscal = DOCINFO_NFEE3BF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIE3BF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIERMB3 Or objNFiscal.iTipoNFiscal = DOCINFO_NFEERMB3 Then
            
                'Verifica se houve devoluções para essa remessa
                lErro = Comando_Executar(alComando(11), "SELECT NumIntItemNFRem FROM ItensNFDevBenefRem WHERE NumIntItemNFRem = ? ", lAux, tItemNF.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 211102
            
                lErro = Comando_BuscarPrimeiro(alComando(11))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211103
                
                If lErro = AD_SQL_SUCESSO Then gError 211104
                
                lErro = Comando_Executar(alComando(10), "DELETE FROM ItensNFRemBenef WHERE NumIntItemNF = ?", tItemNF.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 211105
                
            End If
            
            lErro = CF("ItemNFEnt_Cancelar_Customizado", objNFiscal, objItemNF)
            If lErro <> SUCESSO Then gError 199732
            
            lErro2 = Comando_BuscarProximo(alComando(2))
            If lErro2 <> SUCESSO And lErro2 <> AD_SQL_SEM_DADOS Then gError 67482
        
            'Guarda o número do Item
            lNumIntItemNF = tItemNF.lNumIntDoc
        
        Loop
            
        'Grava a Estatística da Filial Fornecedor
        lErro = CF("FilialFornecedor_Grava_Estatisticas", objNFiscal, CANCELAMENTO_FILIALFORNFILEMP, objFilFornFilEmp)
        If lErro <> SUCESSO Then gError 67476
            
        'Grava a Estatística do Produto com Relação a devolução
        lErro = CF("ProdutoVendido_Grava_Estatisticas", objNFiscal, CANCELAMENTO_DOC)
        If lErro <> SUCESSO Then gError 89775
            
        'Verifica se o modulo de Livros Fiscais está Ativo
        If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                            
            'Grava o Livro Fiscal a partir da Nota Fiscal
            lErro = CF("NotaFiscal_Cancela_Fis", objNFiscal)
            If lErro <> SUCESSO Then gError 81527
        
        End If
        
        'Gera um lançamento extorno para a Nota Fiscal ou se ainda não foi contabilizado Exclui
        lErro = CF("NFiscalEntradaCancela_Contabilidade", objNFiscal, objTipoDocInfo, dtDataCancelamento)
        If lErro <> SUCESSO Then gError 67478
        
    Else
    
        gError 81869
    
    End If
    
    lErro = NFiscal_Exclui_Estorno_Cred(objNFiscal)
    If lErro <> SUCESSO Then gError 67470
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NotaFiscalEntrada_Cancelar_EmTrans = SUCESSO

    Exit Function

Erro_NotaFiscalEntrada_Cancelar_EmTrans:

    NotaFiscalEntrada_Cancelar_EmTrans = gErr

    Select Case gErr

        Case 67461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67463, 67464, 67465, 67466, 206867, 206868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 67467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INTERNA_ENTRADA_NAO_CADASTRADA", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 67468
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", gErr)

        Case 67469
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 67470, 67471, 67472, 67473, 67474, 67475, 67476, 67477, 67478, 67494, 67496, 67502, 67503, 81527, 89424, 89775, 199732

        Case 67479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case 67480, 67481, 67482
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL1", gErr)

        Case 67483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITEMNF", gErr)

        Case 67484, 67485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 67486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO2", gErr, objMovEstoque.iFilialEmpresa, objMovEstoque.lCodigo)

        Case 67488
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_VINCULADA_CANCELAR", gErr, lNumNotaFiscal, iTipoNF)

        Case 67489
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_JA_DEVOLVIDA", gErr)
                    
        Case 81869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFIEIMP_DT_ENT_CANC", Err)
        
        Case 89005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_SEM_ITENS", gErr, objNFiscal.lNumNotaFiscal)
                    
        Case 89425
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
                    
        Case 199073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMNFRESTO", gErr)
                    
        Case 206869
            Call Rotina_Erro(vbOKOnly, "ERRO_CANCNFE_NAO_AUTORIZADA", gErr)
            
        Case 211102, 211103
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFREMBENEF", gErr)
                    
        Case 211104
            Call Rotina_Erro(vbOKOnly, "ERRO_NFREMBENEF_JA_UTILIZADA", gErr)
                    
        Case 211105
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ITEMNFREMBENEF", gErr)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149535)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItemNFItemPC_Exclui_RelacionadaItemNF(ByVal lNumIntItemNF As Long) As Long
'Exlui Registros de ItemNFItemPC associados ao ItemNF de número interno passado em lNumIntItemNF

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1) As Long
Dim lNumIntPedCompra As Long

On Error GoTo Erro_ItemNFItemPC_Exclui_RelacionadaItemNF

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67492
    Next

    'Procura por ItemNFitemPC associado ao ItemNF da Nota Fiscal
    lErro = Comando_ExecutarPos(alComando(0), "SELECT ItemPedCompra FROM ItemNFItemPC WHERE ItemNFiscal = ?", 0, lNumIntPedCompra, lNumIntItemNF)
    If lErro <> AD_SQL_SUCESSO Then gError 67490
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67491
    
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO
    
        'Exlui Registros de ItemNFItemPC associados
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItemNFItemPC", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 67495
        
        'Busca próximo ItemNFItemPC
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67493
        
    Loop
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItemNFItemPC_Exclui_RelacionadaItemNF = SUCESSO
    
    Exit Function
    
Erro_ItemNFItemPC_Exclui_RelacionadaItemNF:

    ItemNFItemPC_Exclui_RelacionadaItemNF = gErr
    
    Select Case gErr
    
        Case 67492
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 67490, 67491, 67493
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFITEMPC", gErr)
        
        Case 67495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMNFITEMPC", gErr, lNumIntItemNF)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149536)
    
    End Select

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function PedidoCompra_Desfaz_Baixa(ByVal objNFiscal As ClassNFiscal, ByVal lNumIntItemNF As Long, ByVal objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Refaz os Pedidos de Compras já Baixados e os ItensPC com
'Status Recebido

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(13) As Long
Dim lNumIntPedCompra As Long
Dim lCodigo As Long
Dim tPedCompra As typePedidoCompras
Dim tItemPedido As typeItemPedCompra
Dim lPedCompra As Long
Dim dQuantNFPC As Double
Dim dQuantRecebidaAtual As Double
Dim iStatus As Integer
Dim lNumIntIPC As Long, lNumIntIRC As Long, dQtdRecebida As Double, dQtdRecebidaAtual As Double

On Error GoTo Erro_PedidoCompra_Desfaz_Baixa
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67504
    Next
    
    'Procura por ItemNFitemPC associado ao ItemNF da Nota Fiscal
    lErro = Comando_ExecutarPos(alComando(0), "SELECT ItemPedCompra, Quantidade FROM ItemNFItemPC WHERE ItemNFiscal = ?", 0, lNumIntPedCompra, dQuantNFPC, lNumIntItemNF)
    If lErro <> AD_SQL_SUCESSO Then gError 67505
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67506
    
    'Enquanto encontrar ItemNFItemPC relacionado ao ItemNF passado
    Do While lErro = AD_SQL_SUCESSO
    
        'Lê PedCompra do Item de Pedido de Compras (Baixados ou não) relacionado ao ItemNF
        lErro = Comando_Executar(alComando(1), "SELECT PedCompra FROM ItensPedCompraTodos WHERE NumIntDoc = ?", lPedCompra, lNumIntPedCompra)
        If lErro <> AD_SQL_SUCESSO Then gError 67508
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67509
        
        'Se o ItemPC não está cadastrado, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 67510
        
        tPedCompra.sContato = String(STRING_CONTATO, 0)
        tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
        tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
        tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)
        
        'Verifica se o Pedido de Compras relacionado ao ItemPC está baixado
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc, FilialEmpresa, Codigo, Fornecedor, Filial, Comprador, Contato, TipoDestino, FornCliDestino, FilialDestino, Data, DataEmissao, DataEnvio, DataAlteracao, CondicaoPagto, OutrasDespesas, ValorFrete, ValorSeguro, ValorDesconto,ValorTotal,ValorIPI,TipoFrete, Transportadora, ProxSeqBloqueio,TipoBaixa, MotivoBaixa, Alcada, Observacao FROM PedidoCompraN WHERE NumIntDoc = ? AND Status = ?", 0, _
        tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.lObservacao, lPedCompra, PEDIDOCOMPRA_STATUS_BAIXADO)
        If lErro <> AD_SQL_SUCESSO Then gError 67511
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67512
                        
        'Se o Pedido de Compras está baixado
        If lErro = AD_SQL_SUCESSO Then
                                                    
            tItemPedido.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
            tItemPedido.sProduto = String(STRING_PRODUTO, 0)
            tItemPedido.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
            tItemPedido.sUM = String(STRING_UM_SIGLA, 0)
            
            'Lê os Itens desse Pedido
            lErro = Comando_ExecutarPos(alComando(4), "SELECT NumIntDoc,PedCompra,DataLimite,Produto,DescProduto,Quantidade,QuantRecebida,QuantRecebimento,UM,PrecoUnitario,ValorDesconto,TipoOrigem,NumIntOrigem,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa,Status,Observacao,ValorIPI,AliquotaIPI,AliquotaICMS FROM ItensPedCompraN WHERE PedCompra = ? AND StatusBaixa = ?", 0, _
            tItemPedido.lNumIntDoc, tItemPedido.lPedCompra, tItemPedido.dtDataLimite, tItemPedido.sProduto, tItemPedido.sDescProduto, tItemPedido.dQuantidade, tItemPedido.dQuantRecebida, tItemPedido.dQuantRecebimento, tItemPedido.sUM, tItemPedido.dPrecoUnitario, tItemPedido.dValorDesconto, tItemPedido.iTipoOrigem, tItemPedido.lNumIntOrigem, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRebebForaFaixa, tItemPedido.iStatus, tItemPedido.lObservacao, tItemPedido.dValorIPI, tItemPedido.dAliquotaIPI, tItemPedido.dAliquotaICMS, lPedCompra, ITENSPEDCOMPRA_STATUS_BAIXADO)
            If lErro <> AD_SQL_SUCESSO Then gError 67513
            
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67514
            
            'Equanto encontrar ItensPC
            Do While lErro = AD_SQL_SUCESSO
            
                'Se o ItemPC é o que está vinculado ao ItemNF
                If tItemPedido.lNumIntDoc = lNumIntPedCompra Then
                            
                    'Guarda a nova quantidade Recebida
                    dQuantRecebidaAtual = tItemPedido.dQuantRecebida - dQuantNFPC
                Else
                    dQuantRecebidaAtual = tItemPedido.dQuantRecebida
                End If
                
                If tItemPedido.iStatus = ITEM_PED_COMPRAS_RECEBIDO Then
        
                    'exclui  o numero de itens de pedido de compras que tiveram seu recebimento completado
                    objFilFornFilEmp.lItensPedidosRecebidos = objFilFornFilEmp.lItensPedidosRecebidos - 1
                    
                    'se houve atraso ==> armazena em lAtrasoAcumulado para excluir
                    If (objNFiscal.dtDataEntrada - tItemPedido.dtDataLimite) > 0 And tItemPedido.dtDataLimite <> DATA_NULA Then
                        objFilFornFilEmp.lAtrasoAcumulado = objFilFornFilEmp.lAtrasoAcumulado - (objNFiscal.dtDataEntrada - tItemPedido.dtDataLimite)
                    End If
                    
                End If
                
                If tItemPedido.dQuantidade > dQuantRecebidaAtual Then
                    iStatus = ITEM_PED_COMPRAS_ABERTO
                Else
                    iStatus = ITEM_PED_COMPRAS_RECEBIDO
                End If
                
                'Atualiza o ItemPCBaixado
                lErro = Comando_ExecutarPos(alComando(5), "UPDATE ItensPedCompraN SET QuantRecebida = ?, Status = ?, StatusBaixa=?", alComando(4), dQuantRecebidaAtual, iStatus, ITENSPEDCOMPRA_STATUS_ATIVO)
                If lErro <> AD_SQL_SUCESSO Then gError 67516
                
                'Busca próximo ItemPC
                lErro = Comando_BuscarProximo(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67515
                        
            Loop
                                         
            'Atualiza o Pedido de Compras para ativo
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE PedidoCompraN SET Status = ?, DataBaixa=?", alComando(2), PEDIDOCOMPRA_STATUS_ATIVO, DATA_NULA)
            If lErro <> AD_SQL_SUCESSO Then gError 67542
            
            objFilFornFilEmp.lPedidosEmAberto = objFilFornFilEmp.lPedidosEmAberto + 1
            objFilFornFilEmp.lPedidosRecebidos = objFilFornFilEmp.lPedidosRecebidos - 1
                                    
            
        'Se o Pedido de Compras relacionado ao ItemPC não está baixado
        Else
                                                                                                   
                'Atualiza a Quantidade Recebida e o Status para Ativo
                lErro = ItensPedCompra_Atualiza(alComando(7), alComando(8), lNumIntPedCompra, objFilFornFilEmp, objNFiscal, dQuantNFPC)
                If lErro <> SUCESSO Then gError 89950
                
        End If
        
        'Altera a quantidade Recebida em ItemRCItemPC -Cancelamento\Exclusão de NF
        lErro = Comando_Executar(alComando(11), "SELECT DISTINCT IRCIPC.ItemPC, IRCIPC.ItemRC, INFIRC.Quantidade " & _
                                                "FROM ItemNFItemRC AS INFIRC, ItensPedCompraTodos AS IPC, ItemRCItemPC AS IRCIPC, ItensReqCompraTodos AS IRC, ItemNFItemPC AS INFIPC " & _
                                                "WHERE IPC.NumIntDoc = IRCIPC.ItemPC AND IRCIPC.ItemRC = IRC.NumIntDoc AND INFIRC.ItemReqCompra = IRC.NumIntDoc AND INFIRC.ItemNFiscal = INFIPC.ItemNFiscal AND INFIPC.ItemNFiscal = ? AND IRCIPC.ItemPC = ?", _
                                                lNumIntIPC, lNumIntIRC, dQtdRecebida, lNumIntItemNF, lNumIntPedCompra)
        lErro = Comando_BuscarPrimeiro(alComando(11))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209350
        
        Do While lErro = AD_SQL_SUCESSO
        
            lErro = Comando_ExecutarPos(alComando(12), "SELECT QuantRecebida FROM ItemRCItemPC WHERE ItemPC = ? AND ItemRC = ? ", 0, dQtdRecebidaAtual, lNumIntIPC, lNumIntIRC)
            If lErro <> AD_SQL_SUCESSO Then gError 209351
            
            lErro = Comando_BuscarPrimeiro(alComando(12))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209352
        
            lErro = Comando_ExecutarPos(alComando(13), "UPDATE ItemRCItemPC SET QuantRecebida = ? ", alComando(12), dQtdRecebidaAtual - dQtdRecebida)
            If lErro <> AD_SQL_SUCESSO Then gError 209353
        
            lErro = Comando_BuscarProximo(alComando(11))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209354
        
        Loop
        
        'Exlui Registros de ItemNFItemPC associados
        lErro = Comando_ExecutarPos(alComando(10), "DELETE FROM ItemNFItemPC", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 67495
        
        'Busca próximo ItemNFItemPC
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67507
    
    Loop
        
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    PedidoCompra_Desfaz_Baixa = SUCESSO
    
    Exit Function

Erro_PedidoCompra_Desfaz_Baixa:

    PedidoCompra_Desfaz_Baixa = gErr

    Select Case gErr
    
        Case 67495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMNFITEMPC", gErr, lNumIntItemNF)
    
        Case 67504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 67505, 67506, 67507, 67544, 67545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFITEMPC", gErr)
                
        Case 67505, 67506, 67507, 67544, 67545, 209350 To 209354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDCOMPRASTODOS", gErr)
        
        Case 67510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", gErr, lNumIntPedCompra)
        
        Case 67511, 67512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRABAIXADO", gErr)
                
        Case 67513, 67514, 67515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)
                
        Case 67516
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDCOMPRA", gErr, tItemPedido.lNumIntDoc)
                    
        Case 67517
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPEDCOMPRA", gErr, tItemPedido.lNumIntDoc)
            
        Case 67542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PEDIDOCOMPRABAIXADO", gErr, tItemPedido.lNumIntDoc)
        
        Case 67543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PEDIDOCOMPRA", gErr, lPedCompra)
        
        Case 67546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNFITEMPC_NAO_CADASTRADO", gErr, lNumIntItemNF, tItemPedido.lNumIntDoc)
                    
        Case 89950
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149537)
    
    End Select
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function ItensPedCompra_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lNumIntPedCompra As Long, ByVal objFilFornFilEmp As ClassFilFornFilEmp, ByVal objNFiscal As ClassNFiscal, ByVal dQuantNFPC As Double) As Long
'Atualiza a Quantidade Recebida e o Status para Ativo

Dim lErro As Long
Dim tItemPedido As typeItemPedCompra

On Error GoTo Erro_ItensPedCompra_Atualiza
                
    'Lê a quantidade do ItemPC que foi Recebida
    lErro = Comando_ExecutarPos(lComando1, "SELECT QuantRecebida, Status, DataLimite FROM ItensPedCompra WHERE NumIntDoc = ?", 0, tItemPedido.dQuantidade, tItemPedido.iStatus, tItemPedido.dtDataLimite, lNumIntPedCompra)
    If lErro <> SUCESSO Then gError 67518
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67519
    
    'Se não encontrou o ItemPC, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67520
    
    If tItemPedido.iStatus = ITEM_PED_COMPRAS_RECEBIDO Then

        'exclui  o numero de itens de pedido de compras que tiveram seu recebimento completado
        objFilFornFilEmp.lItensPedidosRecebidos = objFilFornFilEmp.lItensPedidosRecebidos - 1
        
        'se houve atraso ==> armazena em lAtrasoAcumulado para excluir
        If (objNFiscal.dtDataEntrada - tItemPedido.dtDataLimite) > 0 And tItemPedido.dtDataLimite <> DATA_NULA Then
            objFilFornFilEmp.lAtrasoAcumulado = objFilFornFilEmp.lAtrasoAcumulado - (objNFiscal.dtDataEntrada - tItemPedido.dtDataLimite)
        End If
        
    End If
    
    'Atualiza a quantidade e o Status do ItemPC
    lErro = Comando_ExecutarPos(lComando2, "UPDATE ItensPedCompra Set QuantRecebida = ?, Status = ?", lComando1, tItemPedido.dQuantidade - dQuantNFPC, ITEM_PED_COMPRAS_ABERTO)
    If lErro <> AD_SQL_SUCESSO Then gError 67521

    ItensPedCompra_Atualiza = SUCESSO
    
    Exit Function

Erro_ItensPedCompra_Atualiza:

    ItensPedCompra_Atualiza = gErr
    
    Select Case gErr
    
        Case 67518, 67519
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)
    
        Case 67520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", gErr, lNumIntPedCompra)
    
        Case 67521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDCOMPRA", gErr, lNumIntPedCompra)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149538)
    
    End Select
    
    Exit Function

End Function


Function ItemNFItemRC_Exclui_RelacionadaItemNF(ByVal lNumIntItemNF As Long) As Long
'Exlui Registros de ItemNFItemRC associados ao ItemNF de número interno passado em lNumIntItemNF

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1) As Long
Dim lNumIntReqCompra As Long

On Error GoTo Erro_ItemNFItemRC_Exclui_RelacionadaItemNF

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67497
    Next

    'Procura por ItemNFitemRC associado ao ItemNF da Nota Fiscal
    lErro = Comando_ExecutarPos(alComando(0), "SELECT ItemReqCompra FROM ItemNFItemRC WHERE ItemNFiscal = ?", 0, lNumIntReqCompra, lNumIntItemNF)
    If lErro <> AD_SQL_SUCESSO Then gError 67498
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67499
    
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO
    
        'Exlui Registros de ItemNFItemRC associados
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItemNFItemRC", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 67501
        
        'Busca próximo ItemNFItemRC
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67500
        
    Loop
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItemNFItemRC_Exclui_RelacionadaItemNF = SUCESSO
    
    Exit Function
    
Erro_ItemNFItemRC_Exclui_RelacionadaItemNF:

    ItemNFItemRC_Exclui_RelacionadaItemNF = gErr
    
    Select Case gErr
    
        Case 67497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 67498, 67499, 67500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFITEMRC1", gErr)
        
        Case 67501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMNFITEMRC", gErr, lNumIntItemNF)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149539)
    
    End Select

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function RequisicaoCompra_Desfaz_Baixa(ByVal lNumIntNF As Long, ByVal lNumIntItemNF As Long) As Long
'Refaz os Requisições de Compras já Baixadas e os ItensRC com
'Status Recebido

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(13) As Long
Dim lNumIntReqCompra As Long
Dim lCodigo As Long
Dim tReqCompra As typeRequisicaoCompras
Dim tItemReq As typeItemReqCompra
Dim dQuantidade As Double
Dim lReqCompra As Long
Dim dQuantNFRC As Double
Dim dQuantRecebidaAtual As Double
Dim dQuantPedidaAtual As Double
Dim iStatus As Integer, dQtd1Aux As Double, dQtd2Aux As Double, dQtdPCAux As Double

On Error GoTo Erro_RequisicaoCompra_Desfaz_Baixa

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67522
    Next

    'Procura por ItemNFItemRC associado ao ItemNF da Nota Fiscal
    lErro = Comando_ExecutarPos(alComando(0), "SELECT ItemReqCompra, Quantidade FROM ItemNFItemRC WHERE ItemNFiscal = ?", 0, lNumIntReqCompra, dQuantNFRC, lNumIntItemNF)
    If lErro <> AD_SQL_SUCESSO Then gError 67523

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67524

    'Enquanto encontrar ItemNFItemRC relacionado ao ItemNF passado
    Do While lErro = AD_SQL_SUCESSO

        'Lê ReqCompra do Item de Requisição de Compras (Baixados ou não) relacionado ao ItemNF
        lErro = Comando_Executar(alComando(1), "SELECT ReqCompra FROM ItensReqCompraTodos WHERE NumIntDoc = ?", lReqCompra, lNumIntReqCompra)
        If lErro <> AD_SQL_SUCESSO Then gError 67525

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67526

        'Se o ItemPC não está cadastrado, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 67527

        tReqCompra.sCcl = String(STRING_CCL, 0)
        tReqCompra.sOPCodigo = String(STRING_OPCODIGO, 0)
        tReqCompra.sDigitador = String(STRING_USUARIO_CODIGO, 0)

        'Verifica se a Requisição de Compras relacionado ao ItemRC está baixada
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc, FilialEmpresa, Codigo, Data, DataEnvio, DataLimite, Urgente, Requisitante, Digitador, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao, OPCodigo FROM RequisicaoCompraN WHERE NumIntDoc = ? AND Status = ?", 0, _
        tReqCompra.lNumIntDoc, tReqCompra.iFilialEmpresa, tReqCompra.lCodigo, tReqCompra.dtData, tReqCompra.dtDataEnvio, tReqCompra.dtDataLimite, tReqCompra.lUrgente, tReqCompra.lRequisitante, tReqCompra.sDigitador, tReqCompra.sCcl, tReqCompra.iFilialCompra, tReqCompra.iTipoDestino, tReqCompra.lFornCliDestino, tReqCompra.iFilialDestino, tReqCompra.lObservacao, tReqCompra.iTipoTributacao, tReqCompra.sOPCodigo, lReqCompra, REQUISICAOCOMPRA_STATUS_BAIXADO)
        If lErro <> AD_SQL_SUCESSO Then gError 67528

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67529

        'Se a Requisição de Compras está baixada
        If lErro = AD_SQL_SUCESSO Then
                
            tItemReq.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
            tItemReq.sProduto = String(STRING_PRODUTO, 0)
            tItemReq.sUM = String(STRING_UM_SIGLA, 0)
            tItemReq.sCcl = String(STRING_CCL, 0)
            tItemReq.sContaContabil = String(STRING_CONTA, 0)

            'Lê os Itens dessa Requisição
            lErro = Comando_ExecutarPos(alComando(4), "SELECT NumIntDoc, ReqCompra, Produto, DescProduto, Status, Quantidade, QuantPedida, QuantRecebida, QuantCancelada, UM, Ccl, Almoxarifado, ContaContabil, TipoTributacao, Observacao, Fornecedor, Filial, Exclusivo FROM ItensReqCompraN WHERE ReqCompra = ? AND StatusBaixa=?", 0, _
            tItemReq.lNumIntDoc, tItemReq.lReqCompra, tItemReq.sProduto, tItemReq.sDescProduto, tItemReq.iStatus, tItemReq.dQuantidade, tItemReq.dQuantPedida, tItemReq.dQuantRecebida, tItemReq.dQuantCancelada, tItemReq.sUM, tItemReq.sCcl, tItemReq.iAlmoxarifado, tItemReq.sContaContabil, tItemReq.iTipoTributacao, tItemReq.lObservacao, tItemReq.lFornecedor, tItemReq.iFilial, tItemReq.iExclusivo, lReqCompra, ITENSREQCOMPRA_STATUS_BAIXADO)
            If lErro <> AD_SQL_SUCESSO Then gError 67530

            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67531

            'Enquanto encontrar ItensRC
            Do While lErro = AD_SQL_SUCESSO
                
                If tItemReq.lNumIntDoc = lNumIntReqCompra Then
                
                    'Guarda a nova quantidade Recebida
                    dQuantRecebidaAtual = tItemReq.dQuantRecebida - dQuantNFRC
                    dQuantPedidaAtual = tItemReq.dQuantPedida + dQuantNFRC
                    
                Else
                    
                    dQuantRecebidaAtual = tItemReq.dQuantRecebida
                    dQuantPedidaAtual = tItemReq.dQuantPedida
                
                End If
                If dQuantPedidaAtual < 0 Then dQuantPedidaAtual = 0
                                
                'se ainda tem quantidade a receber
                If tItemReq.dQuantidade - tItemReq.dQuantCancelada > dQuantRecebidaAtual Then
                    iStatus = ITEM_REQ_ABERTO
                Else
                    iStatus = ITEM_REQ_RECEBIDO
                End If
                                
                'Atualiza o ItemRCBaixado
                lErro = Comando_ExecutarPos(alComando(5), "UPDATE ItensReqCompraN SET QuantRecebida = ?, QuantPedida = ?, Status = ?, StatusBaixa = ?", alComando(4), dQuantRecebidaAtual, dQuantPedidaAtual, iStatus, ITENSREQCOMPRA_STATUS_ATIVO)
                If lErro <> AD_SQL_SUCESSO Then gError 67532

                'Busca próximo ItemRC
                lErro = Comando_BuscarProximo(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67534
            
            Loop

            'Exclui a Requisição de Compras Baixada
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE RequisicaoCompraN SET Status = ?, DataBaixa=?", alComando(2), REQUISICAOCOMPRA_STATUS_ATIVO, DATA_NULA)
            If lErro <> AD_SQL_SUCESSO Then gError 67535

        'Se o Pedido de Compras relacionado ao ItemPC não está baixado
        Else

                'Lê a quantidade do ItemRC que foi Recebida
                lErro = Comando_ExecutarPos(alComando(7), "SELECT QuantRecebida FROM ItensReqCompra WHERE NumIntDoc = ?", 0, dQuantidade, lNumIntReqCompra)
                If lErro <> SUCESSO Then gError 67537

                lErro = Comando_BuscarPrimeiro(alComando(7))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67538

                'Se não encontrou o ItemRC, erro
                If lErro = AD_SQL_SEM_DADOS Then gError 67539
                  
                'Atualiza a quantidade e o Status do ItemPC
                lErro = Comando_ExecutarPos(alComando(8), "UPDATE ItensReqCompra Set QuantRecebida = QuantRecebida - ?, QuantPedida = QuantPedida + ?, Status = ?", alComando(7), dQuantNFRC, dQuantNFRC, ITEM_REQ_ABERTO)
                If lErro <> AD_SQL_SUCESSO Then gError 67540

        End If
        
        'Código para acertar a quantidade pedida, não pode fazer pela diferença utilizando a quantidade recebida
        'porque pode se receber mais do que se pede e a quantidade pedida não fica negativa
        dQtdPCAux = 0
        lErro = Comando_Executar(alComando(11), "SELECT Quantidade, QuantRecebida FROM ItemRCItemPC WHERE ItemRC = ?", dQtd1Aux, dQtd2Aux, lNumIntReqCompra)
        If lErro <> SUCESSO Then gError 67537

        lErro = Comando_BuscarPrimeiro(alComando(11))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67538
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            If dQtd1Aux > dQtd2Aux Then dQtdPCAux = dQtdPCAux + dQtd1Aux - dQtd2Aux
        
            lErro = Comando_BuscarProximo(alComando(11))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67538
        
        Loop
        
        'Lê a quantidade do ItemRC que foi Recebida
        lErro = Comando_ExecutarPos(alComando(12), "SELECT QuantRecebida FROM ItensReqCompra WHERE NumIntDoc = ?", 0, dQuantidade, lNumIntReqCompra)
        If lErro <> SUCESSO Then gError 67537

        lErro = Comando_BuscarPrimeiro(alComando(12))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67538

        'Se não encontrou o ItemRC, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 67539
          
        'Atualiza a quantidade e o Status do ItemPC
        lErro = Comando_ExecutarPos(alComando(13), "UPDATE ItensReqCompra SET QuantPedida = ? ", alComando(12), dQtdPCAux)
        If lErro <> AD_SQL_SUCESSO Then gError 67540

        'Exlui Registros de ItemNFItemRC associados
        lErro = Comando_ExecutarPos(alComando(10), "DELETE FROM ItemNFItemRC", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 67501

        'Busca próximo ItemNFItemRC
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67541

    Loop

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RequisicaoCompra_Desfaz_Baixa = SUCESSO

    Exit Function

Erro_RequisicaoCompra_Desfaz_Baixa:

    RequisicaoCompra_Desfaz_Baixa = gErr

    Select Case gErr

        Case 67501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMNFITEMRC", gErr, lNumIntItemNF)
            
        Case 67522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67523, 67524, 67541, 67547, 67548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFITEMRC1", gErr)

        Case 67525, 67526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRASTODOS", gErr)

        Case 67527, 67539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENSREQCOMPRA_NAO_CADASTRADO2", gErr, lNumIntReqCompra)

        Case 67528, 67529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRABAIXADA", gErr)

        Case 67530, 67531, 67534, 67537, 67538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)

        Case 67532
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSREQCOMPRA", gErr, tItemReq.lNumIntDoc)

        Case 67533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMREQCOMPRA", gErr, tItemReq.lNumIntDoc)
        
        Case 67535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REQUISICAOCOMPRA", gErr, lReqCompra)

        Case 67536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REQUISICAO_COMPRA", gErr, tReqCompra.lCodigo)
            
        Case 67540
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSREQCOMPRA", gErr, tItemReq.lNumIntDoc)
            
        Case 67549
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMNFITEMRC_NAO_CADASTRADO", gErr, lNumIntItemNF, tItemReq.lNumIntDoc)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149540)

    End Select

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RecebMaterialFCom_Grava(ByVal objNFiscal As ClassNFiscal, ByVal colCodPedCompras As Collection) As Long
'Grava a Nota Fiscal de Recebimento de Material de Fornecedor de compras

Dim lErro As Long
Dim lNumIntNFItem As Long
Dim lNumIntDoc As Long
Dim lNumIntDocOrigem As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim objItemNF As ClassItemNF
Dim alComando(15) As Long
Dim iTipo As Integer
Dim dtData As Date
Dim objTipoDocInfo As New ClassTipoDocInfo
'#################################
'Inserido por Wagner 26/01/2006
Dim dFator As Double
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dQuantItemNFUMEst As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
'#################################
Dim objProduto As ClassProduto

On Error GoTo Erro_RecebMaterialFCom_Grava

    'Abre os comandos de Transação
    For iIndice = LBound(alComando) To UBound(alComando) 'iIndice = 0 To 13
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 54586
    Next

    'Abre Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 54587

    'Faz lock dos objetos associados
    lErro = RecebMaterialFCom_Lock_ObjAssociados(objNFiscal, alComando(0), alComando(1), alComando(11), alComando(12), alComando(13), colCodPedCompras)
    If lErro <> SUCESSO Then gError 54589

    dtData = objNFiscal.dtDataEmissao
    objNFiscal.dtDataEmissao = objNFiscal.dtDataEntrada

    If objNFiscal.iTipoNFiscal = DOCINFO_NRFPCO Then
        iTipo = DOCINFO_NF_INT_ENTRADA
    Else
        iTipo = DOCINFO_NF_EXTERNA
        
        'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
        lErro = CF("NFiscalEntrada_Verifica_Existencia", objNFiscal, iTipo)
        If lErro <> SUCESSO Then gError 54592
        
        If objNFiscal.lNumRecebimento = 0 Then
            
            'Verifica a Existencia de um Recebimento  com os Valores de Nota igual
            lErro = CF("RecebimentoFCom_Verifica_Existencia", objNFiscal)
            If lErro <> SUCESSO Then gError 54593
        
        End If
        
    End If
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo de documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 27623 Then gError 54594
        
    'Se não encontrar --> erro
    If lErro = 27623 Then gError 54714
    
    objNFiscal.dtDataEmissao = dtData
    
    '###################################################
    'Inserido por Wagner 25/01/2006
    iIndice = 0
    
    For Each objItemNF In objNFiscal.colItensNF
        
        iIndice = iIndice + 1
        
        If objItemNF.colItensRomaneioGrade.Count > 0 Then
            
            For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                Set objProduto = New ClassProduto
            
                objProduto.sCodigo = objItemRomaneio.sProduto
                
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 23080 Then gError 141600
                
                If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    
                    lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemRomaneio.sUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 141600
    
                    dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
    
                    dSomaAlocacao = 0
                                                                    
                    For Each objReservaItem In objItemRomaneio.colLocalizacao
                                            
                        dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                                                    
                    Next
                    
                    dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)
                    
                    If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > QTDE_ESTOQUE_DELTA Then
                    
                        If dSomaAlocacao < dQuantItemNFUMEst Then gError 141601
                        
                        If dSomaAlocacao > dQuantItemNFUMEst Then gError 141602
                    
                    End If
                    
                End If
                
            Next
            
        Else
                    
            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                If objItemNF.ColAlocacoes.Count = 0 Then gError 141603

                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 141604

                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
    
                dSomaAlocacao = 0
                                                                
                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
                    
                Next

                dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)
                
                If Abs(dSomaAlocacao - dQuantItemNFUMEst) > QTDE_ESTOQUE_DELTA Then
                 
                    If dSomaAlocacao < dQuantItemNFUMEst Then gError 141605
                    
                    If dSomaAlocacao > dQuantItemNFUMEst Then gError 141606
            
                End If
                
            End If
            
        End If
        
    Next
    '###################################################
    
    'Se o código do Recebimento estiver = 0 ==>> INSERÇÃO
    If objNFiscal.lNumRecebimento = 0 Then

        lErro = RecebMaterialFCom_Grava_Insercao(objNFiscal, alComando(3), alComando(4), alComando(10), alComando(14))
        If lErro <> SUCESSO Then gError 54590

    Else 'Se não ==>> ALTERAÇÃO

        lErro = RecebMaterialFCom_Grava_Atualiza(objNFiscal, alComando(2), alComando(5), alComando(6), alComando(7), alComando(8), alComando(9), alComando(15))
        If lErro <> SUCESSO Then gError 54591

    End If

    'Confirma a Transação
    lErro = Transacao_Commit
    If lErro <> SUCESSO Then gError 54588

    'Fecha os comandos
    For iIndice = 0 To 12
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    RecebMaterialFCom_Grava = SUCESSO

    Exit Function

Erro_RecebMaterialFCom_Grava:

    RecebMaterialFCom_Grava = gErr

    Select Case gErr

        Case 54586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 54588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 54589, 54590, 54591, 54592, 54593  'Tratados na Rotinas chamadas
        
        Case 54594
    
        Case 54714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO", gErr, objTipoDocInfo)

        Case 141601, 141605
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INCOMPLETA", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case 141602, 141606
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)
        
        Case 141603
            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INEXISTENTE", gErr, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149541)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = 0 To 12
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function RecebMaterialFCom_Lock_ObjAssociados(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal colCodPedCompras As Collection) As Long
'Faz o Lock dos Objetos associados ao RecebMaterialFCom
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim iControleEstoque As Integer
Dim iIndice As Integer
Dim lErro As Long
Dim lCodFornecedor As Long
Dim iCodigo As Integer
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim objCodigoNome As AdmCodigoNome
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim objPedidoCompras As New ClassPedidoCompras
Dim colCodigoNome As New AdmColCodigoNome
Dim iLocado As Integer

On Error GoTo Erro_Lock_ObjAssociados

    lErro = CF("FilialFornecedor_Lock", objNFiscal.lFornecedor, objNFiscal.iFilialForn)
    If lErro <> SUCESSO And lErro <> 42870 Then Error 54596

    'Se não achou a Filial do Fornecedor --> erro
    If lErro <> SUCESSO Then Error 54599

    'Verifica se o código da Transportadora foi preenchido
    If objNFiscal.iCodTransportadora <> 0 Then

        'Faz Lock shared na Transportadora
        lErro = CF("Transportadora_Lock1", objNFiscal.iCodTransportadora)
        If lErro <> SUCESSO And lErro <> 43038 Then Error 54597

        'Se não achou a Transportadora --> erro
        If lErro <> SUCESSO Then Error 54600

    End If

    'Para cada Item de ObjNFiscal.colItens
    For Each objItemNF In objNFiscal.colItensNF

        objProduto.sCodigo = objItemNF.sProduto

        lErro = CF("Produto_Lock", objProduto)
        If lErro <> SUCESSO Then Error 54598

        'Armazena o ControleEstoque do Produto
        objItemNF.iControleEstoque = objProduto.iControleEstoque

        If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then

            iLocado = False

            For Each objCodigoNome In colCodigoNome
                If objCodigoNome.iCodigo = objItemNF.iAlmoxarifado Then
                    iLocado = True
                    Exit For
                End If
            Next

            If Not iLocado Then

                lErro = CF("Almoxarifado_Lock", objItemNF.iAlmoxarifado, lComando1)
                If lErro <> SUCESSO Then Error 54595

            End If
        End If

    Next

    'Para cada Pedido de Compras na coleção colCodPedCompras
    For iIndice = 1 To colCodPedCompras.Count

        objPedidoCompras.iFilialEmpresa = objNFiscal.iFilialPedido
        objPedidoCompras.lCodigo = colCodPedCompras.Item(iIndice)

        'Faz Lock em Pedido de Compras
        lErro = CF("PedidoCompra_LeMod", objPedidoCompras, lComando5)
        If lErro <> SUCESSO And lErro <> 56259 Then Error 54601
        If lErro = 56259 Then Error 54602
    
    Next

    RecebMaterialFCom_Lock_ObjAssociados = SUCESSO

    Exit Function

Erro_Lock_ObjAssociados:

    RecebMaterialFCom_Lock_ObjAssociados = Err

    Select Case Err

        Case 54595, 54596, 54597, 54598, 54601

        Case 54599
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 54600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objNFiscal.iCodTransportadora)

        Case 54602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objPedidoCompras.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149542)

    End Select

    Exit Function

End Function

Private Function RecebMaterialFCom_Grava_Insercao(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long) As Long
'Faz a Inserção do RecebMaterialFCom na Tabela NFiscal
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lCodigo As Long
Dim lNumIntDoc As Long
Dim dtDataEntrada As Date
Dim vbMsg As VbMsgBoxResult
Dim objItemNF As ClassItemNF
Dim objMovEstoque As ClassMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
'distribuicao
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dFator As Double
Dim iIndice As Integer
Dim dQuantItemNFUMEst As Double

On Error GoTo Erro_RecebMaterialFCom_Grava_Insercao

    'Gera o código do Recebimento
    lErro = CF("Recebimento_Automatico", objNFiscal)
    If lErro <> SUCESSO Then gError 54613
    
    'Lê o Número do Próximo Código Automático Interno.
    lErro = CF("NFiscal_Automatico", lNumIntDoc)
    If lErro <> SUCESSO Then gError 54608
    
    'Preenche o objNFiscal com Código
    objNFiscal.lNumIntDoc = lNumIntDoc

    'Insere no BD objNFiscal em NFiscal
    With objNFiscal
        lErro = Comando_Executar(lComando2, "INSERT INTO NFiscal(DataAtualizacao, NumRecebimento , NumIntDoc, FilialEmpresa, Serie, NumNotaFiscal, Status, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEntrada, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, VolumeNumero, FreteRespons, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,Observacao, FilialPedido, HoraEntrada, Moeda, TaxaMoeda, ChvNFe, HoraEmissao, DataCadastro, HoraCadastro, UsuCadastro) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            gdtDataHoje, .lNumRecebimento, .lNumIntDoc, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .iStatus, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEntrada, .lNumPedidoVenda, _
            .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sVolumeNumero, .iFreteRespons, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sObservacao, .iFilialPedido, CDbl(.dtHoraEntrada), .iMoeda, .dTaxaMoeda, .sChvNFe, CDbl(IIf(.dtDataSaida <> DATA_NULA, IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraSaida, .dtHoraEmissao, .dtHoraSaida), IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraEntrada, .dtHoraEmissao, .dtHoraEntrada))), gdtDataHoje, CDbl(Time), gsUsuario)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 54614

    'Inserção de Itens da Nota Fiscal
    lErro = RecebimentoMatFCom_Insere_ItensNFiscal(objNFiscal, lComando3, lComando4)
    If lErro <> SUCESSO Then gError 54610

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 54612
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 54607
    
    'Se o TipoMovtoEstoque for positivo
    If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
        'Cria MovEstoque
        Set objMovEstoque = New ClassMovEstoque
    
        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then gError 54611
    
        'Preenche o objMovEstoque
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque
        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = objTipoDocInfo.sSigla & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
    
'mario distribuicao substituicao do codigo original acrescentando o tab de distribuicao
    
        'Para cada elemento de colItens de objNFiscal, adiciona elemento em colItens de objMovEstoque
        For Each objItemNF In objNFiscal.colItensNF
    
            iIndice = iIndice + 1
    
            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
           
                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 89688

                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
           
                dSomaAlocacao = 0
                                                                
                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
           
                    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemNF.dQuantidade * objItemNF.dPrecoUnitario) - objItemNF.dValorDesconto) * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                    objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                    objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                    objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
    
                Next
                
                If dSomaAlocacao > dQuantItemNFUMEst Then gError 89689
    
            End If
        Next

        'Chama função de Gravação de MovEstoque
        lErro = CF("MovEstoque_Grava", objMovEstoque) 'Chamada de dentro de Transação
        If lErro <> SUCESSO Then gError 54609

    End If
    
    RecebMaterialFCom_Grava_Insercao = SUCESSO

    Exit Function

Erro_RecebMaterialFCom_Grava_Insercao:

    RecebMaterialFCom_Grava_Insercao = gErr

    Select Case gErr

        Case 54607
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 54608, 54609, 54610, 54611, 54612, 54613, 89688
  
        Case 54614
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 89689
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149543)

    End Select

    Exit Function

End Function

Private Function RecebimentoMatFCom_Insere_ItensNFiscal(ByVal objNFiscal As ClassNFiscal, ByVal lComando As Long, ByVal lComando2 As Long) As Long
'Insere os Itens de Nota Fiscal de RecebMaterialFCom
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lNumIntNFItem As Long
Dim objItemNF As ClassItemNF

On Error GoTo Erro_RecebimentoMatFCom_Insere_ItensNFiscal

    'Para cada Item da Nota Fiscal
    For Each objItemNF In objNFiscal.colItensNF

        'Preenche lNumIntNF com objNFiscal.lNumIntDoc
        objItemNF.lNumIntNF = objNFiscal.lNumIntDoc

        'Lê o número do próximo Código Automático interno do Ítem
        lErro = CF("NFiscalItem_Automatico", lNumIntNFItem)
        If lErro <> SUCESSO Then Error 54617

        objItemNF.lNumIntDoc = lNumIntNFItem

        'Insere no BD em ItensNFiscal
        lErro = Comando_Executar(lComando, "INSERT INTO ItensNFiscal(NumIntNF, Item, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, Almoxarifado, NumIntDocOrig, Status,Ccl) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objItemNF.lNumIntNF, objItemNF.iItem, objItemNF.sProduto, objItemNF.sUnidadeMed, objItemNF.dQuantidade, objItemNF.dPrecoUnitario, objItemNF.dPercDesc, objItemNF.dValorDesconto, objItemNF.dtDataEntrega, objItemNF.sDescricaoItem, objItemNF.dValorAbatComissao, objItemNF.lNumIntPedVenda, objItemNF.lNumIntItemPedVenda, objItemNF.lNumIntDoc, objItemNF.lNumIntTrib, objItemNF.iAlmoxarifado, objItemNF.lNumIntDocOrig, objItemNF.iStatus, objItemNF.sCcl)
        If lErro <> AD_SQL_SUCESSO Then Error 54618
        
        'Inclui em ItemNFItemPC o ItemNF associado a ItemPC
        lErro = ItemNFItemPC_Inclui_Recebimento(objItemNF)
        If lErro <> SUCESSO Then Error 54615
        
        objItemNF.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
        objItemNF.objInfoAdicDocItem.lNumIntDocItem = objItemNF.lNumIntDoc
        lErro = CF("InfoAdicDocItem_Insere", objItemNF.objInfoAdicDocItem, lComando2)
        If lErro <> SUCESSO Then Error 54615
    
    Next

    RecebimentoMatFCom_Insere_ItensNFiscal = SUCESSO

    Exit Function

Erro_RecebimentoMatFCom_Insere_ItensNFiscal:

    RecebimentoMatFCom_Insere_ItensNFiscal = Err

    Select Case Err

        Case 54615, 54617

        Case 54618
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", Err, objItemNF.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149544)

    End Select

    Exit Function

End Function

Private Function RecebMaterialFCom_Grava_Atualiza(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal lComando4 As Long, ByVal lComando5 As Long, ByVal lComando6 As Long, ByVal lComando7 As Long) As Long
'Faz a Atualização do RecebMaterialFCom na Tabela NFiscal
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lCodigo As Long
Dim lNumIntDoc As Long
Dim iTipoDoc As Integer
Dim objItemNF As ClassItemNF
Dim objMovEstoque As ClassMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
'distribuicao
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dFator As Double
Dim iIndice As Integer
Dim dQuantItemNFUMEst As Double

On Error GoTo Erro_RecebMaterialFCom_Grava_Atualiza

    'Testa se o Recebimento já foi baixado
    lErro = CF("Recebimento_Testa_Baixado", objNFiscal)
    If lErro <> SUCESSO Then Error 54604
    
    'Tenta ler o Recebimento na Tabela NFiscal
    lErro = Comando_ExecutarPos(lComando3, "SELECT NumIntDoc FROM NFiscal WHERE FilialEmpresa = ? AND NumRecebimento = ?", 0, lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.lNumRecebimento)
    If lErro <> AD_SQL_SUCESSO Then Error 54677

    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54678
    
    'Faz Lock Exclusive no Recebimento
    lErro = Comando_LockExclusive(lComando3)
    If lErro <> SUCESSO Then Error 54679

    objNFiscal.lNumIntDoc = lNumIntDoc

    'Atualiza o Recebimento
    lErro = Comando_ExecutarPos(lComando2, "UPDATE NFiscal SET ValorTotal = ?, ValorProdutos = ?, ValorSeguro = ?, ValorDesconto = ?, ValorFrete = ?, ValorOutrasDespesas = ?, CodTransportadora = ?, Placa = ?, PlacaUF = ?, VolumeQuant = ?, VolumeEspecie = ?, VolumeMarca = ?, VolumeNumero = ?, FreteRespons = ?, MensagemNota = ?, PesoLiq = ?, PesoBruto = ?, Observacao = ?, Serie = ?, NumNotaFiscal = ?, DataEntrada = ?, TipoNFiscal = ?, Fornecedor = ? , FilialForn = ?, HoraEntrada = ?, DataAtualizacao=? ", lComando3, _
    objNFiscal.dValorTotal, objNFiscal.dValorProdutos, objNFiscal.dValorSeguro, objNFiscal.dValorDesconto, objNFiscal.dValorFrete, objNFiscal.dValorOutrasDespesas, objNFiscal.iCodTransportadora, objNFiscal.sPlaca, objNFiscal.sPlacaUF, objNFiscal.lVolumeQuant, objNFiscal.lVolumeEspecie, objNFiscal.lVolumeMarca, objNFiscal.sVolumeNumero, objNFiscal.iFreteRespons, objNFiscal.sMensagemNota, objNFiscal.dPesoLiq, objNFiscal.dPesoBruto, objNFiscal.sObservacao, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEntrada, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, CDbl(objNFiscal.dtHoraEntrada), gdtDataHoje)
    If lErro <> AD_SQL_SUCESSO Then Error 54493

    '==========================
    'EXCLUSÃO ITENS NOTA FISCAL
    '==========================
    lErro = RecebimentoMatFCom_Exclui_ItensNFiscal(objNFiscal, lComando4, lComando5)
    If lErro <> SUCESSO Then Error 54682

    'Inserção de Itens da Nota Fiscal
    lErro = RecebimentoMatFCom_Insere_ItensNFiscal(objNFiscal, lComando6, lComando7)
    If lErro <> SUCESSO Then Error 54681

    '====================
    'INSERÇÃO MOV ESTOQUE
    '====================
    'Cria objMovEstoque
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 54672
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then Error 54676
    
    'Se o TipoMovtoEstoque for positivo
    If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
        Set objMovEstoque = New ClassMovEstoque
    
        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then Error 54603
    
        'Preenche o objMovEstoque
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque
        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = objTipoDocInfo.sSigla & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
    
    
'inicio distribuicao
        'Para cada elemento de colItens de objNFiscal, adiciona elemento em colItens de objMovEstoque
        For Each objItemNF In objNFiscal.colItensNF
    
            iIndice = iIndice + 1
    
            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
           
                dSomaAlocacao = 0
                                                                
                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 89686
                                                                
                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))

                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
           
                    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemNF.dQuantidade * objItemNF.dPrecoUnitario) - objItemNF.dValorDesconto) * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                    objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                    objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                    objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
    
                Next
                
                If dSomaAlocacao > dQuantItemNFUMEst Then gError 89687
    
            End If
        Next
'fim distribuicao
    
        'Chama função de Gravação de MovEstoque
        lErro = CF("MovEstoque_Grava", objMovEstoque) 'Chamada de dentro de Transação
        If lErro <> SUCESSO Then Error 54680

    End If
    
    RecebMaterialFCom_Grava_Atualiza = SUCESSO

    Exit Function

Erro_RecebMaterialFCom_Grava_Atualiza:

    RecebMaterialFCom_Grava_Atualiza = Err

    Select Case Err

        Case 54603, 54604, 54672, 54680, 54681, 54682, 89686

        Case 54676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)

        Case 54677, 54678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 54679
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", Err, objNFiscal.lNumNotaFiscal)

        Case 54493
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", Err, objNFiscal.lNumNotaFiscal)

        Case 89687
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149545)

    End Select

    Exit Function

End Function

Private Function RecebimentoMatFCom_Exclui_ItensNFiscal(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Exclui os Itens da Nota Fiscal que estão ==>>
'==>>relacionados ao RecebMaterial de Compras Selecionado
'CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando3 As Long
Dim objMovEstoqueExcluir As ClassMovEstoque
Dim dQuantidade As Double
Dim lCodigo As Long
Dim sProduto As String
Dim sSiglaUM As String
Dim iAlmoxarifado As Integer
Dim lNumIntDoc As Long
Dim lNumIntDocOrigem As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim dCusto As Double

On Error GoTo Erro_RecebimentoMatFCom_Exclui_ItensNFiscal

    'Inicializa comando
    lComando3 = Comando_Abrir
    If lComando3 = 0 Then Error 54461

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 54532
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then Error 54724
    
    'Se o TipoMovtoEstoque for positivo
    If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
        'Cria objMovEstoque
        Set objMovEstoqueExcluir = New ClassMovEstoque
    
        'Preenche os atributos exceto código e colItens
        objMovEstoqueExcluir.iTipoMov = objTipoDocInfo.iTipoMovtoEstoque2
        objMovEstoqueExcluir.dtData = objNFiscal.dtDataEntrada
        objMovEstoqueExcluir.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoqueExcluir.lCliente = objNFiscal.lCliente
        objMovEstoqueExcluir.lFornecedor = objNFiscal.lFornecedor
        objMovEstoqueExcluir.sDocOrigem = objTipoDocInfo.sSigla
    
        sProduto = String(STRING_PRODUTO, 0)
        sSiglaUM = String(STRING_UM_SIGLA, 0)
    
        'Para cada Item na Tabela ItensNFiscal que ==>>
        'tem NumIntNF igual ao NumIntDoc do Recebimento selecionado
        lErro = Comando_ExecutarPos(lComando1, "SELECT NumIntDoc, Produto, Quantidade, UnidadeMed, Almoxarifado FROM ItensNFiscal WHERE NumIntNF = ?", 0, lNumIntDocOrigem, sProduto, dQuantidade, sSiglaUM, iAlmoxarifado, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 54457
    
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54458
    
        While lErro <> AD_SQL_SEM_DADOS
    
            lErro = Comando_Executar(lComando3, "SELECT NumIntDoc, Custo FROM MovimentoEstoque WHERE NumIntDocOrigem = ? AND TipoNumIntDocOrigem = ? ", lNumIntDoc, dCusto, lNumIntDocOrigem, TIPO_ORIGEM_ITEMNF)
            If lErro <> AD_SQL_SUCESSO Then Error 54462
    
            lErro = Comando_BuscarPrimeiro(lComando3)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54463
    
            If lErro = SUCESSO Then
            
                'Adiciona ítem em colItens de objMovEstoqueExcluir
                objMovEstoqueExcluir.colItens.Add lNumIntDoc, 0, dCusto, 0, sProduto, "", sSiglaUM, dQuantidade, iAlmoxarifado, "", lNumIntDocOrigem, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
    
            End If
    
            'Desfaz o link entre o item do Recebimento e os itens Pedido de compra
            lErro = CF("ItemNFItemPC_Exclui", lNumIntDocOrigem)
            If lErro <> SUCESSO Then Error 54454
    
            'Exclui ItenNFiscal do BD
            lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ItensNFiscal", lComando1)
            If lErro <> AD_SQL_SUCESSO Then Error 54460
    
            lErro = Comando_BuscarProximo(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54459
    
        Wend
    
        'Faz exclusão de MovEstoque Associados
        lErro = CF("MovEstoque_Grava", objMovEstoqueExcluir)
        If lErro <> SUCESSO Then Error 54536

    End If

    'Libera comando
    Call Comando_Fechar(lComando3)

    RecebimentoMatFCom_Exclui_ItensNFiscal = SUCESSO

    Exit Function

Erro_RecebimentoMatFCom_Exclui_ItensNFiscal:

    RecebimentoMatFCom_Exclui_ItensNFiscal = Err

    Select Case Err

        Case 54532, 54536, 54454

        Case 54724
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)

        Case 54457, 54458, 54459
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", Err)

        Case 54460
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSNFISCAL", Err, objNFiscal.lNumIntDoc)

        Case 54461
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54462, 54463
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE3", Err)

        Case 54464
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO", Err, lNumIntDocOrigem)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149546)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Function RecebMaterialFCom_Exclui(ByVal objNFiscal As ClassNFiscal) As Long
'Exclui o Recebimento de Material de Fornecedor de Compras

Dim lErro As Long
Dim iIndice As Integer
Dim iTipoDoc As Integer
Dim lNumIntDocOrigem As Long
Dim lNumIntDoc As Long
Dim lComando(5) As Long
Dim lTransacao As Long
Dim iStatus As Integer

On Error GoTo Erro_RecebMaterialFCom_Exclui

    'Abre os Comandos
    For iIndice = 0 To 4
        lComando(iIndice) = Comando_Abrir
        If lComando(iIndice) = 0 Then Error 54466
    Next

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 54467

    'Tenta selecionar o Recebimento
    lErro = Comando_ExecutarPos(lComando(0), "SELECT NumIntDoc, Status FROM NFiscal WHERE FilialEmpresa = ? AND NumRecebimento = ?", 0, lNumIntDoc, iStatus, objNFiscal.iFilialEmpresa, objNFiscal.lNumRecebimento)
    If lErro <> AD_SQL_SUCESSO Then Error 54468

    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54469

    'Se não achou o Recebimneto --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 54446

    'se o recebimento existe e já está baixado ==> não pode excluir
    If iStatus = STATUS_BAIXADO Then Error 54453
    
    'Faz o Lock exclusivo no Recebimento
    lErro = Comando_LockExclusive(lComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 54447

    objNFiscal.lNumIntDoc = lNumIntDoc

    'EXCLUSÃO DE ITENS DE RECEBIMENTO
    lErro = RecebimentoMatFCom_Exclui_ItensNFiscal(objNFiscal, lComando(2), lComando(3))
    If lErro <> SUCESSO Then Error 54449

    'EXCLUI RECEBIMENTO
    lErro = Comando_ExecutarPos(lComando(4), "DELETE FROM NFiscal", lComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 54451

    'Confirma Transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then Error 54452

    For iIndice = 0 To 4
    
        lErro = Comando_Fechar(lComando(iIndice))

    Next

    RecebMaterialFCom_Exclui = SUCESSO

    Exit Function

Erro_RecebMaterialFCom_Exclui:

    RecebMaterialFCom_Exclui = Err

    Select Case Err

        Case 54446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEB_MAT_FORN_NAO_CADASTRADO", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.dtDataEntrada)

        Case 54447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", Err, objNFiscal.lNumNotaFiscal)

        Case 54448
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RECEB_MAT_FORN_NF", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 54449, 54450

        Case 54451
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NFISCAL", Err, objNFiscal.lNumNotaFiscal)

        Case 54452
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 54453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RECEB_BAIXADO_FORN", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 54466
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 54468, 54469
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149547)

    End Select

    Call Transacao_Rollback

    For iIndice = 0 To 4
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function ItemNFItemPC_Exclui(ByVal lNumIntItemNFiscal As Long) As Long
'Exclui os Registros da tabela ItemNFItemPC relacionados a itens
'de Nota fiscal passados por lNumIntItemNFiscal

Dim alComando(0 To 3) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantAtual As Double
Dim dQuantidade As Double
Dim dQuantRecebimento As Double
Dim lItemPedCompra As Long

On Error GoTo Erro_ItemNFItemPC_Exclui

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 54470
    Next

    'Pesquisa no BD as ligações do ItemNF com itens pedidos de compras
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Quantidade, ItemPedCompra FROM ItemNFItemPC WHERE ItemNFiscal = ?", 0, dQuantidade, lItemPedCompra, lNumIntItemNFiscal)
    If lErro <> AD_SQL_SUCESSO Then Error 54471

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54472

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        'Lê os itens do pedido de compra
        lErro = Comando_ExecutarPos(alComando(1), "SELECT QuantRecebimento FROM ItensPedCompra WHERE NumIntDoc = ?", 0, dQuantRecebimento, lItemPedCompra)
        If lErro <> AD_SQL_SUCESSO Then Error 54474

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54475
        
        'Se não encontrou --> Erro
        If lErro = AD_SQL_SEM_DADOS Then Error 54703

        'QuantAtual = Quantidade Recebimento - Quantidade Recebida
        dQuantAtual = dQuantRecebimento - dQuantidade
        
        'Faz "LockExclusive" no Item Pedido de Compra
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 54476

        'Altera a quantidade de Recebimento do item pedido de compra
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensPedCompra SET QuantRecebimento = ?, Status = ?", alComando(1), dQuantAtual, ITEM_PED_COMPRAS_ABERTO)
        If lErro <> AD_SQL_SUCESSO Then Error 54477

        'Exclui o ItemNFItemPC
        lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM ItemNFItemPC", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 54478

        'Busca próxima ligação
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54473

    Loop

    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemNFItemPC_Exclui = SUCESSO

    Exit Function

Erro_ItemNFItemPC_Exclui:

    ItemNFItemPC_Exclui = Err

    Select Case Err

        Case 54470
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54471, 54472, 54473
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFITEMPC", Err)

        Case 54474, 54475
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", Err)

        Case 54476
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDCOMPRA", Err, lItemPedCompra)

        Case 54477
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDCOMPRA", Err, lItemPedCompra)

        Case 54478
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMNFITEMPC", Err, lNumIntItemNFiscal)

        Case 54703
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", Err, lItemPedCompra)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149548)

    End Select

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItemNFItemPC_Inclui_Recebimento(ByVal objItemNF As ClassItemNF) As Long
'Insere um registro na tabela ItemNFItemPC para cada item da coleção colItemNFItemPC passada por objItemNF

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim dQuantRecebimento As Double
Dim dQuantAtual As Double
Dim objItemNFItemPC As ClassItemNFItemPC
Dim dQuantidade As Double
Dim dQuantRecebida As Double
Dim iStatus As Integer

On Error GoTo Erro_ItemNFItemPC_Inclui_Recebimento

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 54673
    Next

    'Acumula as quantidades dos objItemNFItemPC
    For Each objItemNFItemPC In objItemNF.colItemNFItemPC

        'Insere na tabela ItemNFItemPC os dados passados por objItemPCInfo
        lErro = Comando_Executar(alComando(0), "INSERT INTO ItemNFItemPC (ItemNFiscal, ItemPedCompra, Quantidade, Taxa) VALUES (?,?,?,?)", objItemNF.lNumIntDoc, objItemNFItemPC.lItemPedCompra, objItemNFItemPC.dQuantidade, objItemNFItemPC.dTaxa)
        If lErro <> AD_SQL_SUCESSO Then Error 54674

        'Pesquisa no BD o item pedido de compras passado por objItemPCInfo
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Quantidade, QuantRecebida, QuantRecebimento FROM ItensPedCompra WHERE NumIntDoc = ?", 0, dQuantidade, dQuantRecebida, dQuantRecebimento, objItemNFItemPC.lItemPedCompra)
        If lErro <> AD_SQL_SUCESSO Then Error 54685

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54686

        'Se não encontrou --> Erro
        If lErro = AD_SQL_SEM_DADOS Then Error 54697

        'Faz "Lock Exclusive" no Item de Pedido de Compras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 54717

        dQuantAtual = dQuantRecebimento + objItemNFItemPC.dQuantidade
        
'        'Se a quantidade a receber do itemPC é igual a quantidade recebida
'        If dQuantidade = dQuantAtual + dQuantRecebida Then
'            iStatus = ITEM_PED_COMPRAS_RECEBIDO
'        Else
'            iStatus = ITEM_PED_COMPRAS_ABERTO
'        End If
        
        'Atualiza a Quantidade Recebida do Item pedido de compras em questão
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensPedCompra SET QuantRecebimento = ?", alComando(1), dQuantAtual)
        If lErro <> AD_SQL_SUCESSO Then Error 54687

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    ItemNFItemPC_Inclui_Recebimento = SUCESSO

    Exit Function

Erro_ItemNFItemPC_Inclui_Recebimento:

    ItemNFItemPC_Inclui_Recebimento = Err

    Select Case Err

        Case 54673
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54674
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMNFITEMPC", Err, objItemNFItemPC.lItemNFiscal, objItemNFItemPC.lItemPedCompra)

         Case 54685, 54686
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", Err)

        Case 54687
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDCOMPRA", Err, objItemNFItemPC.lItemPedCompra)

        Case 54697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", Err, objItemNFItemPC.lItemPedCompra)

        Case 54717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDCOMPRA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149549)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalEntradaCom_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil, ByVal colPedCompras As Collection, ByVal colReqCompras As Collection) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNotaFiscal As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NFiscalEntradaCom_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 65915

    Set objNFiscal.objContabil = objContabil

    lErro = CF("NFiscalEntradaCom_Grava_EmTrans", objNFiscal, objContabil, colPedCompras, colReqCompras)
    If lErro <> SUCESSO And lErro <> 207637 Then gError ERRO_SEM_MENSAGEM
        
    If lErro = SUCESSO Then
    
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 65925

    Else
    
        'Não conseguiu excluir, então altera só algumas coisas
        If lNumNotaFiscal = 0 Then gError ERRO_SEM_MENSAGEM
    
        'Dá a Mensagem de aviso que só será gravado a contabilidade
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
        If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM

        Call Transacao_Rollback

        lErro = CF("NFiscalEntrada_AlteraDados", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    lErro = CF("NFiscal_Trata_Eventos_WorkFlow", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    NFiscalEntradaCom_Grava = SUCESSO

    Exit Function

Erro_NFiscalEntradaCom_Grava:

    NFiscalEntradaCom_Grava = gErr

    Select Case gErr

        Case 65915
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 65925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149550)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalEntradaCom_Grava_BD(ByVal objNFiscal As ClassNFiscal, ByVal colPedCompras As Collection, ByVal colReqCompras As Collection) As Long
'Insere uma Nova Nota Fiscal de Entrada no Banco de Dados

Dim alComando(0 To 7) As Long
Dim lNumIntDoc As Long
Dim lErro As Long
Dim iClasseDocCPR As Integer
Dim lNumIntDocCPR As Long
Dim objItemNF As ClassItemNF
Dim objMovEstoque As ClassMovEstoque
Dim objMovEstoqueBaixa As ClassMovEstoque
Dim iIndice As Integer
Dim iTipoMovtoEstoque As Integer
Dim lCodigo As Long
Dim sDocOrigem As String
Dim iTipoMovtoEstoqueBenef As Integer, objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objRastroItemNF As ClassRastroItemNF
Dim objRastroMovto As ClassRastreamentoMovto
Dim colInsumos As Collection
Dim colRastroMovto As Collection
Dim dFator As Double
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dQuantItemNFUMEst As Double
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double

On Error GoTo Erro_NFiscalEntradaCom_Grava_BD

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 65966
    Next

    sDocOrigem = String(STRING_TIPO_DOC_INFO_SIGLA, 0)

    'Lê TipoMovtoEstoque em TiposDocInfo
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then gError 65967

    iTipoMovtoEstoque = objTipoDocInfo.iTipoMovtoEstoque
    sDocOrigem = objTipoDocInfo.sSigla

    'Gera Conta a Pagar e devolve iClasseDocCPR e lNumIntDocCPR
    lErro = CF("CPR_Gera", objNFiscal, iClasseDocCPR, lNumIntDocCPR)
    If lErro <> SUCESSO Then gError 65968

    objNFiscal.iClasseDocCPR = iClasseDocCPR
    objNFiscal.lNumIntDocCPR = lNumIntDocCPR

    'Obtem o Número Interno da nova Nota Fiscal
    lErro = CF("NFiscal_Automatico", lNumIntDoc)
    If lErro <> SUCESSO Then gError 65969

    objNFiscal.lNumIntDoc = lNumIntDoc
    
    lErro = CF("NFiscal_Grava_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError 65969

    With objNFiscal

        .iNFe = SerieEletronica(.sSerie)

        'Insere a nova Nota Fiscal de Entrada no BD
        lErro = Comando_Executar(alComando(0), "INSERT INTO NFiscal (DataAtualizacao, NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, FilialPedido, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,FornecedorBenef,FilialFornBenef,VolumeNumero,FreteRespons,DataReferencia,Observacao,NumRecebimento, HoraEntrada, Moeda, TaxaMoeda, NFe, ChvNFe, HoraEmissao, DataCadastro, HoraCadastro, UsuCadastro) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            gdtDataHoje, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .iFilialPedido, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lFornecedorBenef, .iFilialFornBenef, .sVolumeNumero, .iFreteRespons, .dtDataReferencia, .sObservacao, .lNumRecebimento, CDbl(.dtHoraEntrada), .iMoeda, .dTaxaMoeda, .iNFe, .sChvNFe, _
            CDbl(IIf(.dtDataSaida <> DATA_NULA, IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraSaida, .dtHoraEmissao, .dtHoraSaida), IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraEntrada, .dtHoraEmissao, .dtHoraEntrada))), gdtDataHoje, CDbl(Time), gsUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 65970
    End With

    lErro = CF("NFiscal_GravaTributacao", objNFiscal, alComando(4), alComando(5))
    If lErro <> SUCESSO Then gError 65971

    'Para cada Ítem da Nota
    For Each objItemNF In objNFiscal.colItensNF

        objItemNF.lNumIntNF = objNFiscal.lNumIntDoc

        'Obtém o Número Interno do Ítem
        lErro = CF("NFiscalItem_Automatico", lNumIntDoc)
        If lErro <> SUCESSO Then gError 65972

        objItemNF.lNumIntDoc = lNumIntDoc

        With objItemNF
            'Insere o Ítem no Banco de dados
            lErro = Comando_Executar(alComando(1), "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, NumIntDocOrig,Almoxarifado, Ccl, PrecoUnitarioMoeda) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lNumIntNF, .iItem, STATUS_LANCADO, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, lNumIntDoc, .lNumIntTrib, 0, .iAlmoxarifado, .sCcl, .dPrecoUnitarioMoeda)
            If lErro <> AD_SQL_SUCESSO Then gError 65973
        End With

        lErro = CF("NFiscal_GravaTributacaoItem", objItemNF, alComando(3), alComando(6))
        If lErro <> SUCESSO Then gError 65974
        
        objItemNF.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
        objItemNF.objInfoAdicDocItem.lNumIntDocItem = objItemNF.lNumIntDoc
        lErro = CF("InfoAdicDocItem_Insere", objItemNF.objInfoAdicDocItem, alComando(7))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Next
    
    lErro = CF("NFe_Grava_Vinculo_Xml", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("NFiscalItens_Grava_Customizado", objNFiscal)
    If lErro <> SUCESSO Then gError 181267

    'Inclui em ItemNFItemPC o ItemNF associado a ItemPC
    lErro = CF("ItemNFItemPC_Inclui", objNFiscal)
    If lErro <> SUCESSO Then gError 65975
        
    'Atualiza as Quantidades dos ItensPC
    lErro = CF("ItensPedCompra_Atualiza_Quantidades", objNFiscal, colPedCompras)
    If lErro <> SUCESSO Then gError 66076
        
    'Inclui em ItemNFItemRC o ItemNF associado a ItemRC
    lErro = CF("ItemNFItemRC_Inclui", objNFiscal)
    If lErro <> SUCESSO Then gError 65976
        
    'Atualiza as quantidades dos Itens de Requisição
    lErro = CF("ItensReqCompra_Atualiza_Quantidades", objNFiscal)
    If lErro <> SUCESSO Then gError 66077
    
    'Faz a baixa de um recebimento
    'a baixa do recebimento deve ocorrer antes da gravacao da nota fiscal para que os valores dos saldos estejam corretos e permitam calcular o novo custo do produto, se for o caso.
    If objNFiscal.lNumRecebimento > 0 Then

        'Lê o Recebimento e Verifica os Dados
        lErro = NFiscal_Valida_Recebimento(objNFiscal)
        If lErro <> SUCESSO Then gError 66118
   
        'Faz a baixa de um recebimento
        'a baixa do recebimento deve ocorrer antes da gravacao da nota fiscal para que os valores dos saldos estejam corretos e permitam calcular o novo custo do produto, se for o caso.
        lErro = CF("NFiscalEntradaCom_Baixa_Recebimento", objNFiscal)
        If lErro <> SUCESSO Then gError 65977
    
    End If
    
    If iTipoMovtoEstoque > 0 Then

        Set objMovEstoque = New ClassMovEstoque

        'Obtém o Código do Movimento
        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then gError 65978

        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.iTipoMov = iTipoMovtoEstoque
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = sDocOrigem & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)

        'Se for para incluir o valor do frete e outras depesas no custo do ItemNF
        If gobjEST.iIncluiFreteNoCusto = INCLUI_FRETE_CUSTO Then
            Call Inclui_Frete_CustoItemNF(objNFiscal)
        End If


'mario distribuicao substituicao do codigo original acrescentando o tab de distribuicao
                        
        'Adiciona itens ao Movimento
        iIndice = 0
        For Each objItemNF In objNFiscal.colItensNF
            
            iIndice = iIndice + 1
            
            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                If objItemNF.ColAlocacoes.Count = 0 Then gError 89552

                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 89551

                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))

                'calcula o custo do item da NF em função do credito do ICMS e do IPI
                lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
                If lErro <> SUCESSO Then gError 65979

                dSomaAlocacao = 0
                                                                
                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
                    
                    Set colInsumos = New Collection
                    Set colRastroMovto = New Collection
                                        
                    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                    
                    'Guarda o Rastreamento dos ItensNF
                    If Not (objNFiscal.objRastreamento Is Nothing) Then
                        lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
                        If lErro <> SUCESSO Then gError 89550
                    End If
                    
'mario
'                    If Not (objItemNF.colRastreamento Is Nothing) Then
'
'                        'Guarda o Rastreamento dos ItensNF
'                        For Each objRastroItemNF In objItemNF.colRastreamento
'
'                            If objRastroItemNF.iAlmoxCodigo = objItemNFAloc.iAlmoxarifado Then
'
'                                Set objRastroMovto = New ClassRastreamentoMovto
'                                objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
'                                objRastroMovto.sLote = objRastroItemNF.sLote
'                                objRastroMovto.sProduto = objItemNF.sProduto
'
'                                'Adiciona objRastroMovto na coleção de Rastreamento
'                                objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
'
'                            End If
'
'                        Next
'
'                    End If
'fim mario
                    
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                    objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                    objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                    objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
                    
                    'se for uma nota de complemento de preço ou outra que nao deve afetar qtde ==> zera a quantidade pois somente o custo será utilizado
'                    If (objNFiscal.iTipoNFiscal = DOCINFO_NFEECPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIECPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEIPICO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEICMCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEICMCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEIPICO) Then
                    If TipoDocInfo_ZeraQtdeMovto(objNFiscal.iTipoNFiscal) Then
                        objItemMovEstoque.dQuantidade = 0
                    End If
                    
                Next
                
                '#############################################################
                'Alterado por Wagner 16/11/04
                'Calcula o número de casas decimais do Formato de Estoque
'                iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                dAcrescimo = 0
'                If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                If Abs(dSomaAlocacao - dQuantItemNFUMEst) > dAcrescimo + QTDE_ESTOQUE_DELTA Then
                
                dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)

                
                If Abs(dSomaAlocacao - dQuantItemNFUMEst) > QTDE_ESTOQUE_DELTA Then
                '#############################################################
                
                    If dSomaAlocacao < dQuantItemNFUMEst Then gError 89553
                    
                    If dSomaAlocacao > dQuantItemNFUMEst Then gError 89554
            
                End If
                
            End If
        Next
                        
                        
'mario distribuicao fim da alteração para incluir o tab de distribuicao


'        'Adiciona os ítem ao Movimento
'        For Each objItemNF In objNFiscal.colItensNF
'
'            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
'
'                'calcula o custo do item da NF em função do credito do ICMS e do IPI
'                lErro = CF("NFiscalEntrada_Calcula_Custo",objItemNF)
'                If lErro <> SUCESSO Then gError 65979
'
'                Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNF.sUnidadeMed, objItemNF.dQuantidade, objItemNF.iAlmoxarifado, objItemNF.sAlmoxarifadoNomeRed, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
'                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'
'                Set objItemMovEstoque.colRastreamentoMovto = New Collection
'                Set objItemMovEstoque.colApropriacaoInsumo = New Collection
'
'                'Guarda o Rastreamento dos ItensNF
'                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
'                If lErro <> SUCESSO Then gError 83571
'
''                'Guarda o Rastreamento dos ItensNF
''                For Each objRastroItemNF In objItemNF.colRastreamento
''
''                    Set objRastroMovto = New ClassRastreamentoMovto
''                    objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
''                    objRastroMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
''                    objRastroMovto.sLote = objRastroItemNF.sLote
''                    objRastroMovto.lNumIntDocOrigem = objItemNF.lNumIntDoc
''                    objRastroMovto.sProduto = objItemNF.sProduto
''                    objRastroMovto.iFilialOP = objRastroItemNF.iLoteFilialOP
''
''                    'Adiciona objRastroMovto na coleção de Rastreamento
''                    objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
''
''                Next
'
'                objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
'                objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
'                objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
'                objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
'
'            End If
'
'        Next

        'Gera um Movimento de Estoque
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then gError 65980

        lErro = CF("NFiscalEntCom_Gera_Reservas", objNFiscal, objMovEstoque)
        If lErro <> SUCESSO Then gError 65983
    
    End If

    'Verifica se existem Pedidos a serem Baixados
    lErro = CF("NFiscalEntCom_BaixaPedidos", colPedCompras, objNFiscal)
    If lErro <> SUCESSO Then gError 65982
        
    'Verifica se existem Requisições a serem Baixadas
    lErro = CF("NFiscalEntCom_BaixaRequisicoes", colReqCompras, objNFiscal)
    If lErro <> SUCESSO Then gError 65983
        
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalEntradaCom_Grava_BD = SUCESSO

    Exit Function

Erro_NFiscalEntradaCom_Grava_BD:

    NFiscalEntradaCom_Grava_BD = gErr

    Select Case gErr

        Case 65966
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65967, 65968, 65969, 65971, 65972, 65974, 65975, 65976, 65977, 65978, 65979, 65980, 65982, 65983, 66076, 66077, 66118, 83571, 89550, 89551, 181267
        
        Case 65970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 65973
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 89552
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INEXISTENTE", gErr, iIndice)

        Case 89553
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INCOMPLETA", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case 89554
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149551)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItemNFItemPC_Inclui(ByVal objNFiscal As ClassNFiscal) As Long
'Insere um registro na tabela ItemNFItemPC para cada item da coleção colItemNFItemPC passada por objItemNF

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objItemNFItemPC As ClassItemNFItemPC
Dim objItemNF As ClassItemNF

On Error GoTo Erro_ItemNFItemPC_Inclui

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 65984
    Next

    'Para cada Item da Nota Fiscal
    For Each objItemNF In objNFiscal.colItensNF
    
        'Para cada objItemNFItemPC da coleção colItemNFItemPC do Item da Nota Fiscal
        For Each objItemNFItemPC In objItemNF.colItemNFItemPC
    
            'Insere na tabela ItemNFItemPC os dados passados por objItemNFItemPC
            lErro = Comando_Executar(alComando(0), "INSERT INTO ItemNFItemPC (ItemNFiscal, ItemPedCompra, Quantidade, Taxa) VALUES (?,?,?,?)", objItemNF.lNumIntDoc, objItemNFItemPC.lItemPedCompra, objItemNFItemPC.dQuantidade, objItemNFItemPC.dTaxa)
            If lErro <> AD_SQL_SUCESSO Then gError 65985
            
        Next
                
    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    ItemNFItemPC_Inclui = SUCESSO

    Exit Function

Erro_ItemNFItemPC_Inclui:

    ItemNFItemPC_Inclui = gErr

    Select Case gErr

        Case 65984
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65985
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMNFITEMPC", gErr, objItemNF.lNumIntDoc, objItemNFItemPC.lItemPedCompra)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149552)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensPedCompra_Atualiza_Quantidades(ByVal objNFiscal As ClassNFiscal, ByVal colPedCompras As Collection) As Long
'Atualiza a Quantidade Recebida

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim dQuantRecebida As Double
Dim dQuantidade As Double
Dim dQuantRecebimento As Double
Dim objItemNFItemPC As ClassItemNFItemPC
Dim dQuantRecebidaAtual As Double
Dim dQuantRecebimentoAtual As Double
Dim iStatus As Integer
Dim objItemNF As ClassItemNF
Dim objItemNFItemRC As ClassItemNFItemRC
Dim dtDataLimite As Date
Dim objPedidoCompra As ClassPedidoCompras
Dim lNumIntDocPC As Long

On Error GoTo Erro_ItensPedCompra_Atualiza_Quantidades

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 65874
    Next
    
    'Para cada Item da Nota Fiscal
    For Each objItemNF In objNFiscal.colItensNF
    
        'Para cada objItemNFItemPC da coleção colItemNFItemPC do Item da Nota Fiscal
        For Each objItemNFItemPC In objItemNF.colItemNFItemPC
        
            'Pesquisa no BD o item pedido de compras passado por objItemNFItemPC
            lErro = Comando_ExecutarPos(alComando(0), "SELECT Quantidade, QuantRecebida, QuantRecebimento, DataLimite, PedCompra FROM ItensPedCompra WHERE NumIntDoc = ?", 0, dQuantidade, dQuantRecebida, dQuantRecebimento, dtDataLimite, lNumIntDocPC, objItemNFItemPC.lItemPedCompra)
            If lErro <> AD_SQL_SUCESSO Then gError 65986
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65987
    
            'Se não encontrou --> Erro
            If lErro = AD_SQL_SEM_DADOS Then gError 65988
    
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> AD_SQL_SUCESSO Then gError 65989
    
            dQuantRecebidaAtual = dQuantRecebida + objItemNFItemPC.dQuantidade
            
            'Se é um Recebimento
            If objNFiscal.lNumRecebimento > 0 Then
                dQuantRecebimentoAtual = dQuantRecebimento - objItemNFItemPC.dQuantidade
            Else
                dQuantRecebimentoAtual = dQuantRecebimento
            End If
            
            If dQuantRecebidaAtual >= dQuantidade Then
                iStatus = ITEM_PED_COMPRAS_RECEBIDO
            Else
                iStatus = ITEM_PED_COMPRAS_ABERTO
            End If
            
            'Atualiza a Quantidade Recebida do Item pedido de compras em questão
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensPedCompra SET QuantRecebida = ?, QuantRecebimento = ?, Status = ?", alComando(0), dQuantRecebidaAtual, dQuantRecebimentoAtual, iStatus)
            If lErro <> AD_SQL_SUCESSO Then gError 65990
    
            If iStatus = ITEM_PED_COMPRAS_RECEBIDO Then
    
                For Each objPedidoCompra In colPedCompras
                    If objPedidoCompra.lNumIntDoc = lNumIntDocPC Then
                    
                        'guarda o numero de itens de pedido de compras que tiveram seu recebimento completado
                        objPedidoCompra.lItensPCRecebidos = objPedidoCompra.lItensPCRecebidos + 1
                        
                        'se houve atraso ==> armazena em lAtrasoAcumulado
                        If (objNFiscal.dtDataEntrada - dtDataLimite) > 0 And dtDataLimite <> DATA_NULA Then
                            objPedidoCompra.lAtrasoAcumulado = objPedidoCompra.lAtrasoAcumulado + (objNFiscal.dtDataEntrada - dtDataLimite)
                        End If
                        
                    End If
                Next
    
            End If
    
        Next
    
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    ItensPedCompra_Atualiza_Quantidades = SUCESSO

    Exit Function

Erro_ItensPedCompra_Atualiza_Quantidades:

    ItensPedCompra_Atualiza_Quantidades = gErr

    Select Case gErr

        Case 65874
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

         Case 65986, 65987
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)

        Case 65990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDCOMPRA", gErr, objItemNFItemPC.lItemPedCompra)

        Case 65988
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", gErr, objItemNFItemPC.lItemPedCompra)

        Case 65989
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSREQCOMPRA1", gErr, objItemNFItemPC.lItemPedCompra)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149553)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

End Function

Function ItemNFItemRC_Inclui(ByVal objNFiscal As ClassNFiscal) As Long
'Insere um registro na tabela ItemNFItemPC para cada item da coleção colItemNFItemPC passada por objItemNF

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objItemNFItemRC As ClassItemNFItemRC
Dim objItemNF As ClassItemNF
Dim dQuantidade As Double
Dim colTrabalho As New Collection
Dim lReq As Long
Dim iIndice2 As Integer

On Error GoTo Erro_ItemNFItemRC_Inclui

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 65991
    Next

    'Para cada Item da Nota Fiscal
    For Each objItemNF In objNFiscal.colItensNF
        Set colTrabalho = New Collection
        For Each objItemNFItemRC In objItemNF.colItemNFItemRC
            colTrabalho.Add objItemNFItemRC
        Next
    
        'dQuantidade = 0
        iIndice = 1
        Do While iIndice <= colTrabalho.Count
            
            Set objItemNFItemRC = colTrabalho.Item(iIndice)
            
            dQuantidade = 0
            For iIndice2 = colTrabalho.Count To iIndice Step -1
                
                If colTrabalho(iIndice2).lItemReqCompra = objItemNFItemRC.lItemReqCompra Then
                    dQuantidade = dQuantidade + colTrabalho(iIndice2).dQuantidade
                    If iIndice <> iIndice2 Then colTrabalho.Remove iIndice2
                End If
            Next
            
            'Insere na tabela ItemNFItemRC os dados passados por objItemNFItemRC
            lErro = Comando_Executar(alComando(0), "INSERT INTO ItemNFItemRC (ItemNFiscal, ItemReqCompra, Quantidade) VALUES (?,?,?)", objItemNF.lNumIntDoc, objItemNFItemRC.lItemReqCompra, dQuantidade)
            If lErro <> AD_SQL_SUCESSO Then gError 65992
            
            iIndice = iIndice + 1
        Loop
        
    Next
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    ItemNFItemRC_Inclui = SUCESSO

    Exit Function

Erro_ItemNFItemRC_Inclui:

    ItemNFItemRC_Inclui = gErr

    Select Case gErr

        Case 65991
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMNFITEMRC", gErr, objItemNF.lNumIntDoc, objItemNFItemRC.lItemReqCompra)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149554)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensReqCompra_Atualiza_Quantidades(ByVal objNFiscal As ClassNFiscal) As Long
'Atualiza Quantidades dos Itens de Requisição de Compras

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim dQuantPedida As Double
Dim dQuantRecebida As Double
Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim objItemNFItemRC As ClassItemNFItemRC
Dim dQuantPedidaAtual As Double
Dim dQuantRecebidaAtual As Double
Dim iStatus As Integer
Dim lNumIntItem As Long
Dim objItemNF As ClassItemNF

On Error GoTo Erro_ItensReqCompra_Atualiza_Quantidades

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 65875
    Next
    
    'Para cada Item da Nota Fiscal
    For Each objItemNF In objNFiscal.colItensNF
    
        'Para cada objItemNFItemRC da coleção colItemNFItemRC do Item da Nota Fiscal
        For Each objItemNFItemRC In objItemNF.colItemNFItemRC
    
            'Pesquisa no BD o Item de Requisição de Compras passado por objItemNFItemPC
            lErro = Comando_ExecutarPos(alComando(0), "SELECT Quantidade, QuantPedida, QuantRecebida, QuantCancelada FROM ItensReqCompra WHERE NumIntDoc = ?", 0, dQuantidade, dQuantPedida, dQuantRecebida, dQuantCancelada, objItemNFItemRC.lItemReqCompra)
            If lErro <> AD_SQL_SUCESSO Then gError 65993
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65994
    
            'Se não encontrou --> Erro
            If lErro = AD_SQL_SEM_DADOS Then gError 65995
    
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> AD_SQL_SUCESSO Then gError 65996
    
            dQuantPedidaAtual = dQuantPedida - objItemNFItemRC.dQuantidade
            dQuantRecebidaAtual = dQuantRecebida + objItemNFItemRC.dQuantidade
            
            If dQuantidade - dQuantCancelada <= dQuantRecebidaAtual Then
                iStatus = ITEM_REQ_RECEBIDO
            Else
                iStatus = ITEM_REQ_ABERTO
            End If
            If dQuantPedidaAtual < 0 Then dQuantPedidaAtual = 0
            
            'Atualiza a Quantidade Recebida do Item de RC ITem PC
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensReqCompra SET QuantPedida = ?, QuantRecebida = ?, Status = ?", alComando(0), dQuantPedidaAtual, dQuantRecebidaAtual, iStatus)
            If lErro <> AD_SQL_SUCESSO Then gError 65997
    
            'Pesquisa no BD o Item RC Item PC passado por objItemNFItemPC
            lErro = Comando_ExecutarPos(alComando(2), "SELECT QuantRecebida FROM ItemRCItemPC WHERE ItemPC = ? AND ItemRC = ?", 0, dQuantRecebida, objItemNFItemRC.lNumIntDocItemPC, objItemNFItemRC.lItemReqCompra)
            If lErro <> AD_SQL_SUCESSO Then gError 89261
    
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89262
    
            'Se não encontrou --> Erro
            If lErro = AD_SQL_SEM_DADOS Then gError 89263
    
            lErro = Comando_LockExclusive(alComando(2))
            If lErro <> AD_SQL_SUCESSO Then gError 89264
    
            'Atualiza a Quantidade Recebida do Item RC Item PC
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItemRCItemPC SET QuantRecebida = QuantRecebida + ?", alComando(2), objItemNFItemRC.dQuantidade)
            If lErro <> AD_SQL_SUCESSO Then gError 89265
    
        Next
        
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    ItensReqCompra_Atualiza_Quantidades = SUCESSO

    Exit Function

Erro_ItensReqCompra_Atualiza_Quantidades:

    ItensReqCompra_Atualiza_Quantidades = gErr

    Select Case gErr

        Case 65875
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

         Case 65993, 65994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)

        Case 65995
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMREQCOMPRA_INEXISTENTE", gErr, objItemNFItemRC.lItemReqCompra)

        Case 65996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSREQCOMPRA1", gErr, objItemNFItemRC.lItemReqCompra)
        
        Case 65997
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSREQCOMPRA", gErr, objItemNFItemRC.lItemReqCompra)
        
         Case 89261, 89262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMRCITEMPC", gErr, objItemNFItemRC.lNumIntDocItemPC, objItemNFItemRC.lItemReqCompra)

        Case 89263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMRCITEMPC_INEXISTENTE", gErr, objItemNFItemRC.lNumIntDocItemPC, objItemNFItemRC.lItemReqCompra)

        Case 89264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMRCITEMPC", gErr, objItemNFItemRC.lNumIntDocItemPC, objItemNFItemRC.lItemReqCompra)
        
        Case 89265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITEMRCITEMPC", gErr, objItemNFItemRC.lNumIntDocItemPC, objItemNFItemRC.lItemReqCompra)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149555)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalEntCom_BaixaPedidos(ByVal colPedCompras As Collection, ByVal objNFiscal As ClassNFiscal) As Long
'Baixa os Pedidos de Compras

Dim lErro As Long
Dim alComando(0 To 8) As Long
Dim iIndice As Integer
Dim lNumIntPed As Long
Dim lNumIntItem As Long
Dim iStatus As Integer
Dim tItemPedido As typeItemPedCompra
Dim tPedCompra As typePedidoCompras
Dim objPedidoCompra As ClassPedidoCompras
Dim objItemNF As ClassItemNF
Dim objItemNFItemPC As ClassItemNFItemPC
Dim iStatusBaixa As Integer

On Error GoTo Erro_NFiscalEntCom_BaixaPedidos
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66013
    Next
    
    'Para cada Pedido de Compras
    For Each objPedidoCompra In colPedCompras
        
        'Lê os itens do Pedido de Compras que possuem STATUS_ABERTO
        lErro = Comando_Executar(alComando(0), "SELECT ItensPedCompra.Status FROM PedidoCompra, ItensPedCompra WHERE PedidoCompra.Codigo = ? AND PedidoCompra.FilialEmpresa = ? AND PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND ItensPedCompra.Status = ?", iStatus, objPedidoCompra.lCodigo, objNFiscal.iFilialPedido, ITEM_PED_COMPRAS_ABERTO)
        If lErro <> AD_SQL_SUCESSO Then gError 66014
    
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66015
            
        'Se não encontrar
        If lErro = AD_SQL_SEM_DADOS Then
        
            'Lê os itens de Pedidos de Compras relacionados ao pedido da coleção colPedCompras
            lErro = Comando_Executar(alComando(1), "SELECT DISTINCT PedidoCompra.NumIntDoc, ItensPedCompra.NumIntDoc FROM PedidoCompra, ItensPedCompra WHERE PedidoCompra.Codigo = ? AND PedidoCompra.FilialEmpresa = ? AND PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra", lNumIntPed, lNumIntItem, objPedidoCompra.lCodigo, objNFiscal.iFilialPedido)
            If lErro <> AD_SQL_SUCESSO Then gError 66016
                                
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66017
            
            'Enquanto encontrar itens
            Do While lErro <> AD_SQL_SEM_DADOS
                        
                tItemPedido.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
                tItemPedido.sProduto = String(STRING_PRODUTO, 0)
                tItemPedido.sObservacao = String(STRING_OBSERVACAO, 0)
                tItemPedido.sUM = String(STRING_UM_SIGLA, 0)
                
                'Lê os dados do Item de Pedido de Compras
                lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc,PedCompra,DataLimite,Produto,DescProduto,Quantidade,QuantRecebida,QuantRecebimento,UM,PrecoUnitario,ValorDesconto,TipoOrigem,NumIntOrigem,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa,Status,Observacao,ValorIPI,AliquotaIPI,AliquotaICMS FROM ItensPedCompraN WHERE NumIntDoc = ? AND PedCompra = ? AND StatusBaixa = ?", 0, tItemPedido.lNumIntDoc, tItemPedido.lPedCompra, _
                    tItemPedido.dtDataLimite, tItemPedido.sProduto, tItemPedido.sDescProduto, tItemPedido.dQuantidade, tItemPedido.dQuantRecebida, tItemPedido.dQuantRecebimento, tItemPedido.sUM, tItemPedido.dPrecoUnitario, tItemPedido.dValorDesconto, tItemPedido.iTipoOrigem, tItemPedido.lNumIntOrigem, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRebebForaFaixa, tItemPedido.iStatus, tItemPedido.lObservacao, tItemPedido.dValorIPI, tItemPedido.dAliquotaIPI, tItemPedido.dAliquotaICMS, lNumIntItem, lNumIntPed, ITENSPEDCOMPRA_STATUS_ATIVO)
                If lErro <> AD_SQL_SUCESSO Then gError 66018

                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66019
                
                'Se não encontrar, erro
                If lErro = AD_SQL_SEM_DADOS Then gError 66020
            
                'Exclui o Item de Pedido de Compras
                lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensPedCompraN SET StatusBaixa=?", alComando(2), ITENSPEDCOMPRA_STATUS_BAIXADO)
                If lErro <> AD_SQL_SUCESSO Then gError 66021
                
                'Busca Próximo Item
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66023
            
            Loop
            
            
            tPedCompra.sContato = String(STRING_CONTATO, 0)
            tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
            tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
            tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)

            'Lê os dados do Pedido de Compras
            lErro = Comando_ExecutarPos(alComando(5), "SELECT NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,Observacao, Status FROM PedidoCompraN WHERE Codigo = ? AND FilialEmpresa=?", 0, tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, _
                tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.lObservacao, iStatusBaixa, objPedidoCompra.lCodigo, objNFiscal.iFilialPedido)
            If lErro <> AD_SQL_SUCESSO Then gError 66024
    
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66025
    
            'se o pedido não estiver cadastrado
            If lErro = AD_SQL_SEM_DADOS Then gError 66026
            
            'se o pedido já estiver baixado
            If iStatusBaixa = PEDIDOCOMPRA_STATUS_BAIXADO Then gError 89220
    
            'Atualiza o Pedido de Compras do BD
            lErro = Comando_ExecutarPos(alComando(6), "UPDATE PedidoCompraN SET Status = ?, TipoBaixa = ?, MotivoBaixa = ?, DataBaixa = ?", alComando(5), PEDIDOCOMPRA_STATUS_BAIXADO, TIPO_BAIXA_NORMAL, tPedCompra.sMotivoBaixa, objNFiscal.dtDataEntrada)
            If lErro <> AD_SQL_SUCESSO Then gError 66027
            
        End If
        
    Next
    
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscalEntCom_BaixaPedidos = SUCESSO
    
    Exit Function
    
Erro_NFiscalEntCom_BaixaPedidos:

    NFiscalEntCom_BaixaPedidos = gErr
    
    Select Case gErr
    
        Case 66013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 66014, 66015, 66024, 66025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
        
        Case 66016, 66017, 66018, 66019, 66023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)
        
        Case 66020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", gErr, lNumIntItem)
        
        Case 66021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDCOMPRA", gErr, lNumIntItem)
            
        Case 66022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPEDCOMPRA", gErr, lNumIntItem)
            
        Case 66026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)
        
        Case 66027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOCOMPRA", gErr, tPedCompra.lCodigo)
        
        Case 66028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PEDIDOCOMPRABAIXADO", gErr, tPedCompra.lCodigo)
        
        Case 89218, 89219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRABAIXADO", gErr)
            
        Case 89220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BAIXADO", gErr, objPedidoCompra.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149556)
    
    End Select
    
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function NFiscalEntCom_BaixaRequisicoes(ByVal colReqCompras As Collection, ByVal objNFiscal As ClassNFiscal) As Long
'Baixa as Requisições de Compras

Dim lErro As Long
Dim alComando(0 To 7) As Long
Dim iIndice As Integer
Dim lNumIntReq As Long
Dim lNumIntItem As Long
Dim iStatus As Integer
Dim tItemReqCompras As typeItemReqCompra
Dim tRequisicaoCompras As typeRequisicaoCompras

On Error GoTo Erro_NFiscalEntCom_BaixaRequisicoes
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66029
    Next
    
    'Para cada Requisição de Compras
    For iIndice = 1 To colReqCompras.Count
        
        'Lê os itens de Requisição de Compras que possuem STATUS_ABERTO ou STATUS_PEDIDO
        lErro = Comando_Executar(alComando(0), "SELECT ItensReqCompra.Status FROM RequisicaoCompra, ItensReqCompra WHERE RequisicaoCompra.Codigo = ? AND RequisicaoCompra.FilialCompra = ? AND RequisicaoCompra.NumIntDoc = ItensReqCompra.ReqCompra AND (ItensReqCompra.Status = ? OR ItensReqCompra.Status = ?)", iStatus, colReqCompras.Item(iIndice), objNFiscal.iFilialPedido, ITEM_REQ_ABERTO, ITEM_REQ_PEDIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 66030
    
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66031
            
        'Se não encontrar
        If lErro = AD_SQL_SEM_DADOS Then
        
            'Lê os itens de Requisição de Compras relacionados as Requisições da coleção colReqCompras
            lErro = Comando_Executar(alComando(1), "SELECT RequisicaoCompra.NumIntDoc, ItensReqCompra.NumIntDoc FROM RequisicaoCompra, ItensReqCompra WHERE RequisicaoCompra.Codigo = ? AND RequisicaoCompra.FilialCompra = ? AND RequisicaoCompra.NumIntDoc = ItensReqCompra.ReqCompra", lNumIntReq, lNumIntItem, colReqCompras.Item(iIndice), objNFiscal.iFilialPedido)
            If lErro <> AD_SQL_SUCESSO Then gError 66033
                                
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66034
            
            'Enquanto encontrar itens
            Do While lErro <> AD_SQL_SEM_DADOS
                        
                tItemReqCompras.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
                tItemReqCompras.sProduto = String(STRING_PRODUTO, 0)
                tItemReqCompras.sUM = String(STRING_UM_SIGLA, 0)
                tItemReqCompras.sCcl = String(STRING_CCL, 0)
                tItemReqCompras.sContaContabil = String(STRING_CONTA, 0)
                
                'Lê o item da Requisição na tabela ItensReqCompras
                lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc, ReqCompra, Produto, DescProduto, Status, Quantidade, QuantPedida, QuantRecebida, QuantCancelada, UM, Ccl, Almoxarifado, ContaContabil, TipoTributacao, Observacao, Fornecedor, Filial, Exclusivo FROM ItensReqCompraN WHERE NumIntDoc = ? AND StatusBaixa = ?", 0, _
                    tItemReqCompras.lNumIntDoc, tItemReqCompras.lReqCompra, tItemReqCompras.sProduto, tItemReqCompras.sDescProduto, tItemReqCompras.iStatus, tItemReqCompras.dQuantidade, tItemReqCompras.dQuantPedida, tItemReqCompras.dQuantRecebida, tItemReqCompras.dQuantCancelada, tItemReqCompras.sUM, tItemReqCompras.sCcl, tItemReqCompras.iAlmoxarifado, tItemReqCompras.sContaContabil, tItemReqCompras.iTipoTributacao, tItemReqCompras.lObservacao, tItemReqCompras.lFornecedor, tItemReqCompras.iFilial, tItemReqCompras.iExclusivo, lNumIntItem, ITENSREQCOMPRA_STATUS_ATIVO)
                If lErro <> AD_SQL_SUCESSO Then gError 66035

                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66036
                
                'Se não encontrar, erro
                If lErro = AD_SQL_SEM_DADOS Then gError 66037
            
                'Atualiza o Item de Requisição de Compras
                lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensReqCompraN SET StatusBaixa = ?", alComando(2), ITENSREQCOMPRA_STATUS_BAIXADO)
                If lErro <> AD_SQL_SUCESSO Then gError 66038
                
                'Busca Próximo Item
                lErro = Comando_BuscarProximo(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66043
            
            Loop
            
            tRequisicaoCompras.sCcl = String(STRING_CCL, 0)
            tRequisicaoCompras.sOPCodigo = String(STRING_OPCODIGO, 0)
            tRequisicaoCompras.sDigitador = String(STRING_USUARIO_CODIGO, 0)
            
            'Lê os dados da Requisição de Compras
            lErro = Comando_ExecutarPos(alComando(5), "SELECT NumIntDoc, FilialEmpresa, Codigo, Data, DataEnvio, DataLimite, Urgente, Requisitante, Digitador, Ccl, OPCodigo, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao FROM RequisicaoCompraN WHERE NumIntDoc = ? AND Status = ?", 0, _
                tRequisicaoCompras.lNumIntDoc, tRequisicaoCompras.iFilialEmpresa, tRequisicaoCompras.lCodigo, tRequisicaoCompras.dtData, tRequisicaoCompras.dtDataEnvio, tRequisicaoCompras.dtDataLimite, tRequisicaoCompras.lUrgente, tRequisicaoCompras.lRequisitante, tRequisicaoCompras.sDigitador, tRequisicaoCompras.sCcl, tRequisicaoCompras.sOPCodigo, tRequisicaoCompras.iFilialCompra, tRequisicaoCompras.iTipoDestino, tRequisicaoCompras.lFornCliDestino, tRequisicaoCompras.iFilialDestino, tRequisicaoCompras.lObservacao, tRequisicaoCompras.iTipoTributacao, lNumIntReq, REQUISICAOCOMPRA_STATUS_ATIVO)
            If lErro <> AD_SQL_SUCESSO Then gError 66032
    
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66044
    
            'Se não encontrar, erro
            If lErro = AD_SQL_SEM_DADOS Then gError 66040
            
            objNFiscal.colReqCompraBaixadas.Add lNumIntReq
    
            'Exclui o Requisição de Compras do BD
            lErro = Comando_ExecutarPos(alComando(6), "UPDATE RequisicaoCompraN SET Status = ?, DataBaixa=?, DataRegBaixa = ?, HoraRegBaixa = ?, UsuRegBaixa = ? ", alComando(5), REQUISICAOCOMPRA_STATUS_BAIXADO, objNFiscal.dtDataEntrada, Date, CDbl(Time), gsUsuario)
            If lErro <> AD_SQL_SUCESSO Then gError 66041
            
        End If
        
    Next
    
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscalEntCom_BaixaRequisicoes = SUCESSO
    
    Exit Function
    
Erro_NFiscalEntCom_BaixaRequisicoes:

    NFiscalEntCom_BaixaRequisicoes = gErr
    
    Select Case gErr
    
        Case 66029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 66030, 66031, 66032, 66044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr, colReqCompras.Item(iIndice))
        
        Case 66033, 66034, 66035, 66036, 66043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)
        
        Case 66037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMREQCOMPRA_INEXISTENTE", gErr, lNumIntItem)
        
        Case 66038
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSREQCOMPRA", gErr, lNumIntItem)
            
        Case 66039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSREQCOMPRABAIXADOS", gErr, tItemReqCompras.lNumIntDoc)
            
        Case 66040
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA1", gErr, lNumIntReq)
        
        Case 66041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REQUISICAOCOMPRA", gErr, lNumIntReq)
        
        Case 66042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REQUISICAOCOMPRABAIXADA", gErr, lNumIntReq)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149557)
    
    End Select
    
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function NFiscalEntradaCom_Baixa_Recebimento(ByVal objNFiscal As ClassNFiscal) As Long
'Faz a baixa de um Recebimento

Dim alComando(0 To 6) As Long
Dim tNFiscal As typeNFiscal
Dim objMovEstoqueBaixa As ClassMovEstoque
Dim iItem As Integer
Dim iIndice As Integer
Dim lErro As Long
Dim tItemNF As typeItemNF
Dim lCodigo As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
Dim iTipoNFiscal As Integer
Dim lItemNF As Long
Dim objItemMovEstoqueINF As ClassItemMovEstoque
Dim colMovEstoqueINF As Collection

On Error GoTo Erro_NFiscalEntradaCom_Baixa_Recebimento
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NFIENCO Or objNFiscal.iTipoNFiscal = DOCINFO_NFEENCO Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEEBEN Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEBEN Then
    
        'Abre os Comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 66045
        Next
    
        Select Case objNFiscal.iTipoNFiscal
        
            Case DOCINFO_NFIENCO
                iTipoNFiscal = DOCINFO_NRFPCO
            Case DOCINFO_NFEENCO, DOCINFO_NFEEBF, DOCINFO_NFEEBEN
                iTipoNFiscal = DOCINFO_NRFFCO
                
        End Select
            
        With tNFiscal
    
            .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
            .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
            .sPlaca = String(STRING_NFISCAL_PLACA, 0)
            .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
            .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
    
            'Removido por Luiz Nogueira em 21/08/03
            '.lVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
            '.lVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)
    
            'Procura no BD um Recebimento com os dados compatíveis aos da Nota Fiscal
            lErro = Comando_ExecutarPos(alComando(0), "SELECT DataEntrada, Status, Cliente, FilialCli, FilialEntrega, DataEmissao, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, NumIntDoc FROM NFiscal WHERE FilialEmpresa = ? AND TipoNFiscal = ? AND NumRecebimento = ?", 0, .dtDataEntrada, .iStatus, .lCliente, _
            .iFilialCli, .iFilialEntrega, .dtDataEmissao, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lNumIntDoc, objNFiscal.iFilialEmpresa, iTipoNFiscal, objNFiscal.lNumRecebimento)
            If lErro <> AD_SQL_SUCESSO Then gError 66046
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66047
            
            'Se encontrar
            If lErro = AD_SQL_SUCESSO Then
            
                'Faz um "lock" no Recebimento
                lErro = Comando_LockExclusive(alComando(0))
                If lErro <> SUCESSO Then gError 66048
                
                objTipoDocInfo.iCodigo = iTipoNFiscal
            
                'Lê o Tipo da Nota Fiscal
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then gError 66049
                
                'Se não achou o Tipo de Documento --> erro
                If lErro = 31415 Then gError 66050
                
                'Se o TipoMovtoEstoque for positivo
                If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
                    Set objMovEstoqueBaixa = New ClassMovEstoque
        
                    'Obtém o Codigo Automático do Movimento de Estoque
                    lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
                    If lErro <> SUCESSO Then gError 66051
        
                    objMovEstoqueBaixa.dtData = objNFiscal.dtDataEntrada
                    objMovEstoqueBaixa.iFilialEmpresa = objNFiscal.iFilialEmpresa
                    objMovEstoqueBaixa.iTipoMov = MOV_EST_BAIXA_RECEBIMENTO_MATERIAL
                    objMovEstoqueBaixa.lCodigo = lCodigo
                    objMovEstoqueBaixa.lCliente = objNFiscal.lCliente
                    objMovEstoqueBaixa.lFornecedor = objNFiscal.lFornecedor
                    objMovEstoqueBaixa.sDocOrigem = objTipoDocInfo.sSigla & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
        
                    tItemNF.sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
                    tItemNF.sProduto = String(STRING_PRODUTO, 0)
                    tItemNF.sUnidadeMed = String(STRING_UM_SIGLA, 0)
        
                    'Pesquisa os Ítens do Recebimento
                    lErro = Comando_ExecutarPos(alComando(2), "SELECT Item,Status,Produto,UnidadeMed,Quantidade,Almoxarifado,PrecoUnitario,PercDesc,ValorDesconto,DataEntrega,DescricaoItem,ValorAbatComissao,NumIntPedVenda,NumIntItemPedVenda,NumIntDoc,NumIntTrib,NumIntDocOrig FROM ItensNFiscal WHERE NumIntNF = ? ORDER BY Item ", 0, tItemNF.iItem, tItemNF.iStatus, tItemNF.sProduto, tItemNF.sUnidadeMed, tItemNF.dQuantidade, tItemNF.iAlmoxarifado, tItemNF.dPrecoUnitario, tItemNF.dPercDesc, tItemNF.dValorDesconto, tItemNF.dtDataEntrega, tItemNF.sDescricaoItem, tItemNF.dValorAbatComissao, tItemNF.lNumIntPedVenda, tItemNF.lNumIntItemPedVenda, tItemNF.lNumIntDoc, tItemNF.lNumIntTrib, tItemNF.lNumIntDocOrig, tNFiscal.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 66052
        
                    lErro = Comando_BuscarPrimeiro(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66053
                    
                    'Se encontrar
                    Do While lErro <> AD_SQL_SEM_DADOS
                        'Faz o "lock"
                        lErro = Comando_LockExclusive(alComando(2))
                        If lErro <> AD_SQL_SUCESSO Then gError 66054
        
                        '#########################################
                        'Inserido por Wagner - 08/07/2008
                        Set colMovEstoqueINF = New Collection
                        Set objItemMovEstoqueINF = New ClassItemMovEstoque
                        
                        objItemMovEstoqueINF.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                        objItemMovEstoqueINF.lNumIntDocOrigem = tItemNF.lNumIntDoc
                        objItemMovEstoqueINF.iFilialEmpresa = objNFiscal.iFilialEmpresa
                        
                        lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoqueINF, colMovEstoqueINF)
                        If lErro <> SUCESSO Then gError 66051
                        
                        For Each objItemMovEstoqueINF In colMovEstoqueINF
                            Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (objItemMovEstoqueINF.dQuantidade * tItemNF.dPrecoUnitario) - tItemNF.dValorDesconto, 0, tItemNF.sProduto, tItemNF.sDescricaoItem, objItemMovEstoqueINF.sSiglaUM, objItemMovEstoqueINF.dQuantidade, objItemMovEstoqueINF.iAlmoxarifado, "", tItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                        Next
                        '#########################################
                            
'                        'Atribui os ítens ao Movimento
'                        With tItemNF
'
'                            Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (.dQuantidade * .dPrecoUnitario) - .dValorDesconto, 0, .sProduto, .sDescricaoItem, .sUnidadeMed, .dQuantidade, .iAlmoxarifado, "", .lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
'                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'
'                        End With
                
                        lErro = Comando_ExecutarPos(alComando(4), "UPDATE ItensNFiscal SET Status = ?", alComando(2), STATUS_BAIXADO)
                        If lErro <> AD_SQL_SUCESSO Then gError 66055
        
'                        'Pesquisa em ItemNFItemPC Itens de Nota Fiscal Relacionados ao Pedido de Compras
'                        lErro = Comando_ExecutarPos(alComando(5), "SELECT ItemNFiscal FROM ItemNFItemPC WHERE ItemNFiscal = ?", 0, lItemNF, tItemNF.lNumIntDoc)
'                        If lErro <> AD_SQL_SUCESSO Then gError 66056
'
'                        lErro = Comando_BuscarPrimeiro(alComando(5))
'                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66057
'
'                        'Enquanto Encontrar
'                        Do While lErro <> AD_SQL_SEM_DADOS
'
'                            'Exclui o ItemNFItemPC
'                            lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM ItemNFItemPC", alComando(5))
'                            If lErro <> AD_SQL_SUCESSO Then gError 66058
'
'                            'Busca próximo ItemNFItemPC
'                            lErro = Comando_BuscarProximo(alComando(5))
'                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66059
'
'                        Loop
                        
                        lErro = Comando_BuscarProximo(alComando(2))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66060
        
                    Loop
                    
                    lErro = Comando_ExecutarPos(alComando(4), "UPDATE NFiscal SET Status = ? , Serie = ?, NumNotaFiscal = ?, DataAtualizacao=?", alComando(0), STATUS_BAIXADO, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, gdtDataHoje)
                    If lErro <> AD_SQL_SUCESSO Then gError 66061
                    
                    'Gera o Movimento de Baixa de Recebimento
                    lErro = CF("MovEstoque_Grava", objMovEstoqueBaixa)
                    If lErro <> SUCESSO Then gError 66062
    
                End If
            End If
        End With
    
        'Fecha os comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    NFiscalEntradaCom_Baixa_Recebimento = SUCESSO

    Exit Function

Erro_NFiscalEntradaCom_Baixa_Recebimento:

    NFiscalEntradaCom_Baixa_Recebimento = gErr

    Select Case gErr

        Case 66045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66046, 66047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL1", gErr, iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 66048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", gErr)

        Case 66050
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, iTipoNFiscal)

        Case 66052, 66053, 66060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)

        Case 66054
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSNFISCAL", gErr)

        Case 66055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSNFISCAL", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 66056, 66057, 66059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFITEMPC", gErr)
        
        Case 66058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMNFITEMPC", gErr, tItemNF.lNumIntDoc)

        Case 66061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)

        Case 66062, 66051, 66049
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149558)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalFatEntradaCom_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil, ByVal colPedCompras As Collection, ByVal colReqCompras As Collection) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNotaFiscal As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NFiscalFatEntradaCom_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 66500

    Set objNFiscal.objContabil = objContabil

    lErro = CF("NFiscalFatEntradaCom_Grava_EmTrans", objNFiscal, objContabil, colPedCompras, colReqCompras)
    If lErro <> SUCESSO And lErro <> 207637 Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then
    
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 66509

    Else
    
        'Não conseguiu excluir, então altera só algumas coisas
        If lNumNotaFiscal = 0 Then gError ERRO_SEM_MENSAGEM
    
        'Dá a Mensagem de aviso que só será gravado a contabilidade
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
        If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM

        Call Transacao_Rollback

        lErro = CF("NFiscalEntrada_AlteraDados", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    lErro = CF("WorkFlow_Trata_Transacao", "EST", objNFiscal.objTela, gsUsuario)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("NFiscal_Trata_Eventos_WorkFlow", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    NFiscalFatEntradaCom_Grava = SUCESSO

    Exit Function

Erro_NFiscalFatEntradaCom_Grava:

    NFiscalFatEntradaCom_Grava = gErr

    Select Case gErr

        Case 66500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 66509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149559)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalFatEntradaCom_Grava_BD(ByVal objNFiscal As ClassNFiscal, ByVal colPedCompras As Collection, ByVal colReqCompras As Collection) As Long
'Insere uma Nova Nota Fiscal de Entrada no Banco de Dados

Dim alComando(0 To 7) As Long
Dim lNumIntDoc As Long
Dim lErro As Long
Dim iClasseDocCPR As Integer
Dim lNumIntDocCPR As Long
Dim objItemNF As ClassItemNF
Dim objMovEstoque As ClassMovEstoque
Dim objMovEstoqueBaixa As ClassMovEstoque
Dim iIndice As Integer
Dim iTipoMovtoEstoque As Integer
Dim lCodigo As Long
Dim sDocOrigem As String
Dim iTipoMovtoEstoqueBenef As Integer, objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objRastroItemNF As ClassRastroItemNF
Dim objRastroMovto As ClassRastreamentoMovto
Dim colInsumos As Collection
Dim colRastroMovto As Collection
Dim dFator As Double
Dim dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dQuantItemNFUMEst As Double
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double

On Error GoTo Erro_NFiscalFatEntradaCom_Grava_BD

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66547
    Next

    sDocOrigem = String(STRING_TIPO_DOC_INFO_SIGLA, 0)

    'Lê TipoMovtoEstoque em TiposDocInfo
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then gError 66548

    iTipoMovtoEstoque = objTipoDocInfo.iTipoMovtoEstoque
    sDocOrigem = objTipoDocInfo.sSigla

    'Gera Conta a Pagar e devolve iClasseDocCPR e lNumIntDocCPR
    lErro = CF("CPR_Gera", objNFiscal, iClasseDocCPR, lNumIntDocCPR)
    If lErro <> SUCESSO Then gError 66549

    objNFiscal.iClasseDocCPR = iClasseDocCPR
    objNFiscal.lNumIntDocCPR = lNumIntDocCPR

    'Obtem o Número Interno da nova Nota Fiscal
    lErro = CF("NFiscal_Automatico", lNumIntDoc)
    If lErro <> SUCESSO Then gError 66550

    objNFiscal.lNumIntDoc = lNumIntDoc
    
    lErro = CF("NFiscal_Grava_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError 66550

    With objNFiscal

        .iNFe = SerieEletronica(.sSerie)

        'Insere a nova Nota Fiscal de Entrada no BD
        lErro = Comando_Executar(alComando(0), "INSERT INTO NFiscal (DataAtualizacao, NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,FornecedorBenef,FilialFornBenef,VolumeNumero,FreteRespons,DataReferencia,Observacao,NumRecebimento, FilialPedido, HoraEntrada, Moeda, TaxaMoeda, NFe, ChvNFe, HoraEmissao, DataCadastro, HoraCadastro, UsuCadastro) " & _
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            gdtDataHoje, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lFornecedorBenef, .iFilialFornBenef, .sVolumeNumero, .iFreteRespons, .dtDataReferencia, .sObservacao, .lNumRecebimento, .iFilialPedido, CDbl(.dtHoraEntrada), .iMoeda, .dTaxaMoeda, .iNFe, .sChvNFe, _
            CDbl(IIf(.dtDataSaida <> DATA_NULA, IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraSaida, .dtHoraEmissao, .dtHoraSaida), IIf(.dtHoraEmissao <> 0 And .dtHoraEmissao < .dtHoraEntrada, .dtHoraEmissao, .dtHoraEntrada))), gdtDataHoje, CDbl(Time), gsUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 66564
    
    End With

    lErro = CF("NFiscal_GravaTributacao", objNFiscal, alComando(4), alComando(5))
    If lErro <> SUCESSO Then gError 66551

    'Para cada Ítem da Nota
    For Each objItemNF In objNFiscal.colItensNF

        objItemNF.lNumIntNF = objNFiscal.lNumIntDoc

        'Obtém o Número Interno do Ítem
        lErro = CF("NFiscalItem_Automatico", lNumIntDoc)
        If lErro <> SUCESSO Then gError 66552

        objItemNF.lNumIntDoc = lNumIntDoc

        With objItemNF
            'Insere o Ítem no Banco de dados
            lErro = Comando_Executar(alComando(1), "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, NumIntDocOrig,Almoxarifado, Ccl, PrecoUnitarioMoeda) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lNumIntNF, .iItem, STATUS_LANCADO, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, lNumIntDoc, .lNumIntTrib, 0, .iAlmoxarifado, .sCcl, .dPrecoUnitarioMoeda)
            If lErro <> AD_SQL_SUCESSO Then gError 66565
        End With

       lErro = CF("NFiscal_GravaTributacaoItem", objItemNF, alComando(3), alComando(6))
       If lErro <> SUCESSO Then gError 66553
       
        objItemNF.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
        objItemNF.objInfoAdicDocItem.lNumIntDocItem = objItemNF.lNumIntDoc
        lErro = CF("InfoAdicDocItem_Insere", objItemNF.objInfoAdicDocItem, alComando(7))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Next

    lErro = CF("NFe_Grava_Vinculo_Xml", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("NFiscal_Grava_Customizado", objNFiscal)
    If lErro <> SUCESSO Then gError 181268

    lErro = CF("NFiscalItens_Grava_Customizado", objNFiscal)
    If lErro <> SUCESSO Then gError 181268
    
    'Inclui em ItemNFItemPC o ItemNF associado a ItemPC
    lErro = CF("ItemNFItemPC_Inclui", objNFiscal)
    If lErro <> SUCESSO Then gError 66554
            
    'Atualiza as Quantidades dos ItensPC
    lErro = CF("ItensPedCompra_Atualiza_Quantidades", objNFiscal, colPedCompras)
    If lErro <> SUCESSO Then gError 66076
    
    'Inclui em ItemNFItemRC o ItemNF associado a ItemRC
    lErro = CF("ItemNFItemRC_Inclui", objNFiscal)
    If lErro <> SUCESSO Then gError 66555
        
    'Atualiza as quantidades dos Itens de Requisição
    lErro = CF("ItensReqCompra_Atualiza_Quantidades", objNFiscal)
    If lErro <> SUCESSO Then Error 66562

    'Faz a baixa de um recebimento
    'a baixa do recebimento deve ocorrer antes da gravacao da nota fiscal para que os valores dos saldos estejam corretos e permitam calcular o novo custo do produto, se for o caso.
    If objNFiscal.lNumRecebimento > 0 Then

        'Lê o Recebimento e Verifica os Dados
        lErro = NFiscal_Valida_Recebimento(objNFiscal)
        If lErro <> SUCESSO Then Error 66563

        lErro = CF("NFiscalFatEntradaCom_Baixa_Recebimento", objNFiscal)
        If lErro <> SUCESSO Then gError 66556

    End If

    If iTipoMovtoEstoque > 0 Then

        Set objMovEstoque = New ClassMovEstoque

        'Obtém o Código do Movimento
        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then gError 66557

        objMovEstoque.dtData = objNFiscal.dtDataEntrada
        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.iTipoMov = iTipoMovtoEstoque
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = sDocOrigem & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)

        'Se for para incluir o valor do frete e outras depesas no custo do ItemNF
        If gobjEST.iIncluiFreteNoCusto = INCLUI_FRETE_CUSTO Then
            Call Inclui_Frete_CustoItemNF(objNFiscal)
        End If

'mario distribuicao substituicao do codigo original acrescentando o tab de distribuicao
                        
        'Adiciona itens ao Movimento
        iIndice = 0
        For Each objItemNF In objNFiscal.colItensNF
            
            iIndice = iIndice + 1
            
            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
                If objItemNF.ColAlocacoes.Count = 0 Then gError 89647

                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 89648
                
                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
                
                'calcula o custo do item da NF em função do credito do ICMS e do IPI
                lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
                If lErro <> SUCESSO Then gError 66588

                dSomaAlocacao = 0
                                                                
                For Each objItemNFAloc In objItemNF.ColAlocacoes
                                        
                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
                    
                    Set colInsumos = New Collection
                    Set colRastroMovto = New Collection
                                        
                    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                    
                    'Guarda o Rastreamento dos ItensNF
                    If Not (objNFiscal.objRastreamento Is Nothing) Then
                        lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
                        If lErro <> SUCESSO Then gError 83574
                    End If
                    
'mario
'                    If Not (objItemNF.colRastreamento Is Nothing) Then
'
'                        'Guarda o Rastreamento dos ItensNF
'                        For Each objRastroItemNF In objItemNF.colRastreamento
'
'                            If objRastroItemNF.iAlmoxCodigo = objItemNFAloc.iAlmoxarifado Then
'
'                                Set objRastroMovto = New ClassRastreamentoMovto
'                                objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
'                                objRastroMovto.sLote = objRastroItemNF.sLote
'                                objRastroMovto.sProduto = objItemNF.sProduto
'
'                                'Adiciona objRastroMovto na coleção de Rastreamento
'                                objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
'
'                            End If
'
'                        Next
'
'                    End If
'fim mario
                    
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                    objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                    objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                    objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
                    
                    'se for uma nota de complemento de preço ou outra que nao deve afetar qtde ==> zera a quantidade pois somente o custo será utilizado
                    If TipoDocInfo_ZeraQtdeMovto(objNFiscal.iTipoNFiscal) Then
                        objItemMovEstoque.dQuantidade = 0
                    End If
                    
                Next
                
                '#################################################
                'Alterado por Wagner 16/11/04
                'Calcula o número de casas decimais do Formato de Estoque
'                iNumCasasDec = Len'APAGAR'(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".") + 1)))
'
'                dAcrescimo = 0
'                If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                If Abs(dSomaAlocacao - dQuantItemNFUMEst) > dAcrescimo + QTDE_ESTOQUE_DELTA Then

                dQuantItemNFUMEst = Arredonda_Estoque(dQuantItemNFUMEst)

                
                If Abs(dSomaAlocacao - dQuantItemNFUMEst) > QTDE_ESTOQUE_DELTA Then
                '#################################################
                
                    If dSomaAlocacao < dQuantItemNFUMEst Then gError 89649
                    
                    If dSomaAlocacao > dQuantItemNFUMEst Then gError 89650
            
                End If
                
            End If
        Next


'mario distribuicao fim da alteração para incluir o tab de distribuicao


'        'Adiciona os ítem ao Movimento
'        For Each objItemNF In objNFiscal.colItensNF
'
'            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
'
'                'calcula o custo do item da NF em função do credito do ICMS e do IPI
'                lErro = CF("NFiscalEntrada_Calcula_Custo",objItemNF)
'                If lErro <> SUCESSO Then gError 66558
'
'                Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNF.sUnidadeMed, objItemNF.dQuantidade, objItemNF.iAlmoxarifado, objItemNF.sAlmoxarifadoNomeRed, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
'                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'
'                Set objItemMovEstoque.colRastreamentoMovto = New Collection
'                Set objItemMovEstoque.colApropriacaoInsumo = New Collection
'
'                'Guarda o Rastreamento dos ItensNF
'                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
'                If lErro <> SUCESSO Then gError 83574
'
''                'Guarda o Rastreamento dos ItensNF
''                For Each objRastroItemNF In objItemNF.colRastreamento
''
''                    Set objRastroMovto = New ClassRastreamentoMovto
''                    objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
''                    objRastroMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
''                    objRastroMovto.sLote = objRastroItemNF.sLote
''                    objRastroMovto.lNumIntDocOrigem = objItemNF.lNumIntDoc
''                    objRastroMovto.sProduto = objItemNF.sProduto
''                    objRastroMovto.iFilialOP = objRastroItemNF.iLoteFilialOP
''
''                    'Adiciona objRastroMovto na coleção de Rastreamento
''                    objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
''
''                Next
'
'                objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
'                objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
'                objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
'                objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
'
'            End If
'
'        Next

        'Gera um Movimento de Estoque
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then gError 66559

        lErro = CF("NFiscalEntCom_Gera_Reservas", objNFiscal, objMovEstoque)
        If lErro <> SUCESSO Then gError 66561

    End If

    'Verifica se existem Pedidos a serem Baixados
    lErro = CF("NFiscalEntCom_BaixaPedidos", colPedCompras, objNFiscal)
    If lErro <> SUCESSO Then gError 66560

    'Verifica se existem Requisições a serem Baixadas
    lErro = CF("NFiscalEntCom_BaixaRequisicoes", colReqCompras, objNFiscal)
    If lErro <> SUCESSO Then gError 66561

    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalFatEntradaCom_Grava_BD = SUCESSO

    Exit Function

Erro_NFiscalFatEntradaCom_Grava_BD:

    NFiscalFatEntradaCom_Grava_BD = gErr

    Select Case gErr

        Case 66547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66548, 66549, 66550, 66551, 66552, 66553, 66554, 66555, 66556, 66557, 66558, 66559, 66560, 66561, 66562, 66563, 83574, 89648, 181268

        Case 66564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 66565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", gErr, objNFiscal.lNumNotaFiscal)

        Case 89647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INEXISTENTE", gErr, iIndice)

        Case 89649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INCOMPLETA", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case 89650
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149560)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalFatEntradaCom_Baixa_Recebimento(ByVal objNFiscal As ClassNFiscal) As Long
'Faz a baixa de um Recebimento

Dim alComando(0 To 6) As Long
Dim tNFiscal As typeNFiscal
Dim objMovEstoqueBaixa As ClassMovEstoque
Dim iItem As Integer
Dim iIndice As Integer
Dim lErro As Long
Dim tItemNF As typeItemNF
Dim lCodigo As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
Dim iTipoNFiscal As Integer
Dim lItemNF As Long
Dim objItemMovEstoqueINF As ClassItemMovEstoque
Dim colMovEstoqueINF As Collection

On Error GoTo Erro_NFiscalFatEntradaCom_Baixa_Recebimento
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEFNCO Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEFNCO Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEEBEN Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEBEN Then
    
        'Abre os Comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 66599
        Next
    
        Select Case objNFiscal.iTipoNFiscal
        
            Case DOCINFO_NFIEFNCO
                iTipoNFiscal = DOCINFO_NRFPCO
            Case DOCINFO_NFEEFNCO, DOCINFO_NFEEBF, DOCINFO_NFEEBEN
                iTipoNFiscal = DOCINFO_NRFFCO
                
        End Select
            
        With tNFiscal
    
            .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
            .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
            .sPlaca = String(STRING_NFISCAL_PLACA, 0)
            .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
            .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
    
            'Removido por Luiz Nogueira em 21/08/03
            '.lVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
            '.lVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)
    
            'Procura no BD um Recebimento com os dados compatíveis aos da Nota Fiscal
            lErro = Comando_ExecutarPos(alComando(0), "SELECT DataEntrada, Status, Cliente, FilialCli, FilialEntrega, DataEmissao, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, NumIntDoc FROM NFiscal WHERE FilialEmpresa = ? AND TipoNFiscal = ? AND NumRecebimento = ?", 0, .dtDataEntrada, .iStatus, .lCliente, _
            .iFilialCli, .iFilialEntrega, .dtDataEmissao, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lNumIntDoc, objNFiscal.iFilialEmpresa, iTipoNFiscal, objNFiscal.lNumRecebimento)
            If lErro <> AD_SQL_SUCESSO Then gError 66600
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66601
            
            'Se encontrar
            If lErro = AD_SQL_SUCESSO Then
            
                'Faz um "lock" no Recebimento
                lErro = Comando_LockExclusive(alComando(0))
                If lErro <> SUCESSO Then gError 66602
                
                objTipoDocInfo.iCodigo = iTipoNFiscal
            
                'Lê o Tipo da Nota Fiscal
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then gError 66616
                
                'Se não achou o Tipo de Documento --> erro
                If lErro = 31415 Then gError 66603
                
                'Se o TipoMovtoEstoque for positivo
                If objTipoDocInfo.iTipoMovtoEstoque > 0 Then
    
                    Set objMovEstoqueBaixa = New ClassMovEstoque
        
                    'Obtém o Codigo Automático do Movimento de Estoque
                    lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
                    If lErro <> SUCESSO Then gError 66615
        
                    objMovEstoqueBaixa.dtData = objNFiscal.dtDataEntrada
                    objMovEstoqueBaixa.iFilialEmpresa = objNFiscal.iFilialEmpresa
                    objMovEstoqueBaixa.iTipoMov = MOV_EST_BAIXA_RECEBIMENTO_MATERIAL
                    objMovEstoqueBaixa.lCodigo = lCodigo
                    objMovEstoqueBaixa.lCliente = objNFiscal.lCliente
                    objMovEstoqueBaixa.lFornecedor = objNFiscal.lFornecedor
                    objMovEstoqueBaixa.sDocOrigem = objTipoDocInfo.sSigla & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
        
                    tItemNF.sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
                    tItemNF.sProduto = String(STRING_PRODUTO, 0)
                    tItemNF.sUnidadeMed = String(STRING_UM_SIGLA, 0)
        
                    'Pesquisa os Ítens do Recebimento
                    lErro = Comando_ExecutarPos(alComando(2), "SELECT Item,Status,Produto,UnidadeMed,Quantidade,Almoxarifado,PrecoUnitario,PercDesc,ValorDesconto,DataEntrega,DescricaoItem,ValorAbatComissao,NumIntPedVenda,NumIntItemPedVenda,NumIntDoc,NumIntTrib,NumIntDocOrig FROM ItensNFiscal WHERE NumIntNF = ? ORDER BY Item ", 0, tItemNF.iItem, tItemNF.iStatus, tItemNF.sProduto, tItemNF.sUnidadeMed, tItemNF.dQuantidade, tItemNF.iAlmoxarifado, tItemNF.dPrecoUnitario, tItemNF.dPercDesc, tItemNF.dValorDesconto, tItemNF.dtDataEntrega, tItemNF.sDescricaoItem, tItemNF.dValorAbatComissao, tItemNF.lNumIntPedVenda, tItemNF.lNumIntItemPedVenda, tItemNF.lNumIntDoc, tItemNF.lNumIntTrib, tItemNF.lNumIntDocOrig, tNFiscal.lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 66604
        
                    lErro = Comando_BuscarPrimeiro(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66605
                    
                    'Se encontrar
                    Do While lErro <> AD_SQL_SEM_DADOS
                        'Faz o "lock"
                        lErro = Comando_LockExclusive(alComando(2))
                        If lErro <> AD_SQL_SUCESSO Then gError 66607
        
                        '#########################################
                        'Inserido por Wagner - 08/07/2008
                        Set colMovEstoqueINF = New Collection
                        Set objItemMovEstoqueINF = New ClassItemMovEstoque
                        
                        objItemMovEstoqueINF.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                        objItemMovEstoqueINF.lNumIntDocOrigem = tItemNF.lNumIntDoc
                        objItemMovEstoqueINF.iFilialEmpresa = objNFiscal.iFilialEmpresa
                        
                        lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoqueINF, colMovEstoqueINF)
                        If lErro <> SUCESSO Then gError 66615
                        
                        For Each objItemMovEstoqueINF In colMovEstoqueINF
                            Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (objItemMovEstoqueINF.dQuantidade * tItemNF.dPrecoUnitario) - tItemNF.dValorDesconto, 0, tItemNF.sProduto, tItemNF.sDescricaoItem, objItemMovEstoqueINF.sSiglaUM, objItemMovEstoqueINF.dQuantidade, objItemMovEstoqueINF.iAlmoxarifado, "", tItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                        Next
                        '#########################################
                            
'                        'Atribui os ítens ao Movimento
'                        With tItemNF
'
'                            Set objItemMovEstoque = objMovEstoqueBaixa.colItens.Add(0, 0, (.dQuantidade * .dPrecoUnitario) - .dValorDesconto, 0, .sProduto, .sDescricaoItem, .sUnidadeMed, .dQuantidade, .iAlmoxarifado, "", .lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
'                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'
'                        End With
                
                        lErro = Comando_ExecutarPos(alComando(4), "UPDATE ItensNFiscal SET Status = ?", alComando(2), STATUS_BAIXADO)
                        If lErro <> AD_SQL_SUCESSO Then gError 66608
        
'                        'Pesquisa em ItemNFItemPC Itens de Nota Fiscal Relacionados ao Pedido de Compras
'                        lErro = Comando_ExecutarPos(alComando(5), "SELECT ItemNFiscal FROM ItemNFItemPC WHERE ItemNFiscal = ?", lItemNF, tItemNF.lNumIntDoc)
'                        If lErro <> AD_SQL_SUCESSO Then gError 66609
'
'                        lErro = Comando_BuscarPrimeiro(alComando(5))
'                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66610
'
'                        'Enquanto Encontrar
'                        Do While lErro <> AD_SQL_SEM_DADOS
'
'                            'Exclui o ItemNFItemPC
'                            lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM ItemNFItemPC", alComando(5))
'                            If lErro <> AD_SQL_SUCESSO Then gError 66612
'
'                            'Busca próximo ItemNFItemPC
'                            lErro = Comando_BuscarProximo(alComando(5))
'                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66611
'
'                        Loop
                        
                        lErro = Comando_BuscarProximo(alComando(2))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66606
        
                    Loop
                    
                    lErro = Comando_ExecutarPos(alComando(4), "UPDATE NFiscal SET Status = ? , Serie = ?, NumNotaFiscal = ?, DataAtualizacao=?", alComando(0), STATUS_BAIXADO, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, gdtDataHoje)
                    If lErro <> AD_SQL_SUCESSO Then gError 66613
                    
                    'Gera o Movimento de Baixa de Recebimento
                    lErro = CF("MovEstoque_Grava", objMovEstoqueBaixa)
                    If lErro <> SUCESSO Then gError 66614
    
                End If
            End If
        End With
    
        'Fecha os comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    NFiscalFatEntradaCom_Baixa_Recebimento = SUCESSO

    Exit Function

Erro_NFiscalFatEntradaCom_Baixa_Recebimento:

    NFiscalFatEntradaCom_Baixa_Recebimento = gErr

    Select Case gErr

        Case 66599
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66600, 66601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL1", gErr, iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 66602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", gErr)

        Case 66603
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, iTipoNFiscal)

        Case 66604, 66605, 66606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)

        Case 66607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSNFISCAL", gErr)

        Case 66608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSNFISCAL", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 66609, 66610, 66611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFITEMPC", gErr)
        
        Case 66612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMNFITEMPC", gErr, tItemNF.lNumIntDoc)

        Case 66613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)

        Case 66614, 66615, 66616
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149561)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalEntrada_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNotaFiscal As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NFiscalEntrada_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 35043

    Set objNFiscal.objContabil = objContabil

    lErro = CF("NFiscalEntrada_Grava_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO And lErro <> 207637 Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then
    
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 35079

    Else
        'Não conseguiu excluir, então altera só algumas coisas
        If lNumNotaFiscal = 0 Then gError ERRO_SEM_MENSAGEM
    
        'Dá a Mensagem de aviso que só será gravado a contabilidade
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
        If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM

        Call Transacao_Rollback

        lErro = CF("NFiscalEntrada_AlteraDados", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    NFiscalEntrada_Grava = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Grava:

    NFiscalEntrada_Grava = gErr

    Select Case gErr

        Case 35043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case 35079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149562)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalFatEntrada_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNotaFiscal As Long
Dim vbMsgRes As VbMsgBoxResult


On Error GoTo Erro_NFiscalFatEntrada_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 42665
    
    Set objNFiscal.objContabil = objContabil
    
    lErro = CF("NFiscalFatEntrada_Grava_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO And lErro <> 207637 Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then
    
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 42684

    Else
        'Não conseguiu excluir, então altera só algumas coisas
        If lNumNotaFiscal = 0 Then gError ERRO_SEM_MENSAGEM
    
        'Dá a Mensagem de aviso que só será gravado a contabilidade
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
        If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM

        Call Transacao_Rollback

        lErro = CF("NFiscalEntrada_AlteraDados", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    NFiscalFatEntrada_Grava = SUCESSO

    Exit Function

Erro_NFiscalFatEntrada_Grava:

    NFiscalFatEntrada_Grava = gErr
    
    Select Case gErr
    
        Case 42665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 42684
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case ERRO_SEM_MENSAGEM 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149563)
            
    End Select

    Call Transacao_Rollback

    Exit Function
    
End Function

Function NFiscalEntradaDevolucao_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a gravação de Uma Nota Fiscal de Entrada-Devolução

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNotaFiscal As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NFiscalEntradaDevolucao_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 35378

    Set objNFiscal.objContabil = objContabil

    lErro = CF("NFiscalEntradaDevolucao_Grava_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO And lErro <> 207637 Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then
    
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 35443

    Else
    
        'Não conseguiu excluir, então altera só algumas coisas
        If lNumNotaFiscal = 0 Then gError ERRO_SEM_MENSAGEM
    
        'Dá a Mensagem de aviso que só será gravado a contabilidade
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
        If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM

        Call Transacao_Rollback

        lErro = CF("NFiscalEntrada_AlteraDados", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    NFiscalEntradaDevolucao_Grava = SUCESSO

    Exit Function

Erro_NFiscalEntradaDevolucao_Grava:

    NFiscalEntradaDevolucao_Grava = gErr

    Select Case gErr

        Case 35378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 35443
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149564)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscalEntradaRemessa_Grava(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz as validações necessárias a gravação de Uma Nota Fiscal de Entrada-Remessa

Dim lErro As Long
Dim lTransacao As Long
Dim lNumNotaFiscal As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NFiscalEntradaRemessa_Grava

    lNumNotaFiscal = objNFiscal.lNumNotaFiscal

    'Abre a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 35603
    
    Set objNFiscal.objContabil = objContabil
    
    lErro = CF("NFiscalEntradaRemessa_Grava_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO And lErro <> 207637 Then gError ERRO_SEM_MENSAGEM
        
    If lErro = SUCESSO Then
    
        'Faz o Commit da Transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 35624

    Else
    
        'Não conseguiu excluir, então altera só algumas coisas
        If lNumNotaFiscal = 0 Then gError ERRO_SEM_MENSAGEM
    
        'Dá a Mensagem de aviso que só será gravado a contabilidade
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
        If vbMsgRes = vbNo Then gError ERRO_SEM_MENSAGEM

        Call Transacao_Rollback

        lErro = CF("NFiscalEntrada_AlteraDados", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    NFiscalEntradaRemessa_Grava = SUCESSO
    
    Exit Function

Erro_NFiscalEntradaRemessa_Grava:

    NFiscalEntradaRemessa_Grava = gErr

    Select Case gErr

        Case 35603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 35624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149565)

    End Select

    Call Transacao_Rollback
    
    Exit Function

End Function

Function Recebimento_Desfaz_Baixa(ByVal objNFiscal As ClassNFiscal) As Long
'Desfaz a baixa do Recebimento vinculado a nota fiscal que foi cancelada e passada no obj

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim iStatus As Integer
Dim objMovEstoqueCancBaixa As New ClassMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iTipoNFiscal As Integer
Dim lCodigo As Long
Dim tItemNF As typeItemNF
Dim lNumIntDoc As Integer
Dim objItemMovEstoque As ClassItemMovEstoque
Dim sUpDate As String
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_Recebimento_Desfaz_Baixa

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 62536
    Next
    
    'Busca o recebimento no BD
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc,TipoNFiscal FROM NFiscal WHERE NumRecebimento = ? AND FilialEmpresa = ? AND TipoNFiscal IN (?,?,?,?,?,?)", 0, lNumIntDoc, iTipoNFiscal, objNFiscal.lNumRecebimento, objNFiscal.iFilialEmpresa, DOCINFO_NRCC, DOCINFO_NRCP, DOCINFO_NRFF, DOCINFO_NRFFCO, DOCINFO_NRFP, DOCINFO_NRFPCO)
    If lErro <> AD_SQL_SUCESSO Then gError 62537
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62538
    If lErro <> AD_SQL_SUCESSO Then gError 62539 'Não encontra
    
    'Faz lock no recebimento
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 62540


    objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objMovEstoque.iTipoMov = MOV_EST_BAIXA_RECEBIMENTO_MATERIAL
    
    'Lê os movimento de estoque relativos a baixa do recebimento de material
    lErro = CF("MovEstoqueItens_Le_Desc2", objMovEstoque, lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 186952 Then gError 186954

    If lErro = SUCESSO Then 'gError 186955

        'Redefine os movimento para fazer a exclusao
        objMovEstoque.iOperacao = EXCLUSAO_MOVIMENTO
        
        'Exclui o movimento de estoque
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then gError 186956
        
    End If

    tItemNF.sProduto = String(STRING_PRODUTO, 0)
    tItemNF.sUnidadeMed = String(STRING_UM_SIGLA, 0)
    tItemNF.sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
    
    'Seleciona o itens da NF de recebimento
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Item,Status,Produto,UnidadeMed,Quantidade,Almoxarifado,PrecoUnitario,PercDesc,ValorDesconto,DataEntrega,DescricaoItem,ValorAbatComissao,NumIntPedVenda,NumIntItemPedVenda,NumIntDoc,NumIntTrib,NumIntDocOrig FROM ItensNFiscal WHERE NumIntNF = ? ORDER BY Item ", 0, tItemNF.iItem, tItemNF.iStatus, tItemNF.sProduto, tItemNF.sUnidadeMed, tItemNF.dQuantidade, tItemNF.iAlmoxarifado, tItemNF.dPrecoUnitario, tItemNF.dPercDesc, tItemNF.dValorDesconto, tItemNF.dtDataEntrega, tItemNF.sDescricaoItem, tItemNF.dValorAbatComissao, tItemNF.lNumIntPedVenda, tItemNF.lNumIntItemPedVenda, tItemNF.lNumIntDoc, tItemNF.lNumIntTrib, tItemNF.lNumIntDocOrig, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 62545
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62546
    
    'Para cada um dos itens
    Do While lErro <> AD_SQL_SEM_DADOS
        'Faz o lock
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> SUCESSO Then gError 62547
        
        'Altera o Status desfazendo a baixa.
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensNFiscal SET Status = ?", alComando(2), STATUS_LANCADO)
        If lErro <> AD_SQL_SUCESSO Then gError 62548
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62549
        
    Loop
    
    sUpDate = "UPDATE NFiscal SET Status =?"
        
    'Altera o Status Desfazendo a baixa
'    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_RECCLI Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_RECFORN Then
'        sUpDate = sUpDate & ", Serie = '', NumNotaFiscal = 0"
'    End If
    
    lErro = Comando_ExecutarPos(alComando(1), sUpDate, alComando(0), STATUS_LANCADO)
    If lErro <> AD_SQL_SUCESSO Then gError 62541
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Recebimento_Desfaz_Baixa = SUCESSO
    
    Exit Function

Erro_Recebimento_Desfaz_Baixa:

    Recebimento_Desfaz_Baixa = gErr
    
    Select Case gErr
    
        Case 62536
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 62537, 62538
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)

        Case 62539
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NAO_CADASTRADO", gErr, objNFiscal.lNumRecebimento)
        
        Case 62540
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", gErr)
            
        Case 62541
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL1", gErr)
        
        Case 62542, 62544, 62550, 186954, 186956
        
        Case 62543
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)
        
        Case 62545, 62546, 62549
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL2", gErr)
            
        Case 62547
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSNFISCAL", gErr)
            
        Case 62548
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITEMNF", gErr)
        
        Case 186955
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO", gErr, lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149566)
            
    End Select
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function
    
End Function

Function NotaFiscalEntrada_Excluir(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz o cancelamento de uma nota fiscal de Saida

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_NotaFiscalEntrada_Excluir
    
    'ABre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 197966

    lErro = CF("NotaFiscalEntrada_Excluir_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 197967
    
    'Confirma a transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 197968
    
    NotaFiscalEntrada_Excluir = SUCESSO

    Exit Function

Erro_NotaFiscalEntrada_Excluir:

    NotaFiscalEntrada_Excluir = gErr

    Select Case gErr

        Case 197966
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 197967
        
        Case 197968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197969)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NotaFiscalEntrada_Excluir_EmTrans(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil, Optional ByVal objNFAlterada As ClassNFiscal) As Long
'Faz a exclusao de uma nota fiscal de Entrada

Dim lErro As Long, iIndice As Integer, lErro1 As Long, iClasseDocCPRAux As Integer
Dim alComando(0 To 15) As Long
Dim iClasseDocCPR As Integer, lNumIntDocCPR As Long
Dim lNumMovtoEstoque As Long
Dim objMovEstoque As New ClassMovEstoque
Dim objMovEstoque1 As New ClassMovEstoque
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lCodigo As Long, lNumNotaFiscal As Long, iTipoNF As Integer, iStatus As Integer
Dim lNumRecebimento As Long
Dim lNumIntItemNF As Long
Dim lNumIntDocEst As Long
Dim objFilFornFilEmp As New ClassFilFornFilEmp
Dim tItemNF As typeItemNF
Dim objProduto As New ClassProduto
Dim objItemNF As ClassItemNF
Dim lNumIntOrigem As Long
Dim iTipoOrigem As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim iIndiceAux As Integer, iEnviado As Integer
Dim sConteudo As String, lAux As Long

On Error GoTo Erro_NotaFiscalEntrada_Excluir_EmTrans
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89007
    Next
    
    'Lê a nota fiscal
    lErro = Comando_ExecutarPos(alComando(0), "SELECT  ClasseDocCPR, NumIntDocCPR, NumRecebimento FROM NFiscal WHERE NumIntDoc = ?", 0, iClasseDocCPR, lNumIntDocCPR, lNumRecebimento, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 89009

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89010
    
    If lErro <> AD_SQL_SUCESSO Then gError 89011 'Não encontrou

    'faz o lock na Nota Fiscal
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 89012

    'se foi comandada uma exclusao pelo botao de exclusao e nao pela gravacao de nota (alteracao) ==> testar se a nota foi enviada
    If objNFiscal.objContabil.objGridInt Is Nothing Then
        
        'verifica se a nota fiscal foi processada pela sefaz ou se foi enviada
        lErro = CF("NFiscal_Testa_Existencia1", objNFiscal)
        If lErro <> SUCESSO And lErro <> 207565 And lErro <> 207624 Then gError 207633

    End If

    If Not (objNFAlterada Is Nothing) Then
    
        objNFAlterada.lNumIntNFAnterior = objNFiscal.lNumIntDoc
    
    Else
    
        'exclui os registros de nfefedlote ligado a nota em questao
        lErro = CF("NFeFedLote_Excluir", objNFiscal)
        If lErro <> SUCESSO Then gError 207634
        
        lErro = CF("RPSWEBLote_Excluir", objNFiscal)
        If lErro <> SUCESSO Then gError 207635

    End If
    
'    lErro = CF("CRFatConfig_Le", "TESTA_EXCLUSAO_NFE", 0, sConteudo)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    If StrParaInt(sConteudo) <> DESMARCADO Then
'
'        lErro = CF("NFeFedLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'        If iEnviado = 1 Then gError 203045
'
'    End If


    If Not objContabil Is Nothing Then
        'exclui os dados contábeis pendentes ou atualizados (se houverem) relativos a lNumIntDocOrigem
        lErro = objContabil.Contabil_Exclui(objNFiscal.lNumIntDoc, CONTABIL_EXCLUI_EFETIVO)
        If lErro <> SUCESSO Then gError 89038
    End If
        
    'Lê o tipodocinfo da nota fiscal
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 89013
    
    If lErro <> SUCESSO Then gError 89014
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPPC And objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPCC Then
    
        lErro = NFiscal_Exclui_NFImportacao(objNFiscal)
        If lErro <> SUCESSO Then gError 184748
    
    End If
    
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 89013
    
    'Se é uma nota que possui vinculação com outra que a sucede. Ex. Uma nota que tem uma devolução vinculada ou uma nota que tem um complemento vinculado.
    'If Len(Trim(objTipoDocInfo.sSiglaNFOriginal)) > 0 Or objTipoDocInfo.iComplementar <> DOCINFO_COMPLEMENTO Then
    
    'verifica se a nota está vinculada a alguma outra (um complemento, uma devolução, etc.)
    lErro = Comando_Executar(alComando(1), "SELECT NumNotaFiscal, TipoNFiscal FROM NFiscal WHERE NumIntNotaOriginal = ?", lNumNotaFiscal, iTipoNF, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 89015

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89016
    
    If lErro = AD_SQL_SUCESSO Then gError 89017 'Se já foi, erro.
    
    'verifica se a nota está vinculada a alguma outra (um complemento, uma devolução, etc.)
    lErro = Comando_Executar(alComando(15), "SELECT N.NumNotaFiscal, N.TipoNFiscal FROM NFiscal AS N, ItensNFiscal AS I WHERE N.NumIntDoc = I.NumIntNF AND I.NumIntDocOrig IN (SELECT NumIntDoc FROM ItensNFiscal WHERE NumIntNF = ?)", lNumNotaFiscal, iTipoNF, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 89015

    lErro = Comando_BuscarPrimeiro(alComando(15))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89016
    
    If lErro = AD_SQL_SUCESSO Then gError 89017 'Se já foi, erro.
    'End If

    If lNumRecebimento > 0 Then
        
        objNFiscal.lNumRecebimento = lNumRecebimento
        
        lErro = CF("Recebimento_Desfaz_Baixa", objNFiscal)
        If lErro <> SUCESSO Then gError 89018
        
    End If
    
    'Se a nota gerou documentos em  CPR
    If iClasseDocCPR > 0 And lNumIntDocCPR > 0 Then
        
        If Not (objNFAlterada Is Nothing) Then
        
            lErro = Comando_Executar(alComando(14), "SELECT TiposDeDocumento.ClasseDocCPR FROM TiposDeDocumento, TiposDocInfo WHERE TiposDeDocumento.Sigla = TiposDocInfo.TipoDocCPR AND TiposDocInfo.Codigo = ?", iClasseDocCPRAux, objNFAlterada.iTipoNFiscal)
            If lErro <> AD_SQL_SUCESSO Then gError 89015
            
            lErro = Comando_BuscarPrimeiro(alComando(14))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89016
            
            If lErro = AD_SQL_SUCESSO And iClasseDocCPR = iClasseDocCPRAux And (iClasseDocCPR = CPR_TITULO_PAGAR Or iClasseDocCPR = CPR_NF_PAGAR) Then
            
                objNFAlterada.iClasseDocCPR = iClasseDocCPR
                objNFAlterada.lNumIntDocCPR = -lNumIntDocCPR
            
            Else
            
                'Desfaz os documentos associados
                lErro = CF("CPR_Exclui", iClasseDocCPR, lNumIntDocCPR, 1)
                If lErro <> SUCESSO Then gError 89019
            
            End If
            
        Else
        
            'Desfaz os documentos associados
            lErro = CF("CPR_Exclui", iClasseDocCPR, lNumIntDocCPR, 1)
            If lErro <> SUCESSO Then gError 89019
        
        End If
        
    End If

    tItemNF.sUnidadeMed = String(STRING_UM_SIGLA, 0)
    tItemNF.sProduto = String(STRING_PRODUTO, 0)

    'Busca os itens da nota associada
    lErro1 = Comando_ExecutarPos(alComando(2), "SELECT Produto, NumIntDoc, Quantidade, UnidadeMed, NumIntDocOrig FROM ItensNFiscal WHERE NumIntNF = ? ", 0, tItemNF.sProduto, tItemNF.lNumIntDoc, tItemNF.dQuantidade, tItemNF.sUnidadeMed, tItemNF.lNumIntDocOrig, objNFiscal.lNumIntDoc)
    If lErro1 <> AD_SQL_SUCESSO Then gError 89020

    lErro1 = Comando_BuscarPrimeiro(alComando(2))
    If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 89021
    
    If lErro1 <> AD_SQL_SUCESSO Then gError 89022
    
    'Guarda o número do Item
    lNumIntItemNF = tItemNF.lNumIntDoc
                
    iTipoOrigem = 0
    
    For iIndiceAux = 1 To objNFiscal.colItensNF.Count
    
        If objNFiscal.colItensNF(iIndiceAux).iPossuiGrade = DESMARCADO Then
            
            If objNFiscal.colItensNF(iIndiceAux).ColAlocacoes.Count > 0 Or objNFiscal.iTipoNFiscal = DOCINFO_DDAI Or objNFiscal.iTipoNFiscal = DOCINFO_FDDAI Then
                lNumIntOrigem = objNFiscal.colItensNF(iIndiceAux).lNumIntDoc 'Alterado por Wagner
                iTipoOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                Exit For
            End If
        Else
        
            For Each objItemRomaneio In objNFiscal.colItensNF(iIndiceAux).colItensRomaneioGrade
                If objItemRomaneio.colLocalizacao.Count > 0 Then
                    iTipoOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                    lNumIntOrigem = objItemRomaneio.lNumIntDoc
                    Exit For
                End If
            Next
            If iTipoOrigem > 0 Then Exit For
        End If
    Next
    
    If iTipoOrigem <> 0 Then
    
        'Lê o movto de estoque da NOta fiscal
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        
        lErro = CF("Customizacoes_MovEstoque", objMovEstoque)
        If lErro <> SUCESSO Then gError 126802
        
        'Verifica se há movimento de estoque ligado ao item
        lErro = Comando_ExecutarPos(alComando(4), "SELECT Codigo, NumIntDocEst FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? AND FilialEmpresa = ?", 0, lNumMovtoEstoque, lNumIntDocEst, iTipoOrigem, lNumIntOrigem, objMovEstoque.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 89023
    
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89024
        
        If lErro = AD_SQL_SUCESSO Then
    
            'Lê o movto de estoque da NOta fiscal
            objMovEstoque.lCodigo = lNumMovtoEstoque
            
            'Lê os movimento dos itens de nota fiscal
            lErro = CF("MovEstoqueItens_Le_Desc1", objMovEstoque)
            If lErro <> SUCESSO And lErro <> 89050 Then gError 89027
        
            If lErro <> SUCESSO Then gError 89026
        
            'Redefine os movimento para fazer o cancelamento
            objMovEstoque.iOperacao = EXCLUSAO_MOVIMENTO
            
            'Grava o movimento de estoque
            lErro = CF("MovEstoque_Grava", objMovEstoque)
            If lErro <> SUCESSO Then gError 89028
        
    ' agora a exclusao do movimento de estorno está sendo feito dentro de MovEstoque_Grava
    '        'se lNumIntDocEst <> 0 indica que o movimento tem um cancelamento/estorno associado que também precisa ser excluido
    '        If lNumIntDocEst <> 0 Then
    '
    '            'Verifica se há movimento de estoque ligado ao item
    '            lErro = Comando_Executar(alComando(10), "SELECT Codigo FROM MovimentoEstoque WHERE NumIntDoc = ?", lCodigo, lNumIntDocEst)
    '            If lErro <> AD_SQL_SUCESSO Then gError 89029
    '
    '            lErro = Comando_BuscarPrimeiro(alComando(10))
    '            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89030
    '
    '            'Lê o movto de estoque da NOta fiscal
    '            objMovEstoque1.iFilialEmpresa = objNFiscal.iFilialEmpresa
    '
    '            objMovEstoque1.lCodigo = lCodigo
    '
    '            'Lê os movimento dos itens de nota fiscal
    '            lErro = CF("MovEstoqueItens_Le_Desc1",objMovEstoque1)
    '            If lErro <> SUCESSO And lErro <> 89050 Then gError 89033
    '
    '            If lErro <> SUCESSO Then gError 89032
    '
    '            'Redefine os movimento para fazer o cancelamento
    '            objMovEstoque1.iOperacao = EXCLUSAO_MOVIMENTO
    '
    '            'Grava o movimento de estoque
    '            lErro = CF("MovEstoque_Grava",objMovEstoque1)
    '            If lErro <> SUCESSO Then gError 89034
    '
    '        End If
        
        End If
    
    End If
    
    'Verifica se o modulo de Livros Fiscais está Ativo
    If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                        
        'Grava o Livro Fiscal a partir da Nota Fiscal
        lErro = CF("NotaFiscal_Exclui_Fis", objNFiscal)
        If lErro <> SUCESSO Then gError 89036
                        
    End If
    
    lErro = CF("NFiscal_ExcluiTributacao", objNFiscal)
    If lErro <> SUCESSO Then gError 89037
    
    lErro = CF("NFe_Exclui_Vinculo_Xml", objNFiscal)
    If lErro <> SUCESSO Then gError 89037
    
    'Enquanto encontrar ItensNF
    Do While lErro1 <> AD_SQL_SEM_DADOS
        
        For Each objItemNF In objNFiscal.colItensNF
            If objItemNF.lNumIntDoc = tItemNF.lNumIntDoc Then Exit For
        Next
        
        objProduto.sCodigo = objItemNF.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 89422
        
        If lErro <> SUCESSO Then gError 89423
        
        objItemNF.iClasseUM = objProduto.iClasseUM
        objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
        objItemNF.sUMVenda = objProduto.sSiglaUMVenda
        
        objItemNF.dQuantidade = tItemNF.dQuantidade
        objItemNF.sUnidadeMed = tItemNF.sUnidadeMed
        objItemNF.lNumIntDocOrig = tItemNF.lNumIntDocOrig
        objItemNF.sProduto = tItemNF.sProduto
        
        'Desfaz baixas de Pedidos de Compras
        lErro = CF("PedidoCompra_Desfaz_Baixa", objNFiscal, lNumIntItemNF, objFilFornFilEmp)
        If lErro <> SUCESSO Then gError 89039
        
        lErro = CF("RequisicaoCompra_Desfaz_Baixa", objNFiscal.lNumIntDoc, lNumIntItemNF)
        If lErro <> SUCESSO Then gError 89040

        lErro = CF("ItensNFGrade_Exclui", tItemNF.lNumIntDoc)
        If lErro <> SUCESSO Then gError 20756
        
        lErro = CF("ItemNFEnt_Excluir_Customizado", objNFiscal, objItemNF)
        If lErro <> SUCESSO Then gError 199707
        
        'Cancela o item
        lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM ItensNFiscal", alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 89041
        
        lErro = CF("MensagensRegra_Exclui", REGRAMSG_TIPODOC_ITEMNF, tItemNF.lNumIntDoc)
        If lErro <> SUCESSO Then gError 89041
                
        objItemNF.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
        objItemNF.objInfoAdicDocItem.lNumIntDocItem = objItemNF.lNumIntDoc
    
        lErro = CF("InfoAdicDocItem_Exclui", objItemNF.objInfoAdicDocItem)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        lErro = Comando_Executar(alComando(11), "DELETE FROM ItemNFResto WHERE NumIntItemNF = ?", _
        tItemNF.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199072
                       
        If objNFiscal.iTipoNFiscal = DOCINFO_NFEE3BF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIE3BF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIERMB3 Or objNFiscal.iTipoNFiscal = DOCINFO_NFEERMB3 Then
        
            'Verifica se houve devoluções para essa remessa
            lErro = Comando_Executar(alComando(13), "SELECT NumIntItemNFRem FROM ItensNFDevBenefRem WHERE NumIntItemNFRem = ? ", lAux, tItemNF.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 211102
        
            lErro = Comando_BuscarPrimeiro(alComando(13))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211103
            
            If lErro = AD_SQL_SUCESSO Then gError 211104
            
            lErro = Comando_Executar(alComando(12), "DELETE FROM ItensNFRemBenef WHERE NumIntItemNF = ?", tItemNF.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 211105
            
        End If
                
        lErro1 = Comando_BuscarProximo(alComando(2))
        If lErro1 <> SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 89042
    
        'Guarda o número do Item
        lNumIntItemNF = tItemNF.lNumIntDoc
    
    Loop
        
    'Grava a Estatística da Filial Fornecedor
    lErro = CF("FilialFornecedor_Grava_Estatisticas", objNFiscal, CANCELAMENTO_FILIALFORNFILEMP, objFilFornFilEmp)
    If lErro <> SUCESSO Then gError 89035
        
    'Grava a Estatística do Produto com Relação a devolução
    lErro = CF("ProdutoVendido_Grava_Estatisticas", objNFiscal, CANCELAMENTO_DOC)
    If lErro <> SUCESSO Then gError 89776
        
    'exclui os movimentos de estoque vinculados a nota (e nao aos itens da nota). Por exemplo, movimentos de beneficiamento.
    lErro = NF_Excluir_MovEstoque(objNFiscal)
    If lErro <> SUCESSO Then gError 126933
        
    'Atualiza o status da nota fiscal
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM NFiscal", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 89043
    
    '############################################
    'Inserido por Wagner 22/03/2006
    lErro = CF("ProjetoInfo_Exclui", PRJ_CR_TIPO_NF, objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO Then gError 181764
    '############################################
    
    lErro = NFiscal_Exclui_Estorno_Cred(objNFiscal)
    If lErro <> SUCESSO Then gError 181764
    
    lErro = CF("InfoAdicionais_Exclui", TIPODOC_INFOADIC_NF, objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO Then gError 181764
        
    lErro = CF("MensagensRegra_Exclui", REGRAMSG_TIPODOC_NF, objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO Then gError 181764
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NotaFiscalEntrada_Excluir_EmTrans = SUCESSO

    Exit Function

Erro_NotaFiscalEntrada_Excluir_EmTrans:

    NotaFiscalEntrada_Excluir_EmTrans = gErr

    Select Case gErr

        Case 20756, 89013, 89018, 89019, 89025, 89027, 89028, 89031, 89033, 89034, 89035, 89036, 89037, 89038, 89039, 89040, 89422, 89776, 126802, 126933, 181764, 207633 To 207635

        Case 89007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89009, 89010, 83915, 83916
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 89011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 89012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", gErr)

        Case 89014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case 89017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_VINCULADA_EXCLUIR", gErr, lNumNotaFiscal, iTipoNF)

        Case 89020, 89021, 89042
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL1", gErr)

        Case 89022
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_SEM_ITENS", gErr, objNFiscal.lNumNotaFiscal)
                    
        Case 89023, 89024, 89029, 89030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 89026, 89032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO2", gErr, objNFiscal.iFilialEmpresa, lCodigo)
        
        Case 89041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMNF", gErr)

        Case 89044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NFISCAL1", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 89423
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
               
        Case 199072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMNFRESTO", gErr)
               
        Case 203045
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)
               
        Case 211102, 211103
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMNFREMBENEF", gErr)
                    
        Case 211104
            Call Rotina_Erro(vbOKOnly, "ERRO_NFREMBENEF_JA_UTILIZADA", gErr)
                    
        Case 211105
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ITEMNFREMBENEF", gErr)
            
        Case ERRO_SEM_MENSAGEM
                              
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149567)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function NF_Excluir_MovEstoque(ByVal objNFiscal As ClassNFiscal) As Long
'exclui os movimentos de estoque ligados a nota fiscal (nao aos itens da nota). Isto envolve os beneficiamentos.

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim objMovEstoque As New ClassMovEstoque
Dim lCodigo As Long

On Error GoTo Erro_NF_Excluir_MovEstoque

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126927
    Next

    'Verifica se há movimento de estoque ligado ao item
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? AND FilialEmpresa = ?", 0, lCodigo, MOVEST_TIPONUMINTDOCORIGEM_NFISCAL, objNFiscal.lNumIntDoc, objNFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 126928

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126929
    
    If lErro = AD_SQL_SUCESSO Then

        'Lê o movto de estoque da NOta fiscal
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        
        'Lê os movimento dos itens de nota fiscal
        lErro = CF("MovEstoqueItens_Le_Desc1", objMovEstoque)
        If lErro <> SUCESSO And lErro <> 89050 Then gError 126930
    
        If lErro <> SUCESSO Then gError 126931
    
        'Redefine os movimento para fazer o cancelamento
        objMovEstoque.iOperacao = EXCLUSAO_MOVIMENTO
        
        'Grava o movimento de estoque
        lErro = CF("MovEstoque_Grava", objMovEstoque)
        If lErro <> SUCESSO Then gError 126932

    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NF_Excluir_MovEstoque = SUCESSO

    Exit Function

Erro_NF_Excluir_MovEstoque:

    NF_Excluir_MovEstoque = gErr

    Select Case gErr

        Case 126927
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 126928, 126929
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 126930, 126932

        Case 126931
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO2", gErr, objNFiscal.iFilialEmpresa, lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149568)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'04/09/01 - Marcelo inclusao da função
Function Embalagem_Exclui(ByVal objEmbalagem As ClassEmbalagem) As Long
'Exclui a embalagem do BD

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 2) As Long
Dim tEmbalagem As typeEmbalagem
Dim iEmbalagem As Integer
Dim sProduto As String
Dim iIndice As Integer

On Error GoTo Erro_Embalagem_Exclui

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 82765
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 82766

    tEmbalagem.sDescricao = String(STRING_EMBALAGEM_DESCRICAO, 0)
    tEmbalagem.sSigla = String(STRING_EMBALAGEM_SIGLA, 0)
    sProduto = String(STRING_PRODUTO, 0)
        
    'Seleciona a Embalagem, com o código passado como parametro, no BD
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Embalagens WHERE Codigo=?", 0, tEmbalagem.iCodigo, objEmbalagem.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 82767
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82768
    
    'Locka a Embalagem
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 82769
    
'    'Verifica se a Embalagem passada como parametro está sendo usada como
'    'embalagem padrao de algum produto
'    lErro = Comando_Executar(alComando(1), "SELECT Embalagem FROM Produtos WHERE Embalagem=?", iEmbalagem, objEmbalagem.iCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 82770
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82771
'
'    'Se a embalagem está associada a algum produto ==> erro
'    If lErro = AD_SQL_SUCESSO Then gError 82772
'
    
    '''04/09/01 - Marcelo - verifica se existe a embalagem na tabela ProdutoEmbalagem
    lErro = Comando_Executar(alComando(1), "SELECT Produto, Embalagem FROM ProdutoEmbalagem WHERE Embalagem=?", sProduto, iEmbalagem, objEmbalagem.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 93574

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93575

    'Se existe a embalagem na tabela ProdutoEmbalagem ==> erro
    If lErro = AD_SQL_SUCESSO Then gError 93576
       
    'por leo
    lErro = EmbalagensExpedicao_Exclui(objEmbalagem)
    If lErro <> AD_SQL_SUCESSO Then gError 103439
    
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM Embalagens", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 82773

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 82774

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Embalagem_Exclui = SUCESSO

    Exit Function
    
Erro_Embalagem_Exclui:

    Embalagem_Exclui = gErr
    
    Select Case gErr
    
        Case 82765
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 82766
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 82767, 82768
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)

        Case 82769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EMBALAGEM", gErr, objEmbalagem.iCodigo)

        Case 82770, 82771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)

        Case 82772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_VINCULADA_PRODUTO", gErr, objEmbalagem.iCodigo)

        Case 82773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EMBALAGEM", gErr, objEmbalagem.iCodigo)

        Case 82774
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 93574, 93575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOEMBALAGEM ", gErr)
            
        Case 93576
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_PRODUTOEMBALAGEM", gErr, sProduto, iEmbalagem)
        
        Case 103439
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149569)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function
'04/09/01 - Marcelo inclusao da função
Function Embalagem_Grava(ByVal objEmbalagem As ClassEmbalagem) As Long
'Grava a embalagem passada como parametro no BD

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim lTransacao As Long
Dim tEmbalagem As typeEmbalagem
Dim iIndice As Integer
Dim iNatureza As Integer
Dim iGerencial As Integer

On Error GoTo Erro_Embalagem_Grava

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 82745
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 82746

    tEmbalagem.sDescricao = String(STRING_EMBALAGEM_DESCRICAO, 0)
    tEmbalagem.sSigla = String(STRING_EMBALAGEM_SIGLA, 0)
    
    'Busca no BD uma embalagem com mesma descrição e com código diferente
    lErro = Comando_Executar(alComando(0), "SELECT Codigo,Descricao FROM Embalagens WHERE Codigo<>? AND Descricao=?", tEmbalagem.iCodigo, tEmbalagem.sDescricao, objEmbalagem.iCodigo, objEmbalagem.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then gError 82747
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82748
    
    'Se encontrou ==> erro
    If lErro = AD_SQL_SUCESSO Then gError 82749
    
    '''04/09/01 - Marcelo - verifica se existe uma embalagem com a mesma sigla e código diferente
    lErro = Comando_Executar(alComando(0), "SELECT Codigo,Sigla FROM Embalagens WHERE Codigo<>? AND Sigla=?", tEmbalagem.iCodigo, tEmbalagem.sSigla, objEmbalagem.iCodigo, objEmbalagem.sSigla)
    If lErro <> AD_SQL_SUCESSO Then gError 93577
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93578
    
    'Se encontrou ==> erro
    If lErro = AD_SQL_SUCESSO Then gError 93579

    'Tenta ler o Produto informado.
    lErro = Comando_ExecutarLockado(alComando(3), "SELECT Natureza, Gerencial FROM Produtos WHERE Codigo=?", iNatureza, iGerencial, objEmbalagem.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 103118
    
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103119
    
    'se não encontrou o Produto, Erro.
    If lErro <> AD_SQL_SUCESSO Then gError 103120
    
    'Se o produto não possui Natureza = Embalagem, Erro.
    If iNatureza <> NATUREZA_PROD_EMBALAGENS Then gError 103121
         
    'Se o produto não possui Natureza = Embalagem, Erro.
    If iGerencial = PRODUTO_GERENCIAL Then gError 103125
         
    'Locka o Produto
    lErro = Comando_LockShared(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then gError 103122
     
    'Busca embalagem com mesmo código no BD
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Embalagens WHERE Codigo=?", 0, tEmbalagem.iCodigo, objEmbalagem.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 82750
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82751
    
    'Se encontrou a Embalagem no BD ==> Alteração
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 82752

        'Atualiza os dados da embalagem
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Embalagens SET Codigo= ?, Sigla= ?, Descricao=?, Capacidade=?, Peso=?, Produto=?", alComando(0), objEmbalagem.iCodigo, objEmbalagem.sSigla, objEmbalagem.sDescricao, objEmbalagem.dCapacidade, objEmbalagem.dPeso, objEmbalagem.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 82753

        'por leo
        lErro = EmbalagensExpedicao_Exclui(objEmbalagem)
        If lErro <> AD_SQL_SUCESSO Then gError 103437
        
    'Se não encontrou a Embalagem, insere
    Else
    
        'Insere a nova embalagem
        lErro = Comando_Executar(alComando(0), "INSERT INTO Embalagens (Codigo, Sigla, Descricao,Capacidade,Peso, Produto) VALUES (?, ?, ?, ?, ?, ?)", objEmbalagem.iCodigo, objEmbalagem.sSigla, objEmbalagem.sDescricao, objEmbalagem.dCapacidade, objEmbalagem.dPeso, objEmbalagem.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 82754
    
    End If
    
    'por leo
    lErro = EmbalagensExpedicao_Inclui(objEmbalagem)
    If lErro <> AD_SQL_SUCESSO Then gError 103438
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 82755

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Embalagem_Grava = SUCESSO

    Exit Function

Erro_Embalagem_Grava:

    Embalagem_Grava = gErr

    Select Case gErr

        Case 82745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 82746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 82747, 82748, 82750, 82751, 93577, 93578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)

        Case 82749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESC_EMBALAGEM_IGUAL", gErr, tEmbalagem.iCodigo)
            
        Case 82752
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EMBALAGEM", gErr, objEmbalagem.iCodigo)

        Case 82753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EMBALAGEM", gErr, objEmbalagem.iCodigo)

        Case 82754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_EMBALAGEM", gErr, objEmbalagem.iCodigo)

        Case 82755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 93579
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_SIGLA_EXISTENTE", gErr, tEmbalagem.iCodigo)
        
        Case 103118, 103119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objEmbalagem.sProduto)
            
        Case 103121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_EMBALAGEM", gErr, objEmbalagem.sProduto)
            
        Case 103120 'nao encontrado
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objEmbalagem.sProduto)
            
        Case 103122
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTO", gErr)
                    
        Case 103125
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objEmbalagem.sProduto)
                                
        Case 103437, 103438
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149570)

    End Select

    Call Transacao_Rollback

    'fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Rotina_Reproc_Zera_SldMesEst(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iMes As Integer
Dim iIndice As Integer
Dim sMes As String
Dim sCampo(1 To 20) As String
Dim iMesInicial As Integer
Dim iMesFinal As Integer

On Error GoTo Erro_Rotina_Reproc_Zera_SldMesEst
    
    'Inicia o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90606
    
    'Preenche o array com o nome dos campos da tabela
    sCampo(1) = "QuantEnt"
    sCampo(2) = "ValorEnt"
    sCampo(3) = "QuantSai"
    sCampo(4) = "QuantCons"
    sCampo(5) = "QuantVend"
    sCampo(6) = "QuantVendConsig3"
    sCampo(7) = "ValorSai"
    sCampo(8) = "CustoMedio"
    sCampo(9) = "CustoProducao"
    sCampo(10) = "CustoStandard"
    sCampo(11) = "Status"
    sCampo(12) = "QuantComp"
    sCampo(13) = "ValorComp"
    sCampo(14) = "ValorVend"
    sCampo(15) = "ValorVendConsig3"
    sCampo(16) = "ValorCons"
    sCampo(17) = "SaldoQuantCusto"
    sCampo(18) = "SaldoValorCusto"
    sCampo(19) = "SaldoQuantRecIndl"
    sCampo(20) = "SaldoValorRecIndl"
    
    'Inicia o comando
    sComandoSQL = "UPDATE SldMesEst SET "
    
    If objReprocessamentoEst.iMes <> 0 Then
        iMesInicial = objReprocessamentoEst.iMes
        iMesFinal = objReprocessamentoEst.iMes
    Else
        iMesInicial = 1
        iMesFinal = 12
    End If
    
    'Para cada campo da tabela
    For iIndice = LBound(sCampo) To UBound(sCampo)
        
        'E para cada mês do ano
        For iMes = iMesInicial To iMesFinal
            sMes = CStr(iMes)
            
            ' Verifica se é o último campo e se é o mês 12
            'Se for, significa que é o último campo do UPDATE e portanto não precisa de vírgula no final
            If iIndice = UBound(sCampo) And iMes = iMesFinal Then
                'Incorpora o mês ao nome do campo e inclui o campo no select sem vírgula
                'E o campo é setado como zero
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0"
            
            'Senão
            Else
                'Incorpora o mês ao nome do campo e inclui o campo no select com a vírgula
                'E o campo é setado como zero
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0, "
            End If
            
        Next
        
    Next
    
    If objReprocessamentoEst.iAno = 0 Then sComandoSQL = sComandoSQL & ",QuantInicial=0,ValorInicial=0,QuantInicialCusto=0,ValorInicialCusto=0"
    
    If objReprocessamentoEst.iAno <> 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & " WHERE Produto = ? AND Ano = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
        'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo, objReprocessamentoEst.iAno)
        If lErro <> SUCESSO Then gError 90604
    
    'Se o produto foi preenchido
    ElseIf Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & " WHERE Produto = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
    
        'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
        If lErro <> SUCESSO Then gError 90604
    
    Else
    
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
        'Executa o comando que irá zerar todos os campos da tabela, sem o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL)
        If lErro <> SUCESSO Then gError 90604
    
    End If

    'Fecha o comando
    Call Comando_Fechar(lComando)
       
    Rotina_Reproc_Zera_SldMesEst = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_SldMesEst:
    
    Rotina_Reproc_Zera_SldMesEst = gErr
    
    Select Case gErr
    
        Case 90606
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90604
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDMESEST", gErr)
        
        Case 90605
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149571)
            
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Rotina_Reproc_Zera_SldMesEst1(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iMes As Integer
Dim iIndice As Integer
Dim sMes As String
Dim sCampo(1 To 15) As String
Dim iMesInicial As Integer
Dim iMesFinal As Integer

On Error GoTo Erro_Rotina_Reproc_Zera_SldMesEst1

    'Inicia o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90609
    
    sCampo(1) = "SaldoQuantConsig3"
    sCampo(2) = "SaldoValorConsig3"
    sCampo(3) = "SaldoQuantDemo3"
    sCampo(4) = "SaldoValorDemo3"
    sCampo(5) = "SaldoQuantConserto3"
    sCampo(6) = "SaldoValorConserto3"
    sCampo(7) = "SaldoQuantOutros3"
    sCampo(8) = "SaldoValorOutros3"
    sCampo(9) = "SaldoQuantBenef3"
    sCampo(10) = "SaldoValorBenef3"
    sCampo(11) = "CustoMedioConsig3"
    sCampo(12) = "CustoMedioConserto3"
    sCampo(13) = "CustoMedioDemo3"
    sCampo(14) = "CustoMedioOutros3"
    sCampo(15) = "CustoMedioBenef3"
    
    sComandoSQL = "UPDATE SldMesEst1 SET "
    
    If objReprocessamentoEst.iMes <> 0 Then
        iMesInicial = objReprocessamentoEst.iMes
        iMesFinal = objReprocessamentoEst.iMes
    Else
        iMesInicial = 1
        iMesFinal = 12
    End If
    
    For iIndice = LBound(sCampo) To UBound(sCampo)
        For iMes = iMesInicial To iMesFinal
            sMes = CStr(iMes)
            If iIndice = UBound(sCampo) And iMes = iMesFinal Then
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0"
            Else
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0, "
            End If
    
        Next
    Next
    
    If objReprocessamentoEst.iAno = 0 Then sComandoSQL = sComandoSQL & ",QuantInicialConsig3=0,ValorInicialConsig3=0,QuantInicialDemo3=0,ValorInicialDemo3=0,QuantInicialConserto3=0,ValorInicialConserto3=0,QuantInicialOutros3=0,ValorInicialOutros3=0,QuantInicialBenef3=0"
    
    If objReprocessamentoEst.iAno <> 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & " WHERE Produto = ? AND Ano = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
    
        'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo, objReprocessamentoEst.iAno)
        If lErro <> SUCESSO Then gError 90610
    
    'Se o produto foi preenchido
    ElseIf Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & "WHERE Produto = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
        'Executa o comando passando o parâmetro de produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
        If lErro <> SUCESSO Then Error 90610
        
    Else
    
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
        'Executa o comando sem o parâmetro de produto
        lErro = Comando_Executar(lComando, sComandoSQL)
        If lErro <> SUCESSO Then Error 90610
    
    End If
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
        
    Rotina_Reproc_Zera_SldMesEst1 = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_SldMesEst1:
    
    Rotina_Reproc_Zera_SldMesEst1 = gErr
    
    Select Case gErr
    
        Case 90609
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90610
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDMESEST1", gErr)
        
        Case 90611
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149572)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Rotina_Reproc_Zera_SldMesEst2(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iMes As Integer
Dim iIndice As Integer
Dim sMes As String
Dim sCampo(1 To 15) As String
Dim iMesInicial As Integer
Dim iMesFinal As Integer

On Error GoTo Erro_Rotina_Reproc_Zera_SldMesEst2

    'Inicia o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90613
    
    sCampo(1) = "SaldoQuantConsig"
    sCampo(2) = "SaldoValorConsig"
    sCampo(3) = "SaldoQuantDemo"
    sCampo(4) = "SaldoValorDemo"
    sCampo(5) = "SaldoQuantConserto"
    sCampo(6) = "SaldoValorConserto"
    sCampo(7) = "SaldoQuantOutros"
    sCampo(8) = "SaldoValorOutros"
    sCampo(9) = "SaldoQuantBenef"
    sCampo(10) = "SaldoValorBenef"
    sCampo(11) = "CustoMedioConsig"
    sCampo(12) = "CustoMedioConserto"
    sCampo(13) = "CustoMedioDemo"
    sCampo(14) = "CustoMedioOutros"
    sCampo(15) = "CustoMedioBenef"
    
    sComandoSQL = "UPDATE SldMesEst2 SET "
    
    If objReprocessamentoEst.iMes <> 0 Then
        iMesInicial = objReprocessamentoEst.iMes
        iMesFinal = objReprocessamentoEst.iMes
    Else
        iMesInicial = 1
        iMesFinal = 12
    End If
    
    For iIndice = LBound(sCampo) To UBound(sCampo)
        For iMes = iMesInicial To iMesFinal
            sMes = CStr(iMes)
            If iIndice = UBound(sCampo) And iMes = iMesFinal Then
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0"
            Else
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0, "
            End If
    
        Next
    Next
    
    If objReprocessamentoEst.iAno = 0 Then sComandoSQL = sComandoSQL & ",QuantInicialConsig=0,ValorInicialConsig=0,QuantInicialDemo=0,ValorInicialDemo=0,QuantInicialConserto=0,ValorInicialConserto=0,QuantInicialOutros=0,ValorInicialOutros=0,QuantInicialBenef=0"
    
    If objReprocessamentoEst.iAno <> 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & " WHERE Produto = ? AND Ano = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
    
        'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo, objReprocessamentoEst.iAno)
        If lErro <> SUCESSO Then gError 90614

    'Se o produto foi preenchido
    ElseIf Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & "WHERE Produto = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
    
        'Executa o comando, passando o parâmetro de produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
        If lErro <> SUCESSO Then gError 90614
    
    Else
        
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
        'Executa o comando, sem o parâmetro de produto
        lErro = Comando_Executar(lComando, sComandoSQL)
        If lErro <> SUCESSO Then gError 90614
    
    End If
    
    Call Comando_Fechar(lComando)
    
    Rotina_Reproc_Zera_SldMesEst2 = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_SldMesEst2:
    
    Rotina_Reproc_Zera_SldMesEst2 = gErr
    
    Select Case gErr
    
        Case 90613
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90614
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDMESEST2", gErr)
        
        Case 90615
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149573)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Rotina_Reproc_Zera_SldMesEstAlm(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iMes As Integer
Dim iIndice As Integer
Dim sMes As String
Dim sCampo(1 To 17) As String
Dim iMesInicial As Integer
Dim iMesFinal As Integer

On Error GoTo Erro_Rotina_Reproc_Zera_SldMesEstAlm

    'Inicia o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90617
    
    sCampo(1) = "QuantEnt"
    sCampo(2) = "ValorEnt"
    sCampo(3) = "QuantSai"
    sCampo(4) = "QuantCons"
    sCampo(5) = "QuantVend"
    sCampo(6) = "QuantVendConsig3"
    sCampo(7) = "ValorSai"
    sCampo(8) = "Status"
    sCampo(9) = "QuantComp"
    sCampo(10) = "ValorComp"
    sCampo(11) = "ValorVend"
    sCampo(12) = "ValorVendConsig3"
    sCampo(13) = "ValorCons"
    sCampo(14) = "SaldoQuantCusto"
    sCampo(15) = "SaldoValorCusto"
    sCampo(16) = "SaldoQuantRecIndl"
    sCampo(17) = "SaldoValorRecIndl"

    sComandoSQL = "UPDATE SldMesEstAlm SET "

    If objReprocessamentoEst.iMes <> 0 Then
        iMesInicial = objReprocessamentoEst.iMes
        iMesFinal = objReprocessamentoEst.iMes
    Else
        iMesInicial = 1
        iMesFinal = 12
    End If

    For iIndice = LBound(sCampo) To UBound(sCampo)
        For iMes = iMesInicial To iMesFinal
            sMes = CStr(iMes)
            If iIndice = UBound(sCampo) And iMes = iMesFinal Then
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0"
            Else
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0, "
            End If
    
        Next
    Next
    
    If objReprocessamentoEst.iAno = 0 Then sComandoSQL = sComandoSQL & ",QuantInicial=0,ValorInicial=0"
    
    If objReprocessamentoEst.iAno <> 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & " WHERE Produto = ? AND Ano = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
    
        'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo, objReprocessamentoEst.iAno)
        If lErro <> SUCESSO Then gError 90618
    
    'Se o produto foi preenchido
    ElseIf Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & "WHERE Produto = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
    
        'Executa o comando, passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
        If lErro <> SUCESSO Then gError 90618
        
    Else
    
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
        
        'Executa o comando, sem o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL)
        If lErro <> SUCESSO Then gError 90618
    
    End If
    
    Call Comando_Fechar(lComando)
    
    Rotina_Reproc_Zera_SldMesEstAlm = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_SldMesEstAlm:
    
    Rotina_Reproc_Zera_SldMesEstAlm = gErr
    
    Select Case gErr
    
        Case 90617
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90618
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDMESESTALM", gErr)
        
        Case 90619
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149574)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
   
End Function

Function Rotina_Reproc_Zera_SldMesEstAlm1(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iMes As Integer
Dim iIndice As Integer
Dim sMes As String
Dim sCampo(1 To 10) As String
Dim iMesInicial As Integer
Dim iMesFinal As Integer

On Error GoTo Erro_Rotina_Reproc_Zera_SldMesEstAlm1

    'Inicia o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90621
    
    sCampo(1) = "SaldoQuantConsig3"
    sCampo(2) = "SaldoValorConsig3"
    sCampo(3) = "SaldoQuantDemo3"
    sCampo(4) = "SaldoValorDemo3"
    sCampo(5) = "SaldoQuantConserto3"
    sCampo(6) = "SaldoValorConserto3"
    sCampo(7) = "SaldoQuantOutros3"
    sCampo(8) = "SaldoValorOutros3"
    sCampo(9) = "SaldoQuantBenef3"
    sCampo(10) = "SaldoValorBenef3"
    
    sComandoSQL = "UPDATE SldMesEstAlm1 SET "
    
    If objReprocessamentoEst.iMes <> 0 Then
        iMesInicial = objReprocessamentoEst.iMes
        iMesFinal = objReprocessamentoEst.iMes
    Else
        iMesInicial = 1
        iMesFinal = 12
    End If
    
    For iIndice = LBound(sCampo) To UBound(sCampo)
        For iMes = iMesInicial To iMesFinal
            sMes = CStr(iMes)
            If iIndice = UBound(sCampo) And iMes = iMesFinal Then
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0"
            Else
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0, "
            End If
    
        Next
    Next
    
    If objReprocessamentoEst.iAno = 0 Then sComandoSQL = sComandoSQL & ",QuantInicialConsig3=0,ValorInicialConsig3=0,QuantInicialDemo3=0,ValorInicialDemo3=0,QuantInicialConserto3=0,ValorInicialConserto3=0,QuantInicialOutros3=0,ValorInicialOutros3=0,QuantInicialBenef3=0"
    
    If objReprocessamentoEst.iAno <> 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & " WHERE Produto = ? AND Ano = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
    
        'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo, objReprocessamentoEst.iAno)
        If lErro <> SUCESSO Then gError 90622
    
    'Se o produto foi preenchido
    ElseIf Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & "WHERE Produto = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
        
        'Executa o comando, passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
        If lErro <> SUCESSO Then gError 90622
        
    Else
        
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
        
        'Executa o comando, sem o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL)
        If lErro <> SUCESSO Then gError 90622
    
    End If
    
    Call Comando_Fechar(lComando)
    
    Rotina_Reproc_Zera_SldMesEstAlm1 = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_SldMesEstAlm1:
    
    Rotina_Reproc_Zera_SldMesEstAlm1 = gErr
    
    Select Case gErr
    
        Case 90621
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90622
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDMESESTALM1", gErr)
        
        Case 90623
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149575)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
   
End Function

Function Rotina_Reproc_Zera_SldMesEstAlm2(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iMes As Integer
Dim iIndice As Integer
Dim sMes As String
Dim sCampo(1 To 10) As String
Dim iMesInicial As Integer
Dim iMesFinal As Integer

On Error GoTo Erro_Rotina_Reproc_Zera_SldMesEstAlm2

    'Inicia o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90625
    
    sCampo(1) = "SaldoQuantConsig"
    sCampo(2) = "SaldoValorConsig"
    sCampo(3) = "SaldoQuantDemo"
    sCampo(4) = "SaldoValorDemo"
    sCampo(5) = "SaldoQuantConserto"
    sCampo(6) = "SaldoValorConserto"
    sCampo(7) = "SaldoQuantOutros"
    sCampo(8) = "SaldoValorOutros"
    sCampo(9) = "SaldoQuantBenef"
    sCampo(10) = "SaldoValorBenef"
    
    sComandoSQL = "UPDATE SldMesEstAlm2 SET "
    
    If objReprocessamentoEst.iMes <> 0 Then
        iMesInicial = objReprocessamentoEst.iMes
        iMesFinal = objReprocessamentoEst.iMes
    Else
        iMesInicial = 1
        iMesFinal = 12
    End If
    
    For iIndice = LBound(sCampo) To UBound(sCampo)
        For iMes = iMesInicial To iMesFinal
            sMes = CStr(iMes)
            If iIndice = UBound(sCampo) And iMes = iMesFinal Then
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0"
            Else
                sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0, "
            End If
    
        Next
    Next
    
    If objReprocessamentoEst.iAno = 0 Then sComandoSQL = sComandoSQL & ",QuantInicialConsig=0,ValorInicialConsig=0,QuantInicialDemo=0,ValorInicialDemo=0,QuantInicialConserto=0,ValorInicialConserto=0,QuantInicialOutros=0,ValorInicialOutros=0,QuantInicialBenef=0"
    
    If objReprocessamentoEst.iAno <> 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & " WHERE Produto = ? AND Ano = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
    
        'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo, objReprocessamentoEst.iAno)
        If lErro <> SUCESSO Then gError 90626
    
    'Se o produto foi preenchido
    ElseIf Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & "WHERE Produto = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
        
        'Executa o comando passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
        If lErro <> SUCESSO Then gError 90626
    
    Else
    
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
        
        'Executa o comando sem o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL)
        If lErro <> SUCESSO Then gError 90626
    
    End If
    
    Call Comando_Fechar(lComando)
    
    Rotina_Reproc_Zera_SldMesEstAlm2 = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_SldMesEstAlm2:
    
    Rotina_Reproc_Zera_SldMesEstAlm2 = gErr
    
    Select Case gErr
    
        Case 90625
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90626
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDMESESTALM2", gErr)
        
        Case 90627
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149576)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
   
End Function

Function Rotina_Reproc_Zera_SldDiaEst(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iIndice As Integer
Dim sCampo(1 To 60) As String
Dim dtDataInicial As Date
Dim dtDataFinal As Date

On Error GoTo Erro_Rotina_Reproc_Zera_SldDiaEst

    'Inicia o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90629
    
    sCampo(1) = "QuantEntrada"
    sCampo(2) = "QuantSaida"
    sCampo(3) = "QuantCons"
    sCampo(4) = "QuantVend"
    sCampo(5) = "QuantVendConsig3"
    sCampo(6) = "ValorEntrada"
    sCampo(7) = "ValorSaida"
    sCampo(8) = "QuantComp"
    sCampo(9) = "ValorComp"
    sCampo(10) = "ValorVend"
    sCampo(11) = "ValorCons"
    sCampo(12) = "ValorVendConsig3"
    sCampo(13) = "QuantEntCusto"
    sCampo(14) = "ValorEntCusto"
    sCampo(15) = "QuantSaiCusto"
    sCampo(16) = "ValorSaiCusto"
    sCampo(17) = "QuantEntConsig3"
    sCampo(18) = "ValorEntConsig3"
    sCampo(19) = "QuantSaiConsig3"
    sCampo(20) = "ValorSaiConsig3"
    sCampo(21) = "QuantEntDemo3"
    sCampo(22) = "ValorEntDemo3"
    sCampo(23) = "QuantSaiDemo3"
    sCampo(24) = "ValorSaiDemo3"
    sCampo(25) = "QuantEntConserto3"
    sCampo(26) = "ValorEntConserto3"
    sCampo(27) = "QuantSaiConserto3"
    sCampo(28) = "ValorSaiConserto3"
    sCampo(29) = "QuantEntOutros3"
    sCampo(30) = "ValorEntOutros3"
    sCampo(31) = "QuantSaiOutros3"
    sCampo(32) = "ValorSaiOutros3"
    sCampo(33) = "QuantEntBenef3"
    sCampo(34) = "ValorEntBenef3"
    sCampo(35) = "QuantSaiBenef3"
    sCampo(36) = "ValorSaiBenef3"
    sCampo(37) = "QuantEntConsig"
    sCampo(38) = "ValorEntConsig"
    sCampo(39) = "QuantSaiConsig"
    sCampo(40) = "ValorSaiConsig"
    sCampo(41) = "QuantEntDemo"
    sCampo(42) = "ValorEntDemo"
    sCampo(43) = "QuantSaiDemo"
    sCampo(44) = "ValorSaiDemo"
    sCampo(45) = "QuantEntConserto"
    sCampo(46) = "ValorEntConserto"
    sCampo(47) = "QuantSaiConserto"
    sCampo(48) = "ValorSaiConserto"
    sCampo(49) = "QuantEntOutros"
    sCampo(50) = "ValorEntOutros"
    sCampo(51) = "QuantSaiOutros"
    sCampo(52) = "ValorSaiOutros"
    sCampo(53) = "QuantEntBenef"
    sCampo(54) = "ValorEntBenef"
    sCampo(55) = "QuantSaiBenef"
    sCampo(56) = "ValorSaiBenef"
    sCampo(57) = "QuantEntRecIndl"
    sCampo(58) = "ValorEntRecIndl"
    sCampo(59) = "QuantSaiRecIndl"
    sCampo(60) = "ValorSaiRecIndl"
      
    sComandoSQL = "UPDATE SldDiaEst SET "
    
    For iIndice = LBound(sCampo) To UBound(sCampo)
        If iIndice = UBound(sCampo) Then
            sComandoSQL = sComandoSQL & sCampo(iIndice) & "=0"
        Else
            sComandoSQL = sComandoSQL & sCampo(iIndice) & "=0, "
        End If
    
    Next
    
    If objReprocessamentoEst.iAno <> 0 Then
    
        dtDataInicial = CDate("01/" & CStr(objReprocessamentoEst.iMes) & "/" & CStr(objReprocessamentoEst.iAno))
        dtDataFinal = DateAdd("m", 1, dtDataInicial) - 1
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & " WHERE Produto = ? AND Data >= ? AND Data <= ? "
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
    
        'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo, dtDataInicial, dtDataFinal)
        If lErro <> SUCESSO Then gError 90630
    
    'Se o produto foi preenchido
    ElseIf Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & "WHERE Produto = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
        'Executa o comando passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
        If lErro <> SUCESSO Then gError 90630
    
    Else
    
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
        'Executa o comando sem o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL)
        If lErro <> SUCESSO Then gError 90630
    
    End If
    
    Call Comando_Fechar(lComando)
    
    Rotina_Reproc_Zera_SldDiaEst = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_SldDiaEst:
    
    Rotina_Reproc_Zera_SldDiaEst = gErr
    
    Select Case gErr
    
        Case 90629
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90630
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDDIAEST", gErr)
        
        Case 90631
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149577)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Rotina_Reproc_Zera_SldDiaEstAlm(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iIndice As Integer
Dim sCampo(1 To 60) As String
Dim dtDataInicial As Date
Dim dtDataFinal As Date

On Error GoTo Erro_Rotina_Reproc_Zera_SldDiaEstAlm

    'Inicia o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90637
    
    sCampo(1) = "QuantEntrada"
    sCampo(2) = "QuantSaida"
    sCampo(3) = "QuantCons"
    sCampo(4) = "QuantVend"
    sCampo(5) = "QuantVendConsig3"
    sCampo(6) = "ValorEntrada"
    sCampo(7) = "ValorSaida"
    sCampo(8) = "QuantComp"
    sCampo(9) = "ValorComp"
    sCampo(10) = "ValorVend"
    sCampo(11) = "ValorCons"
    sCampo(12) = "QuantEntCusto"
    sCampo(13) = "ValorEntCusto"
    sCampo(14) = "QuantSaiCusto"
    sCampo(15) = "ValorSaiCusto"
    sCampo(16) = "ValorVendConsig3"
    sCampo(17) = "QuantEntConsig3"
    sCampo(18) = "ValorEntConsig3"
    sCampo(19) = "QuantSaiConsig3"
    sCampo(20) = "ValorSaiConsig3"
    sCampo(21) = "QuantEntDemo3"
    sCampo(22) = "ValorEntDemo3"
    sCampo(23) = "QuantSaiDemo3"
    sCampo(24) = "ValorSaiDemo3"
    sCampo(25) = "QuantEntConserto3"
    sCampo(26) = "ValorEntConserto3"
    sCampo(27) = "QuantSaiConserto3"
    sCampo(28) = "ValorSaiConserto3"
    sCampo(29) = "QuantEntOutros3"
    sCampo(30) = "ValorEntOutros3"
    sCampo(31) = "QuantSaiOutros3"
    sCampo(32) = "ValorSaiOutros3"
    sCampo(33) = "QuantEntBenef3"
    sCampo(34) = "ValorEntBenef3"
    sCampo(35) = "QuantSaiBenef3"
    sCampo(36) = "ValorSaiBenef3"
    sCampo(37) = "QuantEntConsig"
    sCampo(38) = "ValorEntConsig"
    sCampo(39) = "QuantSaiConsig"
    sCampo(40) = "ValorSaiConsig"
    sCampo(41) = "QuantEntDemo"
    sCampo(42) = "ValorEntDemo"
    sCampo(43) = "QuantSaiDemo"
    sCampo(44) = "ValorSaiDemo"
    sCampo(45) = "QuantEntConserto"
    sCampo(46) = "ValorEntConserto"
    sCampo(47) = "QuantSaiConserto"
    sCampo(48) = "ValorSaiConserto"
    sCampo(49) = "QuantEntOutros"
    sCampo(50) = "ValorEntOutros"
    sCampo(51) = "QuantSaiOutros"
    sCampo(52) = "ValorSaiOutros"
    sCampo(53) = "QuantEntBenef"
    sCampo(54) = "ValorEntBenef"
    sCampo(55) = "QuantSaiBenef"
    sCampo(56) = "ValorSaiBenef"
    sCampo(57) = "QuantEntRecIndl"
    sCampo(58) = "ValorEntRecIndl"
    sCampo(59) = "QuantSaiRecIndl"
    sCampo(60) = "ValorSaiRecIndl"
      
    sComandoSQL = "UPDATE SldDiaEstAlm SET "
    
    For iIndice = LBound(sCampo) To UBound(sCampo)
        
        If iIndice = UBound(sCampo) Then
            sComandoSQL = sComandoSQL & sCampo(iIndice) & "=0"
        Else
            sComandoSQL = sComandoSQL & sCampo(iIndice) & "=0, "
        End If
    
    Next
    
    If objReprocessamentoEst.iAno <> 0 Then
    
        dtDataInicial = CDate("01/" & CStr(objReprocessamentoEst.iMes) & "/" & CStr(objReprocessamentoEst.iAno))
        dtDataFinal = DateAdd("m", 1, dtDataInicial) - 1
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & " WHERE Produto = ? AND Data >= ? AND Data <= ? "
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
    
        'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo, dtDataInicial, dtDataFinal)
        If lErro <> SUCESSO Then gError 90638
    
    'Se o produto foi preenchido
    ElseIf Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & "WHERE Produto = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
        
        'Executa o comando passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
        If lErro <> SUCESSO Then gError 90638
    
    Else
    
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
        
        'Executa o comando sem o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL)
        If lErro <> SUCESSO Then gError 90638
    
    End If

    Call Comando_Fechar(lComando)
    
    Rotina_Reproc_Zera_SldDiaEstAlm = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_SldDiaEstAlm:
    
    Rotina_Reproc_Zera_SldDiaEstAlm = gErr
    
    Select Case gErr
    
        Case 90637
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90638
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDDIAESTALM", gErr)
        
        Case 90639
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149578)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
   
End Function

Function Rotina_Reproc_Zera_EstoqueProduto(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iIndice As Integer
Dim sCampo(1 To 24) As String
Dim sCampoInicial(1 To 24) As String

On Error GoTo Erro_Rotina_Reproc_Zera_EstoqueProduto

    'Inicia o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90633

' *** Campos de onde serão retirados os dados para atualização ***
    
    sCampoInicial(1) = "QuantidadeInicial - QuantReservada"
    sCampoInicial(2) = CAMPOINICIAL_INEXISTENTE
    sCampoInicial(3) = CAMPOINICIAL_INEXISTENTE
    sCampoInicial(4) = CAMPOINICIAL_INEXISTENTE
    sCampoInicial(5) = "QuantInicialConsig3"
    sCampoInicial(6) = "QuantInicialConsig"
    sCampoInicial(7) = "QuantInicialDemo3"
    sCampoInicial(8) = "QuantInicialDemo"
    sCampoInicial(9) = "QuantInicialConserto3"
    sCampoInicial(10) = "QuantInicialConserto"
    sCampoInicial(11) = "QuantInicialOutras3"
    sCampoInicial(12) = "QuantInicialOutras"
    sCampoInicial(13) = "QuantInicialBenef"
    sCampoInicial(14) = "QuantInicialBenef3"
    sCampoInicial(15) = "ValorInicialConsig3"
    sCampoInicial(16) = "ValorInicialConsig"
    sCampoInicial(17) = "ValorInicialDemo3"
    sCampoInicial(18) = "ValorInicialDemo"
    sCampoInicial(19) = "ValorInicialConserto3"
    sCampoInicial(20) = "ValorInicialConserto"
    sCampoInicial(21) = "ValorInicialOutras3"
    sCampoInicial(22) = "ValorInicialOutras"
    sCampoInicial(23) = "ValorInicialBenef"
    sCampoInicial(24) = "ValorInicialBenef3"
'*********************************************************************

' *** Campos que serão atualizados ***

    sCampo(1) = "QuantDispNossa"
    sCampo(2) = "QuantRecIndl"
    sCampo(3) = "QuantIndOutras"
    sCampo(4) = "QuantDefeituosa"
    sCampo(5) = "QuantConsig3"
    sCampo(6) = "QuantConsig"
    sCampo(7) = "QuantDemo3"
    sCampo(8) = "QuantDemo"
    sCampo(9) = "QuantConserto3"
    sCampo(10) = "QuantConserto"
    sCampo(11) = "QuantOutras3"
    sCampo(12) = "QuantOutras"
    sCampo(13) = "QuantBenef"
    sCampo(14) = "QuantBenef3"
    sCampo(15) = "ValorConsig3"
    sCampo(16) = "ValorConsig"
    sCampo(17) = "ValorDemo3"
    sCampo(18) = "ValorDemo"
    sCampo(19) = "ValorConserto3"
    sCampo(20) = "ValorConserto"
    sCampo(21) = "ValorOutras3"
    sCampo(22) = "ValorOutras"
    sCampo(23) = "ValorBenef"
    sCampo(24) = "ValorBenef3"
' ***********************************************
          
    sComandoSQL = "UPDATE EstoqueProduto SET "
    
    For iIndice = LBound(sCampo) To UBound(sCampo)
        
        'Se não é o último campo do select e não existe campo inicial para ele
        If sCampoInicial(iIndice) = CAMPOINICIAL_INEXISTENTE Then
            
            'Seta o campo com valor Zero e inclui vírgula
            sComandoSQL = sComandoSQL & sCampo(iIndice) & "=0,"
            
        'Se não é o último campo do select, mas existe campo inicial para ele
        Else
                
            'Seta o campo com o valor inicial e inclui vírgula
            sComandoSQL = sComandoSQL & sCampo(iIndice) & "=" & sCampoInicial(iIndice) & ","
        
        End If
    
    Next
    
    If right(sComandoSQL, 1) = "," Then
        sComandoSQL = left(sComandoSQL, Len(sComandoSQL) - 1)
    End If
    
    'Se o produto foi preenchido
    If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & " WHERE Produto = ?"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
        
        'Executa o comando passando o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
        If lErro <> SUCESSO Then gError 90634
    
    Else
    
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE Almoxarifado IN (SELECT Codigo FROM Almoxarifado WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & "))"
    
        'Executa o comando sem o parâmetro produto
        lErro = Comando_Executar(lComando, sComandoSQL)
        If lErro <> SUCESSO Then gError 90634
    
    End If
    
    Call Comando_Fechar(lComando)
    
    Rotina_Reproc_Zera_EstoqueProduto = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_EstoqueProduto:
    
    Rotina_Reproc_Zera_EstoqueProduto = gErr
    
    Select Case gErr
    
        Case 90633
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90634
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_ESTOQUEPRODUTO", gErr)
        
        Case 90635
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149579)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Zera_Totalizadores_Estoque(ByVal iAno As Integer, ByVal iMes As Integer) As Long
'zera os totalizadores de estoque SldMesEst, SldMesEst1, SldMesEst2, SldMesEstAlm, SldMesEstAlm1, SldMesEstAlm2
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objReprocessamentoEst As New ClassReprocessamentoEST
Dim sCodigo As String

On Error GoTo Erro_Zera_Totalizadores_Estoque

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92644
    Next

    objReprocessamentoEst.iAno = iAno
    objReprocessamentoEst.iMes = iMes

    sCodigo = String(STRING_PRODUTO, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM Produtos WHERE ControleEstoque <> ? AND Compras = ?", 0, sCodigo, PRODUTO_SEM_ESTOQUE, PRODUTO_PRODUZIVEL)
    If lErro <> AD_SQL_SUCESSO Then gError 92645

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92646

    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92647
    
        objReprocessamentoEst.sProdutoCodigo = sCodigo
        
        lErro = Zera_Totalizadores_Estoque1(objReprocessamentoEst)
        If lErro <> SUCESSO Then gError 92648
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92649
        
    Loop
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Zera_Totalizadores_Estoque = SUCESSO
    
    Exit Function
    
Erro_Zera_Totalizadores_Estoque:

    Zera_Totalizadores_Estoque = gErr

    Select Case gErr
    
        Case 92644
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 92645, 92646, 92649
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
    
        Case 92647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOS", gErr)
    
        Case 92648
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149580)
    
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Zera_Totalizadores_Estoque1(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'zera os totalizadores de estoque SldMesEst, SldMesEst1, SldMesEst2, SldMesEstAlm, SldMesEstAlm1, SldMesEstAlm2
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long

On Error GoTo Erro_Zera_Totalizadores_Estoque1

    lErro = CF("Rotina_Reproc_Zera_SldMesEst", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 92636
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEst1", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 92637
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEst2", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 92638
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEstAlm", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 92639
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEstAlm1", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 92640
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEstAlm2", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 92641
    
    lErro = CF("Rotina_Reproc_Zera_SldDiaEst", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 92642
    
    lErro = CF("Rotina_Reproc_Zera_SldDiaEstAlm", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 92643
    
    lErro = CF("Rotina_Reproc_Zera_SldDiaEstTerc", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 92643
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEst1Terc", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 92643
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEst2Terc", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 92643
    
    Zera_Totalizadores_Estoque1 = SUCESSO
    
    Exit Function
    
Erro_Zera_Totalizadores_Estoque1:

    Zera_Totalizadores_Estoque1 = gErr
    
    Select Case gErr
    
        Case 92636 To 92643
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149581)
        
    End Select
    
    Exit Function

End Function


Sub BaixaItensOP(dtDataFinal As Date)
'Baixa todos os itens de op que tenham pelo menos 65% produzido e tenham tido producao entrada até dtDataFinal
'ex.: Call BaixaItensOP(#08/31/2001#)

Dim lErro As Long, lComando As Long, objItemOP As ClassItemOP, ColItemOP As New Collection
Dim tItemOP As typeItemOP

On Error GoTo Erro_BaixaItensOP

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81741
    
    tItemOP.sCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
    tItemOP.sProduto = String(STRING_PRODUTO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Codigo, Item, Produto FROM ItensOrdemProducao WHERE FilialEmpresa = ? AND QuantidadeProd > (Quantidade * 0.65) AND NOT EXISTS (SELECT NumIntDoc FROM MovimentoEstoque WHERE TipoMov = 7 AND MovimentoEstoque.CodigoOP = ItensOrdemProducao.Codigo AND MovimentoEstoque.FilialEmpresa = ItensOrdemProducao.FilialEmpresa AND MovimentoEstoque.Produto = ItensOrdemProducao.Produto AND Data > ?)", tItemOP.lNumIntDoc, tItemOP.sCodigo, tItemOP.iItem, tItemOP.sProduto, giFilialEmpresa, dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then gError 81742
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81743
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objItemOP = New ClassItemOP
        
        With objItemOP
            .iFilialEmpresa = giFilialEmpresa
            .lNumIntDoc = tItemOP.lNumIntDoc
            .sCodigo = tItemOP.sCodigo
            .iItem = tItemOP.iItem
            .sProduto = tItemOP.sProduto
        End With
        
        Call ColItemOP.Add(objItemOP)
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81744
    
    Loop
    
    lErro = CreateObject("RotinasMATInpal.ClassMATGravaInpal").ItensOP_Baixa(ColItemOP)
    If lErro <> SUCESSO Then gError 81745
    
    Call Comando_Fechar(lComando)

    Exit Sub
     
Erro_BaixaItensOP:

    Select Case gErr
          
        Case 81741
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 81742 To 81745
            Call Rotina_Erro(vbOKOnly, "OUTROS_ERROS_BaixaItensOP", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149582)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Sub
    
End Sub

Sub Integridade_movsEst_EstoqueProduto()
'valida saldos atuais na tabela estoque produto com qtdes iniciais

Dim lErro As Long, iIndice As Integer, alComando(1 To 5) As Long, tProduto As typeProduto
Dim testprod As typeEstoqueProduto, objRegInventario As ClassRegInventario, tItemMovEstoque As typeItemMovEstoque, tTipoMovEstoque As typeTipoMovEst
Dim sComandoSQL As String, dFator As Double, lTransacao As Long, dQuantDif As Double

On Error GoTo Erro_Integridade_movsEst_EstoqueProduto

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69911
    Next
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 69911
    
    lErro = Comando_Executar(alComando(5), "DELETE FROM ErrosEstProd")
    If lErro <> AD_SQL_SUCESSO Then gError 69921
    
    With testprod
    
        .sProduto = String(STRING_PRODUTO, 0)
        tProduto.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    
        lErro = Comando_Executar(alComando(1), "SELECT Produtos.Codigo, Produtos.ClasseUM, Produtos.SiglaUMEstoque, Almoxarifado, " & _
            "QuantidadeInicial , QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantRecIndl, QuantIndOutras, " & _
            "QuantDefeituosa , QuantConsig3, QuantConsig, QuantDemo3, " & _
            "QuantDemo , QuantConserto3, QuantConserto, QuantOutras3, " & _
            "QuantOutras, QuantBenef, QuantBenef3, " & _
            "QuantInicialConsig3 , QuantInicialConsig, QuantInicialDemo3, QuantInicialDemo, " & _
            "QuantInicialConserto3 , QuantInicialConserto, QuantInicialOutras3, QuantInicialOutras, QuantInicialBenef, " & _
            "QuantInicialBenef3 FROM Produtos, EstoqueProduto WHERE Produtos.Codigo = EstoqueProduto.Produto ORDER BY EstoqueProduto.Produto, EstoqueProduto.Almoxarifado", _
            .sProduto, tProduto.iClasseUM, tProduto.sSiglaUMEstoque, .iAlmoxarifado, _
            .dQuantidadeInicial, .dQuantDispNossa, .dQuantReservada, .dQuantReservadaConsig, .dQuantRecIndl, .dQuantInd, _
            .dQuantDefeituosa, .dQuantConsig3, .dQuantConsig, .dQuantDemo3, _
            .dQuantDemo, .dQuantConserto3, .dQuantConserto, .dQuantOutras3, _
            .dQuantOutras, .dQuantBenef, .dQuantBenef3, _
            .dQuantInicialConsig3, .dQuantInicialConsig, .dQuantInicialDemo3, .dQuantInicialDemo, _
            .dQuantInicialConserto3, .dQuantInicialConserto, .dQuantInicialOutras3, .dQuantInicialOutras, .dQuantInicialBenef, _
            .dQuantInicialBenef3)
    End With
        
    If lErro <> AD_SQL_SUCESSO Then gError 69921
        
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69922
    
    Do While lErro = AD_SQL_SUCESSO
                
        Set objRegInventario = New ClassRegInventario
        objRegInventario.sProduto = testprod.sProduto
        objRegInventario.iAlmoxarifado = testprod.iAlmoxarifado
        
        'Inicializa as variaveis
        tItemMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
        tTipoMovEstoque.sEntradaOuSaida = String(STRING_ENTRADAOUSAIDA, 0)
    
        'Lê os Movimentos do par produto-almoxarifado
        sComandoSQL = "SELECT MovimentoEstoque.Quantidade, MovimentoEstoque.SiglaUM, TiposMovimentoEstoque.AtualizaConsig, TiposMovimentoEstoque.AtualizaDemo, TiposMovimentoEstoque.AtualizaConserto, TiposMovimentoEstoque.AtualizaOutras, TiposMovimentoEstoque.AtualizaBenef, TiposMovimentoEstoque.EntradaOuSaida, TiposMovimentoEstoque.ProdutodeTerc, TiposMovimentoEstoque.AtualizaNossaDisp FROM TiposMovimentoEstoque, MovimentoEstoque WHERE TiposMovimentoEstoque.Codigo = MovimentoEstoque.TipoMov AND MovimentoEstoque.Produto = ? AND MovimentoEstoque.Almoxarifado = ? AND TiposMovimentoEstoque.AtualizaSoLote = 0"
    
        lErro = Comando_Executar(alComando(4), sComandoSQL, tItemMovEstoque.dQuantidade, tItemMovEstoque.sSiglaUM, tTipoMovEstoque.iAtualizaConsig, tTipoMovEstoque.iAtualizaDemo, tTipoMovEstoque.iAtualizaConserto, tTipoMovEstoque.iAtualizaOutras, tTipoMovEstoque.iAtualizaBenef, tTipoMovEstoque.sEntradaOuSaida, tTipoMovEstoque.iProdutoDeTerc, tTipoMovEstoque.iAtualizaNossaDisp, objRegInventario.sProduto, objRegInventario.iAlmoxarifado)
        If lErro <> AD_SQL_SUCESSO Then gError 69921
    
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69922
    
        Do While lErro = AD_SQL_SUCESSO
    
            'Fator de conv. de UMs
            lErro = CF("UM_Conversao", tProduto.iClasseUM, tItemMovEstoque.sSiglaUM, tProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 69923
    
            'Verifica em que escaninho ele faz parte
            If tTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
    
                If tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
                    objRegInventario.dQuantConserto = objRegInventario.dQuantConserto + tItemMovEstoque.dQuantidade * dFator
                ElseIf tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
                    objRegInventario.dQuantConserto = objRegInventario.dQuantConserto - tItemMovEstoque.dQuantidade * dFator
                End If
    
                If tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
                    objRegInventario.dQuantConsig = objRegInventario.dQuantConsig + tItemMovEstoque.dQuantidade * dFator
                ElseIf tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
                    objRegInventario.dQuantConsig = objRegInventario.dQuantConsig - tItemMovEstoque.dQuantidade * dFator
                End If
    
                If tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
                    objRegInventario.dQuantDemo = objRegInventario.dQuantDemo + tItemMovEstoque.dQuantidade * dFator
                ElseIf tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
                    objRegInventario.dQuantDemo = objRegInventario.dQuantDemo - tItemMovEstoque.dQuantidade * dFator
                End If
    
                If tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
                    objRegInventario.dQuantOutras = objRegInventario.dQuantOutras + tItemMovEstoque.dQuantidade * dFator
                ElseIf tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
                    objRegInventario.dQuantOutras = objRegInventario.dQuantOutras - tItemMovEstoque.dQuantidade * dFator
                End If
    
                If tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
                    objRegInventario.dQuantBenef = objRegInventario.dQuantBenef + tItemMovEstoque.dQuantidade * dFator
                ElseIf tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
                    objRegInventario.dQuantBenef = objRegInventario.dQuantBenef - tItemMovEstoque.dQuantidade * dFator
                End If
    
                If tTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
                    objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque + tItemMovEstoque.dQuantidade * dFator
                ElseIf tTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
                    objRegInventario.dQuantidadeUMEstoque = objRegInventario.dQuantidadeUMEstoque - tItemMovEstoque.dQuantidade * dFator
                End If
        
            ElseIf tTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS Then
    
                If tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
                    objRegInventario.dQuantConserto3 = objRegInventario.dQuantConserto3 + tItemMovEstoque.dQuantidade * dFator
                ElseIf tTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
                    objRegInventario.dQuantConserto3 = objRegInventario.dQuantConserto3 - tItemMovEstoque.dQuantidade * dFator
                End If
    
                If tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
                    objRegInventario.dQuantConsig3 = objRegInventario.dQuantConsig3 + tItemMovEstoque.dQuantidade * dFator
                ElseIf tTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
                    objRegInventario.dQuantConsig3 = objRegInventario.dQuantConsig3 - tItemMovEstoque.dQuantidade * dFator
                End If
    
                If tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
                    objRegInventario.dQuantDemo3 = objRegInventario.dQuantDemo3 + tItemMovEstoque.dQuantidade * dFator
                ElseIf tTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
                    objRegInventario.dQuantDemo3 = objRegInventario.dQuantDemo3 - tItemMovEstoque.dQuantidade * dFator
                End If
    
                If tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
                    objRegInventario.dQuantOutras3 = objRegInventario.dQuantOutras3 + tItemMovEstoque.dQuantidade * dFator
                ElseIf tTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
                    objRegInventario.dQuantOutras3 = objRegInventario.dQuantOutras3 - tItemMovEstoque.dQuantidade * dFator
                End If
    
                If tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
                    objRegInventario.dQuantBenef3 = objRegInventario.dQuantBenef3 + tItemMovEstoque.dQuantidade * dFator
                ElseIf tTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
                    objRegInventario.dQuantBenef3 = objRegInventario.dQuantBenef3 - tItemMovEstoque.dQuantidade * dFator
                End If
    
            End If
        
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69924
    
        Loop
    
        'compara os saldos atuais de EstoqueProduto com os saldos obtidos pelos movtos mais estoque inicial
        
        dQuantDif = testprod.dQuantidadeInicial - testprod.dQuantDispNossa - testprod.dQuantReservada + objRegInventario.dQuantidadeUMEstoque
        If Abs(dQuantDif) > QTDE_ESTOQUE_DELTA Then
            
            If testprod.dQuantDispNossa + dQuantDif > 0 Then
            
                lErro = Comando_Executar(alComando(2), "UPDATE EstoqueProduto SET QuantDispNossa = QuantDispNossa + ? WHERE Produto = ? AND Almoxarifado = ?", dQuantDif, objRegInventario.sProduto, objRegInventario.iAlmoxarifado)
                If lErro <> AD_SQL_SUCESSO Then gError 69921
                
            End If
                        
            lErro = Comando_Executar(alComando(3), "INSERT INTO ErrosEstProd (Produto,Almoxarifado,TipoDif,QuantDif) VALUES (?,?,?,?)", objRegInventario.sProduto, objRegInventario.iAlmoxarifado, 1, dQuantDif)
            If lErro <> AD_SQL_SUCESSO Then gError 69921
                        
        End If
        
        dQuantDif = testprod.dQuantInicialBenef - testprod.dQuantBenef + objRegInventario.dQuantBenef
        If Abs(dQuantDif) > QTDE_ESTOQUE_DELTA Then
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO ErrosEstProd (Produto,Almoxarifado,TipoDif,QuantDif) VALUES (?,?,?,?)", objRegInventario.sProduto, objRegInventario.iAlmoxarifado, 2, dQuantDif)
            If lErro <> AD_SQL_SUCESSO Then gError 69921
        
        End If
        
        dQuantDif = testprod.dQuantInicialBenef3 - testprod.dQuantBenef3 + objRegInventario.dQuantBenef3
        If Abs(dQuantDif) > QTDE_ESTOQUE_DELTA Then
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO ErrosEstProd (Produto,Almoxarifado,TipoDif,QuantDif) VALUES (?,?,?,?)", objRegInventario.sProduto, objRegInventario.iAlmoxarifado, 3, dQuantDif)
            If lErro <> AD_SQL_SUCESSO Then gError 69921
        
        End If
        
        dQuantDif = testprod.dQuantInicialConserto - testprod.dQuantConserto + objRegInventario.dQuantConserto
        If Abs(dQuantDif) > QTDE_ESTOQUE_DELTA Then
    
            lErro = Comando_Executar(alComando(3), "INSERT INTO ErrosEstProd (Produto,Almoxarifado,TipoDif,QuantDif) VALUES (?,?,?,?)", objRegInventario.sProduto, objRegInventario.iAlmoxarifado, 4, dQuantDif)
            If lErro <> AD_SQL_SUCESSO Then gError 69921
        
        End If
        
        dQuantDif = testprod.dQuantInicialConserto3 - testprod.dQuantConserto3 + objRegInventario.dQuantConserto3
        If Abs(dQuantDif) > QTDE_ESTOQUE_DELTA Then
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO ErrosEstProd (Produto,Almoxarifado,TipoDif,QuantDif) VALUES (?,?,?,?)", objRegInventario.sProduto, objRegInventario.iAlmoxarifado, 5, dQuantDif)
            If lErro <> AD_SQL_SUCESSO Then gError 69921
        
        End If
        
        dQuantDif = testprod.dQuantInicialConsig - testprod.dQuantConsig + objRegInventario.dQuantConsig
        If Abs(dQuantDif) > QTDE_ESTOQUE_DELTA Then
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO ErrosEstProd (Produto,Almoxarifado,TipoDif,QuantDif) VALUES (?,?,?,?)", objRegInventario.sProduto, objRegInventario.iAlmoxarifado, 6, dQuantDif)
            If lErro <> AD_SQL_SUCESSO Then gError 69921
        
        End If
        
        dQuantDif = testprod.dQuantInicialConsig3 - testprod.dQuantConsig3 + objRegInventario.dQuantConsig3
        If Abs(dQuantDif) > QTDE_ESTOQUE_DELTA Then
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO ErrosEstProd (Produto,Almoxarifado,TipoDif,QuantDif) VALUES (?,?,?,?)", objRegInventario.sProduto, objRegInventario.iAlmoxarifado, 7, dQuantDif)
            If lErro <> AD_SQL_SUCESSO Then gError 69921
        
        End If
        
        '??? está faltando
'        If Abs(testprod.dQuantDefeituosa + objRegInventario.dQuantidadeUMEstoque) > QTDE_ESTOQUE_DELTA Then
'
'        End If
        
        dQuantDif = testprod.dQuantInicialDemo - testprod.dQuantDemo + objRegInventario.dQuantDemo
        If Abs(dQuantDif) > QTDE_ESTOQUE_DELTA Then
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO ErrosEstProd (Produto,Almoxarifado,TipoDif,QuantDif) VALUES (?,?,?,?)", objRegInventario.sProduto, objRegInventario.iAlmoxarifado, 8, dQuantDif)
            If lErro <> AD_SQL_SUCESSO Then gError 69921
        
        End If
        
        dQuantDif = testprod.dQuantInicialDemo3 - testprod.dQuantDemo3 + objRegInventario.dQuantDemo3
        If Abs(dQuantDif) > QTDE_ESTOQUE_DELTA Then
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO ErrosEstProd (Produto,Almoxarifado,TipoDif,QuantDif) VALUES (?,?,?,?)", objRegInventario.sProduto, objRegInventario.iAlmoxarifado, 9, dQuantDif)
            If lErro <> AD_SQL_SUCESSO Then gError 69921
        
        End If
        
        dQuantDif = testprod.dQuantInicialOutras - testprod.dQuantOutras + objRegInventario.dQuantOutras
        If Abs(dQuantDif) > QTDE_ESTOQUE_DELTA Then
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO ErrosEstProd (Produto,Almoxarifado,TipoDif,QuantDif) VALUES (?,?,?,?)", objRegInventario.sProduto, objRegInventario.iAlmoxarifado, 10, dQuantDif)
            If lErro <> AD_SQL_SUCESSO Then gError 69921
        
        End If
        
        dQuantDif = testprod.dQuantInicialOutras3 - testprod.dQuantOutras3 + objRegInventario.dQuantOutras3
        If Abs(dQuantDif) > QTDE_ESTOQUE_DELTA Then
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO ErrosEstProd (Produto,Almoxarifado,TipoDif,QuantDif) VALUES (?,?,?,?)", objRegInventario.sProduto, objRegInventario.iAlmoxarifado, 11, dQuantDif)
            If lErro <> AD_SQL_SUCESSO Then gError 69921
        
        End If
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 69924
    
    Loop
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 69921
    
    MsgBox ("correcao de estoqueproduto efetuada com sucesso")
    
    Exit Sub
     
Erro_Integridade_movsEst_EstoqueProduto:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149583)
     
    End Select
     
    Call Transacao_Rollback
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    MsgBox ("Houve erro na correcao de estoqueproduto")
    
    Exit Sub

End Sub

Public Function TipoDocInfo_ZeraQtdeMovto(ByVal iTipoNFiscal As Integer) As Boolean
'retorna em bZera se a qtde do movto associado ao tipo da nf deve ser zerada
'Serve para forçar zerar a qtde de movtos como conhecimento de transporte, complementos, etc.
                    
'If (objNFiscal.iTipoNFiscal = DOCINFO_NFIEICM) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEICM) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEECP) Or (objNFiscal.iTipoNFiscal = DOCINFO_CFECT) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEIPI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIECP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEIPI) Or (objNFiscal.iTipoNFiscal = DOCINFO_CFEV) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFETEL) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFES) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIECS) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEECS) Or (objNFiscal.iTipoNFiscal = DOCINFO_DDAI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEESAC) Then
'If (objNFiscal.iTipoNFiscal = DOCINFO_NFIERCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEERCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEECCP) Then
'If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEFS) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCS) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFICM) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFIPI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEFTEL) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFELUZ) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCS) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFICM) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFIPI) Or (objNFiscal.iTipoNFiscal = DOCINFO_FDDAI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFFEESAC) Then
'If (objNFiscal.iTipoNFiscal = DOCINFO_NFEECPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIECPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEIPICO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEICMCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEICMCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEIPICO) Then
'If (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFIPICO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFICMCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFICMCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFIPICO) Then

'??? no futuro a implementacao desta funcao deve mudar para obter a informacao se zera ou nao do tipodocinfo para nao ter que ficar trocando o codigo a cada noo tipo deste.

Dim bZera As Boolean

    bZera = True
    
    Select Case iTipoNFiscal
    
        Case DOCINFO_NFIEICM, DOCINFO_NFEEICM, DOCINFO_NFEECP, DOCINFO_CFECT, DOCINFO_NFIEIPI, DOCINFO_NFIECP, DOCINFO_NFEEIPI, DOCINFO_CFEV, DOCINFO_NFETEL, DOCINFO_NFES, DOCINFO_NFIECS, DOCINFO_NFEECS, DOCINFO_DDAI, DOCINFO_NFEESAC, DOCINFO_CFECC
    
        Case DOCINFO_NFIERCP, DOCINFO_NFEERCP, DOCINFO_NFEECCP, DOCINFO_CFFECT
    
        Case DOCINFO_NFEEFCP, DOCINFO_NFEFS, DOCINFO_NFEEFCS, DOCINFO_NFEEFICM, DOCINFO_NFEEFIPI, DOCINFO_NFEFTEL, DOCINFO_NFELUZ, DOCINFO_NFIEFCP, DOCINFO_NFIEFCS, DOCINFO_NFIEFICM, DOCINFO_NFIEFIPI, DOCINFO_FDDAI, DOCINFO_NFFEESAC
    
        Case DOCINFO_NFEECPCO, DOCINFO_NFIECPCO, DOCINFO_NFEEIPICO, DOCINFO_NFEEICMCO, DOCINFO_NFIEICMCO, DOCINFO_NFIEIPICO
        
        Case DOCINFO_NFIEFIPICO, DOCINFO_NFIEFICMCO, DOCINFO_NFIEFCPCO, DOCINFO_NFEEFCPCO, DOCINFO_NFEEFICMCO, DOCINFO_NFEEFIPICO
        
        Case Else
            bZera = False
    
    End Select
    
    TipoDocInfo_ZeraQtdeMovto = bZera
    
End Function


Function MovEstoque_ReqProducao_Grava_Trans(ByVal objMovEstoque As ClassMovEstoque, ByVal objContabil As ClassContabil) As Long
'Grava o Movimento de Estoque de Requisição de Produção
'IMPORTANTE: Tem que ser chamado em Transacao

Dim lErro As Long
Dim alComando(3) As Long
Dim objItemMovEstoque As ClassItemMovEstoque
Dim ItemOP As New ClassItemOP
Dim sComando_SQL As String
Dim lNumIntDoc As Long, iIndice As Integer
Dim iSituacao As Integer

On Error GoTo Erro_MovEstoque_ReqProducao_Grava_Trans

    'Abre Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 34881
    Next

    lErro = CF("MovEstoque_Trata_Grade", objMovEstoque)
    If lErro <> SUCESSO Then gError 126665

    For Each objItemMovEstoque In objMovEstoque.colItens

        'Testa o Movimento de Estoque de Requisição de Produção
        lErro = CF("MovEstoque_ReqProducao_Grava1", alComando(0), alComando(1), alComando(2), objItemMovEstoque)
        If lErro <> SUCESSO Then gError 55309

    Next
    
    lErro = CF("MovEstoque_Grava", objMovEstoque)
    If lErro <> SUCESSO Then gError 34887

    Set objItemMovEstoque = objMovEstoque.colItens(1)

    If Not (objContabil Is Nothing) Then
    
        'Grava os dados contábeis (contabilidade)
        lErro = objContabil.Contabil_Gravar_Registro(objItemMovEstoque.lNumIntDoc, 0, 0, objItemMovEstoque.dtData)
        If lErro <> SUCESSO Then gError 36491

    End If
    
    '############################################
    'Inserido por Wagner 22/03/2006
    lErro = CF("ProjetoInfo_Grava", objMovEstoque.objProjetoInfo, PRJ_CR_TIPO_REQPROD, objMovEstoque.lCodigo, "", objMovEstoque.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 185851
    '############################################
    
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    MovEstoque_ReqProducao_Grava_Trans = SUCESSO

    Exit Function

Erro_MovEstoque_ReqProducao_Grava_Trans:

    MovEstoque_ReqProducao_Grava_Trans = gErr

    Select Case gErr

        Case 34881
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 34887, 36491, 55309, 126665, 185851

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149584)

    End Select

    'Fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function MovEstoque_ReqProducao_Grava0(ByVal objMovEstoque As ClassMovEstoque, ByVal objContabil As ClassContabil) As Long
'Grava o Movimento de Estoque de Requisição de Produção

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_MovEstoque_ReqProducao_Grava0

    lErro = CF("MovEstoque_ReqProducao_Grava_Trans", objMovEstoque, objContabil)
    If lErro <> SUCESSO Then gError 105145
    
    MovEstoque_ReqProducao_Grava0 = SUCESSO

    Exit Function

Erro_MovEstoque_ReqProducao_Grava0:

    MovEstoque_ReqProducao_Grava0 = gErr

    Select Case gErr

        Case 105145

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149585)

    End Select

    Exit Function

End Function

'Por Leo
Function EmbalagensExpedicao_Inclui(ByVal objEmbalagem As ClassEmbalagem) As Long
'Grava todas as EmbalagensExpedicao passadas objEmbalagem.colEmbExpedicao

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objEmbalagensExpedicao As ClassEmbalagensExpedicao

On Error GoTo Erro_EmbalagensExpedicao_Inclui

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103430
    Next

    'Para cada objEmbalagem de Expedicao associada a Embalagem
    For Each objEmbalagensExpedicao In objEmbalagem.colEmbExpedicao
    
        'Insere na tabela objEmbalagensExpedicao os dados passados por da Embalagem de Expedição
        lErro = Comando_Executar(alComando(0), "INSERT INTO EmbalagensExpedicao(Embalagem, Sequencial, Produto, Quantidade) VALUES (?,?,?,?)", objEmbalagensExpedicao.iEmbalagem, objEmbalagensExpedicao.iSequencial, objEmbalagensExpedicao.sProduto, objEmbalagensExpedicao.dQuantidade)
        If lErro <> AD_SQL_SUCESSO Then gError 103431
                
    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    EmbalagensExpedicao_Inclui = SUCESSO

    Exit Function

Erro_EmbalagensExpedicao_Inclui:

    EmbalagensExpedicao_Inclui = gErr

    Select Case gErr

        Case 103430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_EMBALAGENSEXPEDICAO", gErr, objEmbalagensExpedicao.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149586)

    End Select

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

End Function

'Por Leo
Function EmbalagensExpedicao_Exclui(ByVal objEmbalagem As ClassEmbalagem) As Long
'Exclui os Registros da tabela EmbalagensExpedicao relacionados a embalagem

Dim alComando(0 To 1) As Long
Dim lErro As Long
Dim sProduto As String
Dim iIndice As Integer

On Error GoTo Erro_EmbalagensExpedicao_Exclui

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103432
    Next
    
    sProduto = String(STRING_PRODUTO, 0)
    
    'Pesquisa no BD as ligações do ItemNF com itens pedidos de compras
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Produto FROM EmbalagensExpedicao WHERE Embalagem = ?", 0, sProduto, objEmbalagem.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103433

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103434

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        'Exclui o ItemNFItemPC
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM EmbalagensExpedicao", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 103435

        'Busca próxima ligação
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103436

    Loop

    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    EmbalagensExpedicao_Exclui = SUCESSO

    Exit Function

Erro_EmbalagensExpedicao_Exclui:

    EmbalagensExpedicao_Exclui = gErr

    Select Case gErr

        Case 103432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103433, 103434, 103436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGENSEXPEDICAO", gErr)

        Case 103435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EMBALAGENSEXPEDICAO", gErr, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149587)

    End Select

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Grade_Grava(ByVal objGrade As ClassGrade) As Long
'Atualiza as tabelas Grade e GradeCategoria no BD com os dados contidos em objGrade,
'realizando alteração se o código do obj já existir no banco, ou inclusão, caso contrário

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 3) As Long
Dim lTransacao As Long
Dim sGrade As String
Dim objGradeCategoria As ClassGradeCategoria
    
On Error GoTo Erro_Grade_Grava

    'Abre transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 122700
    
    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 122701
    Next

    'Efetua Lock Shared na tabela CategoriaProduto dos registros referentes
    'às categorias presentes em objGrade
    lErro = CF("Grade_Grava_Lock", objGrade)
    If lErro <> SUCESSO Then gError 122702
    
    sGrade = String(STRING_GRADE_CODIGO, 0)
    
    'Procura na tabela Grade o Codigo armazenado em ObjGrade
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM Grade WHERE Codigo= ?", 0, sGrade, objGrade.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 122703

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122704
    
    'Se não encontrou insere nova grade no BD
    If lErro <> AD_SQL_SUCESSO Then
        
         lErro = Comando_Executar(alComando(1), "INSERT INTO Grade (Codigo,Descricao, Layout) VALUES (?, ?, ?)", objGrade.sCodigo, objGrade.sDescricao, objGrade.iLayout)
         If lErro <> AD_SQL_SUCESSO Then gError 122705
        
    'Se encontrou
    Else
        
        'Se a Grade de objGrade está sendo utilizada por algum produto -> ERRO
        lErro = CF("Grade_Verifica_Utilizacao", objGrade)
        If lErro <> SUCESSO Then gError 122706
    
        'Atualiza o registro na tabela Grade com os dados de objGrade
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Grade SET Descricao=?, Layout = ?", alComando(1), objGrade.sDescricao, objGrade.iLayout)
        If lErro <> AD_SQL_SUCESSO Then gError 122707

        'Exclui as categorias de GradeCategoria que vinham sendo utilizadas pela grade
        'que foi alterada
        lErro = CF("GradeCategoria_Exclui", objGrade)
        If lErro <> SUCESSO Then gError 122708
  
    End If
          
    'Insere em GradeCategoria as Categorias armazenadas em ObjGrade
    For Each objGradeCategoria In objGrade.colCategoria
        
        lErro = Comando_Executar(alComando(3), "INSERT INTO GradeCategoria (Grade,Categoria, Posicao, Seq) VALUES (?, ?, ?,?)", objGrade.sCodigo, objGradeCategoria.sCategoria, objGradeCategoria.iPosicao, objGradeCategoria.iSeq)
        If lErro <> AD_SQL_SUCESSO Then gError 122709
        
    Next
    
    'Fecha transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 122710

    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Grade_Grava = SUCESSO

    Exit Function

Erro_Grade_Grava:

    Grade_Grava = gErr

    Select Case gErr

        Case 122700
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 122701
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                    
        Case 122702, 122706, 122708
        
        Case 122703, 122704
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRADE", gErr, objGrade.sCodigo)
        
        Case 122705
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GRADE", gErr, objGrade.sCodigo)
            
        Case 122707
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_GRADE", gErr, objGrade.sCodigo)
            
        Case 122709
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GRADECATEGORIA", gErr)
        
        Case 122710
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149588)

    End Select

    'Fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function
    
End Function

Function GradeCategoria_Exclui(ByVal objGrade As ClassGrade) As Long
'Exclui da tabela GradeCategoria os registros referentes às Categorias da grade
'armazenada em ObjGrade

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim sGrade As String

On Error GoTo Erro_GradeCategoria_Exclui

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 122718
    Next
    
    sGrade = String(STRING_GRADE_CODIGO, 0)
    
    'Procura o código da grade que está armazenada em objgrade na tabela GradeCategoria
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Grade FROM GradeCategoria WHERE Grade=?", 0, sGrade, objGrade.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 122719
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122720
            
    Do While lErro = AD_SQL_SUCESSO
                                        
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM GradeCategoria", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 122721
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122722
            
    Loop
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    GradeCategoria_Exclui = SUCESSO
       
    Exit Function
    
Erro_GradeCategoria_Exclui:

    GradeCategoria_Exclui = gErr

    Select Case gErr
    
        Case 122718
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 122719, 122720, 122722
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRADECATEGORIA", gErr)
        
        Case 122721
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GRADECATEGORIA", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149589)
    
    End Select
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function Grade_Exclui(ByVal objGrade As ClassGrade) As Long
'Exclui a Grade passada em objGrade do BD (das tabelas Grade e GradeCategoria)

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim sCodigo As String

On Error GoTo Erro_Grade_Exclui

    'Abre transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 122731
    
    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 122732
    Next
    
    sCodigo = String(STRING_GRADE_CODIGO, 0)
    
    'Busca a Grade contida em objGrade na tabela Grade
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM Grade WHERE Codigo=?", 0, sCodigo, objGrade.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 122733
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122734
    
    If lErro = AD_SQL_SEM_DADOS Then gError 122735
    
    'Efetua lock do registro da tabela Grade referente à Grade contida em objGrade
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 122736
    
    'Verifica se a Grade a ser excluída está sendo utilizada por algum produto
    lErro = CF("Grade_Verifica_Utilizacao", objGrade)
    If lErro <> SUCESSO Then gError 122737
    
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM Grade", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 122738
    
    lErro = CF("GradeCategoria_Exclui", objGrade)
    If lErro <> SUCESSO Then gError 122739
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fecha transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 122740

    Grade_Exclui = SUCESSO
       
    Exit Function
    
Erro_Grade_Exclui:

    Grade_Exclui = gErr

    Select Case gErr
    
        Case 122731
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 122732
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 122733, 122734
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRADE", gErr, objGrade.sCodigo)
        
        Case 122735
            Call Rotina_Erro(vbOKOnly, "ERRO_GRADE_NAO_CADASTRADA", gErr, objGrade.sCodigo)
        
        Case 122736
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_GRADE", gErr, objGrade.sCodigo)
            
        Case 122737, 122739
        
        Case 122738
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GRADE", gErr, objGrade.sCodigo)
            
        Case 122740
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149590)
    
    End Select
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function Reserva_Grava_TipoPedidoGrade(ByVal objReserva As ClassReserva, ParamArray alComando() As Variant) As Long
'Faz as atualizações à Nível de Item de Pedido

Dim lErro As Long
Dim tItemPedido As typeItemPedido

On Error GoTo Erro_Reserva_Grava_TipoPedidoGrade

    tItemPedido.sProduto = String(STRING_PRODUTO, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Produto, Quantidade, QuantCancelada, QuantFaturada FROM ItensPedidoDeVendaGrade WHERE NumIntDoc = ?", 0, tItemPedido.sProduto, tItemPedido.dQuantidade, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, objReserva.lNumIntOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 30529

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30531

    'Se não achou o Item de Pedido --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30595

    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 30904

    tItemPedido.dQuantAFaturar = tItemPedido.dQuantidade - tItemPedido.dQuantFaturada - tItemPedido.dQuantCancelada
    
    If objReserva.dQuantidade > tItemPedido.dQuantAFaturar Then Error 30906

    If objReserva.sProduto <> tItemPedido.sProduto Then Error 55506

    lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensPedidoDeVendaGrade SET QuantReservada = QuantReservada + ?", alComando(1), objReserva.dQuantidade)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30805

    lErro = Comando_ExecutarPos(alComando(3), "SELECT Quantidade, QuantCancelada, QuantFaturada FROM ItensPedidoDeVenda WHERE NumIntDoc = (SELECT NumIntItemPV FROM ItensPedidoDeVendaGrade WHERE NumIntDoc = ?)", 0, tItemPedido.dQuantidade, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, objReserva.lNumIntOrigem)
    If lErro <> AD_SQL_SUCESSO Then Error 30529

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30531

    'Se não achou o Item de Pedido --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30595

    lErro = Comando_LockExclusive(alComando(3))
    If lErro <> AD_SQL_SUCESSO Then Error 30904

    tItemPedido.dQuantAFaturar = tItemPedido.dQuantidade - tItemPedido.dQuantFaturada - tItemPedido.dQuantCancelada
    
    If objReserva.dQuantidade > tItemPedido.dQuantAFaturar Then Error 30906

    lErro = Comando_ExecutarPos(alComando(0), "UPDATE ItensPedidoDeVenda SET QuantReservada = QuantReservada + ?", alComando(3), objReserva.dQuantidade)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30805
    
    lErro = CF("Bloqueio_ReservaGrava", objReserva)
    If lErro <> SUCESSO Then Error 56971

    Reserva_Grava_TipoPedidoGrade = SUCESSO

    Exit Function

Erro_Reserva_Grava_TipoPedidoGrade:

    Reserva_Grava_TipoPedidoGrade = Err

    Select Case Err

        Case 30529, 30531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", Err)

        Case 30595
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", Err, objReserva.lDocOrigem, objReserva.lNumIntOrigem)

        Case 30904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDIDODEVENDA", Err, objReserva.lDocOrigem, objReserva.lNumIntOrigem)

        Case 30906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_RESERVADA_MAIOR_FATURAR", Err, objReserva.dQuantidade, tItemPedido.dQuantAFaturar)

        Case 55506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SEM_PRODUTO", Err, objReserva.sProduto)

        Case 56971
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149591)

    End Select

    Exit Function

End Function


Function Reserva_Exclui_TipoPedidoGrade(ByVal objReserva As ClassReserva, ParamArray alComando() As Variant) As Long
'Exclui o Tipo de Pedido associado à Reserva

Dim lErro As Long
Dim lCodigo As Long
Dim tItemPedido As typeItemPedido
Dim objItemPedido As New ClassItemPedido
Dim dFator As Double
Dim dQuantReservada As Double

On Error GoTo Erro_Reserva_Exclui_TipoPedidoGrade

    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM PedidosDeVenda WHERE Codigo = ? AND FilialEmpresa = ?", 0, lCodigo, objReserva.lDocOrigem, objReserva.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 30012

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30013

    'Se não achou o Pedido de Venda --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30014

    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then Error 30015

    tItemPedido.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)
    tItemPedido.sUMEstoque = String(STRING_UM_SIGLA_UM_BASE, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT ClasseUM, UnidadeMed, QuantReservada FROM ItensPedidoDeVenda WHERE NumIntDoc = (SELECT NumIntItemPV FROM ItensPedidoDeVendaGrade WHERE NumIntDoc = ?) AND FilialEmpresa = ?", 0, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dQuantReservada, objReserva.lNumIntOrigem, objReserva.iFilialEmpresa)
    If lErro <> SUCESSO Then Error 30016

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30017

    'Se não achou o Item do Pedido --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30018

    objItemPedido.iClasseUM = tItemPedido.iClasseUM
    objItemPedido.sUnidadeMed = tItemPedido.sUnidadeMed
    objItemPedido.dQuantReservada = tItemPedido.dQuantReservada

    lErro = Comando_Executar(alComando(2), "SELECT SiglaUMEstoque FROM Produtos WHERE Codigo = ?", tItemPedido.sUMEstoque, objReserva.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 30071
 
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30072

    'Se não achou o Produto --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30073

    objItemPedido.sUMEstoque = tItemPedido.sUMEstoque

    lErro = CF("UM_Conversao", objItemPedido.iClasseUM, objItemPedido.sUnidadeMed, objItemPedido.sUMEstoque, dFator)
    If lErro <> SUCESSO Then Error 30044

    dQuantReservada = (objItemPedido.dQuantReservada - objReserva.dQuantidade) * dFator

    lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensPedidoDeVenda SET QuantReservada=?", alComando(1), dQuantReservada)
    If lErro <> AD_SQL_SUCESSO Then Error 30020

    lErro = Comando_ExecutarPos(alComando(4), "SELECT QuantReservada FROM ItensPedidoDeVendaGrade WHERE NumIntDoc = ? ", 0, tItemPedido.dQuantReservada, objReserva.lNumIntOrigem)
    If lErro <> SUCESSO Then Error 30016

    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30017

    'Se não achou o Item do Pedido --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 30018
    
    dQuantReservada = (tItemPedido.dQuantReservada - objReserva.dQuantidade) * dFator

    lErro = Comando_ExecutarPos(alComando(5), "UPDATE ItensPedidoDeVendaGrade SET QuantReservada=?", alComando(4), dQuantReservada)
    If lErro <> AD_SQL_SUCESSO Then Error 30020
    
    lErro = CF("Bloqueio_ReservaExclui", objReserva)
    If lErro <> SUCESSO Then Error 30021

    Reserva_Exclui_TipoPedidoGrade = SUCESSO

    Exit Function

Erro_Reserva_Exclui_TipoPedidoGrade:

    Reserva_Exclui_TipoPedidoGrade = Err

    Select Case Err

        Case 30012, 30013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDAS1", Err)

        Case 30014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_INEXISTENTE", Err, objReserva.lCodigo)

        Case 30015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOS_DE_VENDA", Err, objReserva.lDocOrigem)

        Case 30016, 30017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", Err)

        Case 30018
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_NAO_CADASTRADO", Err, objReserva.lCodigo)

        Case 30020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", Err, objReserva.lNumIntOrigem)

        Case 30021

        Case 30044

        Case 30071, 30072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case 30073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objReserva.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149592)

    End Select

    Exit Function

End Function

Function NFiscalEntrada_Grava_EmTrans(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lComando As Long
Dim sNome As String
Dim iCodigo As Integer
Dim lNumIntDoc As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataEmissao As Date, bNFComNumero As Boolean
Dim objNFiscalAux As New ClassNFiscal

On Error GoTo Erro_NFiscalEntrada_Grava_EmTrans

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 55599
    
    If lErro <> SUCESSO Then gError 55600
    
    lErro = CF("TabelaPreco_Valida_Compra", objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO Then gError 200270
    
    bNFComNumero = (objNFiscal.lNumNotaFiscal <> 0)
    
    If objNFiscal.lNumNotaFiscal = 0 And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then
    
        'gera o Número para a Nota Fiscal Interna
        lErro = CF("NFiscalNumAuto", objNFiscal)
        If lErro <> SUCESSO Then gError 60445
        
    End If
        
    'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
    lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89723 Then gError 55640
        
    'Se for uma nota fiscal existente
    'vai excluir a nota e inclui-la com o mesmo numero
    If lErro <> SUCESSO Then
            
'        objNFiscalAux.lNumIntDoc = objNFiscal.lNumIntDoc
'        Set objNFiscalAux.objContabil = objNFiscal.objContabil
'
'        lErro = CF("NFiscal_Le", objNFiscalAux)
'        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
'
'        'Lê os itens da nota fiscal
'        lErro = CF("NFiscalItens_Le", objNFiscalAux)
'        If lErro <> SUCESSO Then gError 207636
'
'        'Faz a exclusao de uma nota fiscal de Saida
'        lErro = CF("NotaFiscalEntrada_Excluir_EmTrans", objNFiscalAux, objNFiscalAux.objContabil)
'        If lErro <> SUCESSO Then gError 207637
        
        lErro = NotaFiscalEntrada_Grava_Trata_Exclusao(objNFiscal)
        If lErro <> SUCESSO Then gError 207637

    Else
    
        objNFiscal.lNumIntNFAnterior = 0
        
    End If
    
    If lErro = SUCESSO Then
    
        'se nf veio com numero e é uma nf nova interna
    '    If bNFComNumero And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then gError 81505
            
        lErro = CF("NFiscalEntrada_Lock_Gravacao", objNFiscal)
        If lErro <> SUCESSO Then gError 35663
        
        'Dispara a Gravação da Nota Fiscal
        lErro = CF("NFiscalEntrada_Grava_BD", objNFiscal)
        If lErro <> SUCESSO Then gError 35078
    
        'Grava a Estatística da Filial Fornecedor
        lErro = CF("FilialFornecedor_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_FILIALFORNFILEMP)
        If lErro <> SUCESSO Then gError 52958
            
        'Verifica se o modulo de Livros Fiscais está Ativo
        If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                            
            'Grava o Livro Fiscal a partir da Nota Fiscal
            lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
            If lErro <> SUCESSO Then gError 69551
        
        End If
        
'    Else
'
'        lErro = CF("NFiscal_Alteracao", objNFiscal)
'        If lErro <> SUCESSO Then gError 89739

    End If
    
    lErro = CF("NFiscalEntrada_GravaCompl", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    NFiscalEntrada_Grava_EmTrans = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Grava_EmTrans:

    NFiscalEntrada_Grava_EmTrans = gErr

    Select Case gErr

        Case 35078, 35663, 36728, 36729, 52958, 55606, 55612, 69551, 89739, 181706, 200270, 207636, 207637, 207783

        Case 55599, 55640
            
        Case 55600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr)
        
        Case 60445 'Tratado na Rotina chamada
        
        Case 81505
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_NFISCAL_NUMAUTO", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149593)

    End Select

    Exit Function

End Function

Function NFiscalFatEntrada_Grava_EmTrans(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lComando As Long
Dim lNumIntDoc As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataEmissao As Date
Dim iIndice As Integer
Dim sDescricao As String, bNFComNumero As Boolean
Dim objTituloPagar As New ClassTituloPagar
Dim objNFiscalAux As New ClassNFiscal

On Error GoTo Erro_NFiscalFatEntrada_Grava_EmTrans

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 42666
        
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 42667
    
    If lErro <> SUCESSO Then gError 42668
    
    lErro = CF("TabelaPreco_Valida_Compra", objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO Then gError 200271
    
    bNFComNumero = (objNFiscal.lNumNotaFiscal <> 0)
    
    If objNFiscal.lNumNotaFiscal = 0 And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then
    
        'gera o Número para a Nota Fiscal Interna
        lErro = CF("NFiscalNumAuto", objNFiscal)
        If lErro <> SUCESSO Then gError 60446

    End If
        
    'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
    lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89723 Then gError 55645
    
    'Se for uma nota fiscal existente
    'vai excluir a nota e inclui-la com o mesmo numero
    If lErro <> SUCESSO Then

'        objNFiscalAux.lNumIntDoc = objNFiscal.lNumIntDoc
'        Set objNFiscalAux.objContabil = objNFiscal.objContabil
'
'        lErro = CF("NFiscal_Le", objNFiscalAux)
'        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
'
'        'Lê os itens da nota fiscal
'        lErro = CF("NFiscalItens_Le", objNFiscalAux)
'        If lErro <> SUCESSO Then gError 207640
'
'        'Faz a exclusao de uma nota fiscal de Saida
'        lErro = CF("NotaFiscalEntrada_Excluir_EmTrans", objNFiscalAux, objNFiscal.objContabil)
'        If lErro <> SUCESSO Then gError 207641
        
        lErro = NotaFiscalEntrada_Grava_Trata_Exclusao(objNFiscal)
        If lErro <> SUCESSO Then gError 207637

    Else
    
        objNFiscal.lNumIntNFAnterior = 0
        
    End If
    
    'Se for uma nota nova ou a antiga foi excluida
    If lErro = SUCESSO Then
        
'    'se nf veio com numero e é uma nf nova interna
'    If bNFComNumero And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then gError 81505
        
        lErro = CF("NFiscalFatEntrada_Lock_Gravacao", objNFiscal)
        If lErro <> SUCESSO Then gError 42682
    
        lErro = CF("NFiscalFatEntrada_Grava_BD", objNFiscal)
        If lErro <> SUCESSO Then gError 42683
    
        'Grava a Estatística da Filial Fornecedor
        lErro = CF("FilialFornecedor_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_FILIALFORNFILEMP)
        If lErro <> SUCESSO Then gError 52959
    
        'Verifica se o modulo de Livros Fiscais está Ativo
        If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                            
            'Grava o Livro Fiscal a partir da Nota Fiscal
            lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
            If lErro <> SUCESSO Then gError 69552
        
        End If
    
'    Else
'
'        'horaentrada
'        lErro = CF("NFiscal_Alteracao", objNFiscal)
'        If lErro <> SUCESSO Then gError 83427

    End If
    
    
    lErro = CF("NFiscalEntrada_GravaCompl", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    
'    Set objNFiscal.objContabil = objContabil
'
'    'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
'    lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNFiscal, objTipoDocInfo)
'    If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89723 Then gError 55640
'
'    'Se for uma nota fiscal existente
'    'vai excluir a nota e inclui-la com o mesmo numero
'    If lErro <> SUCESSO Then
'
'        objNFiscalAux.lNumIntDoc = objNFiscal.lNumIntDoc
'        Set objNFiscalAux.objContabil = objNFiscal.objContabil
'
'        lErro = CF("NFiscal_Le", objNFiscalAux)
'        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
'
'        'Lê os itens da nota fiscal
'        lErro = CF("NFiscalItens_Le", objNFiscalAux)
'        If lErro <> SUCESSO Then gError 207636
'
'        'Faz a exclusao de uma nota fiscal de Saida
'        lErro = CF("NotaFiscalSaida_Excluir_EmTrans", objNFiscalAux, objNFiscal.objContabil)
'        If lErro <> SUCESSO Then gError 207637
'
'    End If
'
'
'    'se nf veio com numero e é uma nf nova interna
'    If bNFComNumero And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then gError 81505

    Call Comando_Fechar(lComando)

    NFiscalFatEntrada_Grava_EmTrans = SUCESSO

    Exit Function

Erro_NFiscalFatEntrada_Grava_EmTrans:

    NFiscalFatEntrada_Grava_EmTrans = gErr
    
    Select Case gErr
    
        Case 42666
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 42668
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr)
    
        Case 42671, 42674
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)
        
        Case 36501, 42667, 42682, 42683, 52959, 55645, 83427, 89740, 181706, 200271, 207577, 207640, 207641, 207637
        
        Case 60446, 69552 'Tratado na rotina chamada
        
        Case 81505
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_NFISCAL_NUMAUTO", gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149594)
            
    End Select

    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Function NFiscalEntradaDevolucao_Grava_EmTrans(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz a gravação de Uma Nota Fiscal de Entrada-Devolução

Dim lErro As Long
Dim objItemNFOrig As New ClassItemNF
Dim dQuantDevolvida As Double
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iTipoNFOriginal As Integer, bNFComNumero As Boolean
Dim objNFiscalAux As New ClassNFiscal

On Error GoTo Erro_NFiscalEntradaDevolucao_Grava_EmTrans

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 35379
    Next
    
    'Guarda o Tipo Original de NFiscal para usar em estatística
    iTipoNFOriginal = objNFiscal.iTipoNFiscal
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 42070
    
    'Se não encontrou o Tipo de Documento --> erro
    If lErro <> SUCESSO Then gError 42072
    
    bNFComNumero = (objNFiscal.lNumNotaFiscal <> 0)
    
    If objNFiscal.lNumNotaFiscal = 0 And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then
    
        'gera o Número para a Nota Fiscal Interna
        lErro = CF("NFiscalNumAuto", objNFiscal)
        If lErro <> SUCESSO Then gError 60448

    End If
        
    'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
    lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89723 Then gError 39498
                  
    'Se for uma nota fiscal existente
    'vai excluir a nota e inclui-la com o mesmo numero
    If lErro <> SUCESSO Then
            
'        objNFiscalAux.lNumIntDoc = objNFiscal.lNumIntDoc
'        Set objNFiscalAux.objContabil = objNFiscal.objContabil
'
'        lErro = CF("NFiscal_Le", objNFiscalAux)
'        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
'
'        'Lê os itens da nota fiscal
'        lErro = CF("NFiscalItens_Le", objNFiscalAux)
'        If lErro <> SUCESSO Then gError 207642
'
'        'Faz a exclusao de uma nota fiscal de Saida
'        lErro = CF("NotaFiscalEntrada_Excluir_EmTrans", objNFiscalAux, objNFiscal.objContabil)
'        If lErro <> SUCESSO Then gError 207643
        
        lErro = NotaFiscalEntrada_Grava_Trata_Exclusao(objNFiscal)
        If lErro <> SUCESSO Then gError 207637
        
    Else
    
        objNFiscal.lNumIntNFAnterior = 0
        
    End If
                  
    If lErro = SUCESSO Then
                  
'        'se nf veio com numero e é uma nf nova interna
'        If bNFComNumero And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then gError 81505
            
        If objNFiscal.lFornecedor > 0 Then
            'Lê a Filial Fornecedor
            lErro = CF("FilialFornecedor_Lock", objNFiscal.lFornecedor, objNFiscal.iFilialForn)
            If lErro <> SUCESSO And lErro <> 42870 Then gError 42082
            
            'Se não encontrou a Filial Fornecedor --> erro
            If lErro <> SUCESSO Then gError 35390
            
        Else
            'Lê a Filial Cliente
            lErro = CF("FilialCliente_Lock", objNFiscal.lCliente, objNFiscal.iFilialCli)
            If lErro <> SUCESSO And lErro <> 43050 Then gError 35389
            
            'Se não encontrou a Filial Cliente
            If lErro <> SUCESSO Then gError 35442
            
        End If
    
        'Se a Transportadora estiver preenchida
        If objNFiscal.iCodTransportadora > 0 Then
            
            'Lê a Transportadora
            lErro = CF("Transportadora_Lock1", objNFiscal.iCodTransportadora)
            If lErro <> SUCESSO And lErro <> 43038 Then gError 35392
            
            'Se não encontrou a Transportadora --> erro
            If lErro <> SUCESSO Then gError 35394
            
        End If
        
        'Se Possui Nota Fiscal original
        If objNFiscal.lNumIntNotaOriginal > 0 Then
            
            'Busca a Nota Fiscal Original
            lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM NFiscal WHERE NumIntDoc = ? AND Status <> ? ", 0, lNumIntDoc, objNFiscal.lNumIntNotaOriginal, STATUS_EXCLUIDO)
            If lErro <> AD_SQL_SUCESSO Then gError 35396
    
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 35397
            If lErro = AD_SQL_SEM_DADOS Then gError 35398 'Não encontrou
    
            'Faz o "lock" na Nota Fiscal Original
            lErro = Comando_LockShared(alComando(2))
            If lErro <> AD_SQL_SUCESSO Then gError 35399
    
        End If
    
        'Faz os "locks" dos dados ligados aos Ítens da Nota Fiscal
        lErro = CF("ItensNFiscalEntDev_Lock_Gravacao", objNFiscal)
        If lErro <> SUCESSO Then gError 35414
        
        'Faz a Inserção da Nova Nota Fiscal Entrada-Devolução no BD
        lErro = CF("NFiscalEntradaDevolucao_Grava_BD", objNFiscal)
        If lErro <> SUCESSO Then gError 35415
        
        objNFiscal.iTipoNFiscal = iTipoNFOriginal
        
        'Grava a Estatística da Filial Fornecedor
        lErro = CF("FilialFornecedor_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_FILIALFORNFILEMP)
        If lErro <> SUCESSO Then gError 133156
        
        'Grava a Estatística do Produto com Relação a devolução
        lErro = CF("ProdutoVendido_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_DOC)
        If lErro <> SUCESSO Then gError 52961
    
        'Verifica se o modulo de Livros Fiscais está Ativo
        If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                            
            'Grava o Livro Fiscal a partir da Nota Fiscal
            lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
            If lErro <> SUCESSO Then gError 69553
        
        End If
    
'    Else
'
''horaentrada
'        lErro = CF("NFiscal_Alteracao", objNFiscal)
'        If lErro <> SUCESSO Then gError 83429

    End If
    
    lErro = CF("NFiscalEntrada_GravaCompl", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalEntradaDevolucao_Grava_EmTrans = SUCESSO

    Exit Function

Erro_NFiscalEntradaDevolucao_Grava_EmTrans:

    NFiscalEntradaDevolucao_Grava_EmTrans = gErr

    Select Case gErr

        Case 35090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 35379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 35414, 35415, 36497, 36498, 39498, 42070, 42082, 52961, 69553, 83429, 133156, 181706, 207575, 207642, 207643, 207810, 207637
        
        Case 42072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 35389

        Case 35392

        Case 35394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objNFiscal.iCodTransportadora)

        Case 35396, 35397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 35398
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscal.lNumNotaFiscal)

        Case 35399
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", gErr)

        Case 35442
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objNFiscal.lCliente, objNFiscal.iFilialCli)

        Case 60448 'Tratado na rotina chamada
            
        Case 81505
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_NFISCAL_NUMAUTO", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149595)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalEntradaRemessa_Grava_EmTrans(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Faz as validações necessárias a gravação de Uma Nota Fiscal de Entrada-Remessa

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objProduto As New ClassProduto
Dim objItemNF As ClassItemNF, bNFComNumero As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objNFiscalAux As New ClassNFiscal

On Error GoTo Erro_NFiscalEntradaRemessa_Grava_EmTrans

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 35604
    Next

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 42077
    
    'Se não encontrou o Tipo de Documento --> erro
    If lErro <> SUCESSO Then gError 35388
    
    bNFComNumero = (objNFiscal.lNumNotaFiscal <> 0)
    
    If objNFiscal.lNumNotaFiscal = 0 And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then
    
        'gera o Número para a Nota Fiscal Interna
        lErro = CF("NFiscalNumAuto", objNFiscal)
        If lErro <> SUCESSO Then gError 60449

    End If
        
    'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
    lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89149 And lErro <> 89723 Then gError 35605
    
    'Se for uma nota fiscal existente
    'vai excluir a nota e inclui-la com o mesmo numero
    If lErro <> SUCESSO Then
            
'        objNFiscalAux.lNumIntDoc = objNFiscal.lNumIntDoc
'        Set objNFiscalAux.objContabil = objNFiscal.objContabil
'
'        lErro = CF("NFiscal_Le", objNFiscalAux)
'        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
'
'        'Lê os itens da nota fiscal
'        lErro = CF("NFiscalItens_Le", objNFiscalAux)
'        If lErro <> SUCESSO Then gError 207644
'
'        'Faz a exclusao de uma nota fiscal de Saida
'        lErro = CF("NotaFiscalEntrada_Excluir_EmTrans", objNFiscalAux, objNFiscal.objContabil)
'        If lErro <> SUCESSO Then gError 207645
        
        lErro = NotaFiscalEntrada_Grava_Trata_Exclusao(objNFiscal)
        If lErro <> SUCESSO Then gError 207637
        
    Else
    
        objNFiscal.lNumIntNFAnterior = 0
        
    End If
    
    If lErro = SUCESSO Then
    
    
    '    'significa que encontrou a nota fiscal e ela era de importação e estava gravada com data de entrada nula ==> deve gravar o resto dos dados se a data de entrada estiver preenchida
    '    If lErro = 89149 Then objNFiscal.iStatus = STATUS_GRAVADA_IMPORTACAO_DATA_NULA
    '
    '    'Se for uma nota nova
    '    If lErro = SUCESSO Or lErro = 89149 Then
            
    '        'se nf veio com numero e é uma nf nova interna
    '        If lErro = SUCESSO And bNFComNumero And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then gError 81505
                
        'Chama NFiscalEntradaRemessa_Lock_Gravacao
        lErro = CF("NFiscalEntradaRemessa_Lock_Gravacao", objNFiscal, alComando(0))
        If lErro <> SUCESSO Then gError 35440
    
        'Faz a Inserção da Nova Nota Fiscal Entrada-Remessa no BD
        lErro = CF("NFiscalEntradaRemessa_Grava_BD", objNFiscal)
        If lErro <> SUCESSO Then gError 35623
    
        'se a nota for de importação e não estiver com a data de entrada preenchida ==> não processa a gravação das informações
        'caso contrário, processa
        If (objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMP And objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMPSE And objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMPPC) Or objNFiscal.dtDataEntrada <> DATA_NULA Then
                            
            'para garantir que trate casos excepcionais como nf entrada de importacao
            If objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS Or objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS_DEVOLUCAO Then
            
                'Grava a Estatística da Filial Fornecedor
                lErro = CF("FilialFornecedor_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_FILIALFORNFILEMP)
                If lErro <> SUCESSO Then gError 32275
                
            End If
    
            'Verifica se o modulo de Livros Fiscais está Ativo
            If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                            
                'Grava o Livro Fiscal a partir da Nota Fiscal
                lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
                If lErro <> SUCESSO Then gError 69554
        
            End If
        
        End If
    
    Else

''horaentrada
'        'trata da alteração dos dados da nota fiscal.
'        lErro = CF("NFiscal_Alteracao", objNFiscal)
'        If lErro <> SUCESSO Then gError 83430

    End If
    
    lErro = CF("NFiscalEntrada_GravaCompl", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    
    
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalEntradaRemessa_Grava_EmTrans = SUCESSO
    
    Exit Function

Erro_NFiscalEntradaRemessa_Grava_EmTrans:

    NFiscalEntradaRemessa_Grava_EmTrans = gErr

    Select Case gErr

        Case 35388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)
        
        Case 35604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 35440, 35605, 35623, 36499, 36500, 42077, 83430, 32275, 181706, 207576, 207644, 207645, 207815, 207637

        Case 60449, 69554 'Tratado na rotina chamada
        
        Case 81505
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_NFISCAL_NUMAUTO", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149596)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscalFatEntradaCom_Grava_EmTrans(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil, ByVal colPedCompras As Collection, ByVal colReqCompras As Collection) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lComando As Long
Dim sNome As String
Dim iCodigo As Integer
Dim lNumIntDoc As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataEmissao As Date, bNFComNumero As Boolean
Dim objNFiscalAux As New ClassNFiscal

On Error GoTo Erro_NFiscalFatEntradaCom_Grava_EmTrans

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 66501

    If lErro <> SUCESSO Then gError 66502

    lErro = CF("TabelaPreco_Valida_Compra", objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO Then gError 200265

    bNFComNumero = (objNFiscal.lNumNotaFiscal <> 0)
    
    If objNFiscal.lNumNotaFiscal = 0 And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then
    
        'gera o Número para a Nota Fiscal Interna
        lErro = CF("NFiscalNumAuto", objNFiscal)
        If lErro <> SUCESSO Then gError 66504

    End If
        
    'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
    lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89723 Then gError 66503


    'Se for uma nota fiscal existente
    'vai excluir a nota e inclui-la com o mesmo numero
    If lErro <> SUCESSO Then
            
'        objNFiscalAux.lNumIntDoc = objNFiscal.lNumIntDoc
'        Set objNFiscalAux.objContabil = objNFiscal.objContabil
'
'        lErro = CF("NFiscal_Le", objNFiscalAux)
'        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
'
'        'Lê os itens da nota fiscal
'        lErro = CF("NFiscalItens_Le", objNFiscalAux)
'        If lErro <> SUCESSO Then gError 207646
'
'        'Faz a exclusao de uma nota fiscal de Saida
'        lErro = CF("NotaFiscalEntrada_Excluir_EmTrans", objNFiscalAux, objNFiscal.objContabil)
'        If lErro <> SUCESSO Then gError 207647
        
        lErro = NotaFiscalEntrada_Grava_Trata_Exclusao(objNFiscal)
        If lErro <> SUCESSO Then gError 207637
        
    Else
    
        objNFiscal.lNumIntNFAnterior = 0
        
    End If

    'Se for uma nota nova
    If lErro = SUCESSO Then

'        'se nf veio com numero e é uma nf nova interna
'        If bNFComNumero And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then gError 81505
            
        lErro = CF("NFiscalEntradaCom_Lock_Gravacao", objNFiscal, colPedCompras, colReqCompras)
        If lErro <> SUCESSO Then gError 66505
    
        'Dispara a Gravação da Nota Fiscal
        lErro = CF("NFiscalFatEntradaCom_Grava_BD", objNFiscal, colPedCompras, colReqCompras)
        If lErro <> SUCESSO Then gError 66506
    
        'Grava a Estatística da Filial Fornecedor
        lErro = FilialFornecedor_Grava_EstatisticasCom(objNFiscal, colPedCompras)
        If lErro <> SUCESSO Then gError 66507
    
        'Verifica se o modulo de Livros Fiscais está Ativo
        If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                            
            'Grava o Livro Fiscal a partir da Nota Fiscal
            lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
            If lErro <> SUCESSO Then gError 89951
        
        End If

'    Else
'
''horaentrada
'        'trata da alteração dos dados da nota fiscal.
'        lErro = CF("NFiscal_Alteracao", objNFiscal)
'        If lErro <> SUCESSO Then gError 83431

    End If


    lErro = CF("NFiscalEntrada_GravaCompl", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    
    NFiscalFatEntradaCom_Grava_EmTrans = SUCESSO

    Exit Function

Erro_NFiscalFatEntradaCom_Grava_EmTrans:

    NFiscalFatEntradaCom_Grava_EmTrans = gErr

    Select Case gErr

        Case 66502
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr)

        Case 66501, 66503, 66504, 66505, 66506, 66507, 66508, 83431, 89951, 181706, 200265, 207578, 207646, 207647, 207819, 207637

        Case 81505
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_NFISCAL_NUMAUTO", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149597)

    End Select

    Exit Function

End Function

Function NFiscalEntradaCom_Grava_EmTrans(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil, ByVal colPedCompras As Collection, ByVal colReqCompras As Collection) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lComando As Long
Dim sNome As String
Dim iCodigo As Integer
Dim lNumIntDoc As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataEmissao As Date, bNFComNumero As Boolean
Dim objNFiscalAux As New ClassNFiscal

On Error GoTo Erro_NFiscalEntradaCom_Grava_EmTrans
   
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 65916

    If lErro <> SUCESSO Then gError 65917

    lErro = CF("TabelaPreco_Valida_Compra", objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO Then gError 200264

    bNFComNumero = (objNFiscal.lNumNotaFiscal <> 0)
    
    If objNFiscal.lNumNotaFiscal = 0 And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then
    
        'gera o Número para a Nota Fiscal Interna
        lErro = CF("NFiscalNumAuto", objNFiscal)
        If lErro <> SUCESSO Then gError 65919

    End If
        
        
    'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
    lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89723 Then gError 65918

    'Se for uma nota fiscal existente
    'vai excluir a nota e inclui-la com o mesmo numero
    If lErro <> SUCESSO Then
            
'        objNFiscalAux.lNumIntDoc = objNFiscal.lNumIntDoc
'        Set objNFiscalAux.objContabil = objNFiscal.objContabil
'
'        lErro = CF("NFiscal_Le", objNFiscalAux)
'        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
'
'        'Lê os itens da nota fiscal
'        lErro = CF("NFiscalItens_Le", objNFiscalAux)
'        If lErro <> SUCESSO Then gError 207648
'
'        'Faz a exclusao de uma nota fiscal de Saida
'        lErro = CF("NotaFiscalEntrada_Excluir_EmTrans", objNFiscalAux, objNFiscal.objContabil)
'        If lErro <> SUCESSO Then gError 207649

        lErro = NotaFiscalEntrada_Grava_Trata_Exclusao(objNFiscal)
        If lErro <> SUCESSO Then gError 207637
        
    Else
    
        objNFiscal.lNumIntNFAnterior = 0
        
    End If

    'Se for uma nota nova
    If lErro = SUCESSO Then

'        'se nf veio com numero e é uma nf nova interna
'        If bNFComNumero And (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS) Then gError 81505
            
        lErro = CF("NFiscalEntradaCom_Lock_Gravacao", objNFiscal, colPedCompras, colReqCompras)
        If lErro <> SUCESSO Then gError 65920
    
        'Dispara a Gravação da Nota Fiscal
        lErro = CF("NFiscalEntradaCom_Grava_BD", objNFiscal, colPedCompras, colReqCompras)
        If lErro <> SUCESSO Then gError 65921
    
        'Grava a Estatística da Filial Fornecedor
        lErro = FilialFornecedor_Grava_EstatisticasCom(objNFiscal, colPedCompras)
        If lErro <> SUCESSO Then gError 65922
    
        'Verifica se o modulo de Livros Fiscais está Ativo
        If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                            
            'Grava o Livro Fiscal a partir da Nota Fiscal
            lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
            If lErro <> SUCESSO Then gError 89950
        
        End If

'    Else
'
'        lErro = CF("NFiscal_Alteracao", objNFiscal)
'        If lErro <> SUCESSO Then gError 83428

    End If


    lErro = CF("NFiscalEntrada_GravaCompl", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    NFiscalEntradaCom_Grava_EmTrans = SUCESSO

    Exit Function

Erro_NFiscalEntradaCom_Grava_EmTrans:

    NFiscalEntradaCom_Grava_EmTrans = gErr

    Select Case gErr

        Case 65917
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr)
        
        Case 65916, 65918, 65919, 65920, 65921, 65922, 65923, 65924, 83428, 89950, 181706, 200264, 207574, 207648, 207649, 207802, 207637

        Case 81505
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_NFISCAL_NUMAUTO", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149598)

    End Select

    Exit Function

End Function

Function RelProdPontoPedido_Prepara(lNumIntRel As Long, ByVal sProd_I As String, ByVal sProd_F As String) As Long
'preenche tabela p/relatorio de produtos em ponto de pedido e retorna lNumIntRel
'??? EstoqueProduto_Le_Todos_AlmProdFilial acaba sendo chamada 2 vezes p/cada produto

Dim lErro As Long, lTransacao As Long, objProdutoPtoPedido As ClassProdutoPtoPedido
Dim objGeracaoPtoPedido As ClassGeracaoPtoPedido
Dim dSaldoAtual As Double, lComando As Long
Dim dQuantEmPedidoCompra As Double
Dim dQuantEmReqCompra As Double
Dim dQuantComprar As Double
Dim dSaldoDisponivel As Double
Dim dQuantEmpenhada As Double
Dim dQuantEmEstoque As Double
Dim dQuantReservada As Double

On Error GoTo Erro_RelProdPontoPedido_Prepara

    Set objGeracaoPtoPedido = New ClassGeracaoPtoPedido
    Set objGeracaoPtoPedido.colTipoProduto = New Collection
    objGeracaoPtoPedido.sCodProdutoDe = sProd_I
    objGeracaoPtoPedido.sCodProdutoAte = sProd_F
    
    objGeracaoPtoPedido.iExibePPZerado = gobjEST.iRelPontoPedExibePPZerado
    
    'Lê os Produtos que PODEM ESTAR com estoque abaixo do Ponto de Pedido
    lErro = CF("ProdutoFilial_Le_PontoPedido", objGeracaoPtoPedido)
    If lErro <> SUCESSO Then gError 124236
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 124237

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 124238

    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "MATConfig", "NUMINTREL_RELPONTOPEDIDO", lNumIntRel)
    If lErro <> SUCESSO Then gError 124239
    
    'Para cada Produto
    For Each objProdutoPtoPedido In objGeracaoPtoPedido.colProdutoPtoPedido
        
        'obtem dados de estoque
        lErro = CF("SaldoAtualProdComprado", objProdutoPtoPedido.objProduto.sCodigo, dSaldoAtual, dQuantReservada, dQuantEmReqCompra, dQuantEmPedidoCompra, dSaldoDisponivel, dQuantEmpenhada)
        If lErro <> SUCESSO Then gError 124240
                
        'Se a Quantidade de Ponto de Pedido for maior ou igual ao saldo atual
        If objProdutoPtoPedido.objProdutoFilial.dPontoPedido >= dSaldoAtual Then
    
            '??? EstoqueMinimo está sendo assumido como sendo igual ao EstoqueSeguranca
                
            lErro = Comando_Executar(lComando, "INSERT INTO RelProdPontoPedido (NumIntRel, Produto, QuantEstoque, QuantReservada, QuantEmpenhada, QuantPedVenda, QuantPedCompra, QuantReqCompra, SaldoAtual, PontoPedido, EstoqueMinimo, EstoqueSeguranca) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntRel, objProdutoPtoPedido.objProduto.sCodigo, dSaldoDisponivel, dQuantReservada, dQuantEmpenhada, objProdutoPtoPedido.objEstoqueProduto.dQuantPedido, dQuantEmPedidoCompra, dQuantEmReqCompra, dSaldoAtual, objProdutoPtoPedido.objProdutoFilial.dPontoPedido, objProdutoPtoPedido.objProdutoFilial.dEstoqueSeguranca, objProdutoPtoPedido.objProdutoFilial.dEstoqueSeguranca)
            If lErro <> SUCESSO Then gError 124241
            
        End If
        
    Next
    
    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 124242
    
    Call Comando_Fechar(lComando)
    
    RelProdPontoPedido_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelProdPontoPedido_Prepara:

    RelProdPontoPedido_Prepara = gErr
     
    Select Case gErr
          
        Case 124236, 124239, 124240
        
        Case 124237
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 124238
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 124242
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 124241
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELPONTOPEDIDO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149599)
     
    End Select
     
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PrestServ_Grava(ByVal objPrestServ As ClassPrestServ) As Long
'Grava o PrestServ no BD

Dim lTransacao As Long
Dim alComando(0 To 2) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim lCodigo As Long
Dim sPrestServ As String
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_PrestServ_Grava

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 49068
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 49069

    'Procura por PrestServ com mesmo nome reduzido
     lErro = Comando_Executar(alComando(0), "SELECT Codigo FROM PrestServ WHERE Codigo <> ? AND NomeReduzido = ?", lCodigo, objPrestServ.lCodigo, objPrestServ.sNomeReduzido)
     If lErro <> AD_SQL_SUCESSO Then gError 49070

    'Tenta ler o código deste PrestServ
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 49074

    'Se existir outro PrestServ com o mesmo Nome Reduzido, erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 49075

    'Procura por PrestServ com mesmo nome
    lErro = Comando_Executar(alComando(0), "SELECT Codigo FROM PrestServ WHERE Codigo <> ? AND Nome = ?", lCodigo, objPrestServ.lCodigo, objPrestServ.sNome)
    If lErro <> AD_SQL_SUCESSO Then gError 49288

    'Tenta ler o código deste PrestServ
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 49289

    'Se existir outro PrestServ com o mesmo Nome, erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 49290

    'se o fornecedor foi informado
    If objPrestServ.lFornecedor <> 0 Then

        objFilialFornecedor.lCodFornecedor = objPrestServ.lFornecedor
        objFilialFornecedor.iCodFilial = FILIAL_MATRIZ

        'locka a filial do fornecedor
        lErro = CF("FilialFornecedor_Lock1", objFilialFornecedor)
        If lErro <> AD_SQL_SUCESSO And lErro <> 18900 Then gError 109957

        'não existe filial
        If lErro <> SUCESSO Then gError 49076

    End If

    sPrestServ = String(STRING_PRESTSERV_NOMERED, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NomeReduzido FROM PrestServ WHERE Codigo = ?", 0, sPrestServ, objPrestServ.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 49085

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 49086

    'Encontrou o PrestServ
    If lErro = AD_SQL_SUCESSO Then

        'Lock do PrestServ
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 49080

        'Atualiza os dados do PrestServ
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE PrestServ SET Nome = ?,NomeReduzido = ?,Fornecedor = ?", alComando(1), objPrestServ.sNome, objPrestServ.sNomeReduzido, objPrestServ.lFornecedor)
        If lErro <> AD_SQL_SUCESSO Then gError 49087

    Else
        'Insere o PrestServ
        lErro = Comando_Executar(alComando(2), "INSERT INTO PrestServ(Codigo,Nome,NomeReduzido,Fornecedor) VALUES (?,?,?,?)", objPrestServ.lCodigo, objPrestServ.sNome, objPrestServ.sNomeReduzido, objPrestServ.lFornecedor)
        If lErro <> AD_SQL_SUCESSO Then gError 49088

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 49089

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PrestServ_Grava = SUCESSO

    Exit Function

Erro_PrestServ_Grava:

    PrestServ_Grava = gErr

    Select Case gErr

        Case 49068
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 49069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 49070, 49074, 49085, 49086, 49288, 49289
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRESTSERV", gErr, objPrestServ.lCodigo)

        Case 49075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRESTSERV_NOMERED_DUPLICADO", gErr, objPrestServ.sNomeReduzido)

        Case 109957

        Case 49076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objFilialFornecedor.iCodFilial, objFilialFornecedor.lCodFornecedor)

        Case 49080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRESTSERV", gErr, objPrestServ.lCodigo)

        Case 49087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PRESTSERV", gErr, objPrestServ.lCodigo)

        Case 49088
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRESTSERV", gErr, objPrestServ.lCodigo)

        Case 49089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 49290
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRESTSERV_NOME_DUPLICADO", gErr, objPrestServ.sNome)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149600)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PrestServ_Exclui(ByVal objPrestServ As ClassPrestServ) As Long
'Exclui o PrestServ

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim lTransacao As Long
Dim lTotal As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_PrestServ_Exclui

   'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 49062
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 49063

    'Seleciona o PrestServ a ser excluido
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM PrestServ WHERE Codigo =?", 0, lCodigo, objPrestServ.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49064

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49090
    If lErro = AD_SQL_SEM_DADOS Then Error 49091

    'Lock do PrestServ
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 49065

    lErro = CF("PrestServ_Valida_Exclusao", objPrestServ, alComando)
    If lErro <> SUCESSO Then Error 51351

    'Exclui o PrestServ
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PrestServ", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 49066

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 49067

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PrestServ_Exclui = SUCESSO

    Exit Function

Erro_PrestServ_Exclui:

    PrestServ_Exclui = Err

    Select Case Err

        Case 49062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49063
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 49064, 49090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRESTSERV", Err, objPrestServ.lCodigo)

        Case 49065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRESTSERV", Err, objPrestServ.lCodigo)

        Case 49066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRESTSERV", Err, objPrestServ.lCodigo)

        Case 49067
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 49091
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRESTSERV_NAO_CADASTRADO", Err, objPrestServ.lCodigo)

        Case 51351

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149601)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PrestServ_Automatico(lCodigo As Long) As Long
'Gera o proximo codigo da Tabela de PrestServ

Dim lErro As Long

On Error GoTo Erro_PrestServ_Automatico

    'Chama a rotina que gera o sequencial
    lErro = CF("Config_ObterAutomatico", "ESTConfig", "NUM_PROXIMO_PRESTSERV", "PrestServ", "Codigo", lCodigo)
    If lErro <> SUCESSO Then Error 49071

    PrestServ_Automatico = SUCESSO

    Exit Function

Erro_PrestServ_Automatico:

    PrestServ_Automatico = Err

    Select Case Err

        Case 49071

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149602)

    End Select

    Exit Function

End Function

Public Function ImportarEstoqueInicial() As Long

Dim lErro As Long
Dim colEstoqueProduto As New Collection
Dim objEstoqueProduto As ClassEstoqueProduto
Dim colRastreamento As Collection
Dim objRastroIni As ClassRastroEstIni
Dim objRastroLote As ClassRastreamentoLote
Dim iAlmoxarifadoPadrao As Integer
Dim objProduto As ClassProduto

On Error GoTo Erro_ImportarEstoqueInicial

    iAlmoxarifadoPadrao = DESMARCADO

    lErro = CF("ImportEstoqueInicial_Le_Todos", colEstoqueProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objEstoqueProduto In colEstoqueProduto
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objEstoqueProduto.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

        If lErro <> SUCESSO Then gError 208412
'            gError 99999 'Produto não cadastrado
'        End If
        
        If objProduto.iGerencial = MARCADO Then gError 208413
'            gError 99999 'Produto gerencial
'        End If
        
        If objProduto.iControleEstoque = PRODUTO_SEM_ESTOQUE Then gError 208414
'            gError 99999 'Produto não controla o estoque
'        End If
    
        'Lê o Estoque Produto
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 129860
    
        'Se não está cadastrada, cadastra
        If lErro = 21306 Then
    
            Set colRastreamento = New Collection
        
            lErro = CF("ImportRastroInicial_Le", objEstoqueProduto, colRastreamento)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            For Each objRastroIni In colRastreamento
            
                Set objRastroLote = New ClassRastreamentoLote
                
                objRastroLote.sCodigo = objRastroIni.sLote
                objRastroLote.sProduto = objRastroIni.sProduto
                objRastroLote.iFilialOP = objRastroIni.iFilialOP
                objRastroLote.dtDataEntrada = objRastroIni.dtDataEntrada
            
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 129852
            
                'Se o RastreamentoLote não estiver cadastrado, cadastra
                If lErro = 75710 Then
                    lErro = CF("RastreamentoLote_Grava", objRastroLote)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError ERRO_SEM_MENSAGEM
                
                End If
                
                objRastroIni.lNumIntDocLote = objRastroLote.lNumIntDoc
        
            Next
            
            iAlmoxarifadoPadrao = MARCADO
        
            'Chama EstoqueInicial_Grava
            lErro = CF("EstoqueInicial_Grava", objEstoqueProduto, iAlmoxarifadoPadrao, colRastreamento)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
    Next

    ImportarEstoqueInicial = SUCESSO

    Exit Function

Erro_ImportarEstoqueInicial:

    ImportarEstoqueInicial = gErr

    Select Case gErr
    
        Case 208412
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
    
        Case 208413
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
    
        Case 208414
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", gErr, objProduto.sCodigo)
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149603)

    End Select

    Exit Function

End Function

'######################################################
'Inserido por Wagner
Function NFiscalEntrada_Grava_CamposAlteraveis(ByVal objNFiscal As ClassNFiscal) As Long
    NFiscalEntrada_Grava_CamposAlteraveis = SUCESSO
End Function
'######################################################

'######################################################
'Inserido por Wagner 31/10/2003
Public Function RelBarCode_Prepara(ByVal colRelBarCode As Collection, lNumIntRel As Long) As Long
'Monta a tabela temporária para o relatório de calendário de máquina
 
Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 1) As Long
Dim iIndice As Integer
Dim objRelBarCode As ClassRelBarCode
Dim iSeq As Integer
 
On Error GoTo Erro_RelBarCode_Prepara
 
    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 182827
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182828
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_PROX_REL_BARCODE", lNumIntRel)
    If lErro <> SUCESSO Then gError 182829
    
    iSeq = 0
    For Each objRelBarCode In colRelBarCode
    
        For iIndice = 1 To objRelBarCode.iQuantidade
            
            iSeq = iSeq + 1
        
            'Insere no BD
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelBarCode(NumIntRel, Seq, Produto) VALUES (?,?,?)", lNumIntRel, iSeq, objRelBarCode.sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 182830
    
        Next
    
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 182831
 
    RelBarCode_Prepara = SUCESSO
 
    Exit Function
 
Erro_RelBarCode_Prepara:
 
    RelBarCode_Prepara = gErr
 
    Select Case gErr
 
        Case 182827
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
 
        Case 182828
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 182829
        
        Case 182830
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELBARCODE", gErr)
 
        Case 182831
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152399)
 
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback
 
    Exit Function
 
End Function
'######################################################

Function Produto_Automatico(sCodigo As String) As Long
'retorna o número da proximo produto

Dim lErro As Long
Dim dNumProx As Double

On Error GoTo Erro_Produto_Automatico

    'Gera número automático.
    lErro = CF("Config_ObterAutomatico3", "ESTConfig", "NUM_PROX_PRODUTO", "Produtos", "CodigoNumerico", dNumProx)
    If lErro <> SUCESSO Then gError 187400
    
    sCodigo = CStr(Round(dNumProx))
    
    Produto_Automatico = SUCESSO

    Exit Function

Erro_Produto_Automatico:

    Produto_Automatico = gErr

    Select Case gErr

        Case 187400

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187401)

    End Select
    
    Exit Function

End Function

Function RelRotulo_Customiza(ByVal objOP As ClassOrdemDeProducao, lNumIntRel As Long) As Long
    lNumIntRel = 0
    RelRotulo_Customiza = SUCESSO
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function DIInfo_Grava(ByVal objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_DIInfo_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 184517

    lErro = CF("DIInfo_Grava_EmTrans", objDIInfo)
    If lErro <> SUCESSO Then gError 184518

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 184519

    DIInfo_Grava = SUCESSO

    Exit Function

Erro_DIInfo_Grava:

    DIInfo_Grava = gErr

    Select Case gErr

        Case 184517
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 184518

        Case 184519
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184520)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function DIInfo_Grava_EmTrans(ByVal objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim alComando(0 To 11) As Long
Dim alComando2(0 To 10) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumNF As Long

On Error GoTo Erro_DIInfo_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184522
    Next
    
    'Abertura de Comando
    For iIndice = LBound(alComando2) To UBound(alComando2)
        alComando2(iIndice) = Comando_Abrir()
        If alComando2(iIndice) = 0 Then gError 184522
    Next

    'Le a tabelaDIInfo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM DIInfo WHERE Numero= ? ", _
                0, lNumIntDoc, objDIInfo.sNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 184523

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184524

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 184525

        With objDIInfo
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE DIInfo SET Data= ? , FilialEmpresa= ? , Status= ? , Descricao = ?, FornTrading= ? , " & _
                "FilialFornTrading= ? , ProcessoTrading= ? , Moeda= ? , TaxaMoeda= ? , PesoBrutoKG= ? , " & _
                "PesoLiqKG= ? , ValorMercadoriaMoeda= ? , ValorFreteInternacMoeda= ? , ValorSeguroInternacMoeda= ? , ValorMercadoriaEmReal= ? , " & _
                "ValorFreteInternacEmReal= ? , ValorSeguroInternacEmReal= ? , IIValor= ?, IPIValor= ?, PISValor= ?, COFINSValor= ?, ICMSValor= ?, ValorDespesas= ?, Moeda2 = ?, TaxaMoeda2 = ?, DataDesembaraco = ?, UFDesembaraco = ?, LocalDesembaraco = ?, MoedaMercadoria = ?, MoedaFrete = ?, MoedaSeguro = ?, MoedaItens = ?, CodExportador = ?, ViaTransp = ?, Intermedio = ?, CNPJAdquir = ?, UFAdquir = ?", alComando(0), _
                .dtData, .iFilialEmpresa, .iStatus, .sDescricao, .lFornTrading, _
                .iFilialFornTrading, .sProcessoTrading, .iMoeda1, .dTaxaMoeda1, .dPesoBrutoKG, _
                .dPesoLiqKG, .dValorMercadoriaMoeda, .dValorFreteInternacMoeda, .dValorSeguroInternacMoeda, .dValorMercadoriaEmReal, _
                .dValorFreteInternacEmReal, .dValorSeguroInternacEmReal, .dIIValor, .dIPIValor, .dPISValor, .dCOFINSValor, .dICMSValor, .dValorDespesas, .iMoeda2, .dTaxaMoeda2, .dtDataDesembaraco, .sUFDesembaraco, .sLocalDesembaraco, .iMoedaMercadoria, .iMoedaFrete, .iMoedaSeguro, .iMoedaItens, .sCodExportador, .iViaTransp, .iIntermedio, .sCNPJAdquir, .sUFAdquir)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 184526
        
        objDIInfo.lNumIntDoc = lNumIntDoc
        
        'Le a tabelaDINF
        lErro = Comando_Executar(alComando2(0), "SELECT NF.NumNotaFiscal FROM NFImportacao, NFiscal AS NF WHERE NF.NumIntDoc = NFImportacao.NumIntNF AND NFImportacao.NumIntDI= ? ", _
                    lNumNF, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 196556
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando2(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196557
        
        If lErro = AD_SQL_SUCESSO Then gError 196558
        
        lErro = DIInfo_Exclui_Adicoes(objDIInfo, alComando2)
        If lErro <> SUCESSO Then gError 196559
        
        lErro = DIInfo_Exclui_Despesas(objDIInfo, alComando2)
        If lErro <> SUCESSO Then gError 196560

        lErro = DIInfo_Exclui_ItensPCDI(objDIInfo)
        If lErro <> SUCESSO Then gError 210601

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "CRFATConfig", "NUM_INT_PROX_DIINFO", lNumIntDoc)
        If lErro <> SUCESSO Then gError 184521

        objDIInfo.lNumIntDoc = lNumIntDoc

        With objDIInfo
            lErro = Comando_Executar(alComando(1), "INSERT INTO DIInfo( NumIntDoc, Numero, Data, FilialEmpresa, " & _
                "Status, Descricao, FornTrading, FilialFornTrading, ProcessoTrading, Moeda, " & _
                "TaxaMoeda, PesoBrutoKG, PesoLiqKG, ValorMercadoriaMoeda, ValorFreteInternacMoeda, " & _
                "ValorSeguroInternacMoeda, ValorMercadoriaEmReal, ValorFreteInternacEmReal, ValorSeguroInternacEmReal, " & _
                "IIValor, IPIValor, PISValor, COFINSValor, ICMSValor, ValorDespesas, Moeda2, TaxaMoeda2, DataDesembaraco, UFDesembaraco, LocalDesembaraco, MoedaMercadoria, MoedaFrete, MoedaSeguro, MoedaItens, CodExportador, ViaTransp, Intermedio, CNPJAdquir, UFAdquir) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                .lNumIntDoc, .sNumero, .dtData, .iFilialEmpresa, _
                .iStatus, .sDescricao, .lFornTrading, .iFilialFornTrading, .sProcessoTrading, .iMoeda1, _
                .dTaxaMoeda1, .dPesoBrutoKG, .dPesoLiqKG, .dValorMercadoriaMoeda, .dValorFreteInternacMoeda, _
                .dValorSeguroInternacMoeda, .dValorMercadoriaEmReal, .dValorFreteInternacEmReal, .dValorSeguroInternacEmReal, .dIIValor, .dIPIValor, .dPISValor, .dCOFINSValor, .dICMSValor, .dValorDespesas, .iMoeda2, .dTaxaMoeda2, .dtDataDesembaraco, .sUFDesembaraco, .sLocalDesembaraco, .iMoedaMercadoria, .iMoedaFrete, .iMoedaSeguro, .iMoedaItens, .sCodExportador, .iViaTransp, .iIntermedio, .sCNPJAdquir, .sUFAdquir)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 184527
    
    End If
    
    lErro = DIInfo_Insere_Adicoes(objDIInfo, alComando)
    If lErro <> SUCESSO Then gError 184698

    lErro = DIInfo_Insere_Despesas(objDIInfo, alComando)
    If lErro <> SUCESSO Then gError 184699

    lErro = DIInfo_Insere_ItensPCDI(objDIInfo)
    If lErro <> SUCESSO Then gError 210602

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando2) To UBound(alComando2)
        Call Comando_Fechar(alComando2(iIndice))
    Next

    DIInfo_Grava_EmTrans = SUCESSO

    Exit Function

Erro_DIInfo_Grava_EmTrans:

    DIInfo_Grava_EmTrans = gErr

    Select Case gErr

        Case 184521, 184698, 184699, 210601, 210602

        Case 184522
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 184523, 184524
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIINFO", gErr)

        Case 184525
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_DIINFO", gErr)

        Case 184526
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_DIINFO", gErr)

        Case 184527
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_DIINFO", gErr)

        Case 196556, 196557
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DINF", gErr)
            
        Case 196558
            Call Rotina_Erro(vbOKOnly, "ERRO_DIINFO_VINCULADA_NF", gErr, lNumNF)
        
        Case 196559, 196560

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184528)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando2) To UBound(alComando2)
        Call Comando_Fechar(alComando2(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function DIInfo_Exclui(ByVal objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_DIInfo_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 184506

    lErro = CF("DIInfo_Exclui_EmTrans", objDIInfo)
    If lErro <> SUCESSO Then gError 184507

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 184508

    DIInfo_Exclui = SUCESSO

    Exit Function

Erro_DIInfo_Exclui:

    DIInfo_Exclui = gErr

    Select Case gErr

        Case 184506
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 184507

        Case 184508
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184509)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function DIInfo_Exclui_EmTrans(ByVal objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim alComando(0 To 11) As Long
Dim iIndice As Integer
Dim lNumNF As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_DIInfo_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184510
    Next
    
    'Le a tabelaDINF
    lErro = Comando_Executar(alComando(11), "SELECT NF.NumNotaFiscal FROM NFImportacao AS DI, NFiscal AS NF, DIInfo WHERE NF.NumIntDoc = DI.NumIntNF AND DI.NumIntDI= DIInfo.NumIntDoc AND DIInfo.Numero = ?", _
                lNumNF, objDIInfo.sNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 196727

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(11))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196728
    
    If lErro = AD_SQL_SUCESSO Then gError 196729

    'Le a tabelaDIInfo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM DIInfo WHERE Numero= ? ", _
                0, lNumIntDoc, objDIInfo.sNumero)
    If lErro <> AD_SQL_SUCESSO Then gError 184511

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184512

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 184513
    
    objDIInfo.lNumIntDoc = lNumIntDoc

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 184514

    lErro = DIInfo_Exclui_Adicoes(objDIInfo, alComando)
    If lErro <> SUCESSO Then gError 184665
    
    lErro = DIInfo_Exclui_Despesas(objDIInfo, alComando)
    If lErro <> SUCESSO Then gError 184666
    
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM DIInfo", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 184515

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    DIInfo_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_DIInfo_Exclui_EmTrans:

    DIInfo_Exclui_EmTrans = gErr

    Select Case gErr

        Case 184510
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 184511, 184512
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIINFO", gErr)

        Case 184513
            Call Rotina_Erro(vbOKOnly, "ERRO_DIINFO_NAO_CADASTRADO", gErr, objDIInfo.sNumero)

        Case 184514
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_DIINFO", gErr)

        Case 184515
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DIINFO", gErr)

        Case 184665, 184666
        
        Case 196727, 196728
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DINF", gErr)
        
        Case 196729
            Call Rotina_Erro(vbOKOnly, "ERRO_DIINFO_VINCULADA_NF", gErr, lNumNF)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184516)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
  
Private Function DIInfo_Exclui_Adicoes(ByVal objDIInfo As ClassDIInfo, alComando() As Long) As Long

Dim lErro As Long, lNumIntDoc As Long

On Error GoTo Erro_DIInfo_Exclui_Adicoes

    lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM AdicaoDI WHERE NumIntDI = ?", 0, lNumIntDoc, objDIInfo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 184667
    
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184668
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 184669
        
        lErro = AdicaoDI_Exclui_Itens(lNumIntDoc, alComando)
        If lErro <> SUCESSO Then gError 184670
        
        lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM AdicaoDI", alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 184671
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184672
    
    Loop
    
    DIInfo_Exclui_Adicoes = SUCESSO
    
    Exit Function
    
Erro_DIInfo_Exclui_Adicoes:

    DIInfo_Exclui_Adicoes = gErr

    Select Case gErr

        Case 184667, 184668, 184669, 184671, 184672
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ADICAO_DI", gErr)

        Case 184670
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184673)

    End Select
    
    Exit Function

End Function

Private Function AdicaoDI_Exclui_Itens(ByVal lNumIntAdicaoDI As Long, alComando() As Long) As Long

Dim lErro As Long, lNumIntDoc As Long

On Error GoTo Erro_AdicaoDI_Exclui_Itens

    lErro = Comando_ExecutarPos(alComando(4), "SELECT NumIntDoc FROM ItensAdicaoDI WHERE NumIntAdicaoDI = ?", 0, lNumIntDoc, lNumIntAdicaoDI)
    If lErro <> AD_SQL_SUCESSO Then gError 184674
    
    lErro = Comando_BuscarProximo(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184675
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_LockExclusive(alComando(4))
        If lErro <> AD_SQL_SUCESSO Then gError 184676
    
        lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM ItensAdicaoDI", alComando(4))
        If lErro <> AD_SQL_SUCESSO Then gError 184677
        
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184678
    
    Loop
    
    AdicaoDI_Exclui_Itens = SUCESSO
    
    Exit Function
    
Erro_AdicaoDI_Exclui_Itens:

    AdicaoDI_Exclui_Itens = gErr

    Select Case gErr

        Case 184674, 184675, 184676, 184677, 184678
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEM_ADICAO_DI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184679)

    End Select
    
    Exit Function

End Function

Private Function DIInfo_Exclui_Despesas(ByVal objDIInfo As ClassDIInfo, alComando() As Long) As Long

Dim lErro As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_DIInfo_Exclui_Despesas

    lErro = Comando_ExecutarPos(alComando(6), "SELECT NumIntDoc FROM ImportCompl WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ?", 0, lNumIntDoc, IMPORTCOMPL_ORIGEM_DI, objDIInfo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 184680
    
    lErro = Comando_BuscarProximo(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184681
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_LockExclusive(alComando(6))
        If lErro <> AD_SQL_SUCESSO Then gError 184682
    
        lErro = Comando_ExecutarPos(alComando(7), "DELETE FROM ImportCompl", alComando(6))
        If lErro <> AD_SQL_SUCESSO Then gError 184683
        
        lErro = Comando_BuscarProximo(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184684
    
    Loop
        
    DIInfo_Exclui_Despesas = SUCESSO
    
    Exit Function
    
Erro_DIInfo_Exclui_Despesas:

    DIInfo_Exclui_Despesas = gErr

    Select Case gErr

        Case 184680, 184681, 184682, 184683, 184684
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DESPESA_DI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184685)

    End Select
    
    Exit Function

End Function

Private Function DIInfo_Insere_Adicoes(ByVal objDIInfo As ClassDIInfo, alComando() As Long) As Long

Dim lErro As Long, lNumIntDoc As Long, objAdicao As ClassAdicaoDI

On Error GoTo Erro_DIInfo_Insere_Adicoes

    lErro = CF("Config_ObterNumInt", "CRFATConfig", "NUM_INT_PROX_ADICAODI", lNumIntDoc, objDIInfo.colAdicoesDI.Count)
    If lErro <> SUCESSO Then gError 184687
    
    For Each objAdicao In objDIInfo.colAdicoesDI
    
        With objAdicao
            .lNumIntDI = objDIInfo.lNumIntDoc
            .lNumIntDoc = lNumIntDoc
            lNumIntDoc = lNumIntDoc + 1
        
            lErro = Comando_Executar(alComando(3), "INSERT INTO AdicaoDI (NumIntDoc, NumIntDI, Seq, IPICodigo, IIAliquota, IPIAliquota, PISAliquota, COFINSAliquota, " & _
                "ICMSAliquota, IIValor, IPIValor, PISValor, COFINSValor, ICMSValor, IPIBase, PISBase, COFINSBase, " & _
                "ICMSBase, ValorAduaneiro, DespesaAduaneira, TaxaSiscomex, CodFabricante, ICMSPercRedBase, NumDrawback) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                .lNumIntDoc, .lNumIntDI, .iSeq, .sIPICodigo, .dIIAliquota, .dIPIAliquota, .dPISAliquota, .dCOFINSAliquota, _
                .dICMSAliquota, .dIIValor, .dIPIValor, .dPISValor, .dCOFINSValor, .dICMSValor, .dIPIBase, .dPISBase, .dCOFINSBase, .dICMSBase, .dValorAduaneiro, .dDespesaAduaneira, .dTaxaSiscomex, .sCodFabricante, .dICMSPercRedBase, .sNumDrawback)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 184688
    
        lErro = AdicaoDI_Insere_Itens(objAdicao, alComando)
        If lErro <> SUCESSO Then gError 184689
        
    Next
    
    DIInfo_Insere_Adicoes = SUCESSO
    
    Exit Function
    
Erro_DIInfo_Insere_Adicoes:

    DIInfo_Insere_Adicoes = gErr

    Select Case gErr
        
        Case 184687, 184689
        
        Case 184688
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ADICAODI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184690)

    End Select
    
    Exit Function

End Function

Private Function DIInfo_Insere_Despesas(ByVal objDIInfo As ClassDIInfo, alComando() As Long) As Long

Dim lErro As Long, objDespesa As ClassImportCompl, lNumIntDoc As Long

On Error GoTo Erro_DIInfo_Insere_Despesas

    lErro = CF("Config_ObterNumInt", "CRFATConfig", "NUM_INT_PROX_IMPORTCOMPL", lNumIntDoc, objDIInfo.colDespesasDI.Count)
    If lErro <> SUCESSO Then gError 184691

    For Each objDespesa In objDIInfo.colDespesasDI
    
        With objDespesa
            .iTipoDocOrigem = IMPORTCOMPL_ORIGEM_DI
            .lNumIntDocOrigem = objDIInfo.lNumIntDoc
            .lNumIntDoc = lNumIntDoc
            lNumIntDoc = lNumIntDoc + 1
        
            lErro = Comando_Executar(alComando(2), "INSERT INTO ImportCompl (NumIntDoc, TipoDocOrigem, NumIntDocOrigem, Seq, Tipo, Descricao, Valor, ValorBase, Aliquota, Perc, Dias) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                .lNumIntDoc, .iTipoDocOrigem, .lNumIntDocOrigem, .iSeq, .iTipo, .sDescricao, .dValor, 0, 0, .dPerc, .iDias)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 184692
        
    Next
    
    DIInfo_Insere_Despesas = SUCESSO
    
    Exit Function
    
Erro_DIInfo_Insere_Despesas:

    DIInfo_Insere_Despesas = gErr

    Select Case gErr

        Case 184691

        Case 184692
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_DESPESADI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184693)

    End Select
    
    Exit Function

End Function

Private Function AdicaoDI_Insere_Itens(objAdicao As ClassAdicaoDI, alComando() As Long) As Long

Dim lErro As Long, objItem As ClassItemAdicaoDI, lNumIntDoc As Long

On Error GoTo Erro_AdicaoDI_Insere_Itens

    lErro = CF("Config_ObterNumInt", "CRFATConfig", "NUM_INT_PROX_ITEM_ADICAODI", lNumIntDoc, objAdicao.colItensAdicaoDI.Count)
    If lErro <> SUCESSO Then gError 184694
    
    For Each objItem In objAdicao.colItensAdicaoDI
    
        With objItem
            .lNumIntAdicaoDI = objAdicao.lNumIntDoc
            .lNumIntDoc = lNumIntDoc
            lNumIntDoc = lNumIntDoc + 1
            
            lErro = Comando_Executar(alComando(4), "INSERT INTO ItensAdicaoDI (NumIntDoc, NumIntAdicaoDI, Seq, Produto, UM, " & _
                "Quantidade, ValorUnitFOBNaMoeda, ValorUnitFOBEmReal, ValorUnitCIFNaMoeda, ValorUnitCIFEmReal, ValorTotalFOBNaMoeda, " & _
                "ValorTotalFOBEmReal, ValorTotalCIFNaMoeda, ValorTotalCIFEmReal, TotalCIFEmRealManual, PesoBruto, PesoLiq, Descricao, ValorUnitTrib, IPIUnidadePadraoValor) " & _
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lNumIntDoc, .lNumIntAdicaoDI, .iSeq, .sProduto, .sUM, _
                .dQuantidade, .dValorUnitFOBNaMoeda, .dValorUnitFOBEmReal, .dValorUnitCIFNaMoeda, .dValorUnitCIFEmReal, .dValorTotalFOBNaMoeda, _
                .dValorTotalFOBEmReal, .dValorTotalCIFNaMoeda, .dValorTotalCIFEmReal, .iTotalCIFEmRealManual, .dPesoBruto, .dPesoLiq, .sDescricao, .dValorUnitTrib, .dIPIUnidadePadraoValor)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 184695
            
    Next
    
    AdicaoDI_Insere_Itens = SUCESSO
    
    Exit Function
    
Erro_AdicaoDI_Insere_Itens:

    AdicaoDI_Insere_Itens = gErr

    Select Case gErr

        Case 184694

        Case 184695
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITEM_ADICAODI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184696)

    End Select
    
    Exit Function

End Function

Private Function NFiscal_Grava_NFImportacao(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long, objAdicaoDINF As ClassAdicaoDINF, objItemAdicaoDIItemNF As ClassItemAdicaoDIItemNF
Dim alComando(0 To 8) As Long, objAdicaoDIComplNF As ClassAdicaoDIComplNF, lNumIntDocImportCompl As Long
Dim lNumIntItemNF As Long, objItemNF As ClassItemNF, objImportCompl As ClassImportCompl, iIndice As Integer
Dim colItens As New Collection, objItem As ClassItemAdicaoDIItemNF, colComplNF As Collection
Dim colPedCompras As New Collection, colReqCompras As New Collection

On Error GoTo Erro_NFiscal_Grava_NFImportacao

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184705
    Next

    lErro = Comando_Executar(alComando(0), "INSERT INTO NFImportacao (NumIntNF, NumIntDI) VALUES (?,?)", objNFiscal.lNumIntDoc, objNFiscal.objNFImportacao.objDIInfo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 184706

    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPCC Then
        If objNFiscal.objNFImportacao.colItensNF.Count > 0 Then
            For Each objItemNF In objNFiscal.colItensNF
                Set objItemAdicaoDIItemNF = objNFiscal.objNFImportacao.colItensNF(1)
                Set objItem = New ClassItemAdicaoDIItemNF
                colItens.Add objItem
                objItem.lNumIntItemNF = objItemNF.lNumIntDoc
                objItem.lNumIntItemAdicaoDI = objItemAdicaoDIItemNF.lNumIntItemAdicaoDI
            Next
        End If
        Set colComplNF = New Collection
    Else
        Set colItens = objNFiscal.objNFImportacao.colItensNF
        Set colComplNF = objNFiscal.objNFImportacao.colComplNF
    End If

    For Each objItemAdicaoDIItemNF In colItens 'objNFiscal.objNFImportacao.colItensNF
    
        lNumIntItemNF = 0
        'obter numintitemnf
        For Each objItemNF In objNFiscal.colItensNF
            
            If objItemNF.iItem = objItemAdicaoDIItemNF.iItemNF Then
                lNumIntItemNF = objItemNF.lNumIntDoc
                Exit For
            End If
        
        Next
        If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPCC Then lNumIntItemNF = objItemAdicaoDIItemNF.lNumIntItemNF
        If lNumIntItemNF = 0 Then gError 184709
        
        lErro = Comando_Executar(alComando(2), "INSERT INTO ItemAdicaoDIItemNF(NumIntItemNF, NumIntItemAdicaoDI, ValorAduaneiro, ValorII, DespImpValorRateado) VALUES (?,?,?,?,?)", _
            lNumIntItemNF, objItemAdicaoDIItemNF.lNumIntItemAdicaoDI, objItemAdicaoDIItemNF.dValorAduaneiro, objItemAdicaoDIItemNF.dValorII, objItemAdicaoDIItemNF.dDespImpValorRateado)
        If lErro <> AD_SQL_SUCESSO Then gError 184708
        
    Next

    lErro = CF("Config_ObterNumInt", "CRFATConfig", "NUM_INT_PROX_IMPORTCOMPL", lNumIntDocImportCompl, objNFiscal.objNFImportacao.colComplNF.Count)
    If lErro <> SUCESSO Then gError 184711
    
    For Each objImportCompl In colComplNF 'objNFiscal.objNFImportacao.colComplNF
        
        lErro = Comando_Executar(alComando(3), "INSERT INTO ImportCompl (NumIntDoc, TipoDocOrigem, NumIntDocOrigem, Seq, Tipo, Descricao, ValorBase, Aliquota, Valor, Manual) VALUES (?,?,?,?,?,?,?,?,?,?)", _
            lNumIntDocImportCompl, IMPORTCOMPL_ORIGEM_NF, objNFiscal.lNumIntDoc, objImportCompl.iSeq, objImportCompl.iTipo, objImportCompl.sDescricao, objImportCompl.dValorBase, objImportCompl.dAliquota, objImportCompl.dValor, objImportCompl.iManual)
        If lErro <> AD_SQL_SUCESSO Then gError 184710
    
        lNumIntDocImportCompl = lNumIntDocImportCompl + 1
        
    Next

    'Se é uma importação integral ou a primeira remessa parcial, preenche a coleção de pedidos de compra a serem baixados
    If objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMP Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEIMPSE Then
   
        lErro = NFImportacao_Monta_PC(objNFiscal, colPedCompras, colReqCompras)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    If colPedCompras.Count > 0 Then
        
        'Inclui em ItemNFItemPC o ItemNF associado a ItemPC
        lErro = CF("ItemNFItemPC_Inclui", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'Atualiza as Quantidades dos ItensPC
        lErro = CF("ItensPedCompra_Atualiza_Quantidades", objNFiscal, colPedCompras)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'Inclui em ItemNFItemRC o ItemNF associado a ItemRC
        lErro = CF("ItemNFItemRC_Inclui", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'Atualiza as quantidades dos Itens de Requisição
        lErro = CF("ItensReqCompra_Atualiza_Quantidades", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objNFiscal.iFilialPedido = giFilialEmpresa
        
        'Verifica se existem Pedidos a serem Baixados
        lErro = CF("NFiscalEntCom_BaixaPedidos", colPedCompras, objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'Verifica se existem Requisições a serem Baixadas
        lErro = CF("NFiscalEntCom_BaixaRequisicoes", colReqCompras, objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscal_Grava_NFImportacao = SUCESSO
    
    Exit Function
    
Erro_NFiscal_Grava_NFImportacao:

    NFiscal_Grava_NFImportacao = gErr

    Select Case gErr

        Case 184705
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 184706
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NF_IMPORTACAO", gErr)
        
        Case 184707
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ADICAODICOMPLNF", gErr)
        
        Case 184708
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMADICAODIITEMNF", gErr)
        
        Case 184709
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_NF_IMPORTACAO", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184703)

    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function NFiscal_Exclui_NFImportacao(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long, alComando(1 To 3) As Long, iIndice As Integer

On Error GoTo Erro_NFiscal_Exclui_NFImportacao

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184747
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM NFImportacao WHERE NumIntNF = ?", objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 184744

    lErro = Comando_Executar(alComando(2), "DELETE FROM ItemAdicaoDIItemNF WHERE NumIntItemNF IN (SELECT NumIntDoc FROM ItensNFiscal WHERE NumIntNF = ?)", objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 184745

    lErro = Comando_Executar(alComando(3), "DELETE FROM ImportCompl WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ?", IMPORTCOMPL_ORIGEM_NF, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 184746

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscal_Exclui_NFImportacao = SUCESSO
    
    Exit Function
    
Erro_NFiscal_Exclui_NFImportacao:

    NFiscal_Exclui_NFImportacao = gErr

    Select Case gErr

        Case 184744 To 184747
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NFIMPORTACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184743)

    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Reserva_Grava_TipoPedidoSRVCod(ByVal objReserva As ClassReserva, ByVal lComando1 As Long, ByVal lComando2 As Long) As Long
'Faz as atualizações à Nível de Item de Pedido

Dim lErro As Long
Dim tItemPedido As typeItemPedido
Dim objMapBloqGen As ClassMapeamentoBloqGen
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Reserva_Grava_TipoPedidoSRVCod

    tItemPedido.sProduto = String(STRING_PRODUTO, 0)

    lErro = Comando_ExecutarPos(lComando1, "SELECT CodPedido,FilialEmpresa. Produto, Quantidade, QuantCancelada, QuantFaturada FROM ItensPedidoSRV WHERE NumIntDoc = ?", 0, tItemPedido.lCodPedido, tItemPedido.iFilialEmpresa, tItemPedido.sProduto, tItemPedido.dQuantidade, tItemPedido.dQuantCancelada, tItemPedido.dQuantFaturada, objReserva.lNumIntOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 198539

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198540

    'Se não achou o Item de Pedido --> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 198541

    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 198542

    tItemPedido.dQuantAFaturar = tItemPedido.dQuantidade - tItemPedido.dQuantFaturada - tItemPedido.dQuantCancelada
    
    If objReserva.dQuantidade > tItemPedido.dQuantAFaturar Then gError 198543

    If objReserva.sProduto <> tItemPedido.sProduto Then gError 198544

    lErro = Comando_ExecutarPos(lComando2, "UPDATE ItensPedidoSRV SET QuantReservada = QuantReservada + ?", lComando1, objReserva.dQuantidade)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198545

    objMapBloqGen.iTipoTelaBloqueio = MAPBLOQGEN_TIPOTELA_PEDIDOSRV
    
    lErro = CF("MapeamentoBloqGen_Le", objMapBloqGen)
    If lErro <> SUCESSO Then gError 198546
    
    'Passa os dados do Bloqueio para o Obj
    objPedidoVenda.iFilialEmpresa = tItemPedido.iFilialEmpresa
    objPedidoVenda.lCodigo = tItemPedido.lCodPedido
    
    'Lê o Documento
    lErro = CF(objMapBloqGen.sNomeFuncLeDoc, objPedidoVenda)
    If lErro <> SUCESSO Then gError 198547
    
    lErro = CF("BloqueiosGen_Altera", MAPBLOQGEN_TIPOTELA_PEDIDOSRV, objPedidoVenda)
    If lErro <> SUCESSO Then gError 198548
    
    Reserva_Grava_TipoPedidoSRVCod = SUCESSO

    Exit Function

Erro_Reserva_Grava_TipoPedidoSRVCod:

    Reserva_Grava_TipoPedidoSRVCod = gErr

    Select Case gErr

        Case 198539, 198540
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDOSRV", gErr)

        Case 198541
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE", gErr, objReserva.lDocOrigem, objReserva.lNumIntOrigem)

        Case 198542
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDIDOSRV", gErr, objReserva.lDocOrigem, objReserva.lNumIntOrigem)

        Case 198543
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_RESERVADA_MAIOR_FATURAR", gErr, objReserva.dQuantidade, tItemPedido.dQuantAFaturar)

        Case 198544
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SEM_PRODUTO", gErr, objReserva.sProduto)
        
        Case 198545
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ITENSPEDIDOSRV", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149518)

    End Select

    Exit Function

End Function

Function RelCustoImportacao_Prepara(ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal sProd_I As String, ByVal sProd_F As String, ByVal lFornIni As Long, ByVal lFornFim As Long, ByVal sDI As String, lNumIntRel As Long) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer, alComando(1 To 3) As Long

On Error GoTo Erro_RelCustoImportacao_Prepara

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 184871
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184872
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_PROX_REL_CUSTOIMP", lNumIntRel)
    If lErro <> SUCESSO Then gError 184873
    
    If sDI = "" Then
    
        lErro = RelCustoImportacao_Prepara2(dtDataIni, dtDataFim, sProd_I, sProd_F, lFornIni, lFornFim, lNumIntRel, alComando())
        If lErro <> SUCESSO Then gError 184875
    
    Else
    
        lErro = RelCustoImportacao_Prepara1(sProd_I, sProd_F, lFornIni, lFornFim, sDI, lNumIntRel, alComando())
        If lErro <> SUCESSO Then gError 184875
    
    End If
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 184874
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelCustoImportacao_Prepara = SUCESSO
    
    Exit Function
    
Erro_RelCustoImportacao_Prepara:

    RelCustoImportacao_Prepara = gErr

    Select Case gErr

        Case 184871
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 184872
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 184873, 184875
        
        Case 184874
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184870)

    End Select
    
    Call Transacao_Rollback
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function RelCustoImportacao_Prepara1(ByVal sProd_I As String, ByVal sProd_F As String, ByVal lFornIni As Long, ByVal lFornFim As Long, ByVal sDI As String, ByVal lNumIntRel As Long, alComando() As Long) As Long

Dim lErro As Long, objProdutoFilial As ClassProdutoFilial, iFilialForn As Integer
Dim objDIInfo As ClassDIInfo, objDespesa As ClassImportCompl, objAdicao As ClassAdicaoDI
Dim objItem As ClassItemAdicaoDI, lForn As Long, dValorItemaAux As Double
Dim dValorAduaneiroDI As Double, dValorDespesasLegais As Double, dValorTotalCreditos As Double
Dim dValorAcertosFiscais As Double, dFatorItemAdicao As Double, dFatorItemDI As Double
Dim dFatorCustoFinanceiro As Double, dFatorSeguroProduto As Double, dValorTotalBruto As Double
Dim dOverhead As Double, dComissFornec As Double, dComissFornecNormal As Double, dComissFornecMobimax As Double
Dim objTipoImportCompl As New ClassTipoImportCompl
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_RelCustoImportacao_Prepara1

    Set objDIInfo = New ClassDIInfo
        
    objDIInfo.sNumero = sDI
    
    lErro = CF("DIInfo_Le", objDIInfo)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 184876
    If lErro <> SUCESSO Then gError 184877
    
    dValorAduaneiroDI = Arredonda_Moeda(objDIInfo.dValorMercadoriaEmReal + objDIInfo.dValorFreteInternacEmReal + objDIInfo.dValorSeguroInternacEmReal)
    
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then gError 184877
    
    For Each objDespesa In objDIInfo.colDespesasDI
    
        objTipoImportCompl.iCodigo = objDespesa.iTipo
        lErro = CF("TiposImportCompl_Le", objTipoImportCompl)
        If lErro <> SUCESSO And lErro <> ERRO_ITEM_NAO_CADASTRADO Then gError 184876
        If lErro <> SUCESSO Then gError 184877
        
        Select Case objDespesa.iTipo
        
            Case Else
               
                If objTipoImportCompl.iIncluiNoValorAduaneiro = 0 Then
                    dValorDespesasLegais = dValorDespesasLegais + objDespesa.dValor
                Else
                    dValorAduaneiroDI = Arredonda_Moeda(dValorAduaneiroDI + objDespesa.dValor)
                End If
        
        End Select
    
    Next
    
    For Each objAdicao In objDIInfo.colAdicoesDI
    
        For Each objItem In objAdicao.colItensAdicaoDI
    
            If objItem.sProduto >= sProd_I And (sProd_F = "" Or (objItem.sProduto <= sProd_F)) Then
            
                lForn = objDIInfo.lFornTrading
                            
                If lForn = 0 Or (lForn >= lFornIni And (lFornFim = 0 Or (lForn <= lFornFim))) Then
            
                    If objAdicao.dValorAduaneiro = 0 Then
                        dFatorItemAdicao = 0
                    Else
                        dFatorItemAdicao = objItem.dValorTotalCIFEmReal / objAdicao.dValorAduaneiro
                    End If
                    
                    If dValorAduaneiroDI = 0 Then
                        dFatorItemDI = 0
                    Else
                        dFatorItemDI = objItem.dValorTotalCIFEmReal / dValorAduaneiroDI
                    End If
                    
                    If objFilialEmpresa.iSuperSimples <> 0 Then
                        dValorTotalCreditos = 0
                    Else
                        dValorTotalCreditos = Arredonda_Moeda(dFatorItemAdicao * (IIf(objFilialEmpresa.iContribuinteIPI <> 0, objAdicao.dIPIValor, 0) + IIf(objFilialEmpresa.iPISNaoCumulativo <> 0, objAdicao.dPISValor, 0) + IIf(objFilialEmpresa.iCOFINSNaoCumulativo <> 0, objAdicao.dCOFINSValor, 0) + IIf(objFilialEmpresa.sInscricaoEstadual <> "", objAdicao.dICMSValor, 0)))
                    End If
                    'CIF + impostos + despesas legais
                    dValorItemaAux = Arredonda_Moeda(objItem.dValorTotalCIFEmReal + (dFatorItemAdicao * objAdicao.dIIValor) + Arredonda_Moeda(dFatorItemAdicao * (objAdicao.dIPIValor + objAdicao.dPISValor + objAdicao.dCOFINSValor + objAdicao.dICMSValor)) + (dFatorItemDI * dValorDespesasLegais))
                    dValorTotalBruto = Arredonda_Moeda(dValorItemaAux + (dFatorItemDI * dValorAcertosFiscais) + (dValorItemaAux * (dFatorCustoFinanceiro + dFatorSeguroProduto + dOverhead)) + (objItem.dValorTotalCIFEmReal * dComissFornec))
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO RelCustoImportacaoAux (NumIntRel, DINumIntDoc, " & _
                        "AdicaoDINumIntDoc, ItensAdicaoDINumIntDoc, ValorIIEmReal, ValorIPIEmReal, ValorPISEmReal, ValorCOFINSEmReal, " & _
                        "ValorICMSEmReal, ValorDespesasEmReal, " & _
                        "ValorTotalBruto, ValorTotalCreditos, ValorTotalLiquido, Fornecedor) " & _
                        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                        lNumIntRel, objDIInfo.lNumIntDoc, objAdicao.lNumIntDoc, objItem.lNumIntDoc, Arredonda_Moeda(dFatorItemAdicao * objAdicao.dIIValor), _
                        Arredonda_Moeda(dFatorItemAdicao * objAdicao.dIPIValor), Arredonda_Moeda(dFatorItemAdicao * objAdicao.dPISValor), Arredonda_Moeda(dFatorItemAdicao * objAdicao.dCOFINSValor), Arredonda_Moeda(dFatorItemAdicao * objAdicao.dICMSValor), _
                        Arredonda_Moeda(dFatorItemDI * dValorDespesasLegais), _
                        Arredonda_Moeda(dValorTotalBruto), Arredonda_Moeda(dValorTotalCreditos), Arredonda_Moeda(dValorTotalBruto - dValorTotalCreditos), lForn)
                    If lErro <> AD_SQL_SUCESSO Then gError 184881
            
                End If
                
            End If
    
        Next
        
    Next
    
    RelCustoImportacao_Prepara1 = SUCESSO
    
    Exit Function
    
Erro_RelCustoImportacao_Prepara1:

    RelCustoImportacao_Prepara1 = gErr

    Select Case gErr

        Case 184876, 184878
        
        Case 184877
            Call Rotina_Erro(vbOKOnly, "ERRO_DIINFO_NAO_CADASTRADO", gErr, objDIInfo.sNumero)
        
        Case 184879, 184880
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILFORNCATEG", gErr)
        
        Case 184881
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELCUSTOIMPAUX", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184882)

    End Select
    
    Exit Function

End Function

Private Function RelCustoImportacao_Prepara2(ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal sProd_I As String, ByVal sProd_F As String, ByVal lFornIni As Long, ByVal lFornFim As Long, ByVal lNumIntRel As Long, alComando() As Long) As Long

Dim lErro As Long, sDI As String

On Error GoTo Erro_RelCustoImportacao_Prepara2

    If dtDataFim = DATA_NULA Then dtDataFim = gdtDataAtual
    
    sDI = String(STRING_DI_NUMERO, 0)
    lErro = Comando_Executar(alComando(3), "SELECT Numero FROM DIInfo WHERE Data >= ? AND Data <= ? AND FilialEmpresa = ? ORDER BY Data", sDI, dtDataIni, dtDataFim, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 184884
    
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184885
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = RelCustoImportacao_Prepara1(sProd_I, sProd_F, lFornIni, lFornFim, sDI, lNumIntRel, alComando())
        If lErro <> SUCESSO Then gError 184887
    
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184888
    
    Loop
    
    RelCustoImportacao_Prepara2 = SUCESSO
    
    Exit Function
    
Erro_RelCustoImportacao_Prepara2:

    RelCustoImportacao_Prepara2 = gErr

    Select Case gErr

        Case 184887
        
        Case 184884, 184885, 184888
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIINFO_RELCUSTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184883)

    End Select
    
    Exit Function

End Function

Function INV_Trata_CodigoBarras(objProduto As ClassProduto) As Long
    INV_Trata_CodigoBarras = SUCESSO
End Function

Function ItemNFEnt_Devolucao_Customizado(objNFiscal As ClassNFiscal, objItemNF As ClassItemNF) As Long
    ItemNFEnt_Devolucao_Customizado = SUCESSO
End Function

Function ItemNFEnt_Excluir_Customizado(ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF) As Long
    ItemNFEnt_Excluir_Customizado = SUCESSO
End Function
            
Function ItemNFEnt_Cancelar_Customizado(ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF) As Long
    ItemNFEnt_Cancelar_Customizado = SUCESSO
End Function

Function DV_EAN13(sCodigo As String) As Long
    
Dim i As Integer
Dim iSoma As Integer
Dim iDigito As Integer
    
On Error GoTo Erro_DV_EAN13
    
    If Len(sCodigo) <> 12 Then gError 199865
    
    If Not IsNumeric(sCodigo) Then gError 199866
    
    For i = 1 To 12
        
        iSoma = iSoma + Mid(sCodigo, i, 1) * (1 + (((i + 1) Mod 2) * 2))
    
    Next
    
    iDigito = 10 - (iSoma Mod 10)
    
    If iDigito = 10 Then iDigito = 0
    
    sCodigo = sCodigo & iDigito
        
    DV_EAN13 = SUCESSO
    
    Exit Function

Erro_DV_EAN13:
            
    DV_EAN13 = gErr

    Select Case gErr

        Case 199865
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOBARRAS_TAMANHO_12", gErr)
        
        Case 199866
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_NUMERICO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199867)

    End Select
    
    Exit Function

End Function

Private Function NF_Trata_MovEst_DDAI(ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal objMovEstoque As ClassMovEstoque) As Long
'Rateia custos de uma nf pelos itens da nf original (inicialmente criada para nf de despesas de importacao)

Dim lErro As Long
Dim objNFO As ClassNFiscal, objItemNFO As ClassItemNF, objProduto As ClassProduto
Dim dValorRatear As Double, dValorRateado As Double, dValorBaseRateio As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim dQuantItemNFUMEst As Double, dFator As Double
Dim colInsumos As New Collection, colRastroMovto As New Collection
Dim objItemMovEstoque As ClassItemMovEstoque, bExisteItemSemCtrEst As Boolean
Dim dValorBaseRateioComCred As Double, dValorBaseRateioSemCred As Double
Dim bTrataDifComSemCred As Boolean, bItemComCred As Boolean, bItemOrigComCred As Boolean
Dim dQtdeItemAloc As Double

On Error GoTo Erro_NF_Trata_MovEst_DDAI
    
   'calcula o custo do item da NF em função do credito do ICMS e do IPI
    lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    dValorRatear = objItemNF.dCusto
    
    Set objNFO = New ClassNFiscal

    objNFO.lNumIntDoc = objNFiscal.lNumIntNotaOriginal

    lErro = CF("NFiscal_Le", objNFO)
    If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
    
    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê a parte de tributação dos itens da Nota Fiscal
    lErro = CF("NFiscalItens_Le_Trib", objNFO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'distribuicao
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Rateia o custo pela NF Original
    dValorBaseRateio = 0
    dValorBaseRateioSemCred = 0
    dValorBaseRateioComCred = 0
    For Each objItemNFO In objNFO.colItens
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objItemNFO.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        objItemNFO.iControleEstoque = objProduto.iControleEstoque
        objItemNFO.iClasseUM = objProduto.iClasseUM
        objItemNFO.sUMEstoque = objProduto.sSiglaUMEstoque
        objItemNFO.iApropriacaoProd = objProduto.iApropriacaoCusto
    
        'If objItemNFO.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
            dValorBaseRateio = dValorBaseRateio + Arredonda_Moeda(objItemNFO.dPrecoUnitario * objItemNFO.dQuantidade) - objItemNFO.dValorDesconto
        'End If
        
        If objItemNFO.objTributacao.dICMSCredito = 0 Then
            dValorBaseRateioSemCred = dValorBaseRateioSemCred + Arredonda_Moeda(objItemNFO.dPrecoUnitario * objItemNFO.dQuantidade) - objItemNFO.dValorDesconto
        Else
            dValorBaseRateioComCred = dValorBaseRateioComCred + Arredonda_Moeda(objItemNFO.dPrecoUnitario * objItemNFO.dQuantidade) - objItemNFO.dValorDesconto
        End If
        
        'Acumula a quantidade alocada pois não necessariamente bate com a quantidade do item (NFs Parciais)
        dQtdeItemAloc = 0
        For Each objItemNFAloc In objItemNFO.ColAlocacoes
            If objItemNFAloc.dQuantidade > 0 Then
                dQtdeItemAloc = dQtdeItemAloc + objItemNFAloc.dQuantidade
            End If
        Next
        objItemNFO.dQuantidadeAloc = dQtdeItemAloc
    
    Next
    
    bTrataDifComSemCred = False
    If objNFiscal.iTipoNFiscal = DOCINFO_CFECT And objNFiscal.colItens.Count > 1 Then
        bTrataDifComSemCred = True
        If objItemNF.objTributacao.dICMSCredito = 0 Then
            dValorBaseRateio = dValorBaseRateioSemCred
            bItemComCred = False
        Else
            dValorBaseRateio = dValorBaseRateioComCred
            bItemComCred = True
        End If
    End If
    
    dValorRateado = 0
    bExisteItemSemCtrEst = False
    For Each objItemNFO In objNFO.colItens
    
        If objItemNFO.objTributacao.dICMSCredito = 0 Then
            bItemOrigComCred = False
        Else
            bItemOrigComCred = True
        End If
        
        'Se controla estoque e não trata de forma diferenciada com/sem crédito ou ambos são iguais
        If objItemNFO.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And ((Not bTrataDifComSemCred) Or (bTrataDifComSemCred And bItemComCred = bItemOrigComCred)) Then
            objItemNFO.dCusto = Arredonda_Moeda(dValorRatear * ((Arredonda_Moeda(objItemNFO.dPrecoUnitario * objItemNFO.dQuantidade) - objItemNFO.dValorDesconto) / dValorBaseRateio))
            dValorRateado = dValorRateado + objItemNFO.dCusto
        Else
            objItemNFO.dCusto = 0
            If objItemNFO.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then bExisteItemSemCtrEst = True
        End If
    Next
    
    If Not bExisteItemSemCtrEst Then
        If Abs(dValorRateado - dValorRatear) > DELTA_VALORMONETARIO Then
            For Each objItemNFO In objNFO.colItens
            
                If objItemNFO.objTributacao.dICMSCredito = 0 Then
                    bItemOrigComCred = False
                Else
                    bItemOrigComCred = True
                End If
            
                If objItemNFO.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And ((Not bTrataDifComSemCred) Or (bTrataDifComSemCred And bItemComCred = bItemOrigComCred)) Then
                    objItemNFO.dCusto = objItemNFO.dCusto + dValorRateado - dValorRatear
                    Exit For
                End If
            Next
        End If
    End If

    'Gera o obj com o movimento de estoque
    For Each objItemNFO In objNFO.colItens
        If objItemNFO.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
            
            lErro = CF("UM_Conversao", objItemNFO.iClasseUM, objItemNFO.sUnidadeMed, objItemNFO.sUMEstoque, dFator)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If objNFO.iTipoNFiscal = DOCINFO_NFIEIMPSE Then
                dQuantItemNFUMEst = objItemNFO.dQuantidadeAloc 'Proporcionaliza considerando somente a quantidade alocada total
            Else
                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNFO.dQuantidade * dFator))
            End If
                    
            For Each objItemNFAloc In objItemNFO.ColAlocacoes
            
                If objItemNFAloc.dQuantidade > 0 And objItemNFO.dCusto <> 0 Then
                    
                    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNFO.dCusto * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNFO.sProduto, objItemNFO.sDescricaoItem, objItemNFO.sUnidadeMed, 0, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    objItemMovEstoque.iClasseUM = objItemNFO.iClasseUM
                    objItemMovEstoque.iControleEstoque = objItemNFO.iControleEstoque
                    objItemMovEstoque.sSiglaUMEst = objItemNFO.sUMEstoque
                    objItemMovEstoque.iApropriacaoProd = objItemNFO.iApropriacaoProd
                    
                End If
            
            Next
        End If
    Next
        
    NF_Trata_MovEst_DDAI = SUCESSO
    
    Exit Function

Erro_NF_Trata_MovEst_DDAI:
            
    NF_Trata_MovEst_DDAI = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199867)

    End Select
    
    Exit Function

End Function

Function RelRastroLoteProd_Prepara(lNumIntRel As Long, ByVal sProduto As String, ByVal sLote As String, ByVal iFilial As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim alComando(0 To 1) As Long
Dim lSeq As Long, iNivel As Integer, lSeqCab As Long
Dim sOP As String, iFilialEmp As Integer, dQtd As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_RelRastroLoteProd_Prepara

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 206955

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206956
    Next
       
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "ESTConfig", "NUMINTREL_RELRASTROLOTEPROD", lNumIntRel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objProduto.sCodigo = sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
    If objProduto.iRastro <> PRODUTO_RASTRO_OP Then iFilial = 0
    
    sOP = String(STRING_OPCODIGO, 0)
    
    'Le as Produções entradas
    lErro = Comando_Executar(alComando(1), "SELECT E.CodigoOP, E.FilialEmpresa, SUM(M.Quantidade) FROM MovimentoEstoque AS E,RastreamentoMovto AS M, RastreamentoLote AS L WHERE E.NumIntDoc = M.NumIntDocOrigem AND M.TipoDocOrigem = 0 AND M.NumIntDocLote = L.NumIntDoc AND M.Produto = ? AND L.Lote = ? AND L.FilialOP = ? AND E.TipoMov IN (7) GROUP BY E.CodigoOP, E.FilialEmpresa", sOP, iFilialEmp, dQtd, sProduto, sLote, iFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 206960
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206961

    Do While lErro <> AD_SQL_SEM_DADOS
    
        lSeqCab = lSeqCab + 1
    
        lErro = Comando_Executar(alComando(0), "INSERT INTO RelRastroLoteProdCab (NumIntRel, SeqCab, OP, FilialEmp, Quantidade, Produto, Lote, Filial, Descricao, UM) VALUES (?,?,?,?,?,?,?,?,?,?)", lNumIntRel, lSeqCab, sOP, iFilialEmp, dQtd, sProduto, sLote, iFilial, objProduto.sDescricao, objProduto.sSiglaUMEstoque)
        If lErro <> AD_SQL_SUCESSO Then gError 206957
        
        lErro = RelRastroLoteProd_Prepara_Lote(lNumIntRel, sOP, iFilialEmp, iNivel, lSeqCab, lSeq)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206961
        
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 206958

    RelRastroLoteProd_Prepara = SUCESSO

    Exit Function

Erro_RelRastroLoteProd_Prepara:

    RelRastroLoteProd_Prepara = gErr

    Select Case gErr

        Case 206955
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 206956
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 206957
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELRASTROLOTEPROD", gErr)

        Case 206958
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 206960, 206961
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
            
        Case ERRO_SEM_MENSAGEM
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELRASTROLOTEPROD", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206959)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Function RelRastroLoteProd_Prepara_Lote(ByVal lNumIntRel As Long, ByVal sOP As String, ByVal iFilialEmp As Integer, ByVal iNivel As Integer, ByVal lSeqCab As Long, lSeq As Long, Optional ByVal dFator As Double = 1) As Long
'Recebe a OP e lé as requisições dela

Dim lErro As Long
Dim objProduto As ClassProduto, lFornecedor As Long, iFilialForn As Integer
Dim sProduto As String, sLote As String, iFilialOP As Integer, dQtd As Double
Dim alComando(0 To 4) As Long
Dim iIndice As Integer, sFornecedor As String
Dim sOP2 As String, iFilialEmp2 As Integer, dQtdAux As Double, dFatorAux As Double

On Error GoTo Erro_RelRastroLoteProd_Prepara_Lote

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 206956
    Next
    
    iNivel = iNivel + 1
    
    sProduto = String(STRING_PRODUTO, 0)
    sLote = String(STRING_LOTE_RASTREAMENTO, 0)
        
    'Le as requisições para produção
    lErro = Comando_Executar(alComando(2), "SELECT  M.Produto, L.Lote, L.FilialOP, SUM(M.Quantidade) FROM MovimentoEstoque AS E,RastreamentoMovto AS M, RastreamentoLote AS L WHERE E.NumIntDoc = M.NumIntDocOrigem AND M.TipoDocOrigem = 0 AND M.NumIntDocLote = L.NumIntDoc AND E.CodigoOP = ? AND E.FilialEmpresa = ? AND E.TipoMov = 6 GROUP BY M.Produto, L.Lote, L.FilialOP", sProduto, sLote, iFilialOP, dQtd, sOP, iFilialEmp)
    If lErro <> AD_SQL_SUCESSO Then gError 206963
    
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206964
    
    Do While lErro <> AD_SQL_SEM_DADOS
   
        lSeq = lSeq + 1
        
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        If objProduto.iRastro <> PRODUTO_RASTRO_OP Then iFilialOP = 0
       
        sFornecedor = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
                
        'Le o fornecedor da primeira NF de entrada
        lErro = Comando_Executar(alComando(4), "SELECT  N.Fornecedor, F.NomeReduzido, N.FilialForn FROM MovimentoEstoque AS E,RastreamentoMovto AS M, RastreamentoLote AS L, ItensNFiscal AS I, NFiscal AS N, Fornecedores AS F WHERE E.NumIntDoc = M.NumIntDocOrigem AND M.TipoDocOrigem = 0 AND M.NumIntDocLote = L.NumIntDoc AND N.Fornecedor = F.Codigo AND N.NumIntDoc = I.NumIntNF AND I.NumIntDoc = E.NumIntDocOrigem AND E.TipoNumIntDocOrigem = 1 AND M.Produto = ? AND L.Lote = ? AND L.FilialOP = ? ORDER BY N.DataEmissao", lFornecedor, sFornecedor, iFilialForn, sProduto, sLote, iFilialOP)
        If lErro <> AD_SQL_SUCESSO Then gError 206967
        
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206968
        
        If lErro <> AD_SQL_SUCESSO Then
            sFornecedor = ""
            lFornecedor = 0
            iFilialForn = 0
        End If
        
        lErro = Comando_Executar(alComando(3), "INSERT INTO RelRastroLoteProdDet (NumIntRel, SeqCab, Seq, Produto, Lote, Filial, Descricao, UM, Quantidade, OP, FilialEmp, NomeFornecedor, Fornecedor, Nivel) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntRel, lSeqCab, lSeq, sProduto, sLote, iFilialOP, objProduto.sDescricao, objProduto.sSiglaUMEstoque, dQtd * dFator, sOP, iFilialEmp, sFornecedor, lFornecedor, iNivel)
        If lErro <> AD_SQL_SUCESSO Then gError 206962
    
        sOP2 = String(STRING_OPCODIGO, 0)
        
        'Le as Produções entradas para esse produto lote
        lErro = Comando_Executar(alComando(1), "SELECT E.CodigoOP, E.FilialEmpresa, SUM(M.Quantidade) FROM MovimentoEstoque AS E,RastreamentoMovto AS M, RastreamentoLote AS L WHERE E.NumIntDoc = M.NumIntDocOrigem AND M.TipoDocOrigem = 0 AND M.NumIntDocLote = L.NumIntDoc AND M.Produto = ? AND L.Lote = ? AND L.FilialOP = ? AND E.TipoMov IN (7) GROUP BY E.CodigoOP, E.FilialEmpresa", sOP2, iFilialEmp2, dQtdAux, sProduto, sLote, iFilialOP)
        If lErro <> AD_SQL_SUCESSO Then gError 206960
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206961
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            dFatorAux = (dQtd * dFator) / (dQtdAux)
        
            lErro = RelRastroLoteProd_Prepara_Lote(lNumIntRel, sOP2, iFilialEmp2, iNivel, lSeqCab, lSeq, dFatorAux)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
     
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206965
        
        Loop
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 206965
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelRastroLoteProd_Prepara_Lote = SUCESSO

    Exit Function

Erro_RelRastroLoteProd_Prepara_Lote:

    RelRastroLoteProd_Prepara_Lote = gErr

    Select Case gErr
    
        Case 206956
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 206960, 206961, 206963 To 206966
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
        
        Case 206962
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELRASTROLOTEPROD", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206967)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function ProdutoCNAE_Grava(ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ProdutoCNAE_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 208106

    lErro = CF("ProdutoCNAE_Grava_EmTrans", objProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 208107

    ProdutoCNAE_Grava = SUCESSO

    Exit Function

Erro_ProdutoCNAE_Grava:

    ProdutoCNAE_Grava = gErr

    Select Case gErr

        Case 208106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 208107
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208108)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function ProdutoCNAE_Grava_EmTrans(ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim alComando(0 To 4) As Long
Dim iIndice As Integer
Dim objProdutoCNAE As ClassProdutoCNAE
Dim objCodTribMun As ClassCodTribMun
Dim iSeq As Integer, iAux As Integer

On Error GoTo Erro_ProdutoCNAE_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 208109
    Next
    
    Set objProdutoCNAE = objProduto.objProdutoCNAE
    
    'Le a tabelaProdutoCNAE
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Tipo FROM Produtos WHERE Codigo= ? ", _
                0, iAux, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 208110

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208111
    
    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(2))
    If lErro <> SUCESSO Then gError 208112
    
    lErro = Comando_ExecutarPos(alComando(3), "UPDATE Produtos SET ISSQN= ?, NBS = ? ", alComando(2), objProduto.sISSQN, objProduto.sNBS)
    If lErro <> AD_SQL_SUCESSO Then gError 208118

    'Le a tabela ProdutoCNAE
    lErro = Comando_ExecutarPos(alComando(0), "SELECT LocServCliente FROM ProdutoCNAE WHERE Produto= ? ", _
                0, iAux, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 208113

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208114

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 208115

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE ProdutoCNAE SET CNAE= ?, LocServCliente = ?, LocIncidImpCliente=? ", alComando(0), _
                objProdutoCNAE.sCNAE, objProdutoCNAE.iLocServCliente, objProdutoCNAE.iLocIncidImpCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 208116

    Else

        lErro = Comando_Executar(alComando(1), "INSERT INTO ProdutoCNAE( Produto, CNAE, LocServCliente, LocIncidImpCliente) VALUES (?,?,?,?)", _
                objProduto.sCodigo, objProdutoCNAE.sCNAE, objProdutoCNAE.iLocServCliente, objProdutoCNAE.iLocIncidImpCliente)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208117

    End If
        
    lErro = CodTribMun_Exclui(objProduto.sCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iSeq = 0
    For Each objCodTribMun In objProdutoCNAE.colCidades
        iSeq = iSeq + 1
        lErro = Comando_Executar(alComando(4), "INSERT INTO CodTribMun( Cidade, Produto, CodTribMun, Aliquota, Seq) VALUES (?,?,?,?,?)", _
                objCodTribMun.lCidade, objProduto.sCodigo, objCodTribMun.sCodTribMun, objCodTribMun.dAliquota, iSeq)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208119
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ProdutoCNAE_Grava_EmTrans = SUCESSO

    Exit Function

Erro_ProdutoCNAE_Grava_EmTrans:

    ProdutoCNAE_Grava_EmTrans = gErr

    Select Case gErr

        Case 208109
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 208110, 208111
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case 208112
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOS", gErr)

        Case 208113, 208114
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOCNAE", gErr)

        Case 208115
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRODUTOCNAE", gErr)

        Case 208116
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRODUTOCNAE", gErr)

        Case 208117
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRODUTOCNAE", gErr)

        Case 208118
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_PRODUTOCNAE", gErr)
            
        Case 208119
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CODTRIBMUN", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208120)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function CodTribMun_Exclui(ByVal sProduto As String) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_CodTribMun_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 208121
    Next

    'Le a tabela CodTribMun
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Seq FROM CodTribMun WHERE Produto= ? ", _
                0, iAux, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 208122

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208123

    'Se não existir => ERRO
    Do While lErro <> AD_SQL_SEM_DADOS

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 208124
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CodTribMun", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 208125
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208126
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CodTribMun_Exclui = SUCESSO

    Exit Function

Erro_CodTribMun_Exclui:

    CodTribMun_Exclui = gErr

    Select Case gErr

        Case 208121
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208122, 208123, 208126
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CODTRIBMUN", gErr)

        Case 208124
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_CODTRIBMUN", gErr)

        Case 208125
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CODTRIBMUN", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208127)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ProdutoCNAE_Exclui(ByVal sProduto As String) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_ProdutoCNAE_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 208128
    Next

    'Le a tabela CodTribMun
    lErro = Comando_ExecutarPos(alComando(0), "SELECT LocServCliente FROM ProdutoCNAE WHERE Produto= ? ", _
                0, iAux, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 208129

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208130

    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 208131
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ProdutoCNAE", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 208132
            
        lErro = CodTribMun_Exclui(sProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ProdutoCNAE_Exclui = SUCESSO

    Exit Function

Erro_ProdutoCNAE_Exclui:

    ProdutoCNAE_Exclui = gErr

    Select Case gErr

        Case 208128
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208129, 208130
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOCNAE", gErr)

        Case 208131
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_PRODUTOCNAE", gErr)

        Case 208132
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PRODUTOCNAE", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208133)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ProdutoGrade_Grava(ByVal objProduto As ClassProduto, ByVal colProdutos As Collection, ByVal iAlmoxarifado As Integer, ByVal iNumCat As Integer) As Long
'grava na tabela ESTConfig uma coleção de registros de configuração

Dim lErro As Long, iIndice As Integer
Dim lTransacao As Long
Dim objProdutoAux As ClassProduto, objProdutoBD As ClassProduto
Dim colTabelaPrecoItem As New Collection
Dim bExisteProd As Boolean
Dim objProdutoCategoria As ClassProdutoCategoria, objProdutoCategoriaAux As ClassProdutoCategoria
Dim iAux As Integer, dResiduo As Double
Dim objEstoqueProduto As ClassEstoqueProduto
Dim colRastreamento As New Collection
Dim iAlmoxarifadoPadrao As Integer, iCount As Integer
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde

On Error GoTo Erro_ProdutoGrade_Grava

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 208188

    Set objProdutoBD = New ClassProduto
    objProdutoBD.sCodigo = objProduto.sCodigo

    lErro = CF("Produto_Le", objProdutoBD)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
    If lErro <> SUCESSO Then
    
        'iAux = objProduto.iControleEstoque
        'objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE
    
        'Grava o Produto Gerencial se ele não existir
        lErro = CF("Produto_Grava_Trans", objProduto, colTabelaPrecoItem)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'objProduto.iControleEstoque = iAux
        
    End If

    objFrmAguarde.iTotalItens = colProdutos.Count
    Call objFrmAguardeTela.Inicializa_Progressao(objFrmAguarde)

    iAux = 0
    iIndice = 0
    For Each objProdutoAux In colProdutos
        iIndice = iIndice + 1
        If objProdutoAux.iAtivo = PRODUTO_ATIVO Then
        
            If Len(Trim(objProdutoAux.sNomeReduzido)) > STRING_PRODUTO_NOME_REDUZIDO Then gError 208189
            If Len(Trim(objProdutoAux.sDescricao)) > STRING_PRODUTO_DESCRICAO_TELA Then gError 208190
            If Len(Trim(objProdutoAux.sFigura)) > STRING_PRODUTO_FIGURA Then gError 208191
        
            Set objProdutoBD = New ClassProduto
            objProdutoBD.sCodigo = objProdutoAux.sCodigo
        
            lErro = CF("Produto_Le", objProdutoBD)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

            If lErro = SUCESSO Then gError 208192 'O código já existe
            
            Set objProdutoBD = New ClassProduto
            objProdutoBD.sNomeReduzido = objProdutoAux.sNomeReduzido
            
            lErro = CF("Produto_Le_NomeReduzido", objProdutoBD)
            If lErro <> SUCESSO And lErro <> 26927 Then gError ERRO_SEM_MENSAGEM

            If lErro = SUCESSO Then gError 208193 'Já existe o nome reduzido
            
            lErro = CF("ProdutoGrade_Le_ProdCat", objProdutoAux, bExisteProd)
            If lErro <> SUCESSO And lErro <> 26927 Then gError ERRO_SEM_MENSAGEM
            
            If bExisteProd Then gError 208194 'Já existe o mesmo filho do pai de grade com as mesmas características da grade
            
            'Produto sem Problemas
            'Herda as características do pai
            If iAux = 0 Then iAux = objProdutoAux.iNivel
            
            For iCount = objProduto.colCategoriaItem.Count To iNumCat + 1 Step -1
                objProduto.colCategoriaItem.Remove iCount
            Next

            For Each objProdutoCategoria In objProduto.colCategoriaItem
                Set objProdutoCategoriaAux = New ClassProdutoCategoria
                objProdutoCategoriaAux.sCategoria = objProdutoCategoria.sCategoria
                objProdutoCategoriaAux.sItem = objProdutoCategoria.sItem
                objProdutoAux.colCategoriaItem.Add objProdutoCategoriaAux
            Next
            
            For Each objProdutoCategoriaAux In objProdutoAux.colCategoriaItem
                objProdutoCategoriaAux.sProduto = objProdutoAux.sCodigo
            Next

            objProdutoAux.dIPIAliquota = objProduto.dIPIAliquota
            objProdutoAux.iClasseUM = objProduto.iClasseUM
            objProdutoAux.iCompras = objProduto.iCompras
            objProdutoAux.iControleEstoque = objProduto.iControleEstoque
            objProdutoAux.iFaturamento = objProduto.iFaturamento
            objProdutoAux.iGerencial = NAO_GERENCIAL
            objProdutoAux.iKitBasico = objProduto.iKitBasico
            objProdutoAux.iKitInt = objProduto.iKitInt
            objProdutoAux.iNatureza = objProduto.iNatureza
            objProdutoAux.iNivel = objProduto.iNivel + 1
            objProdutoAux.iOrigemMercadoria = objProduto.iOrigemMercadoria
            objProdutoAux.iPCP = objProduto.iPCP
            objProdutoAux.iTipo = objProduto.iTipo
            objProdutoAux.sIPICodigo = objProduto.sIPICodigo
            objProdutoAux.sSiglaUMCompra = objProduto.sSiglaUMCompra
            objProdutoAux.sSiglaUMEstoque = objProduto.sSiglaUMEstoque
            objProdutoAux.sSiglaUMTrib = objProduto.sSiglaUMTrib
            objProdutoAux.sSiglaUMVenda = objProduto.sSiglaUMVenda
            objProdutoAux.dIPIAliquota = objProduto.dIPIAliquota
            objProdutoAux.iCreditoIPI = objProduto.iCreditoIPI
            objProdutoAux.iIPIAgregaCusto = objProduto.iIPIAgregaCusto
        
            'Grava os Produtos Filhos
            dResiduo = objProdutoAux.dResiduo
            objProdutoAux.dResiduo = objProduto.dResiduo
            lErro = CF("Produto_Grava_Trans", objProdutoAux, colTabelaPrecoItem)
            objProdutoAux.dResiduo = dResiduo
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If iAlmoxarifado <> 0 Then
            
                Set objEstoqueProduto = New ClassEstoqueProduto
                objEstoqueProduto.dtDataInicial = gdtDataHoje
                objEstoqueProduto.iFilialEmpresa = giFilialEmpresa
                objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
                objEstoqueProduto.sProduto = objProdutoAux.sCodigo
                objEstoqueProduto.dQuantidadeInicial = dResiduo
                
                iAlmoxarifadoPadrao = MARCADO
                
                'Chama EstoqueInicial_Grava
                lErro = CF("EstoqueInicial_Grava_Trans", objEstoqueProduto, iAlmoxarifadoPadrao, colRastreamento)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
            End If
        
        End If
    
        Call objFrmAguardeTela.ProcessouItem
        If objFrmAguarde.iCancelar = MARCADO Then Exit For
    
    Next

    Set objFrmAguardeTela = Nothing
    If objFrmAguarde.iCancelar = MARCADO Then gError ERRO_SEM_MENSAGEM

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 208195
    
    If iAux <> 0 Then
        For Each objProdutoAux In colProdutos
            objProdutoAux.iNivel = iAux
        Next
    End If

    ProdutoGrade_Grava = SUCESSO

    Exit Function

Erro_ProdutoGrade_Grava:

    ProdutoGrade_Grava = gErr
   
    Select Case gErr

        Case 208188
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 208189
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NOMERED_MAIOR", gErr, objProdutoAux.sCodigo, iIndice, STRING_PRODUTO_NOME_REDUZIDO)

        Case 208190
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DESCRICAO_MAIOR", gErr, objProdutoAux.sCodigo, iIndice, STRING_PRODUTO_DESCRICAO_TELA)

        Case 208191
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FIGURA_MAIOR", gErr, objProdutoAux.sCodigo, iIndice, STRING_PRODUTO_FIGURA)

        Case 208192
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_CODIGO_JA_EXISTE", gErr, objProdutoBD.sCodigo, iIndice)

        Case 208193
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NOMERED_JA_EXISTE", gErr, objProdutoAux.sCodigo, iIndice, objProdutoBD.sCodigo)

        Case 208194
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_CATGRADE_JA_EXISTE", gErr, objProdutoAux.sCodigo, iIndice, objProdutoBD.sCodigo)

        Case 208195
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208196)

    End Select

    Call Transacao_Rollback
    
    If iAux <> 0 Then
        For Each objProdutoAux In colProdutos
            objProdutoAux.iNivel = iAux
        Next
    End If
    
    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If

    Exit Function

End Function

Function NFiscalEntrada_GravaCompl(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 12) As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_NFiscalEntrada_GravaCompl

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201463
    Next
    
    lErro = CF("ProjetoInfo_Grava", objNFiscal.objProjetoInfo, PRJ_CR_TIPO_NF, objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO Then gError 207776
    
    lErro = CF("NFiscal_Trata_MsgItem", objNFiscal)
    If lErro <> SUCESSO Then gError 207777
    
    lErro = CF("NFiscal_Valida_NFE", objNFiscal)
    If lErro <> SUCESSO Then gError 207778
    
    lErro = NFiscal_Grava_Estorno_Cred(objNFiscal)
    If lErro <> SUCESSO Then gError 207778
    
    If Not (objContabil Is Nothing) Then
        
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then gError 207892
        
        'se a nota for de importação e não estiver com a data de entrada preenchida ==> não processa a gravação das informações
        'caso contrário, processa
        If (objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMP And objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMPSE And objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMPPC) Or objNFiscal.dtDataEntrada <> DATA_NULA Then
    
            If objNFiscal.lCliente <> 0 Then
        
                'Grava os dados contábeis (contabilidade)
                lErro = objContabil.Contabil_Gravar_Registro(objNFiscal.lNumIntDoc, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEntrada, NAO_AVALIA_PELO_CUSTO_REAL_PRODUCAO, NAO_AVISA_LANCAMENTOS_CONTABILIZADOS, objNFiscal.lNumNotaFiscal)
                If lErro <> SUCESSO Then gError 207779
                
            Else
            
                'Grava os dados contábeis (contabilidade)
                lErro = objContabil.Contabil_Gravar_Registro(objNFiscal.lNumIntDoc, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada, NAO_AVALIA_PELO_CUSTO_REAL_PRODUCAO, NAO_AVISA_LANCAMENTOS_CONTABILIZADOS, objNFiscal.lNumNotaFiscal, IIf(objTipoDocInfo.sTipoDocCPR <> "", objNFiscal.dValorLivroAux, 0))
                If lErro <> SUCESSO Then gError 207780
        
            End If
        
        End If
        
    End If
    
    objNFiscal.objInfoAdic.lNumIntDoc = objNFiscal.lNumIntDoc
    objNFiscal.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_NF
    lErro = CF("InfoAdicionais_Grava", objNFiscal.objInfoAdic)
    If lErro <> SUCESSO Then gError 207781

    If objNFiscal.lNumIntDoc <> objNFiscal.lNumIntNFAnterior And objNFiscal.lNumIntNFAnterior <> 0 Then
    
        lErro = Comando_Executar(alComando(1), "UPDATE ImportNFeXml SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201490
        
        lErro = Comando_Executar(alComando(2), "UPDATE NFeFedDenegada SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201491
        
        lErro = Comando_Executar(alComando(3), "UPDATE NFeFedLote SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201492
        
        lErro = Comando_Executar(alComando(4), "UPDATE NFeFedLoteLog SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201493
        
        lErro = Comando_Executar(alComando(5), "UPDATE NFeFedProtNFe SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201494
        
        lErro = Comando_Executar(alComando(6), "UPDATE NFeFedRetCancNFe SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201495
        
        lErro = Comando_Executar(alComando(7), "UPDATE NFeFedRetEnvCCe SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201496
        
        lErro = Comando_Executar(alComando(8), "UPDATE NFeFedRetEnvEventoCanc SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201497
        
        lErro = Comando_Executar(alComando(9), "UPDATE RPSWEBLote SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201498
        
        lErro = Comando_Executar(alComando(10), "UPDATE RPSWEBLoteLog SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201499
        
        lErro = Comando_Executar(alComando(11), "UPDATE RPSWEBProt SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201500
        
        lErro = Comando_Executar(alComando(12), "UPDATE RPSWEBRetCanc SET NumIntNF = ? WHERE NumIntNF = ?", objNFiscal.lNumIntDoc, objNFiscal.lNumIntNFAnterior)
        If lErro <> AD_SQL_SUCESSO Then gError 201501
        
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscalEntrada_GravaCompl = SUCESSO

    Exit Function
    
Erro_NFiscalEntrada_GravaCompl:

    NFiscalEntrada_GravaCompl = gErr
    
    Select Case gErr
    
        Case 207776 To 207781, 207892
    
        Case 201463
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207782)
            
    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function NFiscalEntrada_AlteraDados(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Altera NFiscal Fatura

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NFiscalEntrada_AlteraDados


    'Inicia Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 207790
    
    lErro = CF("NFiscalEntrada_AlteraDados_EmTrans", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 207791
    
    'Faz o Commit da Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 207792
    
    NFiscalEntrada_AlteraDados = SUCESSO

    Exit Function
    
Erro_NFiscalEntrada_AlteraDados:

    NFiscalEntrada_AlteraDados = gErr
    
    Select Case gErr
    
        Case 207790
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 207791
        
        Case 207792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207793)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function NFiscalEntrada_AlteraDados_EmTrans(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
'Altera NFiscal Fatura

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim objNFiscalBD As New ClassNFiscal

On Error GoTo Erro_NFiscalEntrada_AlteraDados_EmTrans

    If objNFiscal.lNumIntDoc = 0 Then
        'NFiscal_Alteracao lê pelo NumIntDoc, logo só é necessário tentar ler sem essa informação
        'se o mesmo não tiver preenchido, coisa que não deve ocorrer já que NFiscalEntrada_Verifica_Existencia2
        'já preencheu a informação anteriormente ao verificar a existente da NF ou da NF com data próxima
        'Ou seja, esse código é só uma segurança a mais

'        'trata da alteração dos dados da nota fiscal.
'        lErro = CF("NFiscal_Le_NumNotaFiscal", objNFiscal)
'        If lErro <> SUCESSO And lErro <> 207677 Then gError 207794
'
'        If lErro <> SUCESSO Then gError 207795

        objNFiscalBD.lNumNotaFiscal = objNFiscal.lNumNotaFiscal
        objNFiscalBD.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objNFiscalBD.sSerie = objNFiscal.sSerie
        objNFiscalBD.lFornecedor = objNFiscal.lFornecedor
        objNFiscalBD.lCliente = objNFiscal.lCliente
        objNFiscalBD.iFilialForn = objNFiscal.iFilialForn
        objNFiscalBD.iFilialCli = objNFiscal.iFilialCli
        objNFiscalBD.dtDataEmissao = objNFiscal.dtDataEmissao
        objNFiscalBD.dtDataEntrada = objNFiscal.dtDataEntrada
        objNFiscalBD.iTipoDocInfo = objNFiscal.iTipoNFiscal

        'trata da alteração dos dados da nota fiscal.
        lErro = CF("NFiscal_Le_SemNumIntDoc", objNFiscalBD)
        If lErro <> SUCESSO And lErro <> 79714 Then gError 207794

        If lErro <> SUCESSO Then gError 207795
        
        objNFiscal.lNumIntDoc = objNFiscalBD.lNumIntDoc
        
    End If

    'trata da alteração dos dados da nota fiscal.
    lErro = CF("NFiscal_Alteracao", objNFiscal)
    If lErro <> SUCESSO Then gError 207796
    
    lErro = CF("NFiscalEntrada_GravaCompl", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 207797
    
    NFiscalEntrada_AlteraDados_EmTrans = SUCESSO

    Exit Function
    
Erro_NFiscalEntrada_AlteraDados_EmTrans:

    NFiscalEntrada_AlteraDados_EmTrans = gErr
    
    Select Case gErr
    
        Case 207794, 207796, 207797
        
        Case 207795
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207798)
            
    End Select
    
    Exit Function
    
End Function

Function Rotina_Reproc_Zera_SldDiaEstTerc(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iIndice As Integer
Dim sCampo(1 To 40) As String
Dim dtDataInicial As Date
Dim dtDataFinal As Date

On Error GoTo Erro_Rotina_Reproc_Zera_SldDiaEstTerc

    If gobjMAT.iTrataEstTercCliForn = MARCADO Then
    
        'Inicia o comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 90629
        
        sCampo(1) = "QuantEntConsig3"
        sCampo(2) = "ValorEntConsig3"
        sCampo(3) = "QuantSaiConsig3"
        sCampo(4) = "ValorSaiConsig3"
        sCampo(5) = "QuantEntDemo3"
        sCampo(6) = "ValorEntDemo3"
        sCampo(7) = "QuantSaiDemo3"
        sCampo(8) = "ValorSaiDemo3"
        sCampo(9) = "QuantEntConserto3"
        sCampo(10) = "ValorEntConserto3"
        sCampo(11) = "QuantSaiConserto3"
        sCampo(12) = "ValorSaiConserto3"
        sCampo(13) = "QuantEntOutros3"
        sCampo(14) = "ValorEntOutros3"
        sCampo(15) = "QuantSaiOutros3"
        sCampo(16) = "ValorSaiOutros3"
        sCampo(17) = "QuantEntBenef3"
        sCampo(18) = "ValorEntBenef3"
        sCampo(19) = "QuantSaiBenef3"
        sCampo(20) = "ValorSaiBenef3"
        sCampo(21) = "QuantEntConsig"
        sCampo(22) = "ValorEntConsig"
        sCampo(23) = "QuantSaiConsig"
        sCampo(24) = "ValorSaiConsig"
        sCampo(25) = "QuantEntDemo"
        sCampo(26) = "ValorEntDemo"
        sCampo(27) = "QuantSaiDemo"
        sCampo(28) = "ValorSaiDemo"
        sCampo(29) = "QuantEntConserto"
        sCampo(30) = "ValorEntConserto"
        sCampo(31) = "QuantSaiConserto"
        sCampo(32) = "ValorSaiConserto"
        sCampo(33) = "QuantEntOutros"
        sCampo(34) = "ValorEntOutros"
        sCampo(35) = "QuantSaiOutros"
        sCampo(36) = "ValorSaiOutros"
        sCampo(37) = "QuantEntBenef"
        sCampo(38) = "ValorEntBenef"
        sCampo(39) = "QuantSaiBenef"
        sCampo(40) = "ValorSaiBenef"
          
        sComandoSQL = "UPDATE SldDiaEstTerc SET "
        
        For iIndice = LBound(sCampo) To UBound(sCampo)
            If iIndice = UBound(sCampo) Then
                sComandoSQL = sComandoSQL & sCampo(iIndice) & "=0"
            Else
                sComandoSQL = sComandoSQL & sCampo(iIndice) & "=0, "
            End If
        
        Next
        
        If objReprocessamentoEst.iAno <> 0 Then
        
            dtDataInicial = CDate("01/" & CStr(objReprocessamentoEst.iMes) & "/" & CStr(objReprocessamentoEst.iAno))
            dtDataFinal = DateAdd("m", 1, dtDataInicial) - 1
        
            'Inclui o filtro de produto no select
            sComandoSQL = sComandoSQL & " WHERE Produto = ? AND Data >= ? AND Data <= ? "
            If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
            'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
            lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo, dtDataInicial, dtDataFinal)
            If lErro <> SUCESSO Then gError 90630
        
        'Se o produto foi preenchido
        ElseIf Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
        
            'Inclui o filtro de produto no select
            sComandoSQL = sComandoSQL & "WHERE Produto = ?"
            If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
            
            'Executa o comando passando o parâmetro produto
            lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
            If lErro <> SUCESSO Then gError 90630
        
        Else
        
            If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
            
            'Executa o comando sem o parâmetro produto
            lErro = Comando_Executar(lComando, sComandoSQL)
            If lErro <> SUCESSO Then gError 90630
        
        End If
        
        Call Comando_Fechar(lComando)
        
    End If
    
    Rotina_Reproc_Zera_SldDiaEstTerc = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_SldDiaEstTerc:
    
    Rotina_Reproc_Zera_SldDiaEstTerc = gErr
    
    Select Case gErr
    
        Case 90629
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90630
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDDIAEST", gErr)
        
        Case 90631
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149577)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Rotina_Reproc_Zera_SldMesEst1Terc(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iMes As Integer
Dim iIndice As Integer
Dim sMes As String
Dim sCampo(1 To 15) As String
Dim iMesInicial As Integer
Dim iMesFinal As Integer

On Error GoTo Erro_Rotina_Reproc_Zera_SldMesEst1Terc

    If gobjMAT.iTrataEstTercCliForn = MARCADO Then
    
        'Inicia o comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 90609
        
        sCampo(1) = "SaldoQuantConsig3"
        sCampo(2) = "SaldoValorConsig3"
        sCampo(3) = "SaldoQuantDemo3"
        sCampo(4) = "SaldoValorDemo3"
        sCampo(5) = "SaldoQuantConserto3"
        sCampo(6) = "SaldoValorConserto3"
        sCampo(7) = "SaldoQuantOutros3"
        sCampo(8) = "SaldoValorOutros3"
        sCampo(9) = "SaldoQuantBenef3"
        sCampo(10) = "SaldoValorBenef3"
        sCampo(11) = "CustoMedioConsig3"
        sCampo(12) = "CustoMedioConserto3"
        sCampo(13) = "CustoMedioDemo3"
        sCampo(14) = "CustoMedioOutros3"
        sCampo(15) = "CustoMedioBenef3"
        
        sComandoSQL = "UPDATE SldMesEst1Terc SET "
        
        If objReprocessamentoEst.iMes <> 0 Then
            iMesInicial = objReprocessamentoEst.iMes
            iMesFinal = objReprocessamentoEst.iMes
        Else
            iMesInicial = 1
            iMesFinal = 12
        End If
        
        For iIndice = LBound(sCampo) To UBound(sCampo)
            For iMes = iMesInicial To iMesFinal
                sMes = CStr(iMes)
                If iIndice = UBound(sCampo) And iMes = iMesFinal Then
                    sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0"
                Else
                    sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0, "
                End If
        
            Next
        Next
        
        If objReprocessamentoEst.iAno <> 0 Then
        
            'Inclui o filtro de produto no select
            sComandoSQL = sComandoSQL & " WHERE Produto = ? AND Ano = ?"
            If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
            'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
            lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo, objReprocessamentoEst.iAno)
            If lErro <> SUCESSO Then gError 90610
        
        'Se o produto foi preenchido
        ElseIf Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
        
            'Inclui o filtro de produto no select
            sComandoSQL = sComandoSQL & "WHERE Produto = ?"
            If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
            
            'Executa o comando passando o parâmetro de produto
            lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
            If lErro <> SUCESSO Then Error 90610
            
        Else
        
            If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
            
            'Executa o comando sem o parâmetro de produto
            lErro = Comando_Executar(lComando, sComandoSQL)
            If lErro <> SUCESSO Then Error 90610
        
        End If
        
        'Fecha o comando
        Call Comando_Fechar(lComando)
        
    End If
        
    Rotina_Reproc_Zera_SldMesEst1Terc = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_SldMesEst1Terc:
    
    Rotina_Reproc_Zera_SldMesEst1Terc = gErr
    
    Select Case gErr
    
        Case 90609
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90610
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDMESEST1", gErr)
        
        Case 90611
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149572)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Rotina_Reproc_Zera_SldMesEst2Terc(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim iMes As Integer
Dim iIndice As Integer
Dim sMes As String
Dim sCampo(1 To 15) As String
Dim iMesInicial As Integer
Dim iMesFinal As Integer

On Error GoTo Erro_Rotina_Reproc_Zera_SldMesEst2Terc

    If gobjMAT.iTrataEstTercCliForn = MARCADO Then
    
        'Inicia o comando
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 90613
        
        sCampo(1) = "SaldoQuantConsig"
        sCampo(2) = "SaldoValorConsig"
        sCampo(3) = "SaldoQuantDemo"
        sCampo(4) = "SaldoValorDemo"
        sCampo(5) = "SaldoQuantConserto"
        sCampo(6) = "SaldoValorConserto"
        sCampo(7) = "SaldoQuantOutros"
        sCampo(8) = "SaldoValorOutros"
        sCampo(9) = "SaldoQuantBenef"
        sCampo(10) = "SaldoValorBenef"
        sCampo(11) = "CustoMedioConsig"
        sCampo(12) = "CustoMedioConserto"
        sCampo(13) = "CustoMedioDemo"
        sCampo(14) = "CustoMedioOutros"
        sCampo(15) = "CustoMedioBenef"
        
        sComandoSQL = "UPDATE SldMesEst2Terc SET "
        
        If objReprocessamentoEst.iMes <> 0 Then
            iMesInicial = objReprocessamentoEst.iMes
            iMesFinal = objReprocessamentoEst.iMes
        Else
            iMesInicial = 1
            iMesFinal = 12
        End If
        
        For iIndice = LBound(sCampo) To UBound(sCampo)
            For iMes = iMesInicial To iMesFinal
                sMes = CStr(iMes)
                If iIndice = UBound(sCampo) And iMes = iMesFinal Then
                    sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0"
                Else
                    sComandoSQL = sComandoSQL & sCampo(iIndice) & sMes & "=0, "
                End If
        
            Next
        Next
        
        If objReprocessamentoEst.iAno <> 0 Then
        
            'Inclui o filtro de produto no select
            sComandoSQL = sComandoSQL & " WHERE Produto = ? AND Ano = ?"
            If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
            'Executa o comando que irá zerar todos os campos da tabela, passando o parâmetro produto
            lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo, objReprocessamentoEst.iAno)
            If lErro <> SUCESSO Then gError 90614
    
        'Se o produto foi preenchido
        ElseIf Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
        
            'Inclui o filtro de produto no select
            sComandoSQL = sComandoSQL & "WHERE Produto = ?"
            If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
            'Executa o comando, passando o parâmetro de produto
            lErro = Comando_Executar(lComando, sComandoSQL, objReprocessamentoEst.sProdutoCodigo)
            If lErro <> SUCESSO Then gError 90614
        
        Else
            
            If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
            
            'Executa o comando, sem o parâmetro de produto
            lErro = Comando_Executar(lComando, sComandoSQL)
            If lErro <> SUCESSO Then gError 90614
        
        End If
        
        Call Comando_Fechar(lComando)
        
    End If
    
    Rotina_Reproc_Zera_SldMesEst2Terc = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_SldMesEst2Terc:
    
    Rotina_Reproc_Zera_SldMesEst2Terc = gErr
    
    Select Case gErr
    
        Case 90613
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90614
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDMESEST2", gErr)
        
        Case 90615
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149573)
    
    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function InvCliForn_Grava(ByVal objInvCliForn As ClassInvCliForn) As Long
'Inventário de estoque em/de terceiros (cliente ou fornecedor)
'Não movimenta estoque
Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_InvCliForn_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 209586

    lErro = CF("InvCliForn_Grava_EmTrans", objInvCliForn)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 209587

    InvCliForn_Grava = SUCESSO

    Exit Function

Erro_InvCliForn_Grava:

    InvCliForn_Grava = gErr

    Select Case gErr

        Case 209586
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 209587
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209588)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function InvCliForn_Grava_EmTrans(ByVal objInvCliForn As ClassInvCliForn) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_InvCliForn_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209589
    Next

    'Le a tabela InvCliForn
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM InvCliForn WHERE FilialEmpresa = ? AND Data= ?  AND Escaninho= ?  AND TipoCliForn= ?  AND CliForn= ?  AND Filial= ? ", _
                0, lNumIntDoc, objInvCliForn.iFilialEmpresa, objInvCliForn.dtData, objInvCliForn.iEscaninho, objInvCliForn.iTipoCliForn, objInvCliForn.lCliForn, objInvCliForn.iFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 209590

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209591

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 209592

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE InvCliForn SET Usuario= ? , DataGravacao= ? , HoraGravacao= ? , Obs= ? ", alComando(0), _
                gsUsuario, Date, CDbl(Time), objInvCliForn.sObs)
        If lErro <> AD_SQL_SUCESSO Then gError 209593
        
        objInvCliForn.lNumIntDoc = lNumIntDoc
        
        'Excluir os itens acertando os saldos em/de terceiros
        lErro = InvCliFornItens_Exclui(objInvCliForn)
        If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    Else

        'Obter NumIntDoc
        lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_INT_PROX_INVCLIFORN", lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objInvCliForn.lNumIntDoc = lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO InvCliForn( NumIntDoc, FilialEmpresa, Data, Escaninho, TipoCliForn, " & _
                "CliForn, Filial, Usuario, DataGravacao, HoraGravacao, Obs) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                objInvCliForn.lNumIntDoc, objInvCliForn.iFilialEmpresa, objInvCliForn.dtData, objInvCliForn.iEscaninho, objInvCliForn.iTipoCliForn, _
                objInvCliForn.lCliForn, objInvCliForn.iFilial, gsUsuario, Date, CDbl(Time), _
                objInvCliForn.sObs)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209594

    End If

    'Relança os itens
    lErro = InvCliFornItens_Grava(objInvCliForn)
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InvCliForn_Grava_EmTrans = SUCESSO

    Exit Function

Erro_InvCliForn_Grava_EmTrans:

    InvCliForn_Grava_EmTrans = gErr

    Select Case gErr

        Case 209589
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209590, 209591
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVCLIFORN", gErr)

        Case 209592
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_INVCLIFORN", gErr)

        Case 209593
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_INVCLIFORN", gErr)

        Case 209594
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INVCLIFORN", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209595)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function InvCliForn_Atualiza_SldDiaEst(ByVal objTipoMovEstoque As ClassTipoMovEst, ByVal objSldDiaEst As ClassSldDiaEst, ByVal iEscaninho As Integer, ByVal dQtd As Double, ByVal iExclusao As Integer) As Long
'Preenche os objs ClassTipoMovEst e ClassSldDiaEst para atualização das tabelas de saldo
Dim iEntrada As Integer
Dim iTipo As Integer

On Error GoTo Erro_InvCliForn_Atualiza_SldDiaEst

    If iExclusao = MARCADO Then
        dQtd = -dQtd
    End If
    If dQtd > 0 Then
        iEntrada = MARCADO
        iTipo = 1 'Adiciona
    Else
        iEntrada = DESMARCADO
        iTipo = 2 'Subtrai
    End If
    dQtd = Abs(dQtd)
    
    Select Case iEscaninho
        Case 2 '2   Conserto - Nosso em Poder de Terceiros
            objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO
            objTipoMovEstoque.iAtualizaConserto = iTipo
            If iEntrada = MARCADO Then
                objSldDiaEst.dQuantEntConserto = dQtd
            Else
                objSldDiaEst.dQuantSaiConserto = dQtd
            End If
    
        Case 3 '3   Consignação - Nosso em Poder de Terceiros
            objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO
            objTipoMovEstoque.iAtualizaConsig = iTipo
            If iEntrada = MARCADO Then
                objSldDiaEst.dQuantEntConsig = dQtd
            Else
                objSldDiaEst.dQuantSaiConsig = dQtd
            End If
    
        Case 4 '4   Demonstração - Nosso em Poder de Terceiros
            objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO
            objTipoMovEstoque.iAtualizaDemo = iTipo
            If iEntrada = MARCADO Then
                objSldDiaEst.dQuantEntDemo = dQtd
            Else
                objSldDiaEst.dQuantSaiDemo = dQtd
            End If
    
        Case 5 '5   Outros - Nosso em Poder de Terceiros
            objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO
            objTipoMovEstoque.iAtualizaOutras = iTipo
            If iEntrada = MARCADO Then
                objSldDiaEst.dQuantEntOutros = dQtd
            Else
                objSldDiaEst.dQuantSaiOutros = dQtd
            End If
    
        Case 6 '6   Beneficiamento - Nosso em Poder de Terceiros
            objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO
            objTipoMovEstoque.iAtualizaBenef = iTipo
            If iEntrada = MARCADO Then
                objSldDiaEst.dQuantEntBenef = dQtd
            Else
                objSldDiaEst.dQuantSaiBenef = dQtd
            End If
        
        Case 7 '7   Conserto - De Terceiros em Nosso Poder
            objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS
            objTipoMovEstoque.iAtualizaConserto = iTipo
            If iEntrada = MARCADO Then
                objSldDiaEst.dQuantEntConserto3 = dQtd
            Else
                objSldDiaEst.dQuantSaiConserto3 = dQtd
            End If
    
        Case 8 '8   Consignação - De Terceiros em Nosso Poder
            objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS
            objTipoMovEstoque.iAtualizaConsig = iTipo
            If iEntrada = MARCADO Then
                objSldDiaEst.dQuantEntConsig3 = dQtd
            Else
                objSldDiaEst.dQuantSaiConsig3 = dQtd
            End If
    
        Case 9 '9   Demonstração - De Terceiros em Nosso Poder
            objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS
            objTipoMovEstoque.iAtualizaDemo = iTipo
            If iEntrada = MARCADO Then
                objSldDiaEst.dQuantEntDemo3 = dQtd
            Else
                objSldDiaEst.dQuantSaiDemo3 = dQtd
            End If
    
        Case 10 '10  Outros - De Terceiros em Nosso Poder
            objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS
            objTipoMovEstoque.iAtualizaOutras = iTipo
            If iEntrada = MARCADO Then
                objSldDiaEst.dQuantEntOutros3 = dQtd
            Else
                objSldDiaEst.dQuantSaiOutros3 = dQtd
            End If
    
        Case 11 '11  Beneficiamento - De Terceiros em Nosso Poder
            objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTODETERCEIROS
            objTipoMovEstoque.iAtualizaBenef = iTipo
            If iEntrada = MARCADO Then
                objSldDiaEst.dQuantEntBenef3 = dQtd
            Else
                objSldDiaEst.dQuantSaiBenef3 = dQtd
            End If
    
    End Select

    InvCliForn_Atualiza_SldDiaEst = SUCESSO

    Exit Function

Erro_InvCliForn_Atualiza_SldDiaEst:

    InvCliForn_Atualiza_SldDiaEst = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209596)

    End Select

    Exit Function
        
End Function

Private Function InvCliFornItens_Exclui(ByVal objInv As ClassInvCliForn) As Long

Dim lErro As Long
Dim alComando(0 To 9) As Long
Dim iIndice As Integer
Dim objInvItem As ClassInvCliFornItens
Dim sProduto As String, dQtdAcerto As Double

On Error GoTo Erro_InvCliFornItens_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209597
    Next
    
    sProduto = String(STRING_PRODUTO, 0)

    'Le a tabela InvCliFornItens
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Produto, QtdAcerto FROM InvCliFornItens WHERE NumIntInvCliForn= ? ", _
                0, sProduto, dQtdAcerto, objInv.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209598

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209599

    'Enquanto existir
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objInvItem = New ClassInvCliFornItens
        
        objInvItem.sProduto = sProduto
        objInvItem.dQtdAcerto = dQtdAcerto
    
        'Acerta as tabelas de saldo
        lErro = InvCliForn_Atualiza_Item(alComando(2), alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), alComando(8), alComando(9), objInv, objInvItem, MARCADO)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
           
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM InvCliFornItens", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 209600
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209601
        
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InvCliFornItens_Exclui = SUCESSO

    Exit Function

Erro_InvCliFornItens_Exclui:

    InvCliFornItens_Exclui = gErr

    Select Case gErr

        Case 209597
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209598, 209599, 209601
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVCLIFORNITENS", gErr)

        Case 209600
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_INVCLIFORNITENS", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209601)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function InvCliFornItens_Grava(ByVal objInv As ClassInvCliForn) As Long

Dim lErro As Long
Dim alComando(1 To 9) As Long
Dim iIndice As Integer
Dim objInvItem As ClassInvCliFornItens

On Error GoTo Erro_InvCliFornItens_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209602
    Next

    For Each objInvItem In objInv.colItens
    
        objInvItem.lNumIntInvCliForn = objInv.lNumIntDoc

        lErro = Comando_Executar(alComando(1), "INSERT INTO InvCliFornItens( NumIntInvCliForn, Seq, Produto, QtdCliData, " & _
                "QtdEncontCliData, QtdData, QtdAcerto,Obs) VALUES (?,?,?,?,?,?,?,?)", _
                objInvItem.lNumIntInvCliForn, objInvItem.iSeq, objInvItem.sProduto, objInvItem.dQtdCliData, _
                objInvItem.dQtdEncontCliData, objInvItem.dQtdData, objInvItem.dQtdAcerto, objInvItem.sObs)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209603
        
        lErro = InvCliForn_Atualiza_Item(alComando(2), alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), alComando(8), alComando(9), objInv, objInvItem)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InvCliFornItens_Grava = SUCESSO

    Exit Function

Erro_InvCliFornItens_Grava:

    InvCliFornItens_Grava = gErr

    Select Case gErr

        Case 209602
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209603
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INVCLIFORNITENS", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209604)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function InvCliForn_Exclui(ByVal objInvCliForn As ClassInvCliForn) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_InvCliForn_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 209605

    lErro = CF("InvCliForn_Exclui_EmTrans", objInvCliForn)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 209606

    InvCliForn_Exclui = SUCESSO

    Exit Function

Erro_InvCliForn_Exclui:

    InvCliForn_Exclui = gErr

    Select Case gErr

        Case 209605
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 209606
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209607)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function InvCliForn_Exclui_EmTrans(ByVal objInvCliForn As ClassInvCliForn) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_InvCliForn_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209608
    Next

    'Le a tabelaInvCliForn
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM InvCliForn WHERE FilialEmpresa = ? AND Data= ? AND Escaninho= ?  AND TipoCliForn= ?  AND CliForn= ?  AND Filial= ? ", _
                0, lNumIntDoc, objInvCliForn.iFilialEmpresa, objInvCliForn.dtData, objInvCliForn.iEscaninho, objInvCliForn.iTipoCliForn, objInvCliForn.lCliForn, objInvCliForn.iFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 209609

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209610

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 209611

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 209612
    
    objInvCliForn.lNumIntDoc = lNumIntDoc
    
    lErro = InvCliFornItens_Exclui(objInvCliForn)
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM InvCliForn", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 209613

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InvCliForn_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_InvCliForn_Exclui_EmTrans:

    InvCliForn_Exclui_EmTrans = gErr

    Select Case gErr

        Case 209608
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209609, 209610
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVCLIFORN", gErr)

        Case 209611
            Call Rotina_Erro(vbOKOnly, "ERRO_INVCLIFORN_NAO_CADASTRADO", gErr, objInvCliForn.dtData, objInvCliForn.iEscaninho, objInvCliForn.iTipoCliForn, objInvCliForn.lCliForn, objInvCliForn.iFilial)

        Case 209612
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_INVCLIFORN", gErr)

        Case 209613
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_INVCLIFORN", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209614)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function InvCliForn_Atualiza_Item(ByVal lCmd1 As Long, ByVal lCmd2 As Long, ByVal lCmd3 As Long, ByVal lCmd4 As Long, ByVal lCmd5 As Long, ByVal lCmd6 As Long, ByVal lCmd7 As Long, ByVal lCmd8 As Long, ByVal objInv As ClassInvCliForn, ByVal objInvItem As ClassInvCliFornItens, Optional ByVal iExclusao As Integer = DESMARCADO)

Dim lErro As Long
Dim objSldDiaEst As New ClassSldDiaEst, objTipoMovEstoque As New ClassTipoMovEst
Dim lCli As Long, iFilCli As Integer, lForn As Long, iFilForn As Integer
Dim objSldDiaEstAux As New ClassSldDiaEst
Dim objTipoMovEstoqueAux As New ClassTipoMovEst

On Error GoTo Erro_InvCliForn_Atualiza_Item

    objSldDiaEst.iFilialEmpresa = objInv.iFilialEmpresa
    objSldDiaEst.dtData = objInv.dtData
    objSldDiaEst.sProduto = objInvItem.sProduto
    
    objSldDiaEstAux.iFilialEmpresa = objInv.iFilialEmpresa
    objSldDiaEstAux.dtData = objInv.dtData
    objSldDiaEstAux.sProduto = objInvItem.sProduto
    
    If objInv.iTipoCliForn = TIPO_TERC_CLIENTE Then
        lCli = objInv.lCliForn
        iFilCli = objInv.iFilial
        lForn = 0
        iFilForn = 0
    Else
        lCli = 0
        iFilCli = 0
        lForn = objInv.lCliForn
        iFilForn = objInv.iFilial
    End If
    
    '===>Coloca as informações por terceiro
    
    'Preenche objTipoMovEstoque e objSldDiaEst
    lErro = InvCliForn_Atualiza_SldDiaEst(objTipoMovEstoque, objSldDiaEst, objInv.iEscaninho, objInvItem.dQtdAcerto, iExclusao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'atualiza a tabela de saldos diários de movimentação
    lErro = CF("SldDiaEstTerc_Atualiza", lCmd1, lCmd2, objSldDiaEst, lCli, iFilCli, lForn, iFilForn)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
     
    'atualiza a tabela de saldos mensais de movimentação 1
    lErro = CF("SldMesEst1Terc_Atualiza", lCmd3, lCmd4, lCmd5, objSldDiaEst, objTipoMovEstoque, lCli, iFilCli, lForn, iFilForn, 0, 0, 0, 0, 0)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
     
    'atualiza a tabela de saldos mensais de movimentação 2
    lErro = CF("SldMesEst2Terc_Atualiza", lCmd6, lCmd7, lCmd8, objSldDiaEst, objTipoMovEstoque, lCli, iFilCli, lForn, iFilForn, 0, 0, 0, 0, 0)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    
    '===>Retira de não identificado
    
    'Preenche objTipoMovEstoque e objSldDiaEst
    lErro = InvCliForn_Atualiza_SldDiaEst(objTipoMovEstoqueAux, objSldDiaEstAux, objInv.iEscaninho, -objInvItem.dQtdAcerto, iExclusao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'atualiza a tabela de saldos diários de movimentação
    lErro = CF("SldDiaEstTerc_Atualiza", lCmd1, lCmd2, objSldDiaEstAux, 0, 0, 0, 0)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
     
    'atualiza a tabela de saldos mensais de movimentação 1
    lErro = CF("SldMesEst1Terc_Atualiza", lCmd3, lCmd4, lCmd5, objSldDiaEstAux, objTipoMovEstoqueAux, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
     
    'atualiza a tabela de saldos mensais de movimentação 2
    lErro = CF("SldMesEst2Terc_Atualiza", lCmd6, lCmd7, lCmd8, objSldDiaEstAux, objTipoMovEstoqueAux, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    InvCliForn_Atualiza_Item = SUCESSO

    Exit Function

Erro_InvCliForn_Atualiza_Item:

    InvCliForn_Atualiza_Item = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209615)

    End Select

    Exit Function
    
End Function

Public Function Reprocessamento_Ajuste_InvCliForn(ByVal sProd As String) As Long

Dim lErro As Long
Dim alComando(1 To 9) As Long
Dim iIndice As Integer
Dim objInv As ClassInvCliForn, tInv As typeInvCliForn
Dim objInvItem As ClassInvCliFornItens, tInvItem As typeInvCliFornItens

On Error GoTo Erro_Reprocessamento_Ajuste_InvCliForn

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209616
    Next
    
    tInvItem.sProduto = String(STRING_PRODUTO, 0)

    'Le a tabelaInvCliForn
    lErro = Comando_Executar(alComando(1), "SELECT FilialEmpresa, Data, Escaninho, TipoCliForn, CliForn, Filial, Produto, QtdAcerto  FROM InvCliForn AS I, InvCliFornItens AS P WHERE I.NumIntDoc = P.NumIntInvCliForn AND (Produto = ? OR ? = ? ) ", _
                tInv.iFilialEmpresa, tInv.dtData, tInv.iEscaninho, tInv.iTipoCliForn, tInv.lCliForn, tInv.iFilial, tInvItem.sProduto, tInvItem.dQtdAcerto, sProd, sProd, "")
    If lErro <> AD_SQL_SUCESSO Then gError 209617

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209618

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objInv = New ClassInvCliForn
        Set objInvItem = New ClassInvCliFornItens
        
        objInv.dtData = tInv.dtData
        objInv.iEscaninho = tInv.iEscaninho
        objInv.iTipoCliForn = tInv.iTipoCliForn
        objInv.lCliForn = tInv.lCliForn
        objInv.iFilial = tInv.iFilial
        objInvItem.sProduto = tInvItem.sProduto
        objInvItem.dQtdAcerto = tInvItem.dQtdAcerto
        objInv.iFilialEmpresa = tInv.iFilialEmpresa

        lErro = InvCliForn_Atualiza_Item(alComando(2), alComando(3), alComando(4), alComando(5), alComando(6), alComando(7), alComando(8), alComando(9), objInv, objInvItem)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209619
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Reprocessamento_Ajuste_InvCliForn = SUCESSO

    Exit Function

Erro_Reprocessamento_Ajuste_InvCliForn:

    Reprocessamento_Ajuste_InvCliForn = gErr

    Select Case gErr

        Case 209616
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209617 To 209619
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVCLIFORN", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209620)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function NFiscal_Grava_Estorno_Cred(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer, objTipoDocInfo As New ClassTipoDocInfo
Dim objItem As ClassItemNF, dValEstTrib As Double, dValEstTribAlm As Double
Dim dCredMax As Double, sItemCat As String, colRatreamentoMovto As New Collection
Dim objMovEstoque As New ClassMovEstoque, lCodigo As Long
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objItemNFAloc As ClassItemNFAlocacao, dQtdeTotal As Double
Dim objProd As ClassProduto, lNumIntDoc As Long

On Error GoTo Erro_NFiscal_Grava_Estorno_Cred

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209785
    Next
    
    lErro = NFiscal_Exclui_Estorno_Cred(objNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objTipoDocInfo.iCodigo = objNF.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
    
    For Each objItem In objNF.colItens
    
        If objItem.objTributacao.dICMSCredito > 0 Then
        
            Set objProd = New ClassProduto
            
            objProd.sCodigo = objItem.sProduto
        
            lErro = CF("Produto_Le", objProd)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
            dCredMax = 0
            dValEstTrib = 0
            sItemCat = String(STRING_MAXIMO, 0)

            'Le a categoria do produto para pegar o crédito máximo
            lErro = Comando_Executar(alComando(1), "SELECT Item FROM ProdutoCategoria WHERE Produto = ? AND Categoria = ?", sItemCat, objItem.sProduto, "Estorno Tributário")
            If lErro <> AD_SQL_SUCESSO Then gError 209786
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209787
            
            'Verifica se existe a categoria que controla o máximo de crédito
            If lErro = AD_SQL_SUCESSO Then
                If left(UCase(sItemCat), 9) = "CRÉD.MAX." Then
                    dCredMax = StrParaDbl(Replace(Replace(sItemCat, "Créd.Max.", ""), "%", "")) / 100
                    'Se a alíquota aplicada é maior que a máxima -> calcula o valor de crédito a ser estornado
                    If objItem.objTributacao.dICMSAliquota > dCredMax Then
                        dValEstTrib = objItem.objTributacao.dICMSCredito - (objItem.objTributacao.dICMSBase * dCredMax)
                    End If
                End If
            End If
            
            'Se vai ter estorno de crédito
            If dValEstTrib > DELTA_VALORMONETARIO Then
            
                'Grava na tabela para poder exibir em relatórios de apuração de crédito
                lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_PROX_EST_CRED_TRIB_ICMS", lNumIntDoc, 1, EMPRESA_TODA)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                lErro = Comando_Executar(alComando(2), "INSERT INTO EstornoCredTribICMS (NumIntDoc, NumIntDocNF, NumIntDocItemNF, AliquotaMax, EstornoCred) VALUES (?,?,?,?,?)", lNumIntDoc, objNF.lNumIntDoc, objItem.lNumIntDoc, dCredMax, dValEstTrib)
                If lErro <> AD_SQL_SUCESSO Then gError 209788
                
                'Se controla o estoque gera o movimento para acerto do custo
                If objProd.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
            
                    'Preenche as informações básicas do movimento
                    If lCodigo = 0 Then
                    
                        lErro = CF("MovEstoque_Automatico_EmTransacao", objNF.iFilialEmpresa, lCodigo)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                        objMovEstoque.dtData = objNF.dtDataEntrada
                        objMovEstoque.dtHora = objNF.dtHoraEntrada
                        objMovEstoque.iFilialEmpresa = objNF.iFilialEmpresa
                        objMovEstoque.iTipoMov = MOV_EST_ESTORNO_CRED_TRIBUTARIO
                        objMovEstoque.lCodigo = lCodigo
                        objMovEstoque.lCliente = objNF.lCliente
                        objMovEstoque.lFornecedor = objNF.lFornecedor
                        objMovEstoque.iFilialCli = objNF.iFilialCli
                        objMovEstoque.iFilialForn = objNF.iFilialForn
                        objMovEstoque.sDocOrigem = objTipoDocInfo.sSigla & " " & objNF.sSerie & " " & CStr(objNF.lNumNotaFiscal)
                        
                    End If
                    
                    dQtdeTotal = 0
                    For Each objItemNFAloc In objItem.ColAlocacoes
                        dQtdeTotal = dQtdeTotal + objItemNFAloc.dQuantidade
                    Next
                    
                    'Para cada movimento de material gera o movimento de acerto do custo
                    For Each objItemNFAloc In objItem.ColAlocacoes
                    
                        'Calcula o valor de forma proporcional ao movimentado no almoxarifado
                        dValEstTribAlm = Arredonda_Moeda((dValEstTrib / dQtdeTotal) * objItemNFAloc.dQuantidade, 4)
                    
                        Set objItemMovEstoque = objMovEstoque.colItens.Add(0, MOV_EST_ESTORNO_CRED_TRIBUTARIO, dValEstTribAlm, 0, objItem.sProduto, "", objItem.sUMEstoque, 0, objItemNFAloc.iAlmoxarifado, "", objItem.lNumIntDoc, "", 0, "", "", "", "", 0, colRatreamentoMovto, Nothing, DATA_NULA)
                        objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ESTORNO_CRED_TRIB_ICMS
                        objItemMovEstoque.lNumIntDocOrigem = lNumIntDoc
                        objItemMovEstoque.iClasseUM = objItem.iClasseUM
                        objItemMovEstoque.iControleEstoque = objItem.iControleEstoque
                        objItemMovEstoque.sSiglaUMEst = objItem.sUMEstoque
                        objItemMovEstoque.iApropriacaoProd = objItem.iApropriacaoProd
                        
                    Next
                    
                End If
            
            End If
        
        End If
    Next
    
    'Se gerou alguma movimentação, grava no BD
    If objMovEstoque.colItens.Count > 0 Then
        lErro = CF("MovEstoque_Interno_Grava0", objMovEstoque, Nothing)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscal_Grava_Estorno_Cred = SUCESSO

    Exit Function

Erro_NFiscal_Grava_Estorno_Cred:

    NFiscal_Grava_Estorno_Cred = gErr

    Select Case gErr

        Case 209785
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209786 To 209787
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOCATEGORIA", gErr)
            
        Case 209788
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERT_ESTORNOCREDTRIBICMS", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209789)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function NFiscal_Exclui_Estorno_Cred(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer, lNumIntDoc As Long
Dim lCodigoMovEst As Long, objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_NFiscal_Exclui_Estorno_Cred

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209790
    Next
    
    'Le a categoria do produto para pegar o crédito máximo
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM EstornoCredTribICMS WHERE NumIntDocNF = ?", 0, lNumIntDoc, objNF.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209791

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209792
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        If objMovEstoque.lCodigo = 0 Then
    
            'Verifica se há movimento de estoque ligado ao item
            lErro = Comando_Executar(alComando(3), "SELECT Codigo FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? ", lCodigoMovEst, MOVEST_TIPONUMINTDOCORIGEM_ESTORNO_CRED_TRIB_ICMS, lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 209793
        
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209794
            
            If lErro = AD_SQL_SUCESSO Then
        
                'Lê o movto de estoque da NOta fiscal
                objMovEstoque.lCodigo = lCodigoMovEst
                objMovEstoque.iFilialEmpresa = objNF.iFilialEmpresa
                
                'Lê os movimento dos itens de nota fiscal
                lErro = CF("MovEstoqueItens_Le_Desc1", objMovEstoque)
                If lErro <> SUCESSO And lErro <> 89050 Then gError ERRO_SEM_MENSAGEM
            
                If lErro <> SUCESSO Then gError 209795
            
                'Redefine os movimento para fazer o cancelamento
                objMovEstoque.iOperacao = EXCLUSAO_MOVIMENTO
                
                'Grava o movimento de estoque
                lErro = CF("MovEstoque_Grava", objMovEstoque)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            End If
    
        End If
    
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM EstornoCredTribICMS ", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 209796
    
        'Busca Primeiro
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209797
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscal_Exclui_Estorno_Cred = SUCESSO

    Exit Function

Erro_NFiscal_Exclui_Estorno_Cred:

    NFiscal_Exclui_Estorno_Cred = gErr

    Select Case gErr

        Case 209790
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209791, 209792, 209797
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTORNOCREDTRIBICMS", gErr)
            
        Case 209793, 209794
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
            
        Case 209795
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO2", gErr, objNF.iFilialEmpresa, lCodigoMovEst)
            
        Case 209796
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MOVIMENTOESTOQUE", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209798)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function NotaFiscalEntrada_Grava_Trata_Exclusao(ByVal objNF As ClassNFiscal) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim lErro As Long
Dim iAceitaEstoqueNegativo As Integer
Dim bMudouEstoque As Boolean
Dim objNFBD As New ClassNFiscal
Dim objItemNF As ClassItemNF
Dim objItemNFBD As ClassItemNF
Dim objItemAloc As ClassItemNFAlocacao
Dim objItemAlocBD As ClassItemNFAlocacao
Dim colAlocBD As New Collection
Dim colAloc As New Collection
Dim bAchou As Boolean, iIndice As Integer, iIndiceBD As Integer

On Error GoTo Erro_NotaFiscalEntrada_Grava_Trata_Exclusao

    iAceitaEstoqueNegativo = gobjMAT.iAceitaEstoqueNegativo
        
    objNFBD.lNumIntDoc = objNF.lNumIntDoc
    Set objNFBD.objContabil = objNF.objContabil

    lErro = CF("NFiscal_Le", objNFBD)
    If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM

    'Lê os itens da nota fiscal
    lErro = CF("NFiscalItens_Le", objNFBD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("AlocacoesNF_Le", objNFBD)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    bMudouEstoque = False
    
    For Each objItemNFBD In objNFBD.colItens
        For Each objItemAlocBD In objItemNFBD.ColAlocacoes
            objItemAlocBD.sProduto = objItemNFBD.sProduto
            colAlocBD.Add objItemAlocBD
        Next
        If objItemNFBD.colItensRomaneioGrade.Count <> 0 Then bMudouEstoque = True ' Por enquanto não testa a parte de grade, não deixa ficar com estoque negativo
    Next
    
    For Each objItemNF In objNF.colItens
        For Each objItemAloc In objItemNF.ColAlocacoes
            colAloc.Add objItemAloc
        Next
        If objItemNF.colItensRomaneioGrade.Count <> 0 Then bMudouEstoque = True  ' Por enquanto não testa a parte de grade, não deixa ficar com estoque negativo
    Next
    
    If colAloc.Count <> colAlocBD.Count Then bMudouEstoque = True
    
    'Depois esse teste pode ser melhorado para também aceitar estoque negativo se entrou
    'mais produto no almoxarifado do que antes
    'Também tem que mexer para tratar grade
    If Not bMudouEstoque Then
        'Testa para ver se as alocações são iguais e por conta disso poderá aceitar um estoque negativo de
        'maneira temporária
        For iIndice = colAloc.Count To 1 Step -1
            Set objItemAloc = colAloc(iIndice)
            bAchou = False
            For iIndiceBD = colAlocBD.Count To 1 Step -1
                Set objItemAlocBD = colAlocBD(iIndiceBD)
                If objItemAloc.iAlmoxarifado = objItemAlocBD.iAlmoxarifado And _
                    Abs(objItemAloc.dQuantidade - objItemAlocBD.dQuantidade) < QTDE_ESTOQUE_DELTA And _
                    UCase(objItemAloc.sProduto) = UCase(objItemAlocBD.sProduto) And _
                    UCase(objItemAloc.sUnidadeMed) = UCase(objItemAlocBD.sUnidadeMed) Then
                    bAchou = True
                    colAlocBD.Remove (iIndiceBD)
                    Exit For
                End If
            Next
            If Not bAchou Then
                bMudouEstoque = True
                Exit For
            End If
        Next
    End If
    
    'Desfaz a coleção de alocações porque a função de exclusão já lê
    'Se mantivesse ficaria duplicado
    For Each objItemNFBD In objNFBD.colItens
        Set objItemNFBD.ColAlocacoes = New ColAlocacoesItemNF
    Next
    
    'Faz a exclusao de uma nota fiscal de Saida
    If Not bMudouEstoque Then gobjMAT.iAceitaEstoqueNegativo = MARCADO
    lErro = CF("NotaFiscalEntrada_Excluir_EmTrans", objNFBD, objNFBD.objContabil, objNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    gobjMAT.iAceitaEstoqueNegativo = iAceitaEstoqueNegativo
    
    objNF.iStatus = objNFBD.iStatus 'A tela tenta gravar como LANÇADA mas pode já estar baixada, não pode perder o status
    
    NotaFiscalEntrada_Grava_Trata_Exclusao = SUCESSO

    Exit Function

Erro_NotaFiscalEntrada_Grava_Trata_Exclusao:

    NotaFiscalEntrada_Grava_Trata_Exclusao = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149593)

    End Select

    gobjMAT.iAceitaEstoqueNegativo = iAceitaEstoqueNegativo

    Exit Function

End Function

Public Function Importa_Xml_NFe_Cust() As Long
    Importa_Xml_NFe_Cust = SUCESSO
End Function

Public Function NFe_Grava_Vinculo_Xml(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim lNumNF As Long, sSerie As String
Dim objNFBD As New ClassNFiscal

On Error GoTo Erro_NFe_Grava_Vinculo_Xml

    If Len(Trim(objNF.sChvNFe)) > 0 Then

        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 211272
        Next
        
        objNFBD.sChvNFe = objNF.sChvNFe
        objNFBD.iFilialEmpresa = objNF.iFilialEmpresa
        
        lErro = CF("NFiscal_Le_Pela_ChvNFe", objNFBD, objNF.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then gError 211273 'Chv já vinculada a outra NF
                
        sSerie = String(STRING_MAXIMO, 0)
        
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Serie, NumNF FROM ImportNFeXml WHERE chvNFe = ? AND FilialEmpresa = ? ", 0, sSerie, lNumNF, objNF.sChvNFe, objNF.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 211274
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211275
        
        If lErro = AD_SQL_SUCESSO Then
        
            If objNF.lNumNotaFiscal <> lNumNF Then gError 211276 'Num NF não bate
            
            lErro = Comando_LockExclusive(alComando(1))
            If lErro <> SUCESSO Then gError 211277
            
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE ImportNFeXml SET NumIntNF = ?", alComando(1), objNF.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 211278

        End If
        
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If
    
    NFe_Grava_Vinculo_Xml = SUCESSO

    Exit Function

Erro_NFe_Grava_Vinculo_Xml:

    NFe_Grava_Vinculo_Xml = gErr

    Select Case gErr

        Case 211272
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211273
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CHV_VINC_OUTRA_NF", gErr)
            
        Case 211274 To 211275
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
            
        Case 211276
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_XML_OUTRO_NUM", gErr, objNF.lNumNotaFiscal, lNumNF)
            
        Case 211277
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", gErr)
            
        Case 211278
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_IMPORTNFEXML", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211279)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function NFe_Exclui_Vinculo_Xml(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim lNumNF As Long

On Error GoTo Erro_NFe_Exclui_Vinculo_Xml

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211280
    Next
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumNF FROM ImportNFeXml WHERE NumIntNF = ? ", 0, lNumNF, objNF.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 211281

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211282
    
    If lErro = AD_SQL_SUCESSO Then
             
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> SUCESSO Then gError 211283
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ImportNFeXml SET NumIntNF = ?", alComando(1), 0)
        If lErro <> AD_SQL_SUCESSO Then gError 211284

    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFe_Exclui_Vinculo_Xml = SUCESSO

    Exit Function

Erro_NFe_Exclui_Vinculo_Xml:

    NFe_Exclui_Vinculo_Xml = gErr

    Select Case gErr

        Case 211280
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 211281, 211282
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
            
        Case 211283
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_NFISCAL", gErr)
            
        Case 211284
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_IMPORTNFEXML", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211285)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function NF_Trata_MovEst_IMPCC(ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal objMovEstoque As ClassMovEstoque) As Long
'Cria os movimentos de estoque da NF de complemento de custo de importação

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objNFO As New ClassNFiscal, objItemNFO As ClassItemNF
Dim objItemNFAloc As ClassItemNFAlocacao, bAchou As Boolean
Dim dQtdeTotal As Double, dCustoTotal As Double
Dim colInsumos As New Collection, colRastroMovto As New Collection
Dim objItemMovEstoque As ClassItemMovEstoque

On Error GoTo Erro_NF_Trata_MovEst_IMPCC
    
   'calcula o custo do item da NF em função do credito do ICMS e do IPI
    lErro = NFiscalEntrada_Calcula_Custo_ImpCC(objItemNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objNFO.lNumIntDoc = objNFiscal.lNumIntNotaOriginal

    lErro = CF("NFiscal_Le", objNFO)
    If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
    
    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("AlocacoesNF_Le", objNFO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    bAchou = False
    For Each objItemNFO In objNFO.colItens
        If objItemNFO.sProduto = objItemNF.sProduto Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then gError 211333 'Produto não faz parte da NF original
       
    objProduto.sCodigo = objItemNF.sProduto

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                
        dQtdeTotal = 0
        For Each objItemNFAloc In objItemNFO.ColAlocacoes
            dQtdeTotal = dQtdeTotal + objItemNFAloc.dQuantidade
        Next
        
        dCustoTotal = 0
        For Each objItemNFAloc In objItemNFO.ColAlocacoes
        
            If objItemNFAloc.dQuantidade > 0 Then
                            
                Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, Arredonda_Moeda(objItemNF.dCusto * (objItemNFAloc.dQuantidade / dQtdeTotal), 4), APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objProduto.sSiglaUMEstoque, 0, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
            
                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                objItemMovEstoque.iClasseUM = objProduto.iClasseUM
                objItemMovEstoque.iControleEstoque = objProduto.iControleEstoque
                objItemMovEstoque.sSiglaUMEst = objProduto.sSiglaUMEstoque
                objItemMovEstoque.iApropriacaoProd = objProduto.iApropriacaoCusto
                
                dCustoTotal = dCustoTotal + objItemMovEstoque.dCusto
                
            End If
            
        Next
        
        If Abs(dCustoTotal - objItemNF.dCusto) > DELTA_VALORMONETARIO Then
        
             For Each objItemMovEstoque In objMovEstoque.colItens
                objItemMovEstoque.dCusto = objItemMovEstoque.dCusto + objItemNF.dCusto - dCustoTotal
                Exit For
             Next
        
        End If
    
    End If
        
    NF_Trata_MovEst_IMPCC = SUCESSO
    
    Exit Function

Erro_NF_Trata_MovEst_IMPCC:
            
    NF_Trata_MovEst_IMPCC = gErr

    Select Case gErr
    
        Case 211333
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_ORIG_PROD_NAO_ENCONTRADO", gErr, objProduto.sCodigo)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211334)

    End Select
    
    Exit Function

End Function

Function NotaFiscalEntrada_Cancelar_EmTrans1(ByVal objNFiscal As ClassNFiscal, ByVal dtDataCancelamento As Date) As Long
'Faz o cancelamento de uma nota fiscal de Entrada

Dim lErro As Long

On Error GoTo Erro_NotaFiscalEntrada_Cancelar_EmTrans1
    
    lErro = CF("NotaFiscalEntrada_Cancelar_EmTrans", objNFiscal, dtDataCancelamento)
    If lErro <> SUCESSO Then gError 210480
    
    NotaFiscalEntrada_Cancelar_EmTrans1 = SUCESSO

    Exit Function

Erro_NotaFiscalEntrada_Cancelar_EmTrans1:

    NotaFiscalEntrada_Cancelar_EmTrans1 = gErr

    Select Case gErr

        Case 210480
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210481)

    End Select

    Exit Function

End Function


Private Function DIInfo_Exclui_ItensPCDI(ByVal objDIInfo As ClassDIInfo) As Long

Dim lErro As Long, lNumIntDoc As Long
Dim iIndice As Integer
Dim alComando(1 To 1) As Long

On Error GoTo Erro_DIInfo_Exclui_ItensPCDI


    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 210599
    Next

    lErro = Comando_Executar(alComando(1), "DELETE FROM ItensPCDI WHERE NumIntDI = ?", objDIInfo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 210598
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    DIInfo_Exclui_ItensPCDI = SUCESSO
    
    Exit Function
    
Erro_DIInfo_Exclui_ItensPCDI:

    DIInfo_Exclui_ItensPCDI = gErr

    Select Case gErr

        Case 210598
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPCDI", gErr)

        Case 210599
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210600)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function DIInfo_Insere_ItensPCDI(ByVal objDIInfo As ClassDIInfo) As Long

Dim lErro As Long
Dim objItemPCDI As ClassItemPCDI
Dim iIndice As Integer
Dim alComando(1 To 1) As Long

On Error GoTo Erro_DIInfo_Insere_ItensPCDI

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 210599
    Next

    For Each objItemPCDI In objDIInfo.colItensPC
    
        With objItemPCDI
            lErro = Comando_Executar(alComando(1), "INSERT INTO ItensPCDI (NumIntDI, Seq, CodigoPC, DataPC, ProdutoPC, DescProdutoPC, QuantPC, " & _
                "UMPC, FilialEmpresa) VALUES (?,?,?,?,?,?,?,?,?)", _
                objDIInfo.lNumIntDoc, .iSeq, .lCodigoPC, .dtDataPC, .sProdutoPC, .sDescProdPC, .dQuantPC, _
                .sUMPC, .iFilialEmpresa)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 210603
    
    Next
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    DIInfo_Insere_ItensPCDI = SUCESSO
    
    Exit Function
    
Erro_DIInfo_Insere_ItensPCDI:

    DIInfo_Insere_ItensPCDI = gErr

    Select Case gErr
        
        Case 210603
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITENSPCDI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210604)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function NFImportacao_Monta_PC(ByVal objNF As ClassNFiscal, colPedCompras As Collection, colReqCompras As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 1) As Long
Dim objProduto As ClassProduto
Dim objItemNF As ClassItemNF
Dim objItemPCDI As ClassItemPCDI
Dim objPedidoCompra As ClassPedidoCompras
Dim objDIInfo As New ClassDIInfo
Dim objItemNFItemPC As ClassItemNFItemPC
Dim objItemNFItemRC As ClassItemNFItemRC
Dim objItemPC As ClassItemPedCompra
Dim bAchou As Boolean, dFatorRecebto As Double, dQtdeRecebto As Double
Dim dQtdeIRC As Double, lNumIntDocIRC As Long, dFator As Double
Dim dQtdeRecebtoNF As Double, colRCUsadas As New Collection

On Error GoTo Erro_NFImportacao_Monta_PC

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211670
    Next
    
    Set colPedCompras = New Collection
    For Each objItemNF In objNF.colItensNF
        Set objItemNF.colItemNFItemPC = New Collection
        Set objItemNF.colItemNFItemRC = New Collection
    Next
    
    objDIInfo.lNumIntDoc = objNF.objNFImportacao.objDIInfo.lNumIntDoc
    
    lErro = CF("DIInfo_Le_ItensPCDI", objDIInfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Se a DI possui pedidos de compra a serem baixados
    If objDIInfo.colItensPC.Count > 0 Then
    
        'Se tem pedidos ligados a DI lê
        For Each objItemPCDI In objDIInfo.colItensPC
            bAchou = False
            For Each objPedidoCompra In colPedCompras
                If objPedidoCompra.lCodigo = objItemPCDI.lCodigoPC And objPedidoCompra.iFilialEmpresa = objItemPCDI.iFilialEmpresa Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objPedidoCompra = New ClassPedidoCompras
                objPedidoCompra.lCodigo = objItemPCDI.lCodigoPC
                objPedidoCompra.iFilialEmpresa = objItemPCDI.iFilialEmpresa
                
                lErro = CF("PedidoCompra_Le_Todos", objPedidoCompra)
                If lErro <> SUCESSO And lErro <> 68486 Then gError ERRO_SEM_MENSAGEM
                
                'Ignora os pedidos baixados porque não terá o que fazer e estava dando erro em pedidos baixados ligados a DI mas de produtos não contidos na nota que está sendo gravada na
                'leitura ItensPC_Le (erro 56111) e contornando esse daria erro em NFiscalEntCom_BaixaPedidos (89220)
                If objPedidoCompra.iStatus <> PEDIDOCOMPRA_STATUS_BAIXADO Then
                
                    'lErro = CF("ItensPC_Le", objPedidoCompra)
                    'If lErro <> SUCESSO And lErro <> 56111 Then gError ERRO_SEM_MENSAGEM
                    lErro = CF("ItensPC_LeTodos", objPedidoCompra)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                    colPedCompras.Add objPedidoCompra
                    
                End If
            End If
        Next
    
        'Para cada Item da Nota Fiscal
        For Each objItemNF In objNF.colItensNF
        
            'O produto pode se repetir na lista de pedido de compra da DI
            'O produto pode se repetir na lista dos itens da nota fiscal
            
            'Vai distribuir pela quantidade do Item PC da DI mas controlando o saldo restante do item da NF e
            'o saldo no Item PC da DI
            
            dQtdeRecebtoNF = objItemNF.dQuantidade
        
            For Each objItemPCDI In objDIInfo.colItensPC
            
                'Não há repetição de produto no PC, então localiza o item do pedido pelo produto
                If objItemPCDI.sProdutoPC = objItemNF.sProduto Then
                
                    bAchou = False
                    For Each objPedidoCompra In colPedCompras
                        If objPedidoCompra.lCodigo = objItemPCDI.lCodigoPC And objPedidoCompra.iFilialEmpresa = objItemPCDI.iFilialEmpresa Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    
                    'Só trata o pedido da DI se achar na coleção pois senão é porque já está baixado
                    If bAchou Then
                        
                        bAchou = False
                        For Each objItemPC In objPedidoCompra.colItens
                            If objItemPC.sProduto = objItemNF.sProduto And objItemPC.iStatus <> ITENSPEDCOMPRA_STATUS_BAIXADO Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                        
                      End If
                        
                    'Só trata o pedido da DI se achar na coleção e o item não estiver baixado
                    If bAchou Then
                        
                        Set objProduto = New ClassProduto
                    
                        objProduto.sCodigo = objItemPCDI.sProdutoPC
                        
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 23080 Then gError ERRO_SEM_MENSAGEM
                        
                        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemPC.sUM, objItemNF.sUnidadeMed, dFator)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                        Set objItemNFItemPC = New ClassItemNFItemPC
                        
                        objItemNFItemPC.dQuantidade = objItemPCDI.dQuantPC * dFator
                        objItemNFItemPC.lItemPedCompra = objItemPC.lNumIntDoc
                        objItemNFItemPC.dTaxa = objItemPC.dTaxa
                        
                        'Se não tiver toda quantidade associada ao item de pedido a ser baixado, reduz e procura em outro depois o restante
                        '(no caso de ter repetições de um mesmo produto na DI (4 e 2 UN por exemplo), já ter feito uma NF de 2UN e ela ter baixado 2UN do de 4UN)
                        If objItemNFItemPC.dQuantidade > (objItemPC.dQuantidade - objItemPC.dQuantRecebida) * dFator Then
                            objItemNFItemPC.dQuantidade = (objItemPC.dQuantidade - objItemPC.dQuantRecebida) * dFator
                        End If
                        
                        'Alterado para tratar casos onde o produto é repetido na DI/NF para o mesmo PedCompra
                        'Podento também estar repetido para diferentes Pedidos na DI
                        If objItemNFItemPC.dQuantidade > dQtdeRecebtoNF Then
                            objItemNFItemPC.dQuantidade = dQtdeRecebtoNF
                        End If
                        
                       
                        dQtdeRecebtoNF = dQtdeRecebtoNF - objItemNFItemPC.dQuantidade
                        objItemPCDI.dQuantPC = objItemPCDI.dQuantPC - (objItemNFItemPC.dQuantidade / dFator)
                        
                        'Adiciona na coleção o objItemNFItemPC
                        If objItemNFItemPC.dQuantidade > QTDE_ESTOQUE_DELTA Then objItemNF.colItemNFItemPC.Add objItemNFItemPC
                        
                        'dFatorRecebto = objItemPCDI.dQuantPC / objItemPC.dQuantidade
    
                        'Le as requisições ligadas ao item do pedido ligado a DI
                        lErro = Comando_Executar(alComando(1), "SELECT ItemRC, Quantidade - QuantRecebida FROM ItemRCItemPC AS I WHERE I.ItemPC = ?", _
                                                            lNumIntDocIRC, dQtdeIRC, objItemPC.lNumIntDoc)
                        If lErro <> AD_SQL_SUCESSO Then gError 211671
                    
                        lErro = Comando_BuscarPrimeiro(alComando(1))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211672
                        
                        'dQtdeRecebto = objItemPCDI.dQuantPC
                        dQtdeRecebto = objItemNFItemPC.dQuantidade / dFator
                        
                        Do While lErro <> AD_SQL_SEM_DADOS
                           
                            'Abate do saldo já considerado para essa requisição
                            For Each objItemNFItemRC In colRCUsadas
                                If objItemNFItemRC.lItemReqCompra = lNumIntDocIRC And objItemNFItemRC.lNumIntDocItemPC = objItemPC.lNumIntDoc Then
                                    dQtdeIRC = dQtdeIRC - (objItemNFItemRC.dQuantidade / dFator)
                                End If
                            Next
                            
                            Set objItemNFItemRC = New ClassItemNFItemRC
                            
                            If dQtdeIRC > dQtdeRecebto Then
                                dQtdeIRC = dQtdeRecebto
                            End If
                            
                            objItemNFItemRC.dQuantidade = dQtdeIRC * dFator '* dFatorRecebto
                            objItemNFItemRC.lItemReqCompra = lNumIntDocIRC
                            objItemNFItemRC.lNumIntDocItemPC = objItemPC.lNumIntDoc
                            
                            dQtdeRecebto = dQtdeRecebto - dQtdeIRC
                        
                            'Adiciona na coleção objItemNFItemRC
                            If objItemNFItemRC.dQuantidade > QTDE_ESTOQUE_DELTA Then
                                objItemNF.colItemNFItemRC.Add objItemNFItemRC
                                colReqCompras.Add lNumIntDocIRC
                                colRCUsadas.Add objItemNFItemRC
                            End If
                            
                            If dQtdeRecebto < QTDE_ESTOQUE_DELTA Then Exit Do
                                    
                            lErro = Comando_BuscarProximo(alComando(1))
                            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211673
                                    
                        Loop
                    
                    End If
                    
                End If
            
            Next
        Next
        
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFImportacao_Monta_PC = SUCESSO
    
    Exit Function
    
Erro_NFImportacao_Monta_PC:

    NFImportacao_Monta_PC = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 211670
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 211671 To 211673
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMRCITEMPC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211674)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function NFImport_AtualizaConv(ByVal objNFiscal As ClassNFiscal, Optional ByVal bAtualizaProd As Boolean = True) As Long
'atualiza conversoes de campos vindos de xmls de nfes importadas para os campos dos cadastros do corporator

Dim lErro As Long, objItemNF As ClassItemNF
Dim lTransacao As Long, iIndice As Integer
Dim alComando(1 To 3) As Long, lCodCliente As Long, iCodFilial As Integer, lCodFornecedor As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_NFImport_AtualizaConv

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 201461

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201463
    Next
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
        
    If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Or objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
    
        'obter o cliente e filial
        lErro = Comando_Executar(alComando(1), "SELECT CodCliente, CodFilial FROM FiliaisClientes WHERE CGC = ?", _
            lCodCliente, iCodFilial, objNFiscal.sCGCEmitente)
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212016
        If lErro = AD_SQL_SUCESSO Then
        
            objNFiscal.lCliente = lCodCliente
            objNFiscal.iFilialCli = iCodFilial

        End If
        
    Else
    
        'obter o fornecedor e filial
        lErro = Comando_Executar(alComando(2), "SELECT CodFornecedor, CodFilial FROM FiliaisFornecedores WHERE CGC = ?", _
            lCodFornecedor, iCodFilial, objNFiscal.sCGCEmitente)
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 212018
        If lErro = AD_SQL_SUCESSO Then
        
            objNFiscal.lFornecedor = lCodFornecedor
            objNFiscal.iFilialForn = iCodFilial
    
        End If
    
    End If
    
    If bAtualizaProd Then
    
        For Each objItemNF In objNFiscal.colItens
            lErro = Comando_Executar(alComando(3), "UPDATE ImportNFeItensXml SET ProdutoCorporator = ?, NatOpInterna = ?, TipoTributacao = ?, SiglaUMCorporator = ? WHERE ChvNFe = ? AND Produto = ? AND CFOP = ? AND UM = ?", _
                objItemNF.sProduto, objItemNF.objTributacaoItemNF.sNaturezaOp, objItemNF.objTributacaoItemNF.iTipoTributacao, objItemNF.sUnidadeMed, objNFiscal.sChvNFe, objItemNF.sProdutoXml, objItemNF.sCFOPXml, objItemNF.sUnidadeMedXml)
            If lErro <> AD_SQL_SUCESSO Then gError 201464
        Next
        
    End If
        
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 201462
    
    NFImport_AtualizaConv = SUCESSO
    
    Exit Function
    
Erro_NFImport_AtualizaConv:

    NFImport_AtualizaConv = gErr

    Select Case gErr

        Case 201461
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 201462
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 201463
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 201464
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSXML", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201460)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Desfaz Transação
    Call Transacao_Rollback
    
    Exit Function

End Function

Function RelCustoArbitrado_Prepara(lNumIntRel As Long, ByVal iAno As Integer, ByVal iMes As Integer, ByVal iExibicao As Integer, ByVal dPercCusto As Double) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim alComando(0 To 5) As Long
Dim tProd As typeProduto, sAtivo As String
Dim dtDataIniPer As Date, dtDataFimPer As Date
Dim dSaldoEstoque As Double, dValorInicial As Double
Dim bAchouPreco As Boolean, dMaiorPreco As Double, sOrigem As String, sTipoPreco As String
Dim dtDataIniAux As Date, dtDataFimAux As Date, sUMAux As String, dPreco As Double, dFator As Double
Dim sSigla As String, sSerie As String, lNumNF As Long
Dim iTabela As Integer, iFilialEmp As Integer, dtDataVigencia As Date

On Error GoTo Erro_RelCustoArbitrado_Prepara

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213733

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213734
    Next
    
    dtDataIniPer = CDate("01/" & IIf(iMes < 10, "0", "") & CStr(iMes) & "/" & CStr(iAno))
    
    dtDataFimPer = DateAdd("m", 1, dtDataIniPer)
    dtDataFimPer = DateAdd("d", -1, dtDataFimPer)
       
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "ESTConfig", "NUMINTREL_RELCUSTOARBITRADO", lNumIntRel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    tProd.sCodigo = String(STRING_PRODUTO, 0)
    tProd.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tProd.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    tProd.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
       
    'Le os produtos produzidos
    lErro = Comando_Executar(alComando(1), "SELECT P.Codigo, P.Descricao, P.Ativo, P.SiglaUMEstoque, P.IPICodigo, P.ClasseUM FROM Produtos AS P, ProdutosFilial AS F WHERE P.Codigo = F.Produto AND P.Compras = 0 AND P.Natureza = ? AND P.Gerencial = 0 AND F.ProdNaFilial = 1 AND F.FilialEmpresa = ? ORDER BY Codigo", tProd.sCodigo, tProd.sDescricao, tProd.iAtivo, tProd.sSiglaUMEstoque, tProd.sIPICodigo, tProd.iClasseUM, NATUREZA_PROD_PRODUTO_ACABADO, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 213735
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213736

    Do While lErro <> AD_SQL_SEM_DADOS
    
        If tProd.iAtivo = PRODUTO_ATIVO Then
            sAtivo = "A"
        Else
            sAtivo = "I"
        End If
        
        dSaldoEstoque = 0
        dValorInicial = 0
        
        lErro = CF("Produto_ObterRelQtdSldIniFil", tProd.sCodigo, giFilialEmpresa, dtDataFimPer, dtDataFimPer, dSaldoEstoque, dValorInicial, 0, 0, 0, False)
        If lErro <> AD_SQL_SUCESSO And lErro <> 93720 Then gError ERRO_SEM_MENSAGEM
        
        bAchouPreco = False
        For iIndice = 1 To 3
            Select Case iIndice
                Case 1
                    sTipoPreco = "VM" 'Venda no Mês
                    dtDataIniAux = dtDataIniPer
                    dtDataFimAux = dtDataFimPer
                Case 2
                    sTipoPreco = "VA" 'Venda no Ano
                    dtDataIniAux = CDate("01/01/" & CStr(iAno))
                    dtDataFimAux = CDate("31/12/" & CStr(iAno))
                Case 3
                    sTipoPreco = "V" 'venda
                    dtDataIniAux = DATA_NULA
                    dtDataFimAux = DateAdd("y", 50, Date)
            End Select
            
            sUMAux = String(STRING_UM_SIGLA, 0)
        
            lErro = Comando_Executar(alComando(2), "SELECT I.UnidadeMed, MAX(I.PrecoUnitario) AS MaxPreco FROM NFiscal AS N, TiposDocInfo AS T , ItensNFiscal AS I WHERE N.TipoNFiscal = T.Codigo AND N.NumIntDoc = I.NumIntNF AND T.Faturamento = 1 AND N.Status <> 7 AND T.Tipo = 2 AND Produto = ? AND N.DataEmissao BETWEEN ? AND ? AND FilialEmpresa = ? GROUP BY I.UnidadeMed", _
                                                    sUMAux, dPreco, tProd.sCodigo, dtDataIniAux, dtDataFimAux, giFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 213737
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213738
        
            dMaiorPreco = 0
            Do While lErro <> AD_SQL_SEM_DADOS
            
                bAchouPreco = True
                
                If sUMAux <> tProd.sSiglaUMEstoque Then
                
                    lErro = CF("UM_Conversao", tProd.iClasseUM, tProd.sSiglaUMEstoque, sUMAux, dFator)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    dPreco = dPreco * dFator
                
                End If
                
                If dPreco > dMaiorPreco Then dMaiorPreco = dPreco
                
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213739
                
            Loop
            
            If bAchouPreco Then Exit For
        
        Next
        
        sOrigem = ""
        If bAchouPreco Then
        'Se achou uma venda pega os dados dela
        
            sSigla = String(STRING_TIPODOCINFO_SIGLA, 0)
            sSerie = String(STRING_SERIE, 0)
        
            lErro = Comando_Executar(alComando(4), "SELECT T.Sigla, N.Serie, N.NumNotaFiscal FROM NFiscal AS N, TiposDocInfo AS T , ItensNFiscal AS I WHERE N.TipoNFiscal = T.Codigo AND N.NumIntDoc = I.NumIntNF AND T.Faturamento = 1 AND N.Status <> 7 AND T.Tipo = 2 AND Produto = ? AND N.DataEmissao BETWEEN ? AND ? AND I.UnidadeMed = ? AND FilialEmpresa = ? ORDER BY I.PrecoUnitario DESC, N.DataEmissao DESC", _
                                                    sSigla, sSerie, lNumNF, tProd.sCodigo, dtDataIniAux, dtDataFimAux, sUMAux, giFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 213740
            
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213741
            
            sOrigem = sSigla & " " & sSerie & " " & CStr(lNumNF)
            
        Else
        'Sem venda, procura na tabela de preço
        
            sTipoPreco = "TP" 'Tabela de preço
        
            lErro = Comando_Executar(alComando(5), "SELECT CodTabela, FilialEmpresa, Preco, DataVigencia FROM TabelasDePrecoItens WHERE FilialEmpresa IN (?,?) AND CodProduto = ? AND DataVigencia BETWEEN ? AND ? AND Ativo = 0 ORDER BY FilialEmpresa DESC", _
                                                    iTabela, iFilialEmp, dPreco, dtDataVigencia, EMPRESA_TODA, giFilialEmpresa, tProd.sCodigo, dtDataIniPer, dtDataIniPer)
            If lErro <> AD_SQL_SUCESSO Then gError 213742
            
            lErro = Comando_BuscarProximo(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213743
            
            If lErro <> AD_SQL_SUCESSO Then
                sOrigem = ""
                dMaiorPreco = 0
            Else
                sOrigem = "C:" & CStr(iTabela) & " F:" & CStr(iFilialEmp) & " V:" & Format(dtDataVigencia, "dd/mm/yyyy")
                dMaiorPreco = dPreco
            End If
        End If
        
        If (iExibicao = 0 And dSaldoEstoque > QTDE_ESTOQUE_DELTA2) Or (iExibicao = 1 And Abs(dSaldoEstoque) < QTDE_ESTOQUE_DELTA) Or iExibicao = 2 Then
    
            lErro = Comando_Executar(alComando(0), "INSERT INTO RelCustoArbitrado (NumIntRel, Produto, Descricao, UM, Ativo, NCM, SaldoEstoque, MaiorPrecoVenda, TipoPreco, OrigemPreco, CustoArbitrado, CustoTotal) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
            lNumIntRel, tProd.sCodigo, tProd.sDescricao, tProd.sSiglaUMEstoque, sAtivo, left(tProd.sIPICodigo, 8), dSaldoEstoque, dMaiorPreco, sTipoPreco, sOrigem, dMaiorPreco * dPercCusto, dSaldoEstoque * dMaiorPreco * dPercCusto)
            If lErro <> AD_SQL_SUCESSO Then gError 213744

        End If
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213745
        
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213746

    RelCustoArbitrado_Prepara = SUCESSO

    Exit Function

Erro_RelCustoArbitrado_Prepara:

    RelCustoArbitrado_Prepara = gErr

    Select Case gErr

        Case 213733
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 213734
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213735 To 213743, 213745
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case 213744
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_RELCUSTOARBITRADO", gErr)

        Case 213746
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213747)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Public Function NFiscalEntCom_Gera_Reservas(ByVal objNF As ClassNFiscal, objME As ClassMovEstoque) As Long
    NFiscalEntCom_Gera_Reservas = SUCESSO
End Function

Function NF_Grava_Trib_E_MovEst() As Long
'/***********************************************************************/
'/*** ROTINA DE ACERTO DE NFS GRAVADAS COM TRIBUTACAO E CUSTO ERRADOS ***/
'/***********************************************************************/

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim vbResult As VbMsgBoxResult
Dim objNF As ClassNFiscal
Dim objItemNF As ClassItemNF
Dim objNFBD As ClassNFiscal
Dim objItemNFBD As ClassItemNF
Dim dFator As Double, dQtde As Double
Dim sSQL As String, lNumIntDocNF As Long
Dim iNumItensMesmoPreco As Integer
Dim objProduto As ClassProduto

On Error GoTo Erro_NF_Grava_Trib_E_MovEst

    If gobjEST.iPrioridadeProduto = 171 And LCase(left(gsUsuario, 5)) = "super" Then
    
        'Lê as NFs Fetura via Pedido de compra cujo Preço dos itens não bate na tributação porque tributou o item errado
        sSQL = "SELECT N.NumIntDoc FROM NFiscal As N, ItensNFiscal As I, TributacaoDoc AS TD, TributacaoDocItem AS TI"
        sSQL = sSQL & " " & "WHERE N.NumIntDoc = I.NumIntNF AND TD.TipoDoc = 0 AND TD.TipoDoc = TI.TipoDoc AND"
        sSQL = sSQL & " " & "N.NumIntDoc = TD.NumIntDoc AND TD.NumIntDoc = TI.NumIntDoc AND I.NumIntDoc = TI.NumIntDocItem AND"
        sSQL = sSQL & " " & "N.TipoNFiscal = 84 And Abs(TI.PrecoTotal - (I.Quantidade * I.PrecoUnitario - I.ValorDesconto)) > 0.009"
        sSQL = sSQL & " " & "AND N.ChvNFe IN (SELECT ChvNFe FROM ImportNFeXml)"
        sSQL = sSQL & " " & "ORDER BY N.NumIntDoc"

        vbResult = Rotina_Aviso(vbYesNo, "Rotina administrativa, caso não seja um Analista do Corporator escolha não e comunique ao suporte. Deseja continuar ?")
        If vbResult = vbYes Then
    
            'Abre a Transação
            lTransacao = Transacao_Abrir()
            If lTransacao = 0 Then gError 66500
            
            'Abertura de Comando
            For iIndice = LBound(alComando) To UBound(alComando)
                alComando(iIndice) = Comando_Abrir()
                If alComando(iIndice) = 0 Then gError 213734
            Next
        
            lErro = Comando_Executar(alComando(2), sSQL, lNumIntDocNF)
            If lErro <> AD_SQL_SUCESSO Then gError 89023
        
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89024
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                Set objNF = New ClassNFiscal
                Set objNFBD = New ClassNFiscal
                
                If lNumIntDocNF = 1002 Then
                    lErro = lErro
                End If
                
                objNF.lNumIntDoc = lNumIntDocNF
                objNFBD.lNumIntDoc = lNumIntDocNF
                
                'Tenta ler a nota Fiscal passada por parametro
                lErro = CF("NFiscal_Le", objNF)
                If lErro <> SUCESSO Then gError 31425
                
                'Lê os Ítens da Nota Fiscal
                lErro = CF("NFiscalItens_Le", objNF)
                If lErro <> SUCESSO Then gError 31425
                                
                'Lê a parte de tributacao
                lErro = CF("NFiscal_Le_Tributacao", objNF)
                If lErro <> SUCESSO Then gError 39342
            
                'Lê a parte de tributacao dos Ítens da Nota Fiscal
                lErro = CF("NFiscalItens_Le_Trib", objNF)
                If lErro <> SUCESSO Then gError 39343
                
                'Tenta ler a nota Fiscal passada por parametro
                lErro = CF("NFiscal_Le", objNFBD)
                If lErro <> SUCESSO Then gError 31425
                
                'Lê os Ítens da Nota Fiscal
                lErro = CF("NFiscalItens_Le", objNFBD)
                If lErro <> SUCESSO Then gError 31425
                
                'Lê a parte de tributacao
                lErro = CF("NFiscal_Le_Tributacao", objNFBD)
                If lErro <> SUCESSO Then gError 39342
            
                'Lê a parte de tributacao dos Ítens da Nota Fiscal
                lErro = CF("NFiscalItens_Le_Trib", objNFBD)
                If lErro <> SUCESSO Then gError 39343
                
                'Acerta a tributação do item
                For Each objItemNF In objNF.colItens
                
                    iNumItensMesmoPreco = 0
                    If Abs(objItemNF.dValorTotal - objItemNF.objTributacao.dPrecoTotal) > DELTA_VALORMONETARIO Then
                    
                        For Each objItemNFBD In objNFBD.colItens
                            If Abs(objItemNF.dValorTotal - objItemNFBD.objTributacao.dPrecoTotal) < DELTA_VALORMONETARIO Then
                                iNumItensMesmoPreco = iNumItensMesmoPreco + 1
                            End If
                        Next
                        
                        If iNumItensMesmoPreco = 0 Then gError 99997
                        
                        If objItemNF.lNumIntDoc = 3512 Or _
                            objItemNF.lNumIntDoc = 3514 Or _
                            objItemNF.lNumIntDoc = 3649 Or objItemNF.lNumIntDoc = 8514 Then iNumItensMesmoPreco = 1
                        
                        If iNumItensMesmoPreco > 1 Then gError 99998
                        
                        If iNumItensMesmoPreco = 1 Then
                        
                            For Each objItemNFBD In objNFBD.colItens
                                 If Abs(objItemNF.dValorTotal - objItemNFBD.objTributacao.dPrecoTotal) < DELTA_VALORMONETARIO Then
                                     Exit For
                                 End If
                             Next
                             
                             Call objItemNF.objTributacao.Copia2(objItemNFBD.objTributacao)
                             
                            objItemNF.objTributacao.lNumIntDoc = objNF.lNumIntDoc
                            objItemNF.objTributacao.lNumIntDocItem = objItemNF.lNumIntDoc
                            objItemNF.objTributacao.iItem = objItemNF.iItem
                            objItemNF.objTributacao.dDescontoGrid = objItemNF.dValorDesconto
                            objItemNF.objTributacao.dPrecoUnitario = objItemNF.dPrecoUnitario
                            objItemNF.objTributacao.dPrecoUnitarioMoeda = objItemNF.dPrecoUnitarioMoeda
                            objItemNF.objTributacao.dQuantidade = objItemNF.dQuantidade
                            objItemNF.objTributacao.sProduto = objItemNF.sProduto
                            objItemNF.objTributacao.sProdutoDescricao = objItemNF.sDescricaoItem
                            objItemNF.objTributacao.sUnidadeMed = objItemNF.sUnidadeMed
                            objItemNF.objTributacao.iTipoDoc = TIPODOC_TRIB_NF
                        
                        End If
                    
                    End If
                
                Next
                
                For Each objItemNF In objNF.colItens
                
                    Set objProduto = New ClassProduto
                
                    objProduto.sCodigo = objItemNF.sProduto
                
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 42732
                
                    If objItemNF.objTributacao.lNumIntDocItem = 0 Then gError 999999
                    
                    objItemNF.objTributacao.sIPICodProduto = objProduto.sIPICodigo
                    
                    lErro = CF("TributacaoDocItem_Exclui", objItemNF.objTributacao.iTipoDoc, objItemNF.objTributacao.lNumIntDocItem)
                    If lErro <> SUCESSO Then gError 66551
                
                    lErro = CF("TributacaoDocItem_Grava", objItemNF.objTributacao)
                    If lErro <> SUCESSO Then gError 66551
                    
                    lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
                    If lErro <> SUCESSO Then gError 66588
    
                    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 89464
                    
                    lErro = Comando_ExecutarPos(alComando(0), "SELECT Quantidade FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? AND FilialEmpresa = ?", 0, dQtde, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL, objItemNF.lNumIntDoc, objNF.iFilialEmpresa)
                    If lErro <> AD_SQL_SUCESSO Then gError 89023
                
                    lErro = Comando_BuscarPrimeiro(alComando(0))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89024
                    
                    Do While lErro <> AD_SQL_SEM_DADOS
                    
                        lErro = Comando_ExecutarPos(alComando(1), "UPDATE MovimentoEstoque SET Custo = ?", alComando(0), (objItemNF.dCusto / objItemNF.dQuantidade) * (dQtde / dFator))
                        If lErro <> AD_SQL_SUCESSO Then gError 89023
                    
                        lErro = Comando_BuscarProximo(alComando(0))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89024
                    
                    Loop
                   
                Next
                
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89024
                
            Loop
            
            'Fecha Comando
            For iIndice = LBound(alComando) To UBound(alComando)
                Call Comando_Fechar(alComando(iIndice))
            Next
    
            lErro = Transacao_Commit()
            If lErro <> AD_SQL_SUCESSO Then gError 66509
                        
        End If
        
    End If

    NF_Grava_Trib_E_MovEst = SUCESSO

    Exit Function

Erro_NF_Grava_Trib_E_MovEst:

    NF_Grava_Trib_E_MovEst = gErr

    Select Case gErr

        Case 66500
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 66509
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 99997
            Call Rotina_Erro(vbOKOnly, "O item " & CStr(objItemNF.lNumIntDoc) & " está errado e não foi encontrada a tributação certa", gErr)
        
        Case 99998
            Call Rotina_Erro(vbOKOnly, "O item " & CStr(objItemNF.lNumIntDoc) & " está errado e foi encontrada mais de uma tributação compatível", gErr)

        Case 99999
            Call Rotina_Erro(vbOKOnly, "Item não gravado", gErr)

        Case 213734
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149559)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
            
    Call Transacao_Rollback

    Exit Function

End Function

Function ImportXMLCadProd_Grava(ByVal colProdutos As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim objProd As ClassProduto
Dim colTabelaPrecoItem As New Collection
Dim objClassFisc As ClassClassificacaoFiscal

On Error GoTo Erro_ImportXMLCadProd_Grava

    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213827

    For Each objProd In colProdutos
    
        Set objClassFisc = New ClassClassificacaoFiscal
    
        objClassFisc.sCodigo = objProd.sIPICodigo
        objClassFisc.sDescricao = objProd.sIPICodigo
        
        If Len(Trim(objClassFisc.sCodigo)) > 0 Then
        
            lErro = CF("ClassificacaoFiscal_Le", objClassFisc)
            If lErro <> SUCESSO And lErro <> 123494 Then gError ERRO_SEM_MENSAGEM
            
            If lErro <> SUCESSO Then
            
                lErro = CF("ClassificacaoFiscal_Grava", objClassFisc, True)
                If lErro <> SUCESSO And lErro <> 123494 Then gError ERRO_SEM_MENSAGEM
            
            End If
        
        End If
    
        lErro = CF("Produto_Grava_Trans", objProd, colTabelaPrecoItem)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 213828

    ImportXMLCadProd_Grava = SUCESSO

    Exit Function

Erro_ImportXMLCadProd_Grava:

    ImportXMLCadProd_Grava = gErr

    Select Case gErr

        Case 213827
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case 213828
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213829)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Rotina_Reproc_Calc_SldIni(ByVal objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long
Dim alComando(1 To 14) As Long
Dim sComandoSQL As String
Dim iIndice As Integer
Dim iFilialEmp As Integer
Dim tEst As typeEstoqueProduto
Dim bComSldIni As Boolean, bComSldIni1 As Boolean, bComSldIni2 As Boolean
Dim bComSldIniFilial As Boolean, bComSldIni1Filial As Boolean, bComSldIni2Filial As Boolean
Dim objEstFilial As ClassEstoqueProduto
Dim objEstoqueMes As ClassEstoqueMes, objEstoqueProduto As ClassEstoqueProduto
Dim bForcado As Boolean, dSldMesQtdIni As Double, dSldMesVlrIni As Double
Dim dCusto As Double, dCustoMedio As Double, objProduto As ClassProduto

On Error GoTo Erro_Rotina_Reproc_Calc_SldIni
    
    'Inicia o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216073
    Next
    
    'Lê os valores iniciais em estoque produto
    sComandoSQL = "SELECT E.Produto, E.Almoxarifado, A.FilialEmpresa, E.DataInicial, E.QuantidadeInicial, E.SaldoInicial," & _
                    "E.QuantInicialConsig3, E.QuantInicialConsig, E.QuantInicialDemo3, E.QuantInicialDemo," & _
                    "E.QuantInicialConserto3, E.QuantInicialConserto, E.QuantInicialOutras3, E.QuantInicialOutras," & _
                    "E.QuantInicialBenef, E.QuantInicialBenef3, E.ValorInicialConsig3, E.ValorInicialConsig," & _
                    "E.ValorInicialDemo3, E.ValorInicialDemo, E.ValorInicialConserto3, E.ValorInicialConserto," & _
                    "E.ValorInicialOutras3, E.ValorInicialOutras, E.ValorInicialBenef, E.ValorInicialBenef3 FROM EstoqueProduto AS E, Almoxarifado AS A WHERE E.Almoxarifado = A.Codigo "
                    
    tEst.sProduto = String(STRING_PRODUTO, 0)
    
    'Se teve filtro por produto
    If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
    
        'Inclui o filtro de produto no select
        sComandoSQL = sComandoSQL & " AND Produto = ? ORDER BY FilialEmpresa, DataInicial"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
    
        lErro = Comando_Executar(alComando(1), sComandoSQL, tEst.sProduto, tEst.iAlmoxarifado, iFilialEmp, tEst.dtDataInicial, tEst.dQuantidadeInicial, tEst.dSaldoInicial, _
                                                tEst.dQuantInicialConsig3, tEst.dQuantInicialConsig, tEst.dQuantInicialDemo3, tEst.dQuantInicialDemo, _
                                                tEst.dQuantInicialConserto3, tEst.dQuantInicialConserto, tEst.dQuantInicialOutras3, tEst.dQuantInicialOutras, _
                                                tEst.dQuantInicialBenef, tEst.dQuantInicialBenef3, tEst.dValorInicialConsig3, tEst.dValorInicialConsig, _
                                                tEst.dValorInicialDemo3, tEst.dValorInicialDemo, tEst.dValorInicialConserto3, tEst.dValorInicialConserto, _
                                                tEst.dValorInicialOutras3, tEst.dValorInicialOutras, tEst.dValorInicialBenef, tEst.dValorInicialBenef3, objReprocessamentoEst.sProdutoCodigo)
        If lErro <> SUCESSO Then gError 216074
    
    Else
    
        sComandoSQL = sComandoSQL & " ORDER BY Produto, FilialEmpresa, DataInicial"
        If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then sComandoSQL = sComandoSQL & " WHERE FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
        
        lErro = Comando_Executar(alComando(1), sComandoSQL, tEst.sProduto, tEst.iAlmoxarifado, iFilialEmp, tEst.dtDataInicial, tEst.dQuantidadeInicial, tEst.dSaldoInicial, _
                                                tEst.dQuantInicialConsig3, tEst.dQuantInicialConsig, tEst.dQuantInicialDemo3, tEst.dQuantInicialDemo, _
                                                tEst.dQuantInicialConserto3, tEst.dQuantInicialConserto, tEst.dQuantInicialOutras3, tEst.dQuantInicialOutras, _
                                                tEst.dQuantInicialBenef, tEst.dQuantInicialBenef3, tEst.dValorInicialConsig3, tEst.dValorInicialConsig, _
                                                tEst.dValorInicialDemo3, tEst.dValorInicialDemo, tEst.dValorInicialConserto3, tEst.dValorInicialConserto, _
                                                tEst.dValorInicialOutras3, tEst.dValorInicialOutras, tEst.dValorInicialBenef, tEst.dValorInicialBenef3)
        If lErro <> SUCESSO Then gError 216075

    End If
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216076
    
    'Inicializa dados por produto/filial com o 1o lido
    Set objEstFilial = New ClassEstoqueProduto
    objEstFilial.sProduto = tEst.sProduto
    objEstFilial.iFilialEmpresa = iFilialEmp
    objEstFilial.dtDataInicial = tEst.dtDataInicial
    objEstFilial.iAlmoxarifado = tEst.iAlmoxarifado

    bComSldIniFilial = False
    bComSldIni1Filial = False
    bComSldIni2Filial = False
    
    bForcado = False
    Do While (lErro <> AD_SQL_SEM_DADOS Or bForcado)
       
        'Se mudou de produto ou filial empresa ou então já acabou de ler e vai fazer o ajuste do último (bForcado)
        If objEstFilial.sProduto <> tEst.sProduto Or objEstFilial.iFilialEmpresa <> iFilialEmp Or bForcado Then
        
            dCusto = 0
            dCustoMedio = 0
            
            'Calcula o custo
            Set objProduto = New ClassProduto
            objProduto.sCodigo = objEstFilial.sProduto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
               
            If objProduto.iApropriacaoCusto = APROPR_CUSTO_REAL Then
                 If objEstFilial.dQuantidadeInicial > 0 Then dCusto = objEstFilial.dSaldoInicial / objEstFilial.dQuantidadeInicial
             ElseIf objProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO Then
                 If objEstFilial.dQuantidadeInicial > 0 Then dCustoMedio = objEstFilial.dSaldoInicial / objEstFilial.dQuantidadeInicial
             End If
        
            'Soma as quantidades e valores da filial (dele com ele e dele em 3os)
            dSldMesQtdIni = objEstFilial.dQuantidadeInicial + objEstFilial.dQuantInicialConsig3 + objEstFilial.dQuantInicialDemo3 + objEstFilial.dQuantInicialConserto3 + objEstFilial.dQuantInicialOutras3 + objEstFilial.dQuantInicialBenef3
            dSldMesVlrIni = objEstFilial.dSaldoInicial + objEstFilial.dValorInicialConsig3 + objEstFilial.dValorInicialDemo3 + objEstFilial.dValorInicialConserto3 + objEstFilial.dValorInicialOutras3 + objEstFilial.dValorInicialBenef3
        
            If bComSldIniFilial Then
                lErro = Comando_Executar(alComando(6), "UPDATE SldMesEst SET QuantInicial = ?, ValorInicial = ?,  QuantInicialCusto = ?, ValorInicialCusto = ?, CustoMedioProducaoInicial =?, CustoMedio" + CStr(Month(objEstFilial.dtDataInicial)) + "=?  WHERE FilialEmpresa=? AND Ano=? AND Produto=? ", dSldMesQtdIni, dSldMesVlrIni, objEstFilial.dQuantidadeInicial, objEstFilial.dSaldoInicial, dCusto, dCustoMedio, objEstFilial.iFilialEmpresa, Year(objEstFilial.dtDataInicial), objEstFilial.sProduto)
                If lErro <> SUCESSO Then gError 216077
            End If
        
            If bComSldIni1Filial Then
                lErro = Comando_Executar(alComando(7), "UPDATE SldMesEst1 SET QuantInicialConsig3=?,ValorInicialConsig3=?,QuantInicialDemo3=?,ValorInicialDemo3=?,QuantInicialConserto3=?,ValorInicialConserto3=?,QuantInicialOutros3=?,ValorInicialOutros3=?,QuantInicialBenef3=? WHERE FilialEmpresa=? AND Ano=? AND Produto=? ", _
                                                        objEstFilial.dQuantInicialConsig3, objEstFilial.dValorInicialConsig3, objEstFilial.dQuantInicialDemo3, objEstFilial.dValorInicialDemo3, objEstFilial.dQuantInicialConserto3, objEstFilial.dValorInicialConserto3, objEstFilial.dQuantInicialOutras3, objEstFilial.dValorInicialOutras3, objEstFilial.dQuantInicialBenef3, objEstFilial.iFilialEmpresa, Year(objEstFilial.dtDataInicial), objEstFilial.sProduto)
                If lErro <> SUCESSO Then gError 216078
            End If
            
            If bComSldIni2Filial Then
                lErro = Comando_Executar(alComando(8), "UPDATE SldMesEst2 SET QuantInicialConsig=?,ValorInicialConsig=?,QuantInicialDemo=?,ValorInicialDemo=?,QuantInicialConserto=?,ValorInicialConserto=?,QuantInicialOutros=?,ValorInicialOutros=?,QuantInicialBenef=? WHERE FilialEmpresa=? AND Ano=? AND Produto=? ", _
                                                        objEstFilial.dQuantInicialConsig, objEstFilial.dValorInicialConsig, objEstFilial.dQuantInicialDemo, objEstFilial.dValorInicialDemo, objEstFilial.dQuantInicialConserto, objEstFilial.dValorInicialConserto, objEstFilial.dQuantInicialOutras, objEstFilial.dValorInicialOutras, objEstFilial.dQuantInicialBenef, objEstFilial.iFilialEmpresa, Year(objEstFilial.dtDataInicial), objEstFilial.sProduto)
                If lErro <> SUCESSO Then gError 216079
            End If
            
            'Apaga saldos anteriores ao inicial
            lErro = Comando_Executar(alComando(9), "DELETE FROM SldMesEstAlm WHERE Almoxarifado=? AND Produto=? AND Ano<? ", objEstFilial.iAlmoxarifado, objEstFilial.sProduto, Year(objEstFilial.dtDataInicial))
            If lErro <> SUCESSO Then gError 216086
            
            lErro = Comando_Executar(alComando(10), "DELETE FROM SldMesEstAlm1 WHERE Almoxarifado=? AND Produto=? AND Ano<? ", objEstFilial.iAlmoxarifado, objEstFilial.sProduto, Year(objEstFilial.dtDataInicial))
            If lErro <> SUCESSO Then gError 216087
            
            lErro = Comando_Executar(alComando(11), "DELETE FROM SldMesEstAlm2 WHERE Almoxarifado=? AND Produto=? AND Ano<? ", objEstFilial.iAlmoxarifado, objEstFilial.sProduto, Year(objEstFilial.dtDataInicial))
            If lErro <> SUCESSO Then gError 216088
            
            lErro = Comando_Executar(alComando(12), "DELETE FROM SldMesEst WHERE FilialEmpresa=? AND Ano<? AND Produto=? ", objEstFilial.iFilialEmpresa, Year(objEstFilial.dtDataInicial), objEstFilial.sProduto)
            If lErro <> SUCESSO Then gError 216089
            
            lErro = Comando_Executar(alComando(13), "DELETE FROM SldMesEst1 WHERE FilialEmpresa=? AND Ano<? AND Produto=? ", objEstFilial.iFilialEmpresa, Year(objEstFilial.dtDataInicial), objEstFilial.sProduto)
            If lErro <> SUCESSO Then gError 216090
            
            lErro = Comando_Executar(alComando(14), "DELETE FROM SldMesEst2 WHERE FilialEmpresa=? AND Ano<? AND Produto=? ", objEstFilial.iFilialEmpresa, Year(objEstFilial.dtDataInicial), objEstFilial.sProduto)
            If lErro <> SUCESSO Then gError 216091
        
            'Reinicializa dados por produto/filial
            Set objEstFilial = New ClassEstoqueProduto
            objEstFilial.sProduto = tEst.sProduto
            objEstFilial.iFilialEmpresa = iFilialEmp
            objEstFilial.dtDataInicial = tEst.dtDataInicial
            objEstFilial.iAlmoxarifado = tEst.iAlmoxarifado

            bComSldIniFilial = False
            bComSldIni1Filial = False
            bComSldIni2Filial = False
        
            If bForcado Then Exit Do
            
        End If
        
        'Soma as quantidades e valores da filial (dele com ele e dele em 3os)
        dSldMesQtdIni = tEst.dQuantidadeInicial + tEst.dQuantInicialConsig3 + tEst.dQuantInicialDemo3 + tEst.dQuantInicialConserto3 + tEst.dQuantInicialOutras3 + tEst.dQuantInicialBenef3
        dSldMesVlrIni = tEst.dSaldoInicial + tEst.dValorInicialConsig3 + tEst.dValorInicialDemo3 + tEst.dValorInicialConserto3 + tEst.dValorInicialOutras3 + tEst.dValorInicialBenef3
        
        bComSldIni = False
        bComSldIni1 = False
        bComSldIni2 = False
        If dSldMesQtdIni > QTDE_ESTOQUE_DELTA Or dSldMesVlrIni > DELTA_VALORMONETARIO Then
            bComSldIni = True
            bComSldIniFilial = True
        End If
        If (tEst.dQuantInicialConsig3 + tEst.dQuantInicialDemo3 + tEst.dQuantInicialConserto3 + tEst.dQuantInicialOutras3 + tEst.dQuantInicialBenef3 > QTDE_ESTOQUE_DELTA) Or _
            (tEst.dValorInicialConsig3 + tEst.dValorInicialDemo3 + tEst.dValorInicialConserto3 + tEst.dValorInicialOutras3 + tEst.dValorInicialBenef3 > DELTA_VALORMONETARIO) Then
            bComSldIni1 = True
            bComSldIni1Filial = True
        End If
        If (tEst.dQuantInicialConsig + tEst.dQuantInicialDemo + tEst.dQuantInicialConserto + tEst.dQuantInicialOutras + tEst.dQuantInicialBenef > QTDE_ESTOQUE_DELTA) Or _
            (tEst.dValorInicialConsig + tEst.dValorInicialDemo + tEst.dValorInicialConserto + tEst.dValorInicialOutras + tEst.dValorInicialBenef > DELTA_VALORMONETARIO) Then
            bComSldIni2 = True
            bComSldIni2Filial = True
        End If
        
        'Acumula por produto/filial independente do almoxarifado
        objEstFilial.dQuantidadeInicial = objEstFilial.dQuantidadeInicial + tEst.dQuantidadeInicial
        objEstFilial.dSaldoInicial = objEstFilial.dSaldoInicial + tEst.dSaldoInicial
        
        objEstFilial.dQuantInicialConsig3 = objEstFilial.dQuantInicialConsig3 + tEst.dQuantInicialConsig3
        objEstFilial.dQuantInicialDemo3 = objEstFilial.dQuantInicialDemo3 + tEst.dQuantInicialDemo3
        objEstFilial.dQuantInicialConserto3 = objEstFilial.dQuantInicialConserto3 + tEst.dQuantInicialConserto3
        objEstFilial.dQuantInicialOutras3 = objEstFilial.dQuantInicialOutras3 + tEst.dQuantInicialOutras3
        objEstFilial.dQuantInicialBenef3 = objEstFilial.dQuantInicialBenef3 + tEst.dQuantInicialBenef3
        objEstFilial.dValorInicialConsig3 = objEstFilial.dValorInicialConsig3 + tEst.dValorInicialConsig3
        objEstFilial.dValorInicialDemo3 = objEstFilial.dValorInicialDemo3 + tEst.dValorInicialDemo3
        objEstFilial.dValorInicialConserto3 = objEstFilial.dValorInicialConserto3 + tEst.dValorInicialConserto3
        objEstFilial.dValorInicialOutras3 = objEstFilial.dValorInicialOutras3 + tEst.dValorInicialOutras3
        objEstFilial.dValorInicialBenef3 = objEstFilial.dValorInicialBenef3 + tEst.dValorInicialBenef3

        objEstFilial.dQuantInicialConsig = objEstFilial.dQuantInicialConsig + tEst.dQuantInicialConsig
        objEstFilial.dQuantInicialDemo = objEstFilial.dQuantInicialDemo + tEst.dQuantInicialDemo
        objEstFilial.dQuantInicialConserto = objEstFilial.dQuantInicialConserto + tEst.dQuantInicialConserto
        objEstFilial.dQuantInicialOutras = objEstFilial.dQuantInicialOutras + tEst.dQuantInicialOutras
        objEstFilial.dQuantInicialBenef = objEstFilial.dQuantInicialBenef + tEst.dQuantInicialBenef
        objEstFilial.dValorInicialConsig = objEstFilial.dValorInicialConsig + tEst.dValorInicialConsig
        objEstFilial.dValorInicialDemo = objEstFilial.dValorInicialDemo + tEst.dValorInicialDemo
        objEstFilial.dValorInicialConserto = objEstFilial.dValorInicialConserto + tEst.dValorInicialConserto
        objEstFilial.dValorInicialOutras = objEstFilial.dValorInicialOutras + tEst.dValorInicialOutras
        objEstFilial.dValorInicialBenef = objEstFilial.dValorInicialBenef + tEst.dValorInicialBenef
        
        'Se tem saldo e não começou na mesma data do 1o estoque, muda a data
        If (bComSldIni Or bComSldIni1 Or bComSldIni2) And objEstFilial.dtDataInicial <> tEst.dtDataInicial Then
            
            lErro = Comando_Executar(alComando(2), "UPDATE EstoqueProduto SET DataInicial = ? WHERE Produto = ? AND Almoxarifado = ?", objEstFilial.dtDataInicial, tEst.sProduto, tEst.iAlmoxarifado)
            If lErro <> SUCESSO Then gError 216080
                    
        End If
        
        Set objEstoqueMes = New ClassEstoqueMes
        Set objEstoqueProduto = New ClassEstoqueProduto
        
        objEstoqueMes.iFilialEmpresa = iFilialEmp
        objEstoqueMes.iAno = Year(objEstFilial.dtDataInicial) - 1
        
        objEstoqueProduto.sProduto = tEst.sProduto
        objEstoqueProduto.iAlmoxarifado = tEst.iAlmoxarifado
        objEstoqueProduto.iFilialEmpresa = iFilialEmp
        
        'Caso tenha registros faltando pela antecipação do estoque inicial -> Acerta
        lErro = CF("EstoqueInicial_Gera_Info_Anos_Abertos", objEstoqueMes, objEstoqueProduto)
        If lErro <> SUCESSO Then gError 92941

        If bComSldIni Then
            lErro = Comando_Executar(alComando(3), "UPDATE SldMesEstAlm SET QuantInicial = ?, ValorInicial = ? WHERE Almoxarifado=? AND Produto=? AND Ano=? ", dSldMesQtdIni, dSldMesVlrIni, tEst.iAlmoxarifado, tEst.sProduto, Year(objEstFilial.dtDataInicial))
            If lErro <> SUCESSO Then gError 216081
        End If
    
        If bComSldIni1 Then
            lErro = Comando_Executar(alComando(4), "UPDATE SldMesEstAlm1 SET QuantInicialConsig3=?,ValorInicialConsig3=?,QuantInicialDemo3=?,ValorInicialDemo3=?,QuantInicialConserto3=?,ValorInicialConserto3=?,QuantInicialOutros3=?,ValorInicialOutros3=?,QuantInicialBenef3=? WHERE Almoxarifado=? AND Produto=? AND Ano=? ", _
                                                    tEst.dQuantInicialConsig3, tEst.dValorInicialConsig3, tEst.dQuantInicialDemo3, tEst.dValorInicialDemo3, tEst.dQuantInicialConserto3, tEst.dValorInicialConserto3, tEst.dQuantInicialOutras3, tEst.dValorInicialOutras3, tEst.dQuantInicialBenef3, tEst.iAlmoxarifado, tEst.sProduto, Year(objEstFilial.dtDataInicial))
            If lErro <> SUCESSO Then gError 216082
        End If
        
        If bComSldIni2 Then
            lErro = Comando_Executar(alComando(5), "UPDATE SldMesEstAlm2 SET QuantInicialConsig=?,ValorInicialConsig=?,QuantInicialDemo=?,ValorInicialDemo=?,QuantInicialConserto=?,ValorInicialConserto=?,QuantInicialOutros=?,ValorInicialOutros=?,QuantInicialBenef=? WHERE Almoxarifado=? AND Produto=? AND Ano=? ", _
                                                    tEst.dQuantInicialConsig, tEst.dValorInicialConsig, tEst.dQuantInicialDemo, tEst.dValorInicialDemo, tEst.dQuantInicialConserto, tEst.dValorInicialConserto, tEst.dQuantInicialOutras, tEst.dValorInicialOutras, tEst.dQuantInicialBenef, tEst.iAlmoxarifado, tEst.sProduto, Year(objEstFilial.dtDataInicial))
            If lErro <> SUCESSO Then gError 216083
        End If
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216084
        
        If lErro = AD_SQL_SEM_DADOS Then bForcado = True
    
    Loop

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
       
    Rotina_Reproc_Calc_SldIni = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Calc_SldIni:
    
    Rotina_Reproc_Calc_SldIni = gErr
    
    Select Case gErr
    
        Case 216073
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 216074 To 216084, 216086 To 216091
            Call Rotina_Erro(vbOKOnly, "ERRO_ZERAR_SLDMESEST", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 216085)
            
    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function EstoqueInicial_Gera_Zerado(ByVal objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 1) As Long, iAlmoxarifadoPadrao As Integer
Dim objItemMovEstoque As ClassItemMovEstoque, iAlmoxarifado As Integer
Dim sAlmoxarifadoNomeRed As String, colRastreamento As New Collection
Dim objEstoqueProduto As ClassEstoqueProduto, iFilialEmpresa1 As Integer

On Error GoTo Erro_EstoqueInicial_Gera_Zerado

    iAlmoxarifadoPadrao = MARCADO

    'Inicia o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 216159
    Next
    
    iFilialEmpresa1 = objMovEstoque.iFilialEmpresa
    lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa1)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
    
    lErro = Comando_Executar(alComando(0), "SELECT Codigo, NomeReduzido FROM Almoxarifado WHERE FilialEmpresa = ? ORDER BY Codigo", _
        iAlmoxarifado, sAlmoxarifadoNomeRed, iFilialEmpresa1)
    If lErro <> AD_SQL_SUCESSO Then gError 216160

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216161

    'Nenhum almoxarifado cadastrado para filial
    If lErro <> SUCESSO Then gError 216162
    
    For Each objItemMovEstoque In objMovEstoque.colItens
    
        'Controla estoque mas não tem o estoque cadastrado
        If objItemMovEstoque.iAlmoxarifado = 0 And objItemMovEstoque.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                  
            Set objEstoqueProduto = New ClassEstoqueProduto
            objEstoqueProduto.dtDataInicial = CDate("01/01/" & CStr(Year(objMovEstoque.dtData)))
            objEstoqueProduto.iFilialEmpresa = iFilialEmpresa1
            objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
            objEstoqueProduto.sProduto = objItemMovEstoque.sProduto
            objEstoqueProduto.dQuantidadeInicial = 0
            
            'Chama EstoqueInicial_Grava
            lErro = CF("EstoqueInicial_Grava_Trans", objEstoqueProduto, iAlmoxarifadoPadrao, colRastreamento)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            objItemMovEstoque.iAlmoxarifado = iAlmoxarifado
            objItemMovEstoque.sAlmoxarifadoNomeRed = sAlmoxarifadoNomeRed
            
        End If
        
    Next
        
    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    EstoqueInicial_Gera_Zerado = SUCESSO

    Exit Function

Erro_EstoqueInicial_Gera_Zerado:

    EstoqueInicial_Gera_Zerado = gErr
   
    Select Case gErr

        Case 216159
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
           
        Case 216160, 216161
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO2", gErr)
           
        Case 216162
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ALMOXARIFADO_BD", gErr)
           
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216163)

    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function EstoqueMes_Trata_Abertura() As Long

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes
Dim iMesAtual As Integer, iAnoAtual As Integer
Dim iMesAux As Integer, iAnoAux As Integer
Dim bDefasado As Boolean, vbResult As VbMsgBoxResult, bAvisoDado As Boolean

On Error GoTo Erro_EstoqueMes_Trata_Abertura

    bDefasado = True
    bAvisoDado = False

    If giFilialEmpresa <> EMPRESA_TODA And gobjEST.iTratamentoAberturaMesIni <> 0 Then
    
        iAnoAtual = Year(Date)
        iMesAtual = Month(Date)
        
        '0=Não faz
        '1=Pergunta se é para fazer a abertura
        '2=Faz sem avisar
        If gobjEST.iTratamentoAberturaMesIni = 2 Then bAvisoDado = True
        
        Do While bDefasado
        
            bDefasado = False
            
            Set objEstoqueMes = New ClassEstoqueMes
            
            objEstoqueMes.iFilialEmpresa = giFilialEmpresa
            objEstoqueMes.iFechamento = ESTOQUEMES_FECHAMENTO_ABERTO
    
            'Ler o mês e o ano que esta aberto passando como parametro filialEmpresa  e Fechamento
            lErro = CF("EstoqueMes_Le_Mes1", objEstoqueMes)
            If lErro <> SUCESSO And lErro <> 60861 Then gError ERRO_SEM_MENSAGEM
            
            'Se continua na mesma sai do loop pois não fechou por conta de um erro anterior
            If iAnoAux = objEstoqueMes.iAno And iMesAux = objEstoqueMes.iMes Then Exit Do
        
            If (objEstoqueMes.iAno < iAnoAtual) Or (objEstoqueMes.iAno = iAnoAtual And objEstoqueMes.iMes < iMesAtual) Then bDefasado = True
            
            If Not bDefasado Then Exit Do
            
            If Not bAvisoDado Then
            
                bAvisoDado = True
            
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_ABERTURA_MES_ESTOQUE")
                If vbResult = vbNo Then Exit Do
                
            End If
            
            iAnoAux = objEstoqueMes.iAno
            iMesAux = objEstoqueMes.iMes
                        
            'chama a tela FechamentoMesEst1
            Call Chama_Tela_Modal("FechamentoMesEst1", objEstoqueMes, True)
             
         Loop
    
    End If

    EstoqueMes_Trata_Abertura = SUCESSO

    Exit Function

Erro_EstoqueMes_Trata_Abertura:

    EstoqueMes_Trata_Abertura = gErr
   
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216163)

    End Select


    Exit Function

End Function
