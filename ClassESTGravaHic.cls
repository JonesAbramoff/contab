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

'Function NFiscalFatEntrada_Grava_BD(ByVal objNFiscal As ClassNFiscal) As Long
''Insere uma Nova Nota Fiscal Fatura de Entrada no Banco de Dados
'
'Dim alComando(0 To 6) As Long
'Dim lNumIntDoc As Long
'Dim lErro As Long
'Dim iClasseDocCPR As Integer
'Dim lNumIntDocCPR As Long
'Dim objItemNF As ClassItemNF
'Dim objMovEstoque As ClassMovEstoque
'Dim objMovEstoqueBaixa As ClassMovEstoque
'Dim iIndice As Integer
'Dim iTipoMovtoEstoque As Integer
'Dim lCodigo As Long
'Dim sDocOrigem As String
'Dim iTipoMovtoEstoqueBenef As Integer, objTipoDocInfo As New ClassTipoDocInfo
'Dim objItemMovEstoque As ClassItemMovEstoque
'Dim objRastroItemNF As ClassRastroItemNF
'Dim objRastroMovto As ClassRastreamentoMovto
'Dim colInsumos As Collection
'Dim colRastroMovto As Collection
'Dim dFator As Double
'Dim dSomaAlocacao As Double
'Dim objItemNFAloc As ClassItemNFAlocacao
'Dim dQuantItemNFUMEst As Double
'Dim objItemRomaneio As ClassItemRomaneioGrade
'Dim objReservaItem As ClassReservaItem
'Dim iNumCasasDec As Integer
'Dim dAcrescimo As Double
'Dim objNFiscalHic As New ClassNFiscalHic 'Inserido por Wagner
'
'On Error GoTo Erro_NFiscalFatEntrada_Grava_BD
'
'    'Abre os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then Error 42709
'    Next
'
'    '#################################
'    'Inserido por Wagner
'    If Not (objNFiscal.objInfoUsu Is Nothing) Then
'        Set objNFiscalHic = objNFiscal.objInfoUsu
'    End If
'    '#################################
'
'    'Busca iClasseDocCPR e lNumIntDocCPR
'    lErro = CF("CPR_Gera", objNFiscal, iClasseDocCPR, lNumIntDocCPR)
'    If lErro <> SUCESSO Then Error 42710
'
'    objNFiscal.iClasseDocCPR = iClasseDocCPR
'    objNFiscal.lNumIntDocCPR = lNumIntDocCPR
'
'    sDocOrigem = String(STRING_TIPO_DOC_INFO_SIGLA, 0)
'
'    'Lê TipoMovtoEstoque em TiposDocInfo
'
'    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
'    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
'    If lErro <> SUCESSO Then Error 59260
'
'    iTipoMovtoEstoque = objTipoDocInfo.iTipoMovtoEstoque
'    sDocOrigem = objTipoDocInfo.sSigla
'    iTipoMovtoEstoqueBenef = objTipoDocInfo.iTipoMovtoEstoqueBenef
'
'    'Obtem o Número Interno da nova Nota Fiscal
'    lErro = CF("NFiscal_Automatico", lNumIntDoc)
'    If lErro <> SUCESSO Then Error 42714
'
'    objNFiscal.lNumIntDoc = lNumIntDoc
'
'    With objNFiscal
'
'        .iNFe = SerieEletronica(.sSerie)
'
'        '###############################
'        'Alterado por Wagner
'        'Insere a nova Nota Fiscal de Entrada no BD
'        lErro = Comando_Executar(alComando(0), "INSERT INTO NFiscal (DataAtualizacao, NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, FornecedorBenef, FilialFornBenef, VolumeNumero, FreteRespons,DataReferencia,Observacao, NumRecebimento, HoraEntrada, Moeda, TaxaMoeda, Historico, SubConta, NFe) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            gdtDataHoje, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lFornecedorBenef, .iFilialFornBenef, .sVolumeNumero, .iFreteRespons, .dtDataReferencia, .sObservacao, .lNumRecebimento, CDbl(.dtHoraEntrada), .iMoeda, .dTaxaMoeda, objNFiscalHic.lHistorico, objNFiscalHic.lSubConta, .iNFe)
'        If lErro <> AD_SQL_SUCESSO Then Error 42715
'        '###############################
'    End With
'
'    lErro = CF("NFiscal_GravaTributacao", objNFiscal, alComando(4), alComando(5))
'    If lErro <> SUCESSO Then Error 42716
'
'    'Para cada Ítem da Nota
'    For Each objItemNF In objNFiscal.ColItensNF
'
'        objItemNF.lNumIntNF = objNFiscal.lNumIntDoc
'
'        'Obtém o Número Interno do Ítem
'        lErro = CF("NFiscalItem_Automatico", lNumIntDoc)
'        If lErro <> SUCESSO Then Error 42717
'
'        objItemNF.lNumIntDoc = lNumIntDoc
'
'        With objItemNF
'            'Insere o Ítem no Banco de dados
'            lErro = Comando_Executar(alComando(1), "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, NumIntDocOrig,Almoxarifado,Ccl) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lNumIntNF, .iItem, STATUS_LANCADO, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, lNumIntDoc, .lNumIntTrib, 0, .iAlmoxarifado, .sCcl)
'            If lErro <> AD_SQL_SUCESSO Then Error 42718
'        End With
'
'        lErro = CF("ItensNFGrade_Inclui", objItemNF)
'        If lErro <> SUCESSO Then Error 30739
'
'        lErro = CF("NFiscal_GravaTributacaoItem", objItemNF, alComando(3), alComando(6))
'        If lErro <> SUCESSO Then Error 42719
'
'    Next
'
'    'Faz a baixa de um recebimento
'    'a baixa do recebimento deve ocorrer antes da gravacao da nota fiscal para que os valores dos saldos estejam corretos e permitam calcular o novo custo do produto, se for o caso.
'
'    If objNFiscal.lNumRecebimento > 0 Then
'
'        'Lê o Recebimento e Verifica os Dados
'        lErro = CF("NFiscal_Valida_Recebimento", objNFiscal)
'        If lErro <> SUCESSO Then Error 64249
'
'        lErro = CF("NFiscalFatEntrada_Baixa_Recebimento", objNFiscal)
'        If lErro <> SUCESSO Then Error 42722
'
'    End If
'
'    If iTipoMovtoEstoque > 0 Then
'
'        Set objMovEstoque = New ClassMovEstoque
'
'        'Obtém o Código do Movimento
'        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
'        If lErro <> SUCESSO Then Error 42720
'
'        objMovEstoque.dtData = objNFiscal.dtDataEntrada
'        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
'        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
'        objMovEstoque.iTipoMov = iTipoMovtoEstoque
'        objMovEstoque.lCodigo = lCodigo
'        objMovEstoque.lCliente = objNFiscal.lCliente
'        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
'        objMovEstoque.sDocOrigem = sDocOrigem & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
'
'        'Se for para incluir o valor do frete e outras depesas no custo do ItemNF
'        If gobjEST.iIncluiFreteNoCusto = INCLUI_FRETE_CUSTO Then
'            Call CF("Inclui_Frete_CustoItemNF", objNFiscal)
'        End If
'
'
'
''mario distribuicao substituicao do codigo original acrescentando o tab de distribuicao
'
'        'Adiciona itens ao Movimento
'        For Each objItemNF In objNFiscal.ColItensNF
'
'            iIndice = iIndice + 1
'
'            If objItemNF.colItensRomaneioGrade.Count > 0 Then
'
'                For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
'
'                    If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
'
'                        lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemRomaneio.sUMEstoque, dFator)
'                        If lErro <> SUCESSO Then gError 89464
'
'                        dQuantItemNFUMEst = objItemRomaneio.dQuantidade * dFator
'
'                        'calcula o custo do item da NF em função do credito do ICMS e do IPI
'                        lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
'                        If lErro <> SUCESSO Then gError 60787
'
'                        dSomaAlocacao = 0
'
'                        For Each objReservaItem In objItemRomaneio.colLocalizacao
'
'                            dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
'
'                            Set colInsumos = New Collection
'                            Set colRastroMovto = New Collection
'
'                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, ((objItemRomaneio.dQuantidade * objItemNF.dPrecoUnitario) - (objItemNF.dValorDesconto * (objItemRomaneio.dQuantAFaturar / objItemNF.dQuantidade))) * objItemRomaneio.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objItemRomaneio.sUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
''?? Rastro
'
'                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
'                            objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
'                            objItemMovEstoque.iControleEstoque = objItemRomaneio.iControleEstoque
'                            objItemMovEstoque.sSiglaUMEst = objItemRomaneio.sUMEstoque
'''''                            objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
'
'                            If CF("TipoDocInfo_ZeraQtdeMovto", objNFiscal.iTipoNFiscal) Then
'
'                                objItemMovEstoque.dQuantidade = 0
'
'                            End If
'
'                        Next
'
'                        'Calcula o número de casas decimais do Formato de Estoque
'                        iNumCasasDec = Len(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                        dAcrescimo = 0
'                        If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                            'alteracao por tulio070303...
'                        If Abs(dSomaAlocacao - StrParaDbl(Formata_Estoque(dQuantItemNFUMEst))) > dAcrescimo + QTDE_ESTOQUE_DELTA Then
'
'                            If dSomaAlocacao < dQuantItemNFUMEst Then gError 89536
'
'                            If dSomaAlocacao > dQuantItemNFUMEst Then gError 89537
'
'                        End If
'
'                    End If
'                Next
'            Else
'
'                If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
'
'                    If objItemNF.ColAlocacoes.Count = 0 Then gError 89535
'
'                    lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
'                    If lErro <> SUCESSO Then gError 89539
'
'                    dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
'
'                    'calcula o custo do item da NF em função do credito do ICMS e do IPI
'                    lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
'                    If lErro <> SUCESSO Then gError 60787
'
'                    dSomaAlocacao = 0
'
'                    For Each objItemNFAloc In objItemNF.ColAlocacoes
'
'                        dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
'
'                        Set colInsumos = New Collection
'                        Set colRastroMovto = New Collection
'
'                        Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
'
'                        'Guarda o Rastreamento dos ItensNF
'                        If Not (objNFiscal.objRastreamento Is Nothing) Then
'                            lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
'                            If lErro <> SUCESSO Then gError 89538
'                        End If
'
'                        objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'                        objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
'                        objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
'                        objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
'                        objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
'
'                        'se for uma nota de complemento de preço ou outra que nao deve afetar qtde ==> zera a quantidade pois somente o custo será utilizado
'    '                    If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEFS) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCS) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFICM) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFIPI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEFTEL) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFELUZ) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCS) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFICM) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFIPI) Or (objNFiscal.iTipoNFiscal = DOCINFO_FDDAI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFFEESAC) Then
'                        If CF("TipoDocInfo_ZeraQtdeMovto", objNFiscal.iTipoNFiscal) Then
'
'                            objItemMovEstoque.dQuantidade = 0
'
'                        End If
'
'                    Next
'
'                    'Calcula o número de casas decimais do Formato de Estoque
'                    iNumCasasDec = Len(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                    dAcrescimo = 0
'                    If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                    If Abs(dSomaAlocacao - dQuantItemNFUMEst) > dAcrescimo + QTDE_ESTOQUE_DELTA Then
'
'                        If dSomaAlocacao < dQuantItemNFUMEst Then gError 89536
'
'                        If dSomaAlocacao > dQuantItemNFUMEst Then gError 89537
'
'                    End If
'                End If
'            End If
'        Next
'
'
''mario distribuicao fim da alteração para incluir o tab de distribuicao
'
'
''        'Adiciona os ítem ao Movimento
''        For Each objItemNF In objNFiscal.colItensNF
''
''            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
''
''                'calcula o custo do item da NF em função do credito do ICMS e do IPI
''                lErro = CF("NFiscalEntrada_Calcula_Custo",objItemNF)
''                If lErro <> SUCESSO Then Error 60788
''
''                Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNF.sUnidadeMed, objItemNF.dQuantidade, objItemNF.iAlmoxarifado, objItemNF.sAlmoxarifadoNomeRed, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
''                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
''
''                Set objItemMovEstoque.colRastreamentoMovto = New Collection
''                Set objItemMovEstoque.colApropriacaoInsumo = New Collection
''
''                'Guarda o Rastreamento dos ItensNF
''                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
''                If lErro <> SUCESSO Then gError 83573
''
'''                'Guarda o Rastreamento dos ItensNF
'''                For Each objRastroItemNF In objItemNF.colRastreamento
'''
'''                    Set objRastroMovto = New ClassRastreamentoMovto
'''                    objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
'''                    objRastroMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
'''                    objRastroMovto.sLote = objRastroItemNF.sLote
'''                    objRastroMovto.lNumIntDocOrigem = objItemNF.lNumIntDoc
'''                    objRastroMovto.sProduto = objItemNF.sProduto
'''                    objRastroMovto.iFilialOP = objRastroItemNF.iLoteFilialOP
'''
'''                    'Adiciona objRastroMovto na coleção de Rastreamento
'''                    objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
'''
'''                Next
''
''                objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
''                objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
''                objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
''                objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
''
''                'se for uma nota de complemento de preço ou outra que nao deve afetar qtde ==> zera a quantidade pois somente o custo será utilizado
''                If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCP) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCPCO) Then
''                    objItemMovEstoque.dQuantidade = 0
''                End If
''
''            End If
''
''        Next
'
'        'Gera um Movimento de Estoque
'        lErro = CF("MovEstoque_Grava", objMovEstoque)
'        If lErro <> SUCESSO Then Error 42721
'
'    End If
'
'    'se for entrada de material
'    If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFBEN) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFBEN) Then
'
'        'Gera um Movimento de Estoque
'        lErro = CF("NFiscalEntrada_Grava_MovimentosEst", objNFiscal, iTipoMovtoEstoqueBenef, sDocOrigem)
'        If lErro <> SUCESSO Then Error 55453
'
'    End If
'
'
'    'Fecha os Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    NFiscalFatEntrada_Grava_BD = SUCESSO
'
'    Exit Function
'
'Erro_NFiscalFatEntrada_Grava_BD:
'
'    NFiscalFatEntrada_Grava_BD = gErr
'
'    Select Case gErr
'
'        Case 30739, 42710, 42717, 42714, 42721, 42722, 42720, 42716, 42719, 55453, _
'            59260, 60787, 60788, 64249, 83573, 89464, 89538, 89539
'
'        Case 42709
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 42715
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
'
'        Case 42718
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", gErr, objNFiscal.lNumNotaFiscal)
'
'        Case 42711, 42712
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO", gErr, objNFiscal.iTipoNFiscal)
'
'        Case 42713
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)
'
'        Case 89535
'            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INEXISTENTE", gErr, iIndice)
'
'        Case 89536
'            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INCOMPLETA", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)
'
'        Case 89537
'            Call Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179176)
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
'Function NFiscalFatEntradaCom_Grava_BD(ByVal objNFiscal As ClassNFiscal, ByVal colPedCompras As Collection, ByVal colReqCompras As Collection) As Long
''Insere uma Nova Nota Fiscal de Entrada no Banco de Dados
'
'Dim alComando(0 To 6) As Long
'Dim lNumIntDoc As Long
'Dim lErro As Long
'Dim iClasseDocCPR As Integer
'Dim lNumIntDocCPR As Long
'Dim objItemNF As ClassItemNF
'Dim objMovEstoque As ClassMovEstoque
'Dim objMovEstoqueBaixa As ClassMovEstoque
'Dim iIndice As Integer
'Dim iTipoMovtoEstoque As Integer
'Dim lCodigo As Long
'Dim sDocOrigem As String
'Dim iTipoMovtoEstoqueBenef As Integer, objTipoDocInfo As New ClassTipoDocInfo
'Dim objItemMovEstoque As ClassItemMovEstoque
'Dim objRastroItemNF As ClassRastroItemNF
'Dim objRastroMovto As ClassRastreamentoMovto
'Dim colInsumos As Collection
'Dim colRastroMovto As Collection
'Dim dFator As Double
'Dim dSomaAlocacao As Double
'Dim objItemNFAloc As ClassItemNFAlocacao
'Dim dQuantItemNFUMEst As Double
'Dim iNumCasasDec As Integer
'Dim dAcrescimo As Double
'Dim objNFiscalHic As ClassNFiscalHic 'Inserido por Wagner
'
'On Error GoTo Erro_NFiscalFatEntradaCom_Grava_BD
'
'    'Abre os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 66547
'    Next
'
'    '#################################
'    'Inserido por Wagner
'    If Not (objNFiscal.objInfoUsu Is Nothing) Then
'        Set objNFiscalHic = objNFiscal.objInfoUsu
'    End If
'    '#################################
'
'    sDocOrigem = String(STRING_TIPO_DOC_INFO_SIGLA, 0)
'
'    'Lê TipoMovtoEstoque em TiposDocInfo
'    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
'    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
'    If lErro <> SUCESSO Then gError 66548
'
'    iTipoMovtoEstoque = objTipoDocInfo.iTipoMovtoEstoque
'    sDocOrigem = objTipoDocInfo.sSigla
'
'    'Gera Conta a Pagar e devolve iClasseDocCPR e lNumIntDocCPR
'    lErro = CF("CPR_Gera", objNFiscal, iClasseDocCPR, lNumIntDocCPR)
'    If lErro <> SUCESSO Then gError 66549
'
'    objNFiscal.iClasseDocCPR = iClasseDocCPR
'    objNFiscal.lNumIntDocCPR = lNumIntDocCPR
'
'    'Obtem o Número Interno da nova Nota Fiscal
'    lErro = CF("NFiscal_Automatico", lNumIntDoc)
'    If lErro <> SUCESSO Then gError 66550
'
'    objNFiscal.lNumIntDoc = lNumIntDoc
'
'    With objNFiscal
'
'        .iNFe = SerieEletronica(.sSerie)
'
'        'Insere a nova Nota Fiscal de Entrada no BD
'        '###############################
'        'Alterado por Wagner
'        lErro = Comando_Executar(alComando(0), "INSERT INTO NFiscal (DataAtualizacao, NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,FornecedorBenef,FilialFornBenef,VolumeNumero,FreteRespons,DataReferencia,Observacao,NumRecebimento, FilialPedido, HoraEntrada, Moeda, TaxaMoeda, Historico, SubConta, NFe) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            gdtDataHoje, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lFornecedorBenef, .iFilialFornBenef, .sVolumeNumero, .iFreteRespons, .dtDataReferencia, .sObservacao, .lNumRecebimento, .iFilialPedido, CDbl(.dtHoraEntrada), .iMoeda, .dTaxaMoeda, objNFiscalHic.lHistorico, objNFiscalHic.lSubConta, .iNFe)
'        If lErro <> AD_SQL_SUCESSO Then gError 66564
'        '###############################
'
'    End With
'
'    lErro = CF("NFiscal_GravaTributacao", objNFiscal, alComando(4), alComando(5))
'    If lErro <> SUCESSO Then gError 66551
'
'    'Para cada Ítem da Nota
'    For Each objItemNF In objNFiscal.ColItensNF
'
'        objItemNF.lNumIntNF = objNFiscal.lNumIntDoc
'
'        'Obtém o Número Interno do Ítem
'        lErro = CF("NFiscalItem_Automatico", lNumIntDoc)
'        If lErro <> SUCESSO Then gError 66552
'
'        objItemNF.lNumIntDoc = lNumIntDoc
'
'        With objItemNF
'            'Insere o Ítem no Banco de dados
'            lErro = Comando_Executar(alComando(1), "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, NumIntDocOrig,Almoxarifado, Ccl) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lNumIntNF, .iItem, STATUS_LANCADO, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, lNumIntDoc, .lNumIntTrib, 0, .iAlmoxarifado, .sCcl)
'            If lErro <> AD_SQL_SUCESSO Then gError 66565
'        End With
'
'       lErro = CF("NFiscal_GravaTributacaoItem", objItemNF, alComando(3), alComando(6))
'       If lErro <> SUCESSO Then gError 66553
'
'    Next
'
'    'Inclui em ItemNFItemPC o ItemNF associado a ItemPC
'    lErro = CF("ItemNFItemPC_Inclui", objNFiscal)
'    If lErro <> SUCESSO Then gError 66554
'
'    'Atualiza as Quantidades dos ItensPC
'    lErro = CF("ItensPedCompra_Atualiza_Quantidades", objNFiscal, colPedCompras)
'    If lErro <> SUCESSO Then gError 66076
'
'    'Inclui em ItemNFItemRC o ItemNF associado a ItemRC
'    lErro = CF("ItemNFItemRC_Inclui", objNFiscal)
'    If lErro <> SUCESSO Then gError 66555
'
'    'Atualiza as quantidades dos Itens de Requisição
'    lErro = CF("ItensReqCompra_Atualiza_Quantidades", objNFiscal)
'    If lErro <> SUCESSO Then Error 66562
'
'    'Faz a baixa de um recebimento
'    'a baixa do recebimento deve ocorrer antes da gravacao da nota fiscal para que os valores dos saldos estejam corretos e permitam calcular o novo custo do produto, se for o caso.
'    If objNFiscal.lNumRecebimento > 0 Then
'
'        'Lê o Recebimento e Verifica os Dados
'        lErro = CF("NFiscal_Valida_Recebimento", objNFiscal)
'        If lErro <> SUCESSO Then Error 66563
'
'        lErro = CF("NFiscalFatEntradaCom_Baixa_Recebimento", objNFiscal)
'        If lErro <> SUCESSO Then gError 66556
'
'    End If
'
'    If iTipoMovtoEstoque > 0 Then
'
'        Set objMovEstoque = New ClassMovEstoque
'
'        'Obtém o Código do Movimento
'        lErro = CF("MovEstoque_Automatico_EmTransacao", objNFiscal.iFilialEmpresa, lCodigo)
'        If lErro <> SUCESSO Then gError 66557
'
'        objMovEstoque.dtData = objNFiscal.dtDataEntrada
'        objMovEstoque.dtHora = objNFiscal.dtHoraEntrada
'        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
'        objMovEstoque.iTipoMov = iTipoMovtoEstoque
'        objMovEstoque.lCodigo = lCodigo
'        objMovEstoque.lCliente = objNFiscal.lCliente
'        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
'        objMovEstoque.sDocOrigem = sDocOrigem & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)
'
'        'Se for para incluir o valor do frete e outras depesas no custo do ItemNF
'        If gobjEST.iIncluiFreteNoCusto = INCLUI_FRETE_CUSTO Then
'            Call CF("Inclui_Frete_CustoItemNF", objNFiscal)
'        End If
'
''mario distribuicao substituicao do codigo original acrescentando o tab de distribuicao
'
'        'Adiciona itens ao Movimento
'        For Each objItemNF In objNFiscal.ColItensNF
'
'            iIndice = iIndice + 1
'
'            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
'
'                If objItemNF.ColAlocacoes.Count = 0 Then gError 89647
'
'                lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
'                If lErro <> SUCESSO Then gError 89648
'
'                dQuantItemNFUMEst = CDbl(Formata_Estoque(objItemNF.dQuantidade * dFator))
'
'                'calcula o custo do item da NF em função do credito do ICMS e do IPI
'                lErro = CF("NFiscalEntrada_Calcula_Custo", objItemNF)
'                If lErro <> SUCESSO Then gError 66588
'
'                dSomaAlocacao = 0
'
'                For Each objItemNFAloc In objItemNF.ColAlocacoes
'
'                    dSomaAlocacao = dSomaAlocacao + objItemNFAloc.dQuantidade
'
'                    Set colInsumos = New Collection
'                    Set colRastroMovto = New Collection
'
'                    Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto * objItemNFAloc.dQuantidade / dQuantItemNFUMEst, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
'
'                    'Guarda o Rastreamento dos ItensNF
'                    If Not (objNFiscal.objRastreamento Is Nothing) Then
'                        lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
'                        If lErro <> SUCESSO Then gError 83574
'                    End If
'
''mario
''                    If Not (objItemNF.colRastreamento Is Nothing) Then
''
''                        'Guarda o Rastreamento dos ItensNF
''                        For Each objRastroItemNF In objItemNF.colRastreamento
''
''                            If objRastroItemNF.iAlmoxCodigo = objItemNFAloc.iAlmoxarifado Then
''
''                                Set objRastroMovto = New ClassRastreamentoMovto
''                                objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
''                                objRastroMovto.sLote = objRastroItemNF.sLote
''                                objRastroMovto.sProduto = objItemNF.sProduto
''
''                                'Adiciona objRastroMovto na coleção de Rastreamento
''                                objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
''
''                            End If
''
''                        Next
''
''                    End If
''fim mario
'
'                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'                    objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
'                    objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
'                    objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
'                    objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
'
'                    'se for uma nota de complemento de preço ou outra que nao deve afetar qtde ==> zera a quantidade pois somente o custo será utilizado
'                    If (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFIPICO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFICMCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFCPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCPCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFICMCO) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFIPICO) Then
'                        objItemMovEstoque.dQuantidade = 0
'                    End If
'
'                Next
'
'                'Calcula o número de casas decimais do Formato de Estoque
'                iNumCasasDec = Len(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".") + 1)))
'
'                dAcrescimo = 0
'                If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                If Abs(dSomaAlocacao - dQuantItemNFUMEst) > dAcrescimo + QTDE_ESTOQUE_DELTA Then
'
'                    If dSomaAlocacao < dQuantItemNFUMEst Then gError 89649
'
'                    If dSomaAlocacao > dQuantItemNFUMEst Then gError 89650
'
'                End If
'
'            End If
'        Next
'
'
''mario distribuicao fim da alteração para incluir o tab de distribuicao
'
'
''        'Adiciona os ítem ao Movimento
''        For Each objItemNF In objNFiscal.colItensNF
''
''            If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
''
''                'calcula o custo do item da NF em função do credito do ICMS e do IPI
''                lErro = CF("NFiscalEntrada_Calcula_Custo",objItemNF)
''                If lErro <> SUCESSO Then gError 66558
''
''                Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dCusto, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNF.sUnidadeMed, objItemNF.dQuantidade, objItemNF.iAlmoxarifado, objItemNF.sAlmoxarifadoNomeRed, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA)
''                objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
''
''                Set objItemMovEstoque.colRastreamentoMovto = New Collection
''                Set objItemMovEstoque.colApropriacaoInsumo = New Collection
''
''                'Guarda o Rastreamento dos ItensNF
''                lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
''                If lErro <> SUCESSO Then gError 83574
''
'''                'Guarda o Rastreamento dos ItensNF
'''                For Each objRastroItemNF In objItemNF.colRastreamento
'''
'''                    Set objRastroMovto = New ClassRastreamentoMovto
'''                    objRastroMovto.dQuantidade = objRastroItemNF.dLoteQdtAlocada
'''                    objRastroMovto.iTipoDocOrigem = TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE
'''                    objRastroMovto.sLote = objRastroItemNF.sLote
'''                    objRastroMovto.lNumIntDocOrigem = objItemNF.lNumIntDoc
'''                    objRastroMovto.sProduto = objItemNF.sProduto
'''                    objRastroMovto.iFilialOP = objRastroItemNF.iLoteFilialOP
'''
'''                    'Adiciona objRastroMovto na coleção de Rastreamento
'''                    objItemMovEstoque.colRastreamentoMovto.Add objRastroMovto
'''
'''                Next
''
''                objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
''                objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
''                objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
''                objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
''
''            End If
''
''        Next
'
'        'Gera um Movimento de Estoque
'        lErro = CF("MovEstoque_Grava", objMovEstoque)
'        If lErro <> SUCESSO Then gError 66559
'
'    End If
'
'    'Verifica se existem Pedidos a serem Baixados
'    lErro = CF("NFiscalEntCom_BaixaPedidos", colPedCompras, objNFiscal)
'    If lErro <> SUCESSO Then gError 66560
'
'    'Verifica se existem Requisições a serem Baixadas
'    lErro = CF("NFiscalEntCom_BaixaRequisicoes", colReqCompras, objNFiscal)
'    If lErro <> SUCESSO Then gError 66561
'
'    'Fecha os Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    NFiscalFatEntradaCom_Grava_BD = SUCESSO
'
'    Exit Function
'
'Erro_NFiscalFatEntradaCom_Grava_BD:
'
'    NFiscalFatEntradaCom_Grava_BD = gErr
'
'    Select Case gErr
'
'        Case 66547
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 66548, 66549, 66550, 66551, 66552, 66553, 66554, 66555, 66556, 66557, 66558, 66559, 66560, 66561, 66562, 66563, 83574, 89648
'
'        Case 66564
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
'
'        Case 66565
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", gErr, objNFiscal.lNumNotaFiscal)
'
'        Case 89647
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INEXISTENTE", gErr, iIndice)
'
'        Case 89649
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_ITEM_INCOMPLETA", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)
'
'        Case 89650
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DISTRIBUICAO_MAIOR_QUE_ITEM", gErr, iIndice, dQuantItemNFUMEst, dSomaAlocacao)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179177)
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

Function NFiscalEntrada_Grava_CamposAlteraveis(ByVal objNFiscal As ClassNFiscal) As Long
'Atualiza os campos Historico e SubConta

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim lNumNotaFiscal As Long
Dim iFilialEmpresa As Integer
Dim iIndice As Integer
Dim sSerie As String
Dim objNFiscalHic As ClassNFiscalHic

On Error GoTo Erro_NFiscalEntrada_Grava_CamposAlteraveis

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 132356
    Next

    If Not (objNFiscal.objInfoUsu Is Nothing) Then

        Set objNFiscalHic = objNFiscal.objInfoUsu

        sSerie = String(STRING_SERIE, 0)
    
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Serie, NumNotaFiscal, FilialEmpresa FROM NFiscal WHERE NumIntDoc = ?", 0, sSerie, lNumNotaFiscal, iFilialEmpresa, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 132357
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132358
    
        If lErro = AD_SQL_SEM_DADOS Then gError 132359
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE NFiscal SET Historico = ?, SubConta = ?", alComando(1), objNFiscalHic.lHistorico, objNFiscalHic.lSubConta)
        If lErro <> AD_SQL_SUCESSO Then gError 132360
        
    End If
        
    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalEntrada_Grava_CamposAlteraveis = SUCESSO

    Exit Function

Erro_NFiscalEntrada_Grava_CamposAlteraveis:

    NFiscalEntrada_Grava_CamposAlteraveis = gErr

    Select Case gErr

        Case 132356
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 132357, 132358
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 132359
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 132360
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", gErr, sSerie, lNumNotaFiscal, iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179178)

    End Select

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function
