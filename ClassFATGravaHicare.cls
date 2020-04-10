VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Inserido por Wagner
'#############################
Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long
'##############################

Public Function MapaMensalVenda_Prepara(lNumIntRel As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal iMes As Integer, ByVal sPrevisao As String, ByVal sProd_I As String, ByVal sProd_F As String, ByVal bTodasCategorias As Boolean, ByVal sCategoria As String, ByVal sValorInicial As String, ByVal sValorFinal As String) As Long
'Realiza a seleção dos Produtos verificando o seu faturamento no mês passado e o Ano também passado
'Realiza a gravação dos Produtos selecionados na Tabela MapaMensalVendas

Dim lErro As Long, iIndice As Integer
Dim lTransacao As Long, alComando(1 To 7) As Long
Dim sProduto As String, dMetaQtde As Double, dMetaValor As Double
Dim dMediaAnterior As Double, dPrecoMedio As Double
Dim dQuantidade As Double, dQuantidadeAnterior As Double, dQuantidadePedida As Double, dQuantidadeFaturar As Double
Dim dPrecoUnitario As Double
Dim dValor As Double, dValorAnterior As Double, dValorFaturar As Double, dValorPedido As Double
Dim sMes As String, sMesAnterior As String
Dim dtDataEntregaInicio As Date, dtDataEntregaFim As Date
Dim iClasseUM As Integer, sUM As String, sUMPadrao As String, dFator As Double
Dim sSQL As String 'Inserido por Wagner
        
On Error GoTo Erro_MapaMensalVenda_Prepara

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 128162
    Next
    
    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 128163
 
    'Obtêm o NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_MAPAMENSALVENDA", lNumIntRel)
    If lErro <> SUCESSO Then gError 128176
        
    'Modifica o tipo do mês para string
    sMes = CStr(iMes)
    sMesAnterior = CStr(iMes - 1)
    If sMesAnterior = 0 Then sMesAnterior = 12 'Inserido por Wagner
    
    'Preenche o período de entrega
    If iMes <> 12 Then
        dtDataEntregaFim = CDate("01/" & (iMes + 1) & "/" & iAno) - 1
    Else
        dtDataEntregaFim = CDate("31/12/" & iAno)
    End If
        
    dtDataEntregaInicio = CDate("01/" & iMes & "/" & iAno)
    
    'Inicializa a variável
    sProduto = String(STRING_PRODUTO, 0)
    sUMPadrao = String(STRING_UM_SIGLA, 0)
        
    'Busca os Produtos, ClasseUM, SiglaUMVenda
    'Alterado por Wagner
    '#############################################
    
    Call MapaMensalVendaSQL_Prepara(sProd_I, sProd_F, bTodasCategorias, sCategoria, sValorInicial, sValorFinal, sSQL)
    
    lErro = MapaMensalVendaInt_Prepara(alComando(1), sProduto, iClasseUM, sUMPadrao, iFilialEmpresa, PRODUTO_VENDAVEL, sProd_I, sProd_F, bTodasCategorias, sCategoria, sValorInicial, sValorFinal, sSQL)
'    lErro = Comando_Executar(alComando(1), "SELECT Codigo, ClasseUM, SiglaUMVenda FROM Produtos, ProdutosFilial WHERE Produtos.Codigo = ProdutosFilial.Produto AND ProdutosFilial.FilialEmpresa = ? AND Faturamento = ? ORDER BY Produtos.Codigo", sProduto, iClasseUM, sUMPadrao, iFilialEmpresa, PRODUTO_VENDAVEL)
    If lErro <> SUCESSO Then gError 128164
    '#############################################

    'Busca o primeiro elemento da seleção
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128165
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = CF("Produto_ObtemPrevVenda", sPrevisao, iFilialEmpresa, iAno, iMes, sProduto, dMetaQtde, dMetaValor, alComando(7))
        If lErro <> SUCESSO Then gError 124265

        'Inicializa a string
        sUM = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        
        'Busca as Unidades de Medidas e a QuantidadePedida e o Valor à Faturar
        lErro = Comando_Executar(alComando(2), "SELECT ItensPedidoDeVenda.UnidadeMed, SUM(ItensPedidoDeVenda.Quantidade - ItensPedidoDeVenda.QuantFaturada - ItensPedidoDeVenda.QuantCancelada), SUM((ItensPedidoDeVenda.Quantidade - ItensPedidoDeVenda.QuantFaturada - ItensPedidoDeVenda.QuantCancelada)*ItensPedidoDeVenda.PrecoUnitario) FROM ItensPedidoDeVenda WHERE ItensPedidoDeVenda.Produto = ? AND ItensPedidoDeVenda.Status = ? AND (ItensPedidoDeVenda.DataEntrega BETWEEN ? AND ? OR ItensPedidoDeVenda.DataEntrega = ?) GROUP BY ItensPedidoDeVenda.UnidadeMed", sUM, dQuantidadePedida, dValorPedido, sProduto, STATUS_NAO_ATENDIDO, dtDataEntregaInicio, dtDataEntregaFim, DATA_NULA)
        If lErro <> AD_SQL_SUCESSO Then gError 128166
    
        'Busca o primeiro elemento da seleção
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128167
        
        Do While lErro = AD_SQL_SUCESSO
        
            'Realiza a converção para uma mesma UM
            lErro = CF("UM_Conversao_Trans", iClasseUM, sUM, sUMPadrao, dFator)
            If lErro <> SUCESSO Then gError 128168
                        
            'Preenche a quantidade
            dQuantidadeFaturar = dQuantidadeFaturar + (dQuantidadePedida * dFator)
            dValorFaturar = dValorFaturar + dValorPedido
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128169
                
        Loop
        
        dValor = 0
        dQuantidade = 0
        dValorAnterior = 0
        dQuantidadeAnterior = 0
            
        'se o mês for o de Janeiro
        If iMes = 1 Then
    
            '-> Busca os valores para o mês de Janeiro do ano
            lErro = Comando_Executar(alComando(3), "SELECT ValorFaturado1, QuantFaturada1 FROM SldMesFat WHERE FilialEmpresa = ? AND Ano = ? AND Produto =?", dValor, dQuantidade, iFilialEmpresa, iAno, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 128170
    
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128171
            
            '-> Busca os valores para o mês de Dezembro do ano anterior
            lErro = Comando_Executar(alComando(4), "SELECT ValorFaturado12, QuantFaturada12 FROM SldMesFat WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ?", dValorAnterior, dQuantidadeAnterior, iFilialEmpresa, iAno - 1, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 128172
    
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128173
            
        Else
        
            '-> Busca as informações do mês passado e do mês anterior na Tabela SldMesFat
            lErro = Comando_Executar(alComando(5), "SELECT ValorFaturado" & sMesAnterior & ", QuantFaturada" & sMesAnterior & ", ValorFaturado" & sMes & ", QuantFaturada" & sMes & " FROM SldMesFat WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ?", dValorAnterior, dQuantidadeAnterior, dValor, dQuantidade, iFilialEmpresa, iAno, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 128174
        
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128175
        
        End If
                
        If dQuantidadeAnterior <> 0 Then
            dMediaAnterior = dValorAnterior / dQuantidadeAnterior
        Else
            dMediaAnterior = 0
        End If
        If dQuantidade <> 0 Then
            dPrecoMedio = dValor / dQuantidade
        Else
            dPrecoMedio = 0
        End If
    
        'Realiza a gravação na Tabela MapaMensalVendas
        lErro = Comando_Executar(alComando(6), "INSERT INTO MapaMensalVendas (NumIntRel, Produto, ValorAnterior, MediaAnterior, Quantidade, Valor, PrecoMedio, QuantidadeFaturar, ValorFaturar, MetaQuantidade, MetaValor) VALUES (?,?,?,?,?,?,?,?,?,?,?)", lNumIntRel, sProduto, dValorAnterior, dMediaAnterior, dQuantidade, dValor, dPrecoMedio, dQuantidadeFaturar, dValorFaturar, dMetaQtde, dMetaValor)
        If lErro <> AD_SQL_SUCESSO Then gError 128177
        
        'Busca o Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128178
        
    Loop
    
    'Fecha a Transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 128179
    
    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    MapaMensalVenda_Prepara = SUCESSO

    Exit Function

Erro_MapaMensalVenda_Prepara:

    MapaMensalVenda_Prepara = gErr
    
    Select Case gErr
          
        Case 128162
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 128163
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 128164, 128165, 128170 To 128175, 128178
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDFATMES", gErr, iAno, iFilialEmpresa, sProduto)
          
        Case 128166, 128167, 128169
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)
        
        Case 128168, 128176, 124265
        
        Case 128177
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MAPAMENSALVENDAS", gErr, lNumIntRel, sProduto)
          
        Case 128179
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179179)
     
    End Select
    
    'Transação Rollback
    Transacao_Rollback
     
    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
     
    Exit Function

End Function

Public Function ContratoPropaganda_Grava(ByVal objContratoPropaganda As ClassContratoPropag) As Long
'Grava ou Atualiza o ContratoPropaganda no BD

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim dPercentual As Double

On Error GoTo Erro_GuiaICMS_Grava
 
    'abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = SUCESSO Then gError 128084
    Next
 
    'Abre a transação
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 128085
      
    'Verifica se o ContratoPropaganda já existe
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Percentual FROM ContratoPropaganda WHERE FilialEmpresa = ? AND Cliente = ? AND PeriodoDe = ? AND PeriodoAte = ?", 0, dPercentual, objContratoPropaganda.lCliente, objContratoPropaganda.iFilialEmpresa, objContratoPropaganda.dtPeriodoDe, objContratoPropaganda.dtPeriodoAte)
    If lErro <> AD_SQL_SUCESSO Then gError 128086
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128087
    
    If lErro = AD_SQL_SUCESSO Then
    
        'Realiza o lock
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 128088
        
        'Atualiza os dados de ContratoPropaganda
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ContratoPropaganda SET Percentual = ?", alComando(1), objContratoPropaganda.dPercentual)
        If lErro <> AD_SQL_SUCESSO Then gError 128089
        
    Else
    
        'Realiza a inserção do ContratoPropaganda
        lErro = Comando_Executar(alComando(3), "INSERT INTO ContratoPropaganda (Cliente, FilialEmpresa, PeriodoDe, PeriodoAte, Percentual) VALUES (?,?,?,?,?)", objContratoPropaganda.lCliente, objContratoPropaganda.iFilialEmpresa, objContratoPropaganda.dtPeriodoDe, objContratoPropaganda.dtPeriodoAte, objContratoPropaganda.dPercentual)
        If lErro <> AD_SQL_SUCESSO Then gError 128090
        
    End If
    
    'Encerra a Transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 128091
    
    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ContratoPropaganda_Grava = SUCESSO
    
    Exit Function
    
Erro_GuiaICMS_Grava:

    ContratoPropaganda_Grava = gErr
    
    Select Case gErr
        
        Case 128084
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 128085
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 128086, 128087
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTRATOPROPAGANDA", gErr, objContratoPropaganda.lCliente, objContratoPropaganda.dtPeriodoDe, objContratoPropaganda.dtPeriodoAte)
        
        Case 128088
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTRATOPROPAGANDA", gErr)
        
        Case 128089
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONTRATOPROPAGANDA", gErr, objContratoPropaganda.lCliente, objContratoPropaganda.dtPeriodoDe, objContratoPropaganda.dtPeriodoAte)
            
        Case 128090
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONTRATOPROPAGANDA", gErr, objContratoPropaganda.lCliente, objContratoPropaganda.dtPeriodoDe, objContratoPropaganda.dtPeriodoAte)
        
        Case 128091
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179180)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
        
End Function

Public Function ContratoPropaganda_Exclui(ByVal objContratoPropaganda As ClassContratoPropag) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 3)  As Long
Dim iIndice As Integer
Dim dPercentual As Double

On Error GoTo Erro_ContratoPropaganda_Exclui

    'abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = SUCESSO Then gError 128092
    Next

    'Abre a Transacao
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 128093
    
    'Verifica se o ContratoPropaganda já existe
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Percentual FROM ContratoPropaganda WHERE Cliente = ? AND FilialEmpresa = ? AND PeriodoDe = ? AND PeriodoAte = ?", 0, dPercentual, objContratoPropaganda.lCliente, objContratoPropaganda.iFilialEmpresa, objContratoPropaganda.dtPeriodoDe, objContratoPropaganda.dtPeriodoAte)
    If lErro <> AD_SQL_SUCESSO Then gError 128094
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128095
    
    'ContratoPropaganda não encontrado --> ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 128096
    
    'Realiza o Lock do Contrato
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 128097
    
    'Exclui o ContratoPropaganda
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ContratoPropaganda", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 128098
    
    'Feacha a Transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 128099
    
    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ContratoPropaganda_Exclui = SUCESSO
    
    Exit Function
    
Erro_ContratoPropaganda_Exclui:

    ContratoPropaganda_Exclui = gErr
    
    Select Case gErr
    
        Case 128092
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 128093
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 128094, 128095
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTRATOPROPAGANDA", gErr)
        
        Case 128096 'registro não encontrado
        
        Case 128097
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTRATOPROPAGANDA", gErr)
            
        Case 128098
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTRATOPROPAGANDA", gErr)
    
        Case 128099
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSACAO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179181)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
        
End Function

'Public Function NFiscal_Grava_NFiscal(alComando() As Long, ByVal objNFiscal As ClassNFiscal) As Long
''Grava a nota fiscal, seus itens e a tributacao respectiva
''@@@@@ customizacao: tabelas de itens e nf tem mais campos
'
'Dim lErro As Long
'Dim lNumIntDoc As Long
'Dim objItemNF As ClassItemNF
'
'On Error GoTo Erro_NFiscal_Grava_NFiscal
'
'    'Obtem o Número Interno da nova Nota Fiscal
'    lErro = CF("NFiscal_Automatico1", alComando(1), alComando(2), alComando(3), lNumIntDoc)
'    If lErro <> SUCESSO Then gError 35997
'
'    objNFiscal.lNumIntDoc = lNumIntDoc
'
'    lErro = CF("NFiscal_Grava_Msg", objNFiscal)
'    If lErro <> SUCESSO Then gError 35997
'
'    With objNFiscal
'
'        .iNFe = SerieEletronica(.sSerie)
'
'        If .objInfoUsu Is Nothing Then Set .objInfoUsu = New ClassNFiscalHic
'        'Insere a nova Nota Fiscal de Saida no BD
'        lErro = Comando_Executar(alComando(4), "INSERT INTO NFiscal (DataAtualizacao, NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, DataReferencia, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,FilialPedido, VolumeNumero, FreteRespons, HoraSaida, CodTranspRedesp, DetPagFrete, Moeda, TaxaMoeda, VolumeTotal, SemDataSaida, NaoImpCobranca, ValorFrete1,MensagemCorpoNota,RPS, NFe) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
'            gdtDataHoje, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .dtDataReferencia, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .iFilialPedido, .sVolumeNumero, .iFreteRespons, CDbl(.dtHoraSaida), .iCodTranspRedesp, .iDetPagFrete, .iMoeda, .dTaxaMoeda, .dVolumeTotal, .iSemDataSaida, .iNaoImpCobranca, .objInfoUsu.dValorFrete1, .sMensagemCorpoNota, .iRecibo, .iNFe)
'        If lErro <> AD_SQL_SUCESSO Then gError 35998
'
'    End With
'
'    'grava a tributacao da nota como um todo
'    lErro = CF("NFiscal_GravaTributacao", objNFiscal, alComando(5), alComando(6))
'    If lErro <> SUCESSO Then gError 39042
'
'    'Obtém o Número Interno dos itens da NF
'    lErro = CF("CRFATConfig_ObterNumInt", alComando(7), alComando(8), "NUM_PROX_ITEM_NOTA_FISCAL", lNumIntDoc, objNFiscal.ColItensNF.Count)
'    If lErro <> SUCESSO Then gError 35999
'
'    'Para cada item da Nota
'    For Each objItemNF In objNFiscal.ColItensNF
'
'        objItemNF.lNumIntNF = objNFiscal.lNumIntDoc
'        objItemNF.lNumIntDoc = lNumIntDoc
'
'        With objItemNF
'
'            'Insere o item no Banco de dados
'            lErro = Comando_Executar(alComando(9), "INSERT INTO ItensNFiscal (NumIntNF, Item, Status, Produto, UnidadeMed, Quantidade, PrecoUnitario, PercDesc, ValorDesconto, DataEntrega, DescricaoItem, ValorAbatComissao, NumIntPedVenda, NumIntItemPedVenda, NumIntDoc, NumIntTrib, NumIntDocOrig, Almoxarifado, PrecoUnitarioMoeda) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lNumIntNF, .iItem, STATUS_LANCADO, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, lNumIntDoc, .lNumIntTrib, 0, .iAlmoxarifado, .dPrecoUnitarioMoeda)
'            If lErro <> AD_SQL_SUCESSO Then gError 35212
'
'        End With
'
'        lErro = CF("ItensNFGrade_Inclui", objItemNF)
'        If lErro <> SUCESSO Then gError 86365
'
'        'grava a tributacao relativa a cada item da nota
'        lErro = CF("NFiscal_GravaTributacaoItem", objItemNF, alComando(10), alComando(11))
'        If lErro <> SUCESSO Then gError 39043
'
'        lNumIntDoc = lNumIntDoc + 1
'
'    Next
'
'    NFiscal_Grava_NFiscal = SUCESSO
'
'    Exit Function
'
'Erro_NFiscal_Grava_NFiscal:
'
'    NFiscal_Grava_NFiscal = gErr
'
'    Select Case gErr
'
'        Case 35212
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSNFISCAL", gErr, objNFiscal.lNumNotaFiscal)
'
'        Case 35997, 35999, 39042, 39043, 86365
'
'        Case 35998
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NFISCAL_SAIDA", gErr, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179182)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function RelRentabilidadeCli_Prepara(lNumIntRel As Long, ByVal lClienteDe As Long, ByVal lClienteAte As Long, ByVal iMes As Integer, ByVal iAno As Integer, ByVal iRegiaoDe As Integer, ByVal iRegiaoAte As Integer, ByVal iFilialEmpresa As Integer) As Long
'Realiza a leitura do GastoContrato, ValorComissão, ValoFaturado e do CustoReposição
'através das informações passadas como parâmetro
'Preenche a tabela RentabilidadeCli

Dim lErro As Long
Dim lTransacao As Long, alComando(1 To 7) As Long
Dim dValorComissao As Double, dValorFaturado As Double, dValorFaturadoTotal As Double
Dim dCustoReposicao As Double, dCustoReposicaoTotal As Double
Dim dtDataInicio As Date, dtDataFim As Date
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim lCliente As Long
Dim sMes As String
Dim sProduto As String
Dim dPercentual As Double

On Error GoTo Erro_RelRentabilidadeCli_Prepara
    
    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 128180

    Next

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 128181
    
    dtDataInicio = CDate("01/" & iMes & "/" & iAno)
    'Alterado por Wagner
    dtDataFim = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & iMes & "/" & iAno)))
    
    'Modifica o tipo do mês
    sMes = CStr(iMes)
    
    'Obtêm o NumIntRel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_RENTABILIDADECLI", lNumIntRel)
    If lErro <> SUCESSO Then gError 128182

    'Faz a seleção do Código do Cliente
    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Clientes WHERE Codigo BETWEEN ? AND ? ORDER BY Codigo", lCliente, lClienteDe, lClienteAte)
    If lErro <> AD_SQL_SUCESSO Then gError 128183

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128184

    Do While lErro = AD_SQL_SUCESSO
            
        'Inicializa as variáveis
        sProduto = String(STRING_PRODUTO, 0)
        
        dValorFaturadoTotal = 0
        dCustoReposicaoTotal = 0
    
        'Realiza o somatório da Comissão
        lErro = Comando_Executar(alComando(2), "SELECT SUM(ComissoesNF.Valor) FROM ComissoesNF, FiliaisClientes, NFiscal WHERE NFiscal.Cliente = FiliaisClientes.CodCliente AND NFiscal.FilialCli = FiliaisClientes.CodFilial AND ComissoesNF.CodVendedor = FiliaisClientes.Vendedor AND ComissoesNF.NumIntDoc = NFiscal.NumIntDoc AND NFiscal.FilialEmpresa = ? AND FiliaisClientes.Nome = ? AND FiliaisClientes.CodCliente = ? AND NFiscal.DataEmissao BETWEEN ? AND ? AND FiliaisClientes.Regiao BETWEEN ? AND ?", dValorComissao, iFilialEmpresa, MATRIZ, lCliente, dtDataInicio, dtDataFim, iRegiaoDe, iRegiaoAte)
        If lErro <> AD_SQL_SUCESSO Then gError 128185
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128186
    
        'Busca o percentual na tabela ContratoPropaganda
        lErro = Comando_Executar(alComando(3), "SELECT Percentual FROM ContratoPropaganda WHERE FilialEmpresa = ? AND Cliente = ? AND PeriodoDe > ? AND PeriodoAte < ?", dPercentual, iFilialEmpresa, lCliente, dtDataInicio, dtDataFim)
        If lErro <> AD_SQL_SUCESSO Then gError 128187
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128188
        
        'Realiza a seleção de Produtos
        lErro = Comando_Executar(alComando(4), "SELECT ItensNFiscal.Produto FROM NFiscal, ItensNFiscal, FiliaisClientes WHERE NFiscal.Cliente = FiliaisClientes.CodCliente AND NFiscal.FilialCli = FiliaisClientes.CodFilial AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.FilialEmpresa = ? AND FiliaisClientes.Nome = ? AND FiliaisClientes.CodCliente = ?  AND NFiscal.DataEmissao BETWEEN ? AND ? AND FiliaisClientes.Regiao BETWEEN ? AND ? GROUP BY ItensNFiscal.Produto ORDER BY ItensNFiscal.Produto", sProduto, iFilialEmpresa, MATRIZ, lCliente, dtDataInicio, dtDataFim, iRegiaoDe, iRegiaoAte)
        If lErro <> AD_SQL_SUCESSO Then gError 128189
        
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128190
        
        Do While lErro = AD_SQL_SUCESSO
        
            'Busca o ValorFaturado do mês passado
            lErro = Comando_Executar(alComando(5), "SELECT ValorFaturado" & sMes & " FROM SldMesFat WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ?", dValorFaturado, iFilialEmpresa, iAno, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 128191
            
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128192
            
            'Busca o CustoReposição do mês passado
            lErro = Comando_Executar(alComando(6), "SELECT CustoReposicao" & sMes & " FROM ProdutoMes WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ?", dCustoReposicao, iFilialEmpresa, iAno, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 128193
            
            lErro = Comando_BuscarPrimeiro(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128194
            
            'Realiza o somatório do ValorFaturado e do CustoReposição
            dValorFaturadoTotal = dValorFaturadoTotal + dValorFaturado
            dCustoReposicaoTotal = dCustoReposicaoTotal + dCustoReposicao
            
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128195
            
        Loop
        
        'Realiza a gravação na Tabela RentabilidadeCli
        lErro = Comando_Executar(alComando(7), "INSERT INTO RelRentabilidadeCli (NumIntRel, Cliente, ValorFaturado, Frete, GastosFinanceiro, GastosContrato, GastosComissao, GastosPropaganda, CustoReposicao, Rentabilidade, Lucro) VALUES (?,?,?,?,?,?,?,?,?,?,?)", lNumIntRel, lCliente, dValorFaturadoTotal, 0, 0, 0, dValorComissao, 0, dCustoReposicaoTotal, 0, 0)
        If lErro <> AD_SQL_SUCESSO Then gError 128196
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 128197

    Loop
    
    'Fecha a Transação
    lErro = Transacao_Commit
    If lErro <> AD_SQL_SUCESSO Then gError 128198

    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelRentabilidadeCli_Prepara = SUCESSO

    Exit Function
    
Erro_RelRentabilidadeCli_Prepara:

    RelRentabilidadeCli_Prepara = gErr
    
    Select Case gErr
    
        Case 128180
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 128181
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 128182

        Case 128183, 128184, 128197
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
            
        Case 128185, 128186
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESNF", gErr)
            
        Case 128187, 128188
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTRATOPROPAGANDA", gErr)
            
        Case 128189, 128190, 128195
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 128191, 128192
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFAT", gErr)
            
        Case 128193, 128194
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOMES", gErr)
            
        Case 128196
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELRENTABILIDADECLI", gErr)
            
        Case 128198
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179183)

    End Select

    'Transação Rollback
    Transacao_Rollback

    'Fecha o Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function CancelamentoNFiscal_Atualiza_Pedido(ByVal objNFiscal As ClassNFiscal) As Long
'Atualiza o Pedido de Venda após o cancelamento da NF passada em objNFiscal.

Dim lErro As Long
Dim alComando(0 To 15) As Long
Dim iIndice As Integer
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objItemNF As ClassItemNF
Dim sSelect As String
Dim tItemPedido As typeItemPedido
Dim iBloqueio As Integer
Dim dQuantItem As Double
Dim iItemPVStatus As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim dQuantFaturada As Double
Dim dFator As Double
Dim objProduto As ClassProduto

On Error GoTo Erro_CancelamentoNFiscal_Atualiza_Pedido

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 62208
    Next
    
    objPedidoVenda.lCodigo = objNFiscal.lNumPedidoVenda
    objPedidoVenda.iFilialEmpresa = objNFiscal.iFilialPedido
    
    'BUsca o Pedido na tabela de Pedidos de Venda
    lErro = CF("PedidoDeVenda_Le_Lock", alComando(0), objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 36986 Then gError 62210
    
    'Se encontrou
    If lErro = SUCESSO Then
        
        'BUsca os itens do Pedido
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc,QuantFaturada FROM ItensPedidoDeVenda WHERE CodPedido = ? AND FilialEmpresa = ?", 0, tItemPedido.lNumIntDoc, tItemPedido.dQuantFaturada, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 62212
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62213
        
        'Para cada item do pedido
        Do While lErro <> AD_SQL_SEM_DADOS
            
            dQuantItem = 0
            
            'Verifica quanto foi faturado desse item por essa nf
            For Each objItemNF In objNFiscal.ColItensNF
                If objItemNF.lNumIntItemPedVenda = tItemPedido.lNumIntDoc Then
                    dQuantItem = objItemNF.dQuantidade
                    
                    For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                        
                        lErro = Comando_ExecutarPos(alComando(14), "SELECT QuantFaturada FROM ItensPedidoDeVendaGrade WHERE NumIntDoc =?", 0, dQuantItem, objItemRomaneio.lNumIntItemPV)
                        If lErro <> AD_SQL_SUCESSO Then gError 86369
                        
                        lErro = Comando_BuscarPrimeiro(alComando(14))
                        If lErro <> AD_SQL_SUCESSO Then gError 86370
                        
                        lErro = Comando_ExecutarPos(alComando(15), "UPDATE ItensPedidoDeVendaGrade SET Quantfaturada = QuantFaturada - ?", alComando(14), objItemRomaneio.dQuantidade)
                        If lErro <> AD_SQL_SUCESSO Then gError 86371
                    Next
            
                    'Atualiza a qtd faturada o item de Pedido
                    lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensPedidoDeVenda SET QuantFaturada = QuantFaturada - ?", alComando(2), objItemNF.dQuantidade)
                    If lErro <> AD_SQL_SUCESSO Then gError 62219
            
                    Exit For
                End If
            Next
            
            'BUsca o próximo item
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 62270
        
        Loop
        
        If giTipoVersao = VERSAO_FULL Then
            'seta a variável p\ uma possível criação de 1 bloqueio parcial p\ o PV
            iBloqueio = BLOQUEIO_PARCIAL
        End If
        
    'Se não encontrou
    Else
        
        'BUsca o pedido da tabela dos Pedidos de Venda Baixados
        lErro = CF("PedidoDeVendaBaixado_LeMod", alComando(0), objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 62205 Then gError 62211
        
        'Se não encontrou ---> erro
        If lErro <> SUCESSO Then gError 62218
            
        'Desfaz a baixa do pedido de venda
        lErro = CF("PedidoDeVenda_Insere", objPedidoVenda, alComando(2))
        If lErro <> SUCESSO Then gError 62220
        
        With tItemPedido
            
            .sProduto = String(STRING_PRODUTO, 0)
            .sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
            .sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
            .sUnidadeMed = String(STRING_UM_SIGLA, 0)
            
            'BUsca os itens baixados do pedido
            sSelect = "SELECT NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC,PercDesc1,TipoDesc1,PercDesc2,TipoDesc2,PercDesc3,TipoDesc3 FROM ItensPedidoDeVendaBaixados WHERE  FilialEmpresa = ? AND CodPedido = ?"
            lErro = Comando_ExecutarPos(alComando(3), sSelect, 0, .lNumIntDoc, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .iClasseUM, .sUnidadeMed, .dValorDesconto, .dtDataEntrega, .sDescricao, .sLote, .iStatus, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .dQuantFaturada, .dQuantOP, .dQuantSC, .dPercDesc1, .iTipoDesc1, .dPercDesc2, .iTipoDesc2, .dPercDesc3, .iTipoDesc3, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
            If lErro <> AD_SQL_SUCESSO Then gError 62214
                            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62215
            
            If lErro <> AD_SQL_SUCESSO Then gError 62217 'Não encontrou itens
            
            'Para cada item do Pedido
            Do While lErro <> AD_SQL_SEM_DADOS
            
                iItemPVStatus = .iStatus
                
                'Atualiza a quant faturada
                For Each objItemNF In objNFiscal.ColItensNF
                    If objItemNF.lNumIntItemPedVenda = .lNumIntDoc Then
                    
                        '#####################################################
                        'Inserido por Wagner
                        Set objProduto = New ClassProduto
                        
                        objProduto.sCodigo = objItemNF.sProduto
                        
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 23080 Then gError 138549
                        
                        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemNF.sUnidadeMed, .sUnidadeMed, dFator)
                        If lErro <> SUCESSO Then gError 138550
                        '####################################################
                    
                        .dQuantFaturada = StrParaDbl(Formata_Estoque(.dQuantFaturada - (objItemNF.dQuantidade * dFator))) 'Alterado por Wagner
                        iItemPVStatus = STATUS_ABERTO
                        
                        For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                            
                            lErro = Comando_ExecutarPos(alComando(14), "SELECT QuantFaturada FROM ItensPedidoDeVendaGrade WHERE NumIntDoc =?", 0, dQuantItem, objItemRomaneio.lNumIntItemPV)
                            If lErro <> AD_SQL_SUCESSO Then gError 86369
                            
                            lErro = Comando_BuscarPrimeiro(alComando(14))
                            If lErro <> AD_SQL_SUCESSO Then gError 86370
                            
                            lErro = Comando_ExecutarPos(alComando(15), "UPDATE ItensPedidoDeVendaGrade SET Quantfaturada = QuantFaturada - ?", alComando(14), objItemRomaneio.dQuantidade)
                            If lErro <> AD_SQL_SUCESSO Then gError 86371
                        Next
                
                        Exit For
                    End If
                Next
                
                objPedidoVenda.colItensPedido.Add .iFilialEmpresa, .lCodPedido, 0, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReservaItem, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, .dQuantAFaturar, .sDescricao, iItemPVStatus, .iControleEstoque, .dQuantOP, .dQuantSC, .dPercDesc1, .iTipoDesc1, .dPercDesc2, .iTipoDesc2, .dPercDesc3, .iTipoDesc3, .lNumIntDoc
                
                'Desfaz a baixa do item de pedido de venda
                lErro = Comando_Executar(alComando(4), "INSERT INTO ItensPedidoDeVenda (FilialEmpresa, CodPedido, NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC,PercDesc1,TipoDesc1,PercDesc2,TipoDesc2,PercDesc3,TipoDesc3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, .lNumIntDoc, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .iClasseUM, .sUnidadeMed, .dValorDesconto, .dtDataEntrega, .sDescricao, .sLote, iItemPVStatus, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .dQuantFaturada, .dQuantOP, .dQuantSC, .dPercDesc1, .iTipoDesc1, .dPercDesc2, .iTipoDesc2, .dPercDesc3, .iTipoDesc3)
                If lErro <> AD_SQL_SUCESSO Then gError 62221
            
                'Exclui o item da tabela dos itens baixados
                lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM ItensPedidoDeVendaBaixados", alComando(3))
                If lErro <> AD_SQL_SUCESSO Then gError 62222
            
                'BUsca o próximo item
                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62269
                
            Loop
            
        End With
            
        'Exclui o pedido de venda da tabela dos baixados
        lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM PedidosDeVendaBaixados", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 62223
        
        'Desfaz a baixa da tributação do pedido
        lErro = CF("TributacaoPV_Desfaz_Baixa", objPedidoVenda)
        If lErro <> SUCESSO Then gError 62245
        
        'Desfaz a baixa das parcelas do pedido
        lErro = CF("ParcelasPV_Desfaz_Baixa", objPedidoVenda, alComando(7), alComando(8))
        If lErro <> SUCESSO Then gError 62248

        'Desfaz a baixa das comissoes do pedido
        lErro = CF("ComissoesPV_Desfaz_Baixa", objPedidoVenda, alComando(9), alComando(10))
        If lErro <> SUCESSO Then gError 62249
        
        If giTipoVersao = VERSAO_FULL Then
            'Desfaz a baixa dos bloqueios do pedido
            lErro = CF("BloqueiosPV_Desfaz_Baixa", objPedidoVenda, alComando(11), alComando(12))
            If lErro <> SUCESSO Then gError 62250
            
            'seta a variável p\ uma possível criação de 1 bloqueio parcial p\ o PV
            iBloqueio = BLOQUEIO_NAO_RESERVA
        End If
        
    End If

    If giTipoVersao = VERSAO_FULL Then
        
        'Inclui, se necessário, um bloqueio no tipo passado para o pedido de venda
        'faturado pela NF passada
        lErro = CF("CancelaNFiscal_Inclui_BloqueioPV", alComando(13), objNFiscal, objPedidoVenda, iBloqueio)
        If lErro <> SUCESSO Then gError 62268
        
    End If
    
    'mario 19/01/01
    'Troca o sinal da quantidade para que a função ProdutoFilial_Subtrai_NF_QuantPedida adicione as quantidades dos itens as quantidade em pedido de cada produto
    For Each objItemNF In objNFiscal.ColItensNF
        objItemNF.dQuantidade = -objItemNF.dQuantidade
    Next
    
    'Atualiza, para cada produto, a quantidade em pedido em ProdutosFilial
    lErro = CF("ProdutoFilial_Subtrai_NF_QuantPedida", objNFiscal)
    If lErro <> SUCESSO Then gError 83542
    
    'Retorna ao  sinal da quantidade original
    For Each objItemNF In objNFiscal.ColItensNF
        objItemNF.dQuantidade = -objItemNF.dQuantidade
    Next
    'fim mario 19/01/01
    
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
       
    CancelamentoNFiscal_Atualiza_Pedido = SUCESSO
        
    Exit Function
   
Erro_CancelamentoNFiscal_Atualiza_Pedido:
    
    CancelamentoNFiscal_Atualiza_Pedido = gErr
    
    Select Case gErr
    
        Case 62208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 62209, 62210, 62211, 62246, 62247, 62248, 62249, 62250, 62268, 62270, 83542, 62220, 138549, 138550
        
        Case 62212, 62213, 62214, 62215, 86369, 86370
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)
        
        Case 62216, 62217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_NAO_CADASTRADO", gErr, objItemNF.lNumIntItemPedVenda)
        
        Case 62218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
        
        Case 62219, 86371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDVENDA", gErr, objPedidoVenda.lCodigo)
        
        Case 62221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPV", gErr)
        
        Case 62222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPEDIDODEVENDA", gErr, objPedidoVenda.lCodigo)
            
        Case 62223
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PEDIDODEVENDA", gErr, objPedidoVenda.lCodigo)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179184)
            
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Sub MapaMensalVendaSQL_Prepara(ByVal vsProd_I As Variant, ByVal vsProd_F As Variant, ByVal vbTodasCategorias As Variant, ByVal vsCategoria As Variant, ByVal vsValorInicial As Variant, ByVal vsValorFinal As Variant, sSQL As String)
'monta o comando SQL para obtencao das fretes dinamicamente e retorna.
Dim sSelect As String, sWhere As String, sFrom As String, sOrderBy As String, sGroupBy As String

On Error GoTo Erro_MapaMensalVendaSQL_Prepara

    sSelect = "SELECT  Codigo, " & _
                    "ClasseUM,  " & _
                    "SiglaUMVenda  "
                    
    sFrom = "FROM Produtos, " & _
                "ProdutosFilial "

    sWhere = "WHERE Produtos.Codigo = ProdutosFilial.Produto AND " & _
                    "ProdutosFilial.FilialEmpresa = ? AND " & _
                    "Faturamento = ? "
     
    sOrderBy = "ORDER BY Produtos.Codigo "
    
    sGroupBy = "GROUP BY Codigo, " & _
                    "ClasseUM,  " & _
                    "SiglaUMVenda  "
                    
    If Len(Trim(vsCategoria)) > 0 Or Len(Trim(vsValorInicial)) > 0 Or Len(Trim(vsValorFinal)) > 0 Then
   
        sFrom = sFrom & ", ProdutoCategoria "
        sWhere = sWhere & " AND Produtos.Codigo = ProdutoCategoria.Produto " _

    End If
   
    If Len(Trim(vsProd_I)) > 0 Then
        sWhere = sWhere & "AND Produtos.Codigo >= ? "
    End If
    
    If Len(Trim(vsProd_F)) > 0 Then
        sWhere = sWhere & "AND Produtos.Codigo <= ? "
    End If
    
    If Not vbTodasCategorias Then
        sWhere = sWhere & "AND ProdutoCategoria.Categoria = ? "
    End If
    
    If Len(Trim(vsValorInicial)) > 0 Then
        sWhere = sWhere & "AND ProdutoCategoria.Item >= ? "
    End If
    
    If Len(Trim(vsValorFinal)) > 0 Then
        sWhere = sWhere & "AND ProdutoCategoria.Item <= ? "
    End If
    
    sSQL = sSelect & sFrom & sWhere & sGroupBy & sOrderBy

    Exit Sub

Erro_MapaMensalVendaSQL_Prepara:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179185)

    End Select

    Exit Sub

End Sub

Private Function MapaMensalVendaInt_Prepara(ByVal lComando As Long, vsProduto As Variant, viClasseUM As Variant, vsUMPadrao As Variant, ByVal viFilialEmpresa As Variant, ByVal viFaturamento As Variant, ByVal vsProd_I As Variant, ByVal vsProd_F As Variant, ByVal vbTodasCategorias As Variant, ByVal vsCategoria As Variant, ByVal vsValorInicial As Variant, ByVal vsValorFinal As Variant, ByVal sSQL As String) As Long

Dim lErro As Long

On Error GoTo Erro_MapaMensalVendaInt_Prepara

    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129188

    lErro = Comando_BindVarInt(lComando, vsProduto)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129189
    
    lErro = Comando_BindVarInt(lComando, viClasseUM)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129190

    lErro = Comando_BindVarInt(lComando, vsUMPadrao)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129191

    lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129192
    
    lErro = Comando_BindVarInt(lComando, viFaturamento)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129193
      
    If Len(Trim(vsProd_I)) > 0 Then
        lErro = Comando_BindVarInt(lComando, vsProd_I)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129194
    End If
    
    If Len(Trim(vsProd_F)) > 0 Then
        lErro = Comando_BindVarInt(lComando, vsProd_F)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129195
    End If
   
   If Not vbTodasCategorias Then
        lErro = Comando_BindVarInt(lComando, vsCategoria)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129196
   End If
   
    If Len(Trim(vsValorInicial)) > 0 Then
        lErro = Comando_BindVarInt(lComando, vsValorInicial)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129197
    End If
    
    If Len(Trim(vsValorFinal)) > 0 Then
        lErro = Comando_BindVarInt(lComando, vsValorFinal)
        If (lErro <> AD_SQL_SUCESSO) Then gError 129198
    End If

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 129199
    
    MapaMensalVendaInt_Prepara = SUCESSO

    Exit Function

Erro_MapaMensalVendaInt_Prepara:

    MapaMensalVendaInt_Prepara = gErr

    Select Case gErr
    
        Case 129188 To 129199

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179186)

    End Select

    Exit Function

End Function

'#############################################################
'Inserido por Wagner
Function MovEstoque_Gera(alComando() As Long, alComando1() As Long, ByVal objNFiscal As ClassNFiscal) As Long
'gera os movimentos de estoque da NFiscal, se necessários

Dim lErro As Long, iIndice As Integer
Dim iTipoMovtoEstoque As Integer
Dim objMovEstoque As ClassMovEstoque
Dim lCodigo As Long, dFator As Double
Dim objItemNF As ClassItemNF, dSomaAlocacao As Double
Dim objItemNFAloc As ClassItemNFAlocacao
Dim sDocOrigem As String, objTipoDocInfo As New ClassTipoDocInfo
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objRastroItemNF As ClassRastroItemNF
Dim objRastroMovto As ClassRastreamentoMovto
Dim colInsumos As Collection
Dim colRastroMovto As Collection
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objProduto As New ClassProduto
Dim objReservaItem As ClassReservaItem

On Error GoTo Erro_MovEstoque_Gera

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then gError 59261
    
    iTipoMovtoEstoque = objTipoDocInfo.iTipoMovtoEstoque
    sDocOrigem = objTipoDocInfo.sSigla

    If iTipoMovtoEstoque > 0 Then

        Set objMovEstoque = New ClassMovEstoque

        'Obtém o Código do Movimento
        lErro = CF("MovEstoque_Automatico_EmTransacao1", alComando(2), alComando(3), alComando(4), objNFiscal.iFilialEmpresa, lCodigo)
        If lErro <> SUCESSO Then gError 39001

        'Alterado por Wagner
        If objNFiscal.dtDataSaida = DATA_NULA Then
            objMovEstoque.dtData = objNFiscal.dtDataEmissao
        Else
            objMovEstoque.dtData = objNFiscal.dtDataSaida
        End If
        
        objMovEstoque.dtHora = objNFiscal.dtHoraSaida
        objMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objMovEstoque.iTipoMov = iTipoMovtoEstoque
        objMovEstoque.lCodigo = lCodigo
        objMovEstoque.lCliente = objNFiscal.lCliente
        objMovEstoque.lFornecedor = objNFiscal.lFornecedor
        objMovEstoque.sDocOrigem = sDocOrigem & " " & objNFiscal.sSerie & " " & CStr(objNFiscal.lNumNotaFiscal)

        iIndice = 0
        
        'Adiciona itens ao Movimento
        For Each objItemNF In objNFiscal.ColItensNF
            
            iIndice = iIndice + 1
            
            If objItemNF.colItensRomaneioGrade.Count = 0 Then
            
                If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    
                    If objItemNF.ColAlocacoes.Count = 0 Then gError 42367
    
                    dSomaAlocacao = 0
                                                                    
                    For Each objItemNFAloc In objItemNF.ColAlocacoes
                                            
                        dSomaAlocacao = dSomaAlocacao + Arredonda_Estoque(objItemNFAloc.dQuantidade)
                        
                        Set colInsumos = New Collection
                        Set colRastroMovto = New Collection
                                            
                        Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dValorTotal, APROPR_CUSTO_INFORMADO, objItemNF.sProduto, objItemNF.sDescricaoItem, objItemNFAloc.sUnidadeMed, objItemNFAloc.dQuantidade, objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNF.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                        
                        If Not objNFiscal.objRastreamento Is Nothing Then
                        
                            'Guarda o Rastreamento dos ItensNF
                            lErro = objNFiscal.objRastreamento.Move_RastroItemNF_ItemMovEstoque(objItemMovEstoque, objItemNF.colRastreamento, objTipoDocInfo)
                            If lErro <> SUCESSO Then gError 89461
                        
                        End If
    
                        objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                        
                        objItemMovEstoque.iClasseUM = objItemNF.iClasseUM
                        objItemMovEstoque.iControleEstoque = objItemNF.iControleEstoque
                        objItemMovEstoque.sSiglaUMEst = objItemNF.sUMEstoque
                        objItemMovEstoque.iApropriacaoProd = objItemNF.iApropriacaoProd
                        
                    Next
                    
                    lErro = CF("UM_Conversao", objItemNF.iClasseUM, objItemNF.sUMEstoque, objItemNF.sUnidadeMed, dFator)
                    If lErro <> SUCESSO Then gError 41525
                    
                    dSomaAlocacao = Arredonda_Estoque(dSomaAlocacao * dFator)
    
                    If Abs(dSomaAlocacao - objItemNF.dQuantidade) > QTDE_ESTOQUE_DELTA2 Then gError 42368
                
                End If
            Else
                For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                    objProduto.sCodigo = objItemRomaneio.sProduto
                    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 23080 Then gError 86366
                    If lErro <> SUCESSO Then Error 86367
                    
                    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    
                        If objItemRomaneio.colLocalizacao.Count = 0 Then gError 42367
                        
                        dSomaAlocacao = 0
                        
                        For Each objReservaItem In objItemRomaneio.colLocalizacao
                            
                            dSomaAlocacao = dSomaAlocacao + objReservaItem.dQuantidade
                            
                            Set objItemMovEstoque = objMovEstoque.colItens.Add(0, 0, objItemNF.dValorTotal, APROPR_CUSTO_INFORMADO, objItemRomaneio.sProduto, objItemRomaneio.sDescricao, objProduto.sSiglaUMEstoque, objReservaItem.dQuantidade, objReservaItem.iAlmoxarifado, objReservaItem.sAlmoxarifado, objItemRomaneio.lNumIntDoc, "", 0, "", "", "", "", 0, colRastroMovto, colInsumos, DATA_NULA)
                            
                            objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
                            
                            objItemMovEstoque.iClasseUM = objProduto.iClasseUM
                            objItemMovEstoque.iControleEstoque = objProduto.iControleEstoque
                            objItemMovEstoque.sSiglaUMEst = objProduto.sSiglaUMEstoque
                            objItemMovEstoque.iApropriacaoProd = objProduto.iApropriacaoCusto
                            
                        Next
                        
                    End If
                Next
            
            End If
        Next

        'Grava os Movimentos de Estoque relacionados a Nota Fiscal
        lErro = CF("MovEstoque_Grava1", alComando1(), objMovEstoque)
        If lErro <> SUCESSO Then gError 39002
        
    End If

    MovEstoque_Gera = SUCESSO

    Exit Function

Erro_MovEstoque_Gera:

    MovEstoque_Gera = gErr

    Select Case gErr

        Case 35135, 35756
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO1", gErr, objNFiscal.iTipoNFiscal)

        Case 39000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 39001, 39002, 41525, 59261, 89461

        Case 42367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCALIZACAO_ITEM_INEXISTENTE", gErr, iIndice)

        Case 42368
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCALIZACAO_ITEM_INCOMPLETA", gErr, iIndice, objItemNF.dQuantidade, dSomaAlocacao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179187)

    End Select

    Exit Function

End Function

Function BloqueioCredito_Adiciona(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Inclui no BD um bloqueio de Credito, se o cliente não tiver crédito

Dim lComando As Long
Dim lComando1 As Long
Dim lErro As Long
Dim dValorPedido As Double
Dim objItemPV As ClassItemPedido
Dim tCliente As typeCliente
Dim objClienteEstatistica As New ClassFilialClienteEst
Dim objAutorizacaoCredito As New ClassAutorizacaoCredito 'Inserido por Wagner

On Error GoTo Erro_BloqueioCredito_Adiciona

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 42974

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 42979

    dValorPedido = 0

    dValorPedido = objPedidoVenda.dValorTotal

    'Lê os saldos e o limite de credito do Cliente
    lErro = Comando_ExecutarPos(lComando, "SELECT LimiteCredito, Bloqueado FROM Clientes WHERE Codigo = ?", 0, tCliente.dLimiteCredito, tCliente.iBloqueado, objPedidoVenda.lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 42975

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 42976
    If lErro = AD_SQL_SEM_DADOS Then gError 42977 'Não encontrou

    'Faz o lock do Cliente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 42978
    
    objClienteEstatistica.lCodCliente = objPedidoVenda.lCliente
    objClienteEstatistica.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
    
    'Le dinamicamente o Saldo dos Titulos e dos Pedidos Liberados
    lErro = CF("Cliente_Le_Estatistica_Credito", objClienteEstatistica)
    If lErro <> SUCESSO Then gError 52952
        
    'Verifica se a Soma dos Titulos, Pedidos, NFs não faturadas e valor do Pedido ultrapassa o limite do Cliente
    If (tCliente.dLimiteCredito < (objClienteEstatistica.dSaldoTitulos + objClienteEstatistica.dSaldoPedidosLiberados + objClienteEstatistica.dValorNFsNaoFaturadas + dValorPedido)) Then
        
        objAutorizacaoCredito.dValor = objPedidoVenda.dValorTotal
        objAutorizacaoCredito.lCliente = objPedidoVenda.lCliente
        
        'Limite de crédito foi ultrapassado ==> pede autorização de credito ao usuário
        Call Chama_Tela_Modal("AutorizacaoAtraso", objAutorizacaoCredito)
        
        'se o crédito não foi aprovado ==> erro
        If objAutorizacaoCredito.iCreditoAutorizado = CREDITO_RECUSADO Then
        
            'Adiciona o bloqueio de crédito ao Pedido
            objPedidoVenda.colBloqueiosPV.Add giFilialEmpresa, objPedidoVenda.lCodigo, 0, BLOQUEIO_CREDITO, gsUsuario, BLOQUEIO_AUTO_RESP, gdtDataHoje, "", "", DATA_NULA, ""

        End If

    End If

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    BloqueioCredito_Adiciona = SUCESSO

    Exit Function

Erro_BloqueioCredito_Adiciona:

    BloqueioCredito_Adiciona = gErr

    Select Case gErr

        Case 42974, 42979
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 42975, 42976
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)

        Case 42977
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objPedidoVenda.lCliente)

        Case 42978
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CLIENTES", gErr, objPedidoVenda.lCliente)

        Case 42980
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CLIENTE", gErr, objPedidoVenda.lCliente)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179188)

    End Select

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Libera_BloqueioCliente(ByVal lCliente As Long) As Long
'Chama a tela de autorização caso o cliente esteja bloqueado

Dim lComando As Long
Dim lErro As Long
Dim tCliente As typeCliente
Dim objAutorizacaoCredito As New ClassAutorizacaoCredito 'Inserido por Wagner
Dim dSaldoAtraso As Double

On Error GoTo Erro_Libera_BloqueioCliente

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 132300
    
    'Verifica atrasos
    lErro = CF("Cliente_Obtem_SaldoAtraso", lCliente, dSaldoAtraso)
    If lErro <> SUCESSO Then gError 132306
    
    'Lê os saldos e o limite de credito do Cliente
    lErro = Comando_ExecutarPos(lComando, "SELECT LimiteCredito, Bloqueado FROM Clientes WHERE Codigo = ?", 0, tCliente.dLimiteCredito, tCliente.iBloqueado, lCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 132301

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 132302
    If lErro = AD_SQL_SEM_DADOS Then gError 132303 'Não encontrou
    
    'Verifica se o Cliente está bloqueado ou em atraso
    If tCliente.iBloqueado = MARCADO Or dSaldoAtraso <> 0 Then
        
        objAutorizacaoCredito.dValor = 0
        objAutorizacaoCredito.lCliente = lCliente
        
        'pede autorização ao usuário
        Call Chama_Tela_Modal("AutorizacaoAtraso", objAutorizacaoCredito)
        
        'se o crédito não foi aprovado ==> erro
        If objAutorizacaoCredito.iCreditoAutorizado = CREDITO_RECUSADO Then
            If tCliente.iBloqueado = MARCADO Then gError 132304
            If dSaldoAtraso <> 0 Then gError 132305
        End If

    End If

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    Libera_BloqueioCliente = SUCESSO

    Exit Function

Erro_Libera_BloqueioCliente:

    Libera_BloqueioCliente = gErr

    Select Case gErr

        Case 132300
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 132301, 132302
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)

        Case 132303
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, lCliente)
        
        Case 132304
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_BLOQUEADO", gErr, lCliente)
        
        Case 132305
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_COM_ATRASO", gErr)

        Case 132306

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179189)

    End Select

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function
'###################################################################

Function NFiscalItens_Grava_Customizado(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim objItemNF As ClassItemNF
Dim lNumIntDoc As Long

On Error GoTo Erro_NFiscalItens_Grava_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197688
    Next
    
    For Each objItemNF In objNFiscal.ColItensNF
    
        'Pesquisa os ítens vinculados à Nota Fiscal
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ItensNFiscal WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItemNF.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 197689

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197690

        If lErro = AD_SQL_SUCESSO Then
            
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensNFiscal SET PrecoUnitarioMoeda = ?", alComando(0), objItemNF.dPrecoUnitarioMoeda)
            If lErro <> AD_SQL_SUCESSO Then gError 197691

        End If

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalItens_Grava_Customizado = SUCESSO

    Exit Function

Erro_NFiscalItens_Grava_Customizado:

    NFiscalItens_Grava_Customizado = gErr

    Select Case gErr

        Case 197688
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197689, 197690
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)

        Case 197691
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ITENSNFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197692)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function NFiscal_Grava_NFiscal_Cust(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer, lNumIntDoc As Long
Dim objNFHic  As ClassNFiscalHic

On Error GoTo Erro_NFiscal_Grava_NFiscal_Cust

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 97513
    Next
    
    If Not (objNFiscal.objInfoUsu Is Nothing) Then
    
        Set objNFHic = objNFiscal.objInfoUsu
    
        lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM NFiscal WHERE NumIntDoc = ? ", 0, lNumIntDoc, objNFiscal.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 197689
    
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197690
    
        If lErro = AD_SQL_SUCESSO Then
            
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE NFiscal SET ValorFrete1 = ?", alComando(0), objNFHic.dValorFrete1)
            If lErro <> AD_SQL_SUCESSO Then gError 197691
            
        End If

    End If
            
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFiscal_Grava_NFiscal_Cust = SUCESSO

    Exit Function

Erro_NFiscal_Grava_NFiscal_Cust:
    
    NFiscal_Grava_NFiscal_Cust = gErr

    Select Case gErr
        
        Case 97513
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 99096, 99097
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
            
        Case 99099
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function NFiscal_Grava_Customizado(ByVal objNFiscal As ClassNFiscal) As Long
    NFiscal_Grava_Customizado = CF("NFiscalEntrada_Grava_CamposAlteraveis", objNFiscal)
End Function
