VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVGravaHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function PedidoVenda_Grava_BD(ByVal objPedidoVenda As ClassPedidoDeVenda, Optional ByVal lComando1 As Long = 0) As Long
'Insere o Pedido de Venda passado na Tabela de Pedidos de Venda
'ou atualiza caso o Pedido já exista

Dim lErro As Long
Dim alComando(0 To 6) As Long
Dim iIndice As Integer, dSaldoAtraso As Double

On Error GoTo Erro_PedidoVenda_Grava_BD

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 42958
    Next

    'Verifica se é uma Inclusão
    If lComando1 = 0 Then
        
        '#######################################
        'Comentado por Wagner
        'Esse código está sendo tratado na Tela ao trazer o Cliente,
        'onde através de senha o usuário pode permitir a gravação
'        If gsUsuario <> "supervisor" Then
'
'            lErro = CF("Cliente_Obtem_SaldoAtraso", objPedidoVenda.lCliente, dSaldoAtraso)
'            If lErro <> SUCESSO Then gError 124203
'
'            If dSaldoAtraso <> 0 Then gError 124204
'
'        End If
        '#######################################
        
        'Insere o novo Pedido no BD
        lErro = CF("PedidoDeVenda_Insere", objPedidoVenda, alComando(0))
        If lErro <> SUCESSO Then gError 42959
        
    Else 'lComando1 foi passado
        'Altera o Pedido de Venda
        With objPedidoVenda
            lErro = Comando_ExecutarPos(alComando(0), "UPDATE PedidosDeVenda SET DataEmissao = ?, FilialEntrega = ?, CodTransportadora = ?,CondicaoPagto =?, PercAcrescFinanceiro= ?, DataReferencia =?, MensagemNota =?,NaturezaOp = ?, ValorProdutos =?,ValorTotal =?, ValorFrete=?, ValorDesconto=?, ValorSeguro =?, ValorOutrasDespesas =?, PedidoCliente = ?, CanalVenda = ?,TabelaPreco =?, FaturaIntegral = ?, ComissaoAutomatica = ?, CobrancaAutomatica = ?,Placa = ? ,PlacaUF = ? ,VolumeQuant = ?, VolumeEspecie = ? ,VolumeMarca = ?,PesoLiq = ? ,PesoBruto = ?, VolumeNumero = ?, FreteRespons = ?, FilialEmpresaFaturamento = ?, CodTranspRedesp = ?, DetPagFrete = ?, DataEntrega = ?," & _
                "VolumeTotal = ?, Moeda = ?, TaxaMoeda = ?, ValorFrete1 = ?, ValorSeguro1 = ?, ValorOutrasDespesas1 = ?, ValorFrete2 = ?, ValorSeguro2 = ?, ValorOutrasDespesas2 = ?, FlagCompl1 = ?, FlagCompl2 = ?, PedidoRepresentante = ?", lComando1, _
                .dtDataEmissao, .iFilialEntrega, .iCodTransportadora, .iCondicaoPagto, .dPercAcrescFinanceiro, .dtDataReferencia, .sMensagemNota, .sNaturezaOp, .dValorProdutos, .dValorTotal, .dValorFrete, .dValorDesconto, .dValorSeguro, .dValorOutrasDespesas, .sPedidoCliente, .iCanalVenda, .iTabelaPreco, .iFaturaIntegral, .iComissaoAutomatica, .iCobrancaAutomatica, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .dPesoLiq, .dPesoBruto, .sVolumeNumero, .iFreteRespons, _
                .iFilialEmpresaFaturamento, .iCodTranspRedesp, .iDetPagFrete, .dtDataEntrega, _
                .dVolumeTotal, .iMoeda, .dTaxaMoeda, .objInfoUsu.dValorFrete1, .objInfoUsu.dValorSeguro1, .objInfoUsu.dValorOutrasDespesas1, .objInfoUsu.dValorFrete2, .objInfoUsu.dValorSeguro2, .objInfoUsu.dValorOutrasDespesas2, .objInfoUsu.iFlagCompl1, .objInfoUsu.iFlagCompl2, .sPedidoRepresentante)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 46004

        lErro = CF("PedidoVenda_ExcluiTributacao", objPedidoVenda, alComando(1), alComando(2), alComando(5), alComando(6))
        If lErro <> SUCESSO Then gError 42377
    End If

    lErro = CF("PedidoVenda_GravaTributacao", objPedidoVenda, alComando(3), alComando(4))
    If lErro <> SUCESSO Then gError 42375

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoVenda_Grava_BD = SUCESSO

    Exit Function

Erro_PedidoVenda_Grava_BD:

    PedidoVenda_Grava_BD = gErr

    Select Case gErr

        Case 42375, 42377, 124203, 42959

        Case 42958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 46004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDVENDA", gErr, objPedidoVenda.lCodigo)

        Case 124204
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_COM_ATRASO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179218)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensPV_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Insere os itens do Pedido de Venda passado no BD

Dim alComando(0 To 4) As Long
Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim objItemOp As New ClassItemOP
Dim sProduto As String
Dim iStatus As Integer

On Error GoTo Erro_ItensPV_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 42964
    Next

    'Para cada item da coleção
    For Each objItemPV In objPedidoVenda.colItensPedido

        sProduto = String(STRING_PRODUTO, 0)
        
        'Localiza o item no Banco de Dados
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Produto, Status, NumIntDoc FROM ItensPedidoDeVenda WHERE CodPedido = ? AND FilialEmpresa = ? AND Produto = ? ", 0, sProduto, iStatus, lNumIntDoc, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa, objItemPV.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 46045

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 46046
        If lErro <> AD_SQL_SUCESSO Then

            'Obtém o número interno automático do item
            lErro = CF("Config_ObterNumInt", "FatConfig", "NUM_PROX_ITEM_PEDVENDA", lNumIntDoc)
            If lErro <> SUCESSO Then gError 42965

            objItemPV.iFilialEmpresa = giFilialEmpresa
            objItemPV.lCodPedido = objPedidoVenda.lCodigo
            objItemPV.lNumIntDoc = lNumIntDoc
            If objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada <= 0 Then objItemPV.iStatus = STATUS_ATENDIDO
            
            'Insere o item no BD
            With objItemPV
                lErro = Comando_Executar(alComando(0), "INSERT INTO ItensPedidoDeVenda (NumIntDoc,FilialEmpresa,CodPedido,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC,PercDesc1,TipoDesc1,PercDesc2,TipoDesc2,PercDesc3,TipoDesc3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                    lNumIntDoc, .iFilialEmpresa, .lCodPedido, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .iClasseUM, .sUnidadeMed, .dValorDesconto, .dtDataEntrega, .sDescricao, .sLote, .iStatus, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .dQuantFaturada, .dQuantOP, .dQuantSC, .dPercDesc1, .iTipoDesc1, .dPercDesc2, .iTipoDesc2, .dPercDesc3, .iTipoDesc3)
            End With
            If lErro <> AD_SQL_SUCESSO Then gError 42966
        
        'Se for uma alteração
        Else
            
            objItemPV.lNumIntDoc = lNumIntDoc
            
            'Se houve alteracao de Produto
            If sProduto <> objItemPV.sProduto Then
            
                'Verifica se ítem estava vinculado a Ordem Produção
                lErro = CF("ItemOP_Le_ItemPV", objItemOp, objItemPV)
                If lErro <> SUCESSO And lErro <> 46074 Then gError 24461
                If lErro = SUCESSO Then gError 24462
            
            End If
            
            objItemPV.iStatus = iStatus
            If objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada <= 0 Then objItemPV.iStatus = STATUS_ATENDIDO
            
            'Faz a atualização no Banco de Dados
            With objItemPV
                lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensPedidoDeVenda SET Produto =?, Quantidade = ?, PrecoUnitario = ?, PrecoTotal = ?, ClasseUM = ?, UnidadeMed = ?, ValorDesconto =?, DataEntrega = ?, Descricao = ?, QuantCancelada = ?, QuantReservada = ?, QuantFaturada = ?, Status = ?,PercDesc1 = ?,TipoDesc1 = ?,PercDesc2 = ?,TipoDesc2 = ?,PercDesc3 = ?,TipoDesc3 = ?", alComando(0), _
                    .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .iClasseUM, .sUnidadeMed, .dValorDesconto, .dtDataEntrega, .sDescricao, .dQuantCancelada, .dQuantReservada, .dQuantFaturada, .iStatus, _
                    .dPercDesc1, .iTipoDesc1, .dPercDesc2, .iTipoDesc2, .dPercDesc3, .iTipoDesc3)
            End With
            If lErro <> AD_SQL_SUCESSO Then gError 46048

            lErro = CF("ItemPV_ExcluiTributacao", objItemPV, alComando(2), alComando(3))
            If lErro <> SUCESSO Then gError 42985
            
            lErro = CF("ItensPVGrade_Exclui", objItemPV)
            If lErro <> SUCESSO Then gError 86338
        
        End If

        If objItemPV.colItensRomaneioGrade.Count > 0 Then
        
            lErro = CF("ItensPVGrade_Inclui", objItemPV)
            If lErro <> SUCESSO Then gError 86339
    
        End If
        
        lErro = CF("ItemPV_GravaTributacao", objItemPV, alComando(4))
        If lErro <> SUCESSO Then gError 42365

    Next

    'Fecha os comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensPV_Grava = SUCESSO

    Exit Function

Erro_ItensPV_Grava:

    ItensPV_Grava = gErr

    Select Case gErr

        Case 24461, 42365, 42376, 42985, 51386, 86338, 86339
        
        Case 24462
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PV_VINCULADO_ITEM_OP", gErr, objItemPV.iItem, objItemOp.lNumIntDoc)

        Case 42964
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 42965
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPV", gErr)

        Case 46045, 46046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case 46047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_NAO_ENCONTRADO", gErr, objItemPV.lNumIntDoc)

        Case 46048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", gErr, objItemPV.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179219)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoVendaBaixado_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda, faz o lock e o insere na tabela de Pedidos baixados

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim tPedidoVenda As typePedidoVenda, tPedidoVendaHic As typePedidoVendaHic

On Error GoTo Erro_PedidoVendaBaixado_Grava

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46098

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 46099

    tPedidoVenda.sMensagemNota = String(STRING_PEDIDOVENDA_MENSAGEM_NOTA, 0)
    tPedidoVenda.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tPedidoVenda.sPedidoCliente = String(STRING_PEDIDOVENDA_PEDIDO_CLIENTE, 0)
    tPedidoVenda.sPlaca = String(STRING_PEDIDOVENDA_PLACA, 0)
    tPedidoVenda.sPlacaUF = String(STRING_PEDIDOVENDA_PLACA_UF, 0)
    tPedidoVenda.sVolumeNumero = String(STRING_PEDIDOVENDA_VOLUME_NUMERO, 0)
    tPedidoVenda.sPedidoRepresentante = String(STRING_PEDIDOVENDA_PEDIDO_REPRESENTANTE, 0)
    
    'Lê o Pedido de Venda
    With tPedidoVenda
        lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresaFaturamento,Cliente,Filial,FilialEntrega,CodTransportadora,CondicaoPagto,PercAcrescFinanceiro,DataEmissao,DataReferencia,MensagemNota,NaturezaOp,ValorTotal,ValorFrete,ValorDesconto,ValorSeguro,ValorOutrasDespesas,PedidoCliente,CanalVenda,TabelaPreco,ProxSeqBloqueio,FaturaIntegral,ComissaoAutomatica,CobrancaAutomatica,NumIntDoc,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,PesoLiq,PesoBruto,VolumeNumero,FreteRespons,ValorProdutos, CodTranspRedesp, DetPagFrete, DataEntrega, VolumeTotal, Moeda, TaxaMoeda, ValorFrete1, ValorSeguro1, ValorOutrasDespesas1, ValorFrete2, ValorSeguro2, ValorOutrasDespesas2, FlagCompl1, FlagCompl2,PedidoRepresentante FROM PedidosDeVenda WHERE FilialEmpresa = ? AND Codigo = ?", 0, _
            .iFilialEmpresaFaturamento, .lCliente, .iFilial, .iFilialEntrega, .iCodTransportadora, .iCondicaoPagto, .dPercAcrescFinanceiro, .dtDataEmissao, .dtDataReferencia, .sMensagemNota, _
            .sNaturezaOp, .dValorTotal, .dValorFrete, .dValorDesconto, .dValorSeguro, .dValorOutrasDespesas, .sPedidoCliente, .iCanalVenda, .iTabelaPreco, .iProxSeqBloqueio, .iFaturaIntegral, .iComissaoAutomatica, .iCobrancaAutomatica, .lNumIntDoc, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .dPesoLiq, .dPesoBruto, .sVolumeNumero, .iFreteRespons, .dValorProdutos, .iCodTranspRedesp, .iDetPagFrete, .dtDataEntrega, _
            .dVolumeTotal, .iMoeda, .dTaxaMoeda, tPedidoVendaHic.dValorFrete1, tPedidoVendaHic.dValorSeguro1, tPedidoVendaHic.dValorOutrasDespesas1, tPedidoVendaHic.dValorFrete2, tPedidoVendaHic.dValorSeguro2, tPedidoVendaHic.dValorOutrasDespesas2, tPedidoVendaHic.iFlagCompl1, tPedidoVendaHic.iFlagCompl2, .sPedidoRepresentante, _
            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    End With
    If lErro <> AD_SQL_SUCESSO Then Error 46099

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46100
    If lErro = AD_SQL_SEM_DADOS Then Error 46101 'Não encontrou

    'Faz o "lock" no Pedido de Venda
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 46102

    objPedidoVenda.lNumIntDoc = tPedidoVenda.lNumIntDoc
    
    With tPedidoVenda
        tPedidoVenda.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
        tPedidoVenda.lCodigo = objPedidoVenda.lCodigo
        'Insere o Pedido na tabela de Pedidos Baixados
        lErro = Comando_Executar(lComando1, "INSERT INTO PedidosDeVendaBaixados (FilialEmpresa, Codigo,FilialEmpresaFaturamento,Cliente,Filial,FilialEntrega,CodTransportadora,CondicaoPagto,PercAcrescFinanceiro,DataEmissao,DataReferencia,MensagemNota,NaturezaOp,ValorTotal,ValorFrete,ValorDesconto,ValorSeguro,ValorOutrasDespesas,PedidoCliente,CanalVenda,TabelaPreco,ProxSeqBloqueio,FaturaIntegral,ComissaoAutomatica,CobrancaAutomatica,NumIntDoc,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,PesoLiq,PesoBruto,ValorProdutos,VolumeNumero,FreteRespons, CodTranspRedesp, DetPagFrete, DataEntrega, VolumeTotal, Moeda, TaxaMoeda, ValorFrete1, ValorSeguro1, ValorOutrasDespesas1, ValorFrete2, ValorSeguro2, ValorOutrasDespesas2, FlagCompl1, FlagCompl2,PedidoRepresentante) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            .iFilialEmpresa, .lCodigo, .iFilialEmpresaFaturamento, .lCliente, .iFilial, .iFilialEntrega, .iCodTransportadora, .iCondicaoPagto, .dPercAcrescFinanceiro, .dtDataEmissao, .dtDataReferencia, .sMensagemNota, .sNaturezaOp, .dValorTotal, .dValorFrete, .dValorDesconto, .dValorSeguro, .dValorOutrasDespesas, .sPedidoCliente, .iCanalVenda, .iTabelaPreco, .iProxSeqBloqueio, .iFaturaIntegral, .iComissaoAutomatica, .iCobrancaAutomatica, .lNumIntDoc, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .dPesoLiq, .dPesoBruto, .dValorProdutos, .sVolumeNumero, .iFreteRespons, .iCodTranspRedesp, .iDetPagFrete, .dtDataEntrega, .dVolumeTotal, .iMoeda, .dTaxaMoeda, tPedidoVendaHic.dValorFrete1, tPedidoVendaHic.dValorSeguro1, tPedidoVendaHic.dValorOutrasDespesas1, tPedidoVendaHic.dValorFrete2, tPedidoVendaHic.dValorSeguro2, tPedidoVendaHic.dValorOutrasDespesas2, tPedidoVendaHic.iFlagCompl1, tPedidoVendaHic.iFlagCompl2, .sPedidoRepresentante)
        If lErro <> AD_SQL_SUCESSO Then Error 46103
    End With

    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    PedidoVendaBaixado_Grava = SUCESSO

    Exit Function

Erro_PedidoVendaBaixado_Grava:

    PedidoVendaBaixado_Grava = Err

    Select Case Err

        Case 46098, 46099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 46099, 46100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", Err, objPedidoVenda.lCodigo)

        Case 46101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", Err, objPedidoVenda.lCodigo)

        Case 46102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOS_DE_VENDA", Err, objPedidoVenda.lCodigo)

        Case 46103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PEDIDODEVENDABAIXADO", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179220)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function ItensPVBaixados_Grava(ByVal objPedidoVenda As ClassPedidoDeVenda, iFaturaTudo As Integer) As Long

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim tItemPV As typeItemPedido
Dim iTemItens As Integer 'Sinaliza se o Pedido já vem carregado com seus itens

On Error GoTo Erro_ItensPVBaixados_Grava

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46104

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 46105

    tItemPV.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItemPV.sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
    tItemPV.sProduto = String(STRING_PRODUTO, 0)
    tItemPV.sUnidadeMed = String(STRING_UM_SIGLA, 0)

    With tItemPV
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC,PercDesc1,TipoDesc1,PercDesc2,TipoDesc2,PercDesc3,TipoDesc3 FROM ItensPedidoDeVenda WHERE FilialEmpresa = ? AND CodPedido = ?", 0, _
            .lNumIntDoc, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .iClasseUM, .sUnidadeMed, .dValorDesconto, .dtDataEntrega, .sDescricao, .sLote, .iStatus, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .dQuantFaturada, .dQuantOP, .dQuantSC, _
            .dPercDesc1, .iTipoDesc1, .dPercDesc2, .iTipoDesc2, .dPercDesc3, .iTipoDesc3, _
            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    End With
    If lErro <> AD_SQL_SUCESSO Then Error 46106

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46107

    iTemItens = False
    If objPedidoVenda.colItensPedido.Count > 0 Then iTemItens = True

    Do While lErro = AD_SQL_SUCESSO

        With tItemPV
            
            If Not iTemItens Then objPedidoVenda.colItensPedido.Add objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, 0, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReservaItem, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, .dQuantFaturada, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, .dPercDesc1, .iTipoDesc1, .dPercDesc2, .iTipoDesc2, .dPercDesc3, .iTipoDesc3, .lNumIntDoc
            
            .lCodPedido = objPedidoVenda.lCodigo
            .iFilialEmpresa = objPedidoVenda.iFilialEmpresa
            .dQuantReservada = 0

            lErro = Comando_Executar(lComando1, "INSERT INTO ItensPedidoDeVendaBaixados (FilialEmpresa, CodPedido, NumIntDoc,Produto,Quantidade,PrecoUnitario,PrecoTotal,ClasseUM,UnidadeMed,ValorDesconto,DataEntrega,Descricao,Lote,Status,ValorAbatComissao,QuantCancelada,QuantReservada,QuantFaturada,QuantOP,QuantSC,PercDesc1,TipoDesc1,PercDesc2,TipoDesc2,PercDesc3,TipoDesc3) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .iFilialEmpresa, .lCodPedido, .lNumIntDoc, .sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .iClasseUM, .sUnidadeMed, .dValorDesconto, .dtDataEntrega, .sDescricao, .sLote, .iStatus, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .dQuantFaturada, .dQuantOP, .dQuantSC, .dPercDesc1, .iTipoDesc1, .dPercDesc2, .iTipoDesc2, .dPercDesc3, .iTipoDesc3)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46108

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46109

        End With

    Loop

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ItensPVBaixados_Grava = SUCESSO

    Exit Function

Erro_ItensPVBaixados_Grava:

    ItensPVBaixados_Grava = Err

    Select Case Err

        Case 46104, 46105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 46106, 46107, 46109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA2", Err, objPedidoVenda.lCodigo)

        Case 46108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPEDIDODEVENDABAIXADOS", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179221)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function PedidoDeVenda_Insere(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal lComando As Long) As Long
'insere o pv na tabela

Dim lErro As Long

On Error GoTo Erro_PedidoDeVenda_Insere
            
    With objPedidoVenda
        lErro = Comando_Executar(lComando, "INSERT INTO PedidosDeVenda (FilialEmpresa,Codigo,FilialEmpresaFaturamento,Cliente,Filial,FilialEntrega,CodTransportadora,CondicaoPagto,PercAcrescFinanceiro,DataEmissao,DataReferencia,MensagemNota,NaturezaOp,ValorTotal,ValorFrete,ValorDesconto,ValorSeguro,ValorOutrasDespesas,PedidoCliente,CanalVenda,TabelaPreco,ProxSeqBloqueio,FaturaIntegral,ComissaoAutomatica,CobrancaAutomatica,NumIntDoc,Placa,PlacaUF,VolumeQuant,VolumeEspecie,VolumeMarca,PesoLiq,PesoBruto,ValorProdutos,VolumeNumero,FreteRespons, CodTranspRedesp, DetPagFrete, DataEntrega, VolumeTotal, Moeda, TaxaMoeda, ValorFrete1, ValorSeguro1, ValorOutrasDespesas1, ValorFrete2, ValorSeguro2, ValorOutrasDespesas2, FlagCompl1, FlagCompl2,PedidoRepresentante) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresaFaturamento, objPedidoVenda.lCliente, objPedidoVenda.iFilial, _
            .iFilialEntrega, .iCodTransportadora, .iCondicaoPagto, .dPercAcrescFinanceiro, .dtDataEmissao, .dtDataReferencia, .sMensagemNota, .sNaturezaOp, .dValorTotal, .dValorFrete, .dValorDesconto, .dValorSeguro, .dValorOutrasDespesas, .sPedidoCliente, .iCanalVenda, .iTabelaPreco, .iProxSeqBloqueio, .iFaturaIntegral, .iComissaoAutomatica, .iCobrancaAutomatica, .lNumIntDoc, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .dPesoLiq, .dPesoBruto, .dValorProdutos, .sVolumeNumero, .iFreteRespons, .iCodTranspRedesp, .iDetPagFrete, _
            .dtDataEntrega, .dVolumeTotal, .iMoeda, .dTaxaMoeda, .objInfoUsu.dValorFrete1, .objInfoUsu.dValorSeguro1, .objInfoUsu.dValorOutrasDespesas1, .objInfoUsu.dValorFrete2, .objInfoUsu.dValorSeguro2, .objInfoUsu.dValorOutrasDespesas2, .objInfoUsu.iFlagCompl1, .objInfoUsu.iFlagCompl2, .sPedidoRepresentante)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 62220
       
    PedidoDeVenda_Insere = SUCESSO
        
    Exit Function
   
Erro_PedidoDeVenda_Insere:
    
    PedidoDeVenda_Insere = gErr
    
    Select Case gErr
    
        Case 62220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PEDVENDA", gErr, objPedidoVenda.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179222)
            
    End Select

    Exit Function

End Function

