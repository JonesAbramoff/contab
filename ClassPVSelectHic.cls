VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPVSelectHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function PedidoDeVenda_Le_InfoUsu(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long, lComando As Long
Dim tPedidoVendaHic As typePedidoVendaHic

On Error GoTo Erro_PedidoDeVenda_Le_InfoUsu

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 124185
    
    lErro = Comando_Executar(lComando, "SELECT ValorFrete1, ValorSeguro1, ValorOutrasDespesas1, ValorFrete2, ValorSeguro2, ValorOutrasDespesas2, FlagCompl1, FlagCompl2 FROM PedidosDeVenda WHERE  FilialEmpresa = ? AND Codigo = ?", _
        tPedidoVendaHic.dValorFrete1, tPedidoVendaHic.dValorSeguro1, tPedidoVendaHic.dValorOutrasDespesas1, tPedidoVendaHic.dValorFrete2, tPedidoVendaHic.dValorSeguro2, tPedidoVendaHic.dValorOutrasDespesas2, tPedidoVendaHic.iFlagCompl1, tPedidoVendaHic.iFlagCompl2, _
        objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 124186
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124187
    
    'Alterado por Wagner
    '######################
    'Se não encontrar dados busca na tabela de baixados
    If lErro = AD_SQL_SEM_DADOS Then
        lErro = Comando_Executar(lComando, "SELECT ValorFrete1, ValorSeguro1, ValorOutrasDespesas1, ValorFrete2, ValorSeguro2, ValorOutrasDespesas2, FlagCompl1, FlagCompl2 FROM PedidosDeVendaBaixados WHERE  FilialEmpresa = ? AND Codigo = ?", _
            tPedidoVendaHic.dValorFrete1, tPedidoVendaHic.dValorSeguro1, tPedidoVendaHic.dValorOutrasDespesas1, tPedidoVendaHic.dValorFrete2, tPedidoVendaHic.dValorSeguro2, tPedidoVendaHic.dValorOutrasDespesas2, tPedidoVendaHic.iFlagCompl1, tPedidoVendaHic.iFlagCompl2, _
            objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 129166
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129167
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 124188
    '######################

    Set objPedidoVenda.objInfoUsu = New ClassPedidoDeVendaHic
    
    With objPedidoVenda.objInfoUsu
        .dValorFrete1 = tPedidoVendaHic.dValorFrete1
        .dValorSeguro1 = tPedidoVendaHic.dValorSeguro1
        .dValorOutrasDespesas1 = tPedidoVendaHic.dValorOutrasDespesas1
        .dValorFrete2 = tPedidoVendaHic.dValorFrete2
        .dValorSeguro2 = tPedidoVendaHic.dValorSeguro2
        .dValorOutrasDespesas2 = tPedidoVendaHic.dValorOutrasDespesas2
        .iFlagCompl1 = tPedidoVendaHic.iFlagCompl1
        .iFlagCompl2 = tPedidoVendaHic.iFlagCompl2
    End With
    
    Call Comando_Fechar(lComando)
    
    PedidoDeVenda_Le_InfoUsu = SUCESSO
     
    Exit Function
    
Erro_PedidoDeVenda_Le_InfoUsu:

    PedidoDeVenda_Le_InfoUsu = gErr
     
    Select Case gErr
          
        Case 124185
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 124186, 124187, 129166, 129167
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PV_INFOUSU", gErr)
        
        Case 124188
            Call Rotina_Erro(vbOKOnly, "ERRO_PV_SEM_INFOUSU", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179223)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function PedidoDeVendaBaixado_LeMod(ByVal lComando As Long, ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda a partir do Código passado.

Dim lErro As Long, tPedidoVendaHic As typePedidoVendaHic
Dim tPedidoVenda As typePedidoVenda

On Error GoTo Erro_PedidoDeVendaBaixado_LeMod

    tPedidoVenda.sMensagemNota = String(STRING_PEDIDOVENDA_MENSAGEM_NOTA, 0)
    tPedidoVenda.sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
    tPedidoVenda.sPedidoCliente = String(STRING_PEDIDOVENDA_PEDIDO_CLIENTE, 0)
    tPedidoVenda.sPedidoRepresentante = String(STRING_PEDIDOVENDA_PEDIDO_REPRESENTANTE, 0)

    '#########################################
    'Alterado por Wagner
    'Inserido PesoLiq, PesoBruto
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresaFaturamento, Cliente, Filial, FilialEntrega, CodTransportadora, CondicaoPagto, PercAcrescFinanceiro, DataEmissao, MensagemNota, NaturezaOp, ValorTotal, ValorFrete, ValorDesconto, ValorSeguro, ValorOutrasDespesas, PedidoCliente, CanalVenda, TabelaPreco, ProxSeqBloqueio, FaturaIntegral, ComissaoAutomatica, CobrancaAutomatica, DataReferencia, NumIntDoc, ValorProdutos, CodTranspRedesp, DetPagFrete, DataEntrega, VolumeTotal, Moeda, TaxaMoeda, ValorFrete1, ValorSeguro1, ValorOutrasDespesas1, ValorFrete2, ValorSeguro2, ValorOutrasDespesas2, FlagCompl1, FlagCompl2, PesoLiq, PesoBruto,PedidoRepresentante FROM PedidosDeVendaBaixados WHERE FilialEmpresa = ? AND Codigo = ? ", 0, tPedidoVenda.iFilialEmpresaFaturamento, tPedidoVenda.lCliente, tPedidoVenda.iFilial, tPedidoVenda.iFilialEntrega, tPedidoVenda.iCodTransportadora, tPedidoVenda.iCondicaoPagto, _
        tPedidoVenda.dPercAcrescFinanceiro, tPedidoVenda.dtDataEmissao, tPedidoVenda.sMensagemNota, tPedidoVenda.sNaturezaOp, tPedidoVenda.dValorTotal, tPedidoVenda.dValorFrete, tPedidoVenda.dValorDesconto, tPedidoVenda.dValorSeguro, tPedidoVenda.dValorOutrasDespesas, tPedidoVenda.sPedidoCliente, tPedidoVenda.iCanalVenda, tPedidoVenda.iTabelaPreco, tPedidoVenda.iProxSeqBloqueio, tPedidoVenda.iFaturaIntegral, tPedidoVenda.iComissaoAutomatica, tPedidoVenda.iCobrancaAutomatica, tPedidoVenda.dtDataReferencia, tPedidoVenda.lNumIntDoc, tPedidoVenda.dValorProdutos, tPedidoVenda.iCodTranspRedesp, tPedidoVenda.iDetPagFrete, tPedidoVenda.dtDataEntrega, _
        tPedidoVenda.dVolumeTotal, tPedidoVenda.iMoeda, tPedidoVenda.dTaxaMoeda, tPedidoVendaHic.dValorFrete1, tPedidoVendaHic.dValorSeguro1, tPedidoVendaHic.dValorOutrasDespesas1, tPedidoVendaHic.dValorFrete2, tPedidoVendaHic.dValorSeguro2, tPedidoVendaHic.dValorOutrasDespesas2, tPedidoVendaHic.iFlagCompl1, tPedidoVendaHic.iFlagCompl2, tPedidoVenda.dPesoLiq, tPedidoVenda.dPesoBruto, tPedidoVenda.sPedidoRepresentante, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 62203
    '#########################################

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62204

    If lErro = AD_SQL_SEM_DADOS Then Error 62205

    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 62206

    objPedidoVenda.iFilialEmpresaFaturamento = tPedidoVenda.iFilialEmpresaFaturamento
    objPedidoVenda.lCliente = tPedidoVenda.lCliente
    objPedidoVenda.iFilial = tPedidoVenda.iFilial
    objPedidoVenda.iFilialEntrega = tPedidoVenda.iFilialEntrega
    objPedidoVenda.iCodTransportadora = tPedidoVenda.iCodTransportadora
    objPedidoVenda.iCodTranspRedesp = tPedidoVenda.iCodTranspRedesp
    objPedidoVenda.iDetPagFrete = tPedidoVenda.iDetPagFrete
    objPedidoVenda.dtDataEntrega = tPedidoVenda.dtDataEntrega
    objPedidoVenda.iCondicaoPagto = tPedidoVenda.iCondicaoPagto
    objPedidoVenda.dPercAcrescFinanceiro = tPedidoVenda.dPercAcrescFinanceiro
    objPedidoVenda.dtDataEmissao = tPedidoVenda.dtDataEmissao
    objPedidoVenda.sMensagemNota = tPedidoVenda.sMensagemNota
    objPedidoVenda.sNaturezaOp = tPedidoVenda.sNaturezaOp
    objPedidoVenda.dValorTotal = tPedidoVenda.dValorTotal
    objPedidoVenda.dValorFrete = tPedidoVenda.dValorFrete
    objPedidoVenda.dValorDesconto = tPedidoVenda.dValorDesconto
    objPedidoVenda.dValorSeguro = tPedidoVenda.dValorSeguro
    objPedidoVenda.dValorOutrasDespesas = tPedidoVenda.dValorOutrasDespesas
    objPedidoVenda.sPedidoCliente = tPedidoVenda.sPedidoCliente
    objPedidoVenda.iCanalVenda = tPedidoVenda.iCanalVenda
    objPedidoVenda.iTabelaPreco = tPedidoVenda.iTabelaPreco
    objPedidoVenda.iProxSeqBloqueio = tPedidoVenda.iProxSeqBloqueio
    objPedidoVenda.iFaturaIntegral = tPedidoVenda.iFaturaIntegral
    objPedidoVenda.iComissaoAutomatica = tPedidoVenda.iComissaoAutomatica
    objPedidoVenda.iCobrancaAutomatica = tPedidoVenda.iCobrancaAutomatica
    objPedidoVenda.dtDataReferencia = tPedidoVenda.dtDataReferencia
    objPedidoVenda.lNumIntDoc = tPedidoVenda.lNumIntDoc
    objPedidoVenda.dValorProdutos = tPedidoVenda.dValorProdutos
    
    objPedidoVenda.dVolumeTotal = tPedidoVenda.dVolumeTotal
    objPedidoVenda.iMoeda = tPedidoVenda.iMoeda
    objPedidoVenda.dTaxaMoeda = tPedidoVenda.dTaxaMoeda
    
    '#######################################
    'Inserido por Wagner
    objPedidoVenda.dPesoBruto = tPedidoVenda.dPesoBruto
    objPedidoVenda.dPesoLiq = tPedidoVenda.dPesoLiq
    '#######################################
    
    Set objPedidoVenda.objInfoUsu = New ClassPedidoDeVendaHic
    
    With objPedidoVenda.objInfoUsu
        .dValorFrete1 = tPedidoVendaHic.dValorFrete1
        .dValorSeguro1 = tPedidoVendaHic.dValorSeguro1
        .dValorOutrasDespesas1 = tPedidoVendaHic.dValorOutrasDespesas1
        .dValorFrete2 = tPedidoVendaHic.dValorFrete2
        .dValorSeguro2 = tPedidoVendaHic.dValorSeguro2
        .dValorOutrasDespesas2 = tPedidoVendaHic.dValorOutrasDespesas2
        .iFlagCompl1 = tPedidoVendaHic.iFlagCompl1
        .iFlagCompl2 = tPedidoVendaHic.iFlagCompl2
    End With
    
    PedidoDeVendaBaixado_LeMod = SUCESSO

    Exit Function

Erro_PedidoDeVendaBaixado_LeMod:

    PedidoDeVendaBaixado_LeMod = Err

    Select Case Err

        Case 62203, 62204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOSDEVENDA", Err, objPedidoVenda.lCodigo)

        Case 62205 'Tratado na rotina chamadora

        Case 62206
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOS_DE_VENDA", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 179224)

    End Select

    Exit Function

End Function


