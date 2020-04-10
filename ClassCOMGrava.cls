VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCOMGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function Alcada_Grava(ByVal objAlcada As ClassAlcada) As Long
'Grava uma alçada

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim sCodigo As String

On Error GoTo Erro_Alcada_Grava

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 49176
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 49177

    sCodigo = String(STRING_ALCADA_CODUSUARIO, 0)

    'verifica se a alçada está cadastrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodUsuario FROM Alcada WHERE CodUsuario = ?", 0, sCodigo, objAlcada.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 49178

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49179

    'se a alçada estiver cadastrada ==> é uma alteração
    If lErro = AD_SQL_SUCESSO Then

        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 49180

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Alcada SET LimiteMensal = ?, LimiteOperacao = ?", alComando(0), objAlcada.dLimiteMensal, objAlcada.dLimiteOperacao)
        If lErro <> AD_SQL_SUCESSO Then Error 49181

    Else

        'a alçada não estava cadastrada ==> é uma inclusão
        lErro = Comando_ExecutarLockado(alComando(2), "SELECT CodUsuario FROM Usuario WHERE CodUsuario = ?", sCodigo, objAlcada.sCodUsuario)
        If lErro <> AD_SQL_SUCESSO Then Error 49182

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49183

        lErro = Comando_LockShared(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then Error 49184

        lErro = Comando_Executar(alComando(0), "INSERT INTO Alcada (CodUsuario, LimiteOperacao, LimiteMensal) VALUES (?, ?, ?)", objAlcada.sCodUsuario, objAlcada.dLimiteOperacao, objAlcada.dLimiteMensal)
        If lErro <> AD_SQL_SUCESSO Then Error 49185

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 49186

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Alcada_Grava = SUCESSO

    Exit Function

Erro_Alcada_Grava:

    Alcada_Grava = Err

    Select Case Err

        Case 49176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49177
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 49178, 49179
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALCADA", Err, objAlcada.sCodUsuario)

        Case 49180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ALCADA", Err, objAlcada.sCodUsuario)

        Case 49181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ALCADA", Err, objAlcada.sCodUsuario)

        Case 49182, 49183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", Err)

        Case 49184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_USUARIOS_DIC", Err)

        Case 49185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ALCADA", Err, objAlcada.sCodUsuario)

        Case 49186
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147440)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Alcada_Exclui(ByVal objAlcada As ClassAlcada) As Long
'Exclui uma alçada

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim sCodigo As String
Dim iStatus As Integer

On Error GoTo Erro_Alcada_Exclui

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 49187
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 49188

    sCodigo = String(STRING_ALCADA_CODUSUARIO, 0)

    'Seleciona a alcada determinada por sCodUsuario
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodUsuario FROM Alcada WHERE CodUsuario = ?", 0, sCodigo, objAlcada.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 49189

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49190

    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then Error 49191

    sCodigo = String(STRING_ALCADA_CODUSUARIO, 0)

    'Verifica se o usuario da alcada esta vinculado a um Pedido de Compra
    lErro = Comando_Executar(alComando(1), "SELECT Alcada, Status FROM PedidoCompraN WHERE Alcada =? ", sCodigo, iStatus, objAlcada.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 49192

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49193

    'Encontrou usuario de Alcada vinculado a PedidoCompra
    If lErro = AD_SQL_SUCESSO Then
    
        If iStatus = PEDIDOCOMPRA_STATUS_ATIVO Then
            Error 49194
        Else
            Error 49197
        End If
        
    End If

    'Verifica se o usuario esta vinculado a um Valor de Pedido de Compra Liberado
    lErro = Comando_Executar(alComando(1), "SELECT CodUsuario FROM ValorPCLiberado WHERE CodUsuario =? ", sCodigo, objAlcada.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 49198

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49199

    'Encontrou usuario de Alcada vinculado a ValorPCLiberado
    If lErro = AD_SQL_SUCESSO Then Error 49200

    'Exclui a alçada do BD
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Alcada", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 49201

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 49202

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Alcada_Exclui = SUCESSO

    Exit Function

Erro_Alcada_Exclui:

    Alcada_Exclui = Err

    Select Case Err

        Case 49187
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 49189, 49190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALCADA", Err, objAlcada.sCodUsuario)

        Case 49191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ALCADA", Err, objAlcada.sCodUsuario)

        Case 49192, 49193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", Err)

        Case 49194
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALCADA_VINCULADA_PEDIDOCOMPRA", Err, objAlcada.sCodUsuario)

        Case 49195, 49196
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRABAIXADO", Err)

        Case 49197
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALCADA_VINCULADA_PEDIDOCOMPRABAIXADO", Err, objAlcada.sCodUsuario)

        Case 49198, 49199
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VALORPCLIBERADO", Err)

        Case 49200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALCADA_VINCULADA_VALORPCLIBERADO", Err, objAlcada.sCodUsuario)

        Case 49201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ALCADA", Err, objAlcada.sCodUsuario)

        Case 49202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147441)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItemPedCompraBaixado_Altera_ItemRC(ByVal lNumIntDoc As Long) As Long
    
Dim alComando(0 To 2) As Long
Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantRecebida As Double
Dim lItemRC As Long
Dim dQuantPedida As Double
Dim iIndice As Integer

On Error GoTo Erro_ItemPedCompraBaixado_Altera_ItemRC

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 84530
    Next
 
    'Lê de ItemRCItemPC todos os campos cujo
    'ItemPC = lNumIntDoc(tabela ItensPedCompra)
    lErro = Comando_Executar(alComando(0), "SELECT ItemRCItemPC.Quantidade, ItemRCItemPC.QuantRecebida, ItemRCItemPC.ItemRC FROM ItemRCItemPC, ItensReqCompra WHERE ItemPC=?  AND ItensReqCompra.NumIntDoc = ItemRC", dQuantidade, dQuantRecebida, lItemRC, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 84531
   
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84532
      
    Do While lErro = SUCESSO
    
        'Lê de ItensRecCompra o campo QuantPedida cujo
        'NumIntDoc = ItemPC (tabela ItemRCItemPC)
        lErro = Comando_ExecutarPos(alComando(1), "SELECT QuantPedida FROM ItensReqCompra WHERE NumIntDoc=? ", 0, dQuantPedida, lItemRC)
        If lErro <> AD_SQL_SUCESSO Then gError 84533
        
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84534
        If lErro = AD_SQL_SEM_DADOS Then gError 84537
        
        'Quantidade pedida vai ser igual a
        'quantidade pedida menos (quantidade menos quantidade recebida)
        dQuantPedida = dQuantPedida - (dQuantidade - dQuantRecebida)
        If dQuantPedida < 0 Then dQuantPedida = 0
        
        'Altera em ItensRecCompra o valor de QuantPedida
        'para o valor de dQuantPedida(calculado acima)
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensReqCompra SET QuantPedida = ? ", alComando(1), dQuantPedida)
        If lErro <> AD_SQL_SUCESSO Then gError 84535
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84538
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItemPedCompraBaixado_Altera_ItemRC = SUCESSO

    Exit Function

Erro_ItemPedCompraBaixado_Altera_ItemRC:

    ItemPedCompraBaixado_Altera_ItemRC = gErr

    Select Case gErr

        Case 84537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGISTRONAOENCONTRADO_ITENSREQCOMPRA", gErr)
        
        Case 84530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 84531, 84532, 84538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMRCITEMPC", gErr)

        Case 84533, 84534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)

        Case 84535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_ITENSREQCOMPRA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147442)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function



End Function

'Function PedidosAbertos_Grava_Estatistica(objPedCompras As ClassPedidoCompras) As Long
''Atualiza FornecedorProdutoFF pela inclusao de um pedido de compras
'
'Dim lErro As Long
'Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
'Dim objItemPedCompra As New ClassItemPedCompra
'Dim alComando(2) As Long
'Dim dQuantPedAbertos As Double
'Dim objProduto As New ClassProduto
'Dim dFator As Double
'Dim dQuantUMCompra As Double
'Dim iIndice As Integer
'
'On Error GoTo Erro_PedidosAbertos_Grava_Estatistica
'
'    'Abre os Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 67023
'    Next
'
'    objFornecedorProdutoFF.lFornecedor = objPedCompras.lFornecedor
'    objFornecedorProdutoFF.iFilialForn = objPedCompras.iFilial
'    objFornecedorProdutoFF.iFilialEmpresa = objPedCompras.iFilialEmpresa
'
'    'Para cada Item de Pedido de Compra
'    For Each objItemPedCompra In objPedCompras.colItens
'
'        'Le o Item em FornecedorProdutoFF
'        lErro = Comando_ExecutarPos(alComando(1), "SELECT QuantPedAbertos FROM FornecedorProdutoFF WHERE FilialEmpresa = ?, Fornecedor = ?, FilialForn =?, Produto = ?", 0, dQuantPedAbertos, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objItemPedCompra.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 67024
'
'        lErro = Comando_BuscarPrimeiro(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67025
'
'        'Se não Encontrar ---> Erro
'        If lErro = AD_SQL_SEM_DADOS Then gError 67026
'
'        'Faz o "lock"
'        lErro = Comando_LockExclusive(alComando(1))
'        If lErro <> AD_SQL_SUCESSO Then gError 67027
'
'        objProduto.sCodigo = objItemPedCompra.sProduto
'
'        'Lê na tabela de Produto a UM Compra e a Classe UM
'        lErro = CF("Produto_Le",objProduto)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError 67028
'
'        'Se não encontrar o produto  - - ->>> Erro
'        If lErro = 28030 Then gError 67029
'
'        'Calcula fator de conversão para UM_COMPRA
'        lErro = CF("UM_Conversao_Trans",objProduto.iClasseUM, objItemPedCompra.sUM, objProduto.sSiglaUMCompra, dFator)
'        If lErro <> SUCESSO Then gError 67030
'
'        'Converte Quantidade para UM_Compra
'        dQuantUMCompra = (objItemPedCompra.dQuantidade * dFator)
'
'        dQuantPedAbertos = dQuantPedAbertos + dQuantUMCompra
'
'        'Atualiza a qdte em pedidos abertos
'        lErro = Comando_ExecutarPos(alComando(2), "UPDATE FornecedorProdutoFF SET QuantPedAbertos", alComando(1), dQuantPedAbertos)
'        If lErro <> AD_SQL_SUCESSO Then gError 67031
'
'    Next
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    PedidosAbertos_Grava_Estatistica = SUCESSO
'
'    Exit Function
'
'Erro_PedidosAbertos_Grava_Estatistica:
'
'    PedidosAbertos_Grava_Estatistica = gErr
'
'    Select Case gErr
'
'        Case 67023
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 67024, 67025
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)
'
'        Case 67026
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objItemPedCompra.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
'
'        Case 67027
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FORNECEDORPRODUTOFF", gErr)
'
'        Case 67028, 67030
'
'        Case 67029
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objItemPedCompra.sProduto)
'
'        Case 67031
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FORNECEDORPRODUTOFF", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objItemPedCompra.sProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147443)
'
'    End Select
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

'Function FornecedorProduto_Grava_EstatisticasCom(objNFiscal As ClassNFiscal, colItensPedComprasFechados As Collection) As Long
'
'Dim lErro As Long
'Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
'Dim objItemNF As New ClassItemNF
'Dim colFornecedorProdutoFF As New Collection
'
'On Error GoTo Erro_FornecedorProduto_Grava_EstatisticasCom
'
'    objFornecedorProdutoFF.lFornecedor = objNFiscal.lFornecedor
'    objFornecedorProdutoFF.iFilialForn = objNFiscal.iFilialForn
'    objFornecedorProdutoFF.iFilialEmpresa = objNFiscal.iFilialEmpresa
'
'    For Each objItemNF In objNFiscal.ColItensNF
'
'        objFornecedorProdutoFF.sProduto = objItemNF.sProduto
'
'        'Calcula a Quantidade em Pedidos Abertos
'        lErro = QuantEmPedAbertos_Calcula(objFornecedorProdutoFF, objItemNF)
'        If lErro <> SUCESSO Then gError 67032
'
'        'Calcula o Tempo de Ressuprimento
'        lErro = CF("TempoRessup_Calcula",objFornecedorProdutoFF, objItemNF)
'        If lErro <> SUCESSO Then gError 67033
'
'        'Calcula o Preco Total do Produto
'        objFornecedorProdutoFF.dPrecoTotal = (objItemNF.dPrecoUnitario * objItemNF.dQuantidade) - (objItemNF.dPrecoUnitario * objItemNF.dQuantidade * objItemNF.dPercDesc)
'
'        'Calcula o Último Pedido de Compra Fechado
'        lErro = UltimoItemPCFechado_Calcula(objFornecedorProdutoFF, objItemNF, colItensPedComprasFechados, objNFiscal)
'        If lErro <> SUCESSO Then gError 67034
'
'        colFornecedorProdutoFF.Add objFornecedorProdutoFF
'
'    Next
'
'    'Grava a Estatistica calculada em FornecedorProduto
'    lErro = FornecedorProduto_Grava_EstatisticasCom_BD(objNFiscal, colFornecedorProdutoFF)
'    If lErro <> SUCESSO Then gError 67035
'
'    FornecedorProduto_Grava_EstatisticasCom = SUCESSO
'
'    Exit Function
'
'Erro_FornecedorProduto_Grava_EstatisticasCom:
'
'    FornecedorProduto_Grava_EstatisticasCom = gErr
'
'    Select Case gErr
'
'        Case 67032, 67033, 67034, 67035
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147444)
'
'    End Select
'
'    Exit Function
'
'End Function

Function QuantEmPedAbertos_Calcula(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF, ByVal objItemNF As ClassItemNF) As Long
'Soma as Quantidade correspondentes a Pedidos de Compras de um item da Nota Fiscal

Dim lErro As Long
Dim objItemNFItemPC As New ClassItemNFItemPC
Dim objProduto As New ClassProduto
Dim dFator As Double

On Error GoTo Erro_QuantEmPedAbertos_Calcula
    
    '??? Jones: deveria ler produto fora de todos os calculos estatisticos e passar os dados como parametro
        
    objProduto.sCodigo = objItemNF.sProduto
    
    'Le no Produto a Unidade de Compras
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 67036
    
    If lErro = 28030 Then gError 67037
    
    For Each objItemNFItemPC In objItemNF.colItemNFItemPC
                
        'Converte as unidades de medida
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMCompra, dFator)
        If lErro <> SUCESSO Then gError 67038
        
        'Acumula as Quantidade de  Pedidos na UM de Compras
        objFornecedorProdutoFF.dQuantPedAbertos = objFornecedorProdutoFF.dQuantPedAbertos + (objItemNFItemPC.dQuantidade * dFator)
    
    Next
    
    QuantEmPedAbertos_Calcula = SUCESSO
    
    Exit Function
    
Erro_QuantEmPedAbertos_Calcula:

    QuantEmPedAbertos_Calcula = gErr
    
    Select Case gErr
        
        Case 67036, 67038
        
        Case 67037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objItemNF.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147445)
        
    End Select
    
    Exit Function

End Function


'Function UltimoItemPCFechado_Calcula(objFornecedorProdutoFF As ClassFornecedorProdutoFF, objItemNF As ClassItemNF, colItensPedComprasFechados As Collection, objNFiscal As ClassNFiscal) As Long
''Calcula os dados do Último ItemPC Fechado (QuantRecebida, Quantidade, DataEntrada, DataEnvio)
''OBS: Em colItensPedComprasFechados entram os Itens de Pedido que estão sendo Fechados com esta Nota (STATUS = FECHADO)
'
'Dim lErro As Long
'Dim alComando(2) As Long
'Dim objItemNFItemPC As New ClassItemNFItemPC
'Dim objItemPedComprasBaixados As New ClassItemPedCompra
'Dim dtDataEnvio As Date
'Dim dQuantidade As Double
'Dim dQuantRecebida As Double
'Dim sUM As String
'Dim bEncontrouInt As Boolean
'Dim bEncontrou As Boolean
'Dim sSQL As String
'Dim sSelect As String
'Dim sWhere As String
'Dim sOrderBy As String
'Dim objProduto As ClassProduto
'Dim dFator As Double
'Dim iIndice As Integer
'
'On Error GoTo Erro_UltimoItemPCFechado_Calcula
'
'    'Abre os Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 67047
'    Next
'
'    'Inicializa o Select
'    sSelect = "SELECT PedidoCompra.DataEnvio, ItensPedCompra.QuantRecebida, ItensPedCompra.Quantidade, ItensPedCompra.UM FROM PedidoCompra, ItensPedCompra WHERE PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra"
'    bEncontrouInt = False
'
'    For Each objItemNFItemPC In objItemNF.colItemNFItemPC
'
'        bEncontrou = False
'
'        'Verifica se está também na coleção de Itens de Pedido Baixados
'        For Each objItemPedComprasBaixados In colItensPedComprasFechados
'             If objItemNFItemPC.lItemPedCompra = objItemPedComprasBaixados.lNumIntDoc Then
'                bEncontrou = True
'                Exit For
'            End If
'        Next
'
'        If bEncontrou = True Then
'
'            'Adiciona dinamicamente os Itens que tem intersecao das 2 colecoes
'            sWhere = sWhere & " AND ItensPedCompra.NumIntDoc = " & objItemNFItemPC.lItemPedCompra
'            bEncontrouInt = True
'
'        End If
'
'    Next
'
'    If bEncontrouInt = True Then
'
'        'Termina a Montagem do Select para ordenar pela data de Envio do Pedido
'        sOrderBy = " ORDER BY PedidoCompra.DataEnvio"
'        sSQL = sSelect & sWhere & sOrderBy
'
'        lErro = Comando_Executar(alComando(1), sSQL, dtDataEnvio, dQuantRecebida, dQuantidade, sUM)
'        If lErro <> AD_SQL_SUCESSO Then gError 67048
'
'        lErro = Comando_BuscarPrimeiro(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67049
'
'        objProduto.sCodigo = objItemNF.sProduto
'
'        'Le no Produto a Unidade de Compras
'        lErro = CF("Produto_Le",objProduto)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError 67050
'
'        If lErro = 28030 Then gError 67051
'
'        'Converte as unidades de medida
'        lErro = CF("UM_Conversao_Trans",objProduto.iClasseUM, sUM, objProduto.sSiglaUMCompra, dFator)
'        If lErro <> SUCESSO Then gError 67052
'
'        objFornecedorProdutoFF.dQuantRecebida = dQuantRecebida * dFator
'        objFornecedorProdutoFF.dQuantPedida = dQuantidade * dFator
'        objFornecedorProdutoFF.dtDataReceb = objNFiscal.dtDataEntrada
'        objFornecedorProdutoFF.dtDataPedido = dtDataEnvio
'
'    End If
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    UltimoItemPCFechado_Calcula = SUCESSO
'
'    Exit Function
'
'Erro_UltimoItemPCFechado_Calcula:
'
'    UltimoItemPCFechado_Calcula = gErr
'
'    Select Case gErr
'
'        Case 67047
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 67048, 67049
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQL)
'
'        Case 67050, 67052
'
'        Case 67051
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objItemNF.sProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147446)
'
'    End Select
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

'Private Function FornecedorProduto_Grava_EstatisticasCom_BD(objNFiscal As ClassNFiscal, colFornecedorProdutoFF As Collection) As Long
''Grava a Estatistica no BD (Notas Fiscais de Compras)
'
'Dim alComando(2) As Long
'Dim lErro As Long
'Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
'Dim iFilialEmpresaTemp As Integer
'Dim iIndice As Integer
'
'On Error GoTo Erro_FornecedorProduto_Grava_EstatisticasCom_BD
'
'    'Abre os Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 67053
'    Next
'
'    For Each objFornecedorProdutoFF In colFornecedorProdutoFF
'
'        lErro = Comando_ExecutarPos(alComando(1), "SELECT FilialEmpresa FROM FornecedorProdutoFF WHERE FilialEmpresa = ?, Fornecedor = ?, FilialForn =?, Produto = ?", 0, iFilialEmpresaTemp, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 67054
'
'        lErro = Comando_BuscarPrimeiro(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67055
'
'        'Se não Encontrar ---> Erro
'        If lErro = AD_SQL_SEM_DADOS Then gError 67056
'
'        'Faz o "lock"
'        lErro = Comando_LockExclusive(alComando(1))
'        If lErro <> AD_SQL_SUCESSO Then gError 67057
'
'        'Atualiza
'        lErro = Comando_ExecutarPos(alComando(2), "UPDATE FornecedorProdutoFF SET QuantPedAbertos = QuantPedAbertos - ?, DataUltimaCompra =?, TempoRessup = ?, PrecoTotal = ?, QuantPedida = ?, QuantRecebida = ?, DataPedido = ?, DataReceb = ?", alComando(1), objFornecedorProdutoFF.dQuantPedAbertos, objNFiscal.dtDataEntrada, objFornecedorProdutoFF.iTempoRessup, objFornecedorProdutoFF.dPrecoTotal, objFornecedorProdutoFF.dQuantPedida, objFornecedorProdutoFF.dQuantRecebida, objFornecedorProdutoFF.dtDataPedido, objFornecedorProdutoFF.dtDataReceb)
'        If lErro <> AD_SQL_SUCESSO Then gError 67058
'
'    Next
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    FornecedorProduto_Grava_EstatisticasCom_BD = SUCESSO
'
'    Exit Function
'
'Erro_FornecedorProduto_Grava_EstatisticasCom_BD:
'
'    FornecedorProduto_Grava_EstatisticasCom_BD = gErr
'
'    Select Case gErr
'
'        Case 67053
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 67054, 67055
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)
'
'        Case 67056
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
'
'        Case 67057
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FORNECEDORPRODUTOFF", gErr)
'
'        Case 67058
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FORNECEDORPRODUTOFF", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147447)
'
'    End Select
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

'Function FornecedorProduto_Grava_Estatisticas(objNFiscal As ClassNFiscal) As Long
''Calcula a Estastistica e Chama a Função de Gravacao
'
'Dim lErro As Long
'Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
'Dim objItemNF As New ClassItemNF
'Dim colFornecedorProdutoFF As New Collection
'
'On Error GoTo Erro_FornecedorProduto_Grava_Estatisticas
'
'    objFornecedorProdutoFF.lFornecedor = objNFiscal.lFornecedor
'    objFornecedorProdutoFF.iFilialForn = objNFiscal.iFilialForn
'    objFornecedorProdutoFF.iFilialEmpresa = objNFiscal.iFilialEmpresa
'
'    For Each objItemNF In objNFiscal.ColItensNF
'
'        objFornecedorProdutoFF.sProduto = objItemNF.sProduto
'
'        'Calcula o Preco Total do Produto
'        objFornecedorProdutoFF.dPrecoTotal = (objItemNF.dPrecoUnitario * objItemNF.dQuantidade) - (objItemNF.dPrecoUnitario * objItemNF.dQuantidade * objItemNF.dPercDesc)
'
'        colFornecedorProdutoFF.Add objFornecedorProdutoFF
'
'    Next
'
'    'Grava a Estatistica calculada em FornecedorProduto
'    lErro = FornecedorProduto_Grava_Estatisticas_BD(objNFiscal, colFornecedorProdutoFF)
'    If lErro <> SUCESSO Then gError 67059
'
'    FornecedorProduto_Grava_Estatisticas = SUCESSO
'
'    Exit Function
'
'Erro_FornecedorProduto_Grava_Estatisticas:
'
'    FornecedorProduto_Grava_Estatisticas = gErr
'
'    Select Case gErr
'
'        Case 67059
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147448)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function FornecedorProduto_Grava_Estatisticas_BD(objNFiscal As ClassNFiscal, colFornecedorProdutoFF As Collection) As Long
''Grava a Estatistica no BD (Notas Fiscais Versao 1)
'
'Dim lErro As Long
'Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
'Dim alComando(2) As Long
'Dim iFilialEmpresaTemp As Integer
'Dim iIndice As Integer
'
'On Error GoTo Erro_FornecedorProduto_Grava_Estatisticas_BD
'
'    'Abre os Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 67060
'    Next
'
'    For Each objFornecedorProdutoFF In colFornecedorProdutoFF
'
'        lErro = Comando_ExecutarPos(alComando(1), "SELECT FilialEmpresa FROM FornecedorProdutoFF WHERE FilialEmpresa = ?, Fornecedor = ?, FilialForn =?, Produto = ?", 0, iFilialEmpresaTemp, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 67061
'
'        lErro = Comando_BuscarPrimeiro(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67062
'
'        'Se não encontrar --> Erro
'        If lErro = AD_SQL_SEM_DADOS Then gError 67063
'
'        'Faz o "lock"
'        lErro = Comando_LockExclusive(alComando(1))
'        If lErro <> AD_SQL_SUCESSO Then gError 67064
'
'        'Atualiza
'        lErro = Comando_ExecutarPos(alComando(2), "UPDATE FornecedorProdutoFF SET DataUltimaCompra =?, PrecoTotal = ?", alComando(1), objNFiscal.dtDataEntrada, objFornecedorProdutoFF.dPrecoTotal)
'        If lErro <> AD_SQL_SUCESSO Then gError 67065
'
'    Next
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    FornecedorProduto_Grava_Estatisticas_BD = SUCESSO
'
'    Exit Function
'
'Erro_FornecedorProduto_Grava_Estatisticas_BD:
'
'    FornecedorProduto_Grava_Estatisticas_BD = gErr
'
'    Select Case gErr
'
'        Case 67060
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 67061, 67062
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)
'
'        Case 67063
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
'
'        Case 67064
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FORNECEDORPRODUTOFF", gErr)
'
'        Case 67065
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FORNECEDORPRODUTOFF", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147449)
'
'    End Select
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

'Function UltimoItemPCFechado_Grava(colPedCompraBaixados As Collection) As Long
''Calcula a Estatistica para o Pedido de Compra Fechado
'
'Dim objPedCompraBaixado As New ClassPedidoCompras
'Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
'Dim objItemPedCompraBaixado As New ClassItemPedCompra
'Dim objProduto As New ClassProduto
'Dim dFator As Double
'Dim dtDataEntrada As Date
'Dim colFornecedorProdutoFF As New Collection
'Dim alComando(2) As Long
'Dim iIndice As Integer
'Dim lErro As Long
'Dim sSQL As String
'
'On Error GoTo Erro_UltimoItemPCFechado_Grava
'
'    'Abre os Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 67066
'    Next
'
'    For Each objPedCompraBaixado In colPedCompraBaixados
'
'        If objPedCompraBaixado.iTipoDestino = TIPO_DESTINO_EMPRESA Then
'
'            objFornecedorProdutoFF.lFornecedor = objPedCompraBaixado.lFornecedor
'            objFornecedorProdutoFF.iFilialForn = objPedCompraBaixado.iFilial
'            objFornecedorProdutoFF.iFilialEmpresa = objPedCompraBaixado.iFilialDestino
'
'            For Each objItemPedCompraBaixado In objPedCompraBaixado.colItens
'
'                If objItemPedCompraBaixado.iStatus = ITEM_PED_COMPRAS_ABERTO Then
'
'                    objFornecedorProdutoFF.sProduto = objItemPedCompraBaixado.sProduto
'
'                    objProduto.sCodigo = objItemPedCompraBaixado.sProduto
'
'                    'Le no Produto a Unidade de Compras
'                    lErro = CF("Produto_Le",objProduto)
'                    If lErro <> SUCESSO And lErro <> 28030 Then gError 67067
'
'                    If lErro = 28030 Then gError 67068
'
'                    'Converte as unidades de medida
'                    lErro = CF("UM_Conversao_Trans",objProduto.iClasseUM, objItemPedCompraBaixado.sUM, objProduto.sSiglaUMEstoque, dFator)
'                    If lErro <> SUCESSO Then gError 67069
'
'                    objFornecedorProdutoFF.dQuantRecebida = objItemPedCompraBaixado.dQuantRecebida * dFator
'                    objFornecedorProdutoFF.dQuantPedida = objItemPedCompraBaixado.dQuantidade * dFator
'
'                    sSQL = "SELECT NFiscal.DataEntrada FROM NFiscal, ItensNFiscal, ItemNFItemPC WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItensNFiscal.NumIntDoc = ItemNFItemPC.ItemNFiscal AND ItemNFItemPC.ItemPedCompra AND ItemNFItemPC.ItemPedCompra = ? AND (TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ?) ORDER BY NFiscal.DataEntrada DESC"
'                    lErro = Comando_Executar(alComando(1), sSQL, dtDataEntrada, objItemPedCompraBaixado.lNumIntDoc, DOCINFO_NFIENCO, DOCINFO_NFEENCO, DOCINFO_NFIEFNCO, DOCINFO_NFEEFNCO)
'                    If lErro <> AD_SQL_SUCESSO Then gError 67070
'
'                    lErro = Comando_BuscarPrimeiro(alComando(1))
'                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67071
'
'                    If lErro = AD_SQL_SEM_DADOS Then
'                        objFornecedorProdutoFF.dtDataReceb = DATA_NULA
'                    Else
'                        objFornecedorProdutoFF.dtDataReceb = dtDataEntrada
'                    End If
'
'                    objFornecedorProdutoFF.dtDataPedido = objPedCompraBaixado.dtDataEnvio
'
'                End If
'
'            Next
'
'        End If
'
'    Next
'
'    colFornecedorProdutoFF.Add objFornecedorProdutoFF
'
'    lErro = UltimoItemPCFechado_Grava_BD(colFornecedorProdutoFF)
'    If lErro <> SUCESSO Then gError 67072
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    UltimoItemPCFechado_Grava = SUCESSO
'
'    Exit Function
'
'Erro_UltimoItemPCFechado_Grava:
'
'    UltimoItemPCFechado_Grava = gErr
'
'    Select Case gErr
'
'        Case 67066
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 67067, 67069, 67072
'
'        Case 67068
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objItemPedCompraBaixado.sProduto)
'
'        Case 67070, 67071
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", gErr, sSQL)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147450)
'
'    End Select
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Function UltimoItemPCFechado_Grava_BD(colFornecedorProdutoFF As Collection) As Long
''Grava a Estatistica do Pedido de Compra Fechado no BD
'
'Dim lErro As Long
'Dim alComando(2) As Long
'Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
'Dim iFilialEmpresaTemp As Integer
'Dim iIndice As Integer
'
'On Error GoTo Erro_UltimoItemPCFechado_Grava_BD
'
'    'Abre os Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 67073
'    Next
'
'    For Each objFornecedorProdutoFF In colFornecedorProdutoFF
'
'        lErro = Comando_ExecutarPos(alComando(1), "SELECT FilialEmpresa FROM FornecedorProdutoFF WHERE FilialEmpresa = ?, Fornecedor = ?, FilialForn =?, Produto = ?", 0, iFilialEmpresaTemp, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 67074
'
'        lErro = Comando_BuscarPrimeiro(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67075
'
'        'Se não encontrar --> Erro
'        If lErro = AD_SQL_SEM_DADOS Then gError 67076
'
'        'Faz o "lock"
'        lErro = Comando_LockExclusive(alComando(1))
'        If lErro <> AD_SQL_SUCESSO Then gError 67077
'
'        'Atualiza
'        lErro = Comando_ExecutarPos(alComando(2), "UPDATE FornecedorProdutoFF SET QuantPedida = ?, QuantRecebida = ?, DataPedido = ?, DataReceb = ?", alComando(1), objFornecedorProdutoFF.dQuantPedida, objFornecedorProdutoFF.dQuantRecebida, objFornecedorProdutoFF.dtDataPedido, objFornecedorProdutoFF.dtDataReceb)
'        If lErro <> AD_SQL_SUCESSO Then gError 67078
'
'    Next
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    UltimoItemPCFechado_Grava_BD = SUCESSO
'
'    Exit Function
'
'Erro_UltimoItemPCFechado_Grava_BD:
'
'    UltimoItemPCFechado_Grava_BD = gErr
'
'    Select Case gErr
'
'        Case 67073
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 67074, 67075
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)
'
'        Case 67076
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
'
'        Case 67077
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FORNECEDORPRODUTOFF", gErr)
'
'        Case 67078
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FORNECEDORPRODUTOFF", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147451)
'
'    End Select
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

'Function UltimaCotacao_Grava(objPedidoCotacao As ClassPedidoCotacao, objCotacao As ClassCotacao) As Long
''Calcula a Estatistica de Cotacao e Grava no BD
'
'Dim lErro As Long
'Dim objItemPedCotacao As New ClassItemPedCotacao
'Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
'Dim objProduto As New ClassProduto
'Dim dFator As Double
'Dim colFornecedorProdutoFF As New Collection
'Dim objItemCotacao As ClassItemCotacao
'
'On Error GoTo Erro_UltimaCotacao_Grava
'
'    If objCotacao.iTipoDestino = TIPO_DESTINO_AUSENTE Or objCotacao.iTipoDestino = TIPO_DESTINO_EMPRESA Then
'
'        For Each objItemPedCotacao In objPedidoCotacao.colItens
'
'            Set objFornecedorProdutoFF = New ClassFornecedorProdutoFF
'
'            objFornecedorProdutoFF.lFornecedor = objPedidoCotacao.lFornecedor
'            objFornecedorProdutoFF.iFilialForn = objPedidoCotacao.iFilial
'
'            If objCotacao.iTipoDestino = TIPO_DESTINO_AUSENTE Then
'                objFornecedorProdutoFF.iFilialEmpresa = objPedidoCotacao.iFilialEmpresa
'            ElseIf objCotacao.iTipoDestino = TIPO_DESTINO_EMPRESA Then
'                objFornecedorProdutoFF.iFilialEmpresa = objCotacao.iFilialEmpresa
'            End If
'
'            objFornecedorProdutoFF.sProduto = objItemPedCotacao.sProduto
'            objFornecedorProdutoFF.iTipoFreteUltimaCotacao = objPedidoCotacao.iTipoFrete
'
'            objProduto.sCodigo = objItemPedCotacao.sProduto
'
'            'Le no Produto a Unidade de Compras
'            lErro = CF("Produto_Le",objProduto)
'            If lErro <> SUCESSO And lErro <> 28030 Then gError 67079
'
'            If lErro = 28030 Then gError 67080
'
'            'Converte as unidades de medida
'            lErro = CF("UM_Conversao_Trans",objProduto.iClasseUM, objItemPedCotacao.sUM, objProduto.sSiglaUMCompra, dFator)
'            If lErro <> SUCESSO Then gError 67081
'
'            objFornecedorProdutoFF.dQuantUltimaCotacao = objItemPedCotacao.dQuantidade * dFator
'
'            If objItemPedCotacao.colItensCotacao.Count = 1 Then
'
'                objFornecedorProdutoFF.iCondPagto = objItemPedCotacao.colItensCotacao(1).iCondPagto
'                objFornecedorProdutoFF.dUltimaCotacao = objItemPedCotacao.colItensCotacao(1).dPrecoUnitario * dFator
'
'            ElseIf objItemPedCotacao.colItensCotacao.Count = 2 Then
'
'                For Each objItemCotacao In objItemPedCotacao.colItensCotacao
'
'                    If objItemCotacao.iCondPagto = CONDPAGTO_VISTA Then
'
'                        objFornecedorProdutoFF.iCondPagto = COD_A_VISTA
'                        objFornecedorProdutoFF.dUltimaCotacao = objItemCotacao.dPrecoUnitario * dFator
'
'                    End If
'                Next
'
'            End If
'
'            objFornecedorProdutoFF.dtDataUltimaCotacao = objPedidoCotacao.dtDataEmissao
'
'            colFornecedorProdutoFF.Add objFornecedorProdutoFF
'
'        Next
'
'    End If
'
'    lErro = Ultima_Cotacao_Grava_BD(colFornecedorProdutoFF)
'    If lErro <> SUCESSO Then gError 67082
'
'    UltimaCotacao_Grava = SUCESSO
'
'    Exit Function
'
'Erro_UltimaCotacao_Grava:
'
'    UltimaCotacao_Grava = gErr
'
'    Select Case gErr
'
'        Case 67079, 67081
'
'        Case 67080
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objItemPedCotacao.sProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147452)
'
'    End Select
'
'End Function

'Function Ultima_Cotacao_Grava_BD(colFornecedorProdutoFF As Collection) As Long
''Garava a Estatistica de Cotacao no BD
'
'Dim lErro As Long
'Dim alComando(2) As Long
'Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
'Dim iFilialEmpresaTemp As Integer
'Dim iIndice As Integer
'
'On Error GoTo Erro_Ultima_Cotacao_Grava_BD
'
'    'Abre os Comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 67082
'    Next
'
'    For Each objFornecedorProdutoFF In colFornecedorProdutoFF
'
'        lErro = Comando_ExecutarPos(alComando(1), "SELECT FilialEmpresa FROM FornecedorProdutoFF WHERE FilialEmpresa = ?, Fornecedor = ?, FilialForn =?, Produto = ?", 0, iFilialEmpresaTemp, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 67083
'
'        lErro = Comando_BuscarPrimeiro(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67084
'
'        'Se não encontrar --> Erro
'        If lErro = AD_SQL_SEM_DADOS Then gError 67085
'
'        'Faz o "lock"
'        lErro = Comando_LockExclusive(alComando(1))
'        If lErro <> AD_SQL_SUCESSO Then gError 67086
'
'        'Atualiza
'        lErro = Comando_ExecutarPos(alComando(2), "UPDATE FornecedorProdutoFF SET TipoFreteUltimaCotacao = ?, QuantUltimaCotacao = ?, CondPagtoUltimaCotacao = ?, UltimaCotacao = ?, DataUltimaCotacao = ? ", alComando(1), objFornecedorProdutoFF.iTipoFreteUltimaCotacao, objFornecedorProdutoFF.dQuantUltimaCotacao, objFornecedorProdutoFF.iCondPagto, objFornecedorProdutoFF.dUltimaCotacao, objFornecedorProdutoFF.dtDataUltimaCotacao)
'        If lErro <> AD_SQL_SUCESSO Then gError 67087
'
'    Next
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Ultima_Cotacao_Grava_BD = SUCESSO
'
'    Exit Function
'
'Erro_Ultima_Cotacao_Grava_BD:
'
'    Ultima_Cotacao_Grava_BD = gErr
'
'    Select Case gErr
'
'        Case 67082
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 67083, 67084
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)
'
'        Case 67085
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
'
'        Case 67086
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FORNECEDORPRODUTOFF", gErr)
'
'        Case 67087
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FORNECEDORPRODUTOFF", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147453)
'
'    End Select
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Function PedidoComprasBaixar_Batch(ByVal colPedCompras As Collection) As Long
'transfere os Pedidos de Compra p/Pedidos de Compra Baixados

Dim lErro As Long
Dim objPedidoCompra As ClassPedidoCompras
Dim alComando(0 To 5) As Long
Dim iIndice As Integer

On Error GoTo Erro_PedidoComprasBaixar_Batch

    'Percorre a coleção de Pedidos de Compra
    For Each objPedidoCompra In colPedCompras

        'Abre os comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then Error 53152
        Next
        
        'Baixa o Pedido de Compra
        lErro = CF("PedidoCompra_Baixar1", objPedidoCompra, alComando())
        If lErro <> SUCESSO Then Error 57006

        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    Next

    PedidoComprasBaixar_Batch = SUCESSO

    Exit Function

Erro_PedidoComprasBaixar_Batch:

    PedidoComprasBaixar_Batch = Err

    Select Case Err

        Case 57006
        
        Case 53152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147454)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoCompra_Baixar1(ByVal objPedidoCompra As ClassPedidoCompras, alComando() As Long) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim lCodigo As Long
Dim lNumIntDoc As Long
Dim iValorAlterado As Integer
Dim objItemPC As New ClassItemPedCompra
Dim tItemPedido As typeItemPedCompra
Dim iTipoBaixa As Integer
Dim sMotivoBaixa As String
Dim objFilFornFilEmp As New ClassFilFornFilEmp
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_PedidoCompra_Baixar1

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 53153

    iTipoBaixa = objPedidoCompra.iTipoBaixa
    sMotivoBaixa = objPedidoCompra.sMotivoBaixa
    
    'le e locka o pedido de compra
    lErro = CF("PedidoCompra_LeMod", objPedidoCompra, alComando(0))
    If lErro <> SUCESSO And lErro <> 62646 Then gError 53154

    'Nao encontrou ---> erro
    If lErro <> SUCESSO Then gError 53155

    objPedidoCompra.dtDataBaixa = gdtDataAtual
    objPedidoCompra.iTipoBaixa = iTipoBaixa
    objPedidoCompra.sMotivoBaixa = sMotivoBaixa
    
    tItemPedido.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedido.sProduto = String(STRING_PRODUTO, 0)
    tItemPedido.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tItemPedido.sUM = String(STRING_UM_SIGLA, 0)

    'Ler os itens relacionados aquele pedido
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc,DataLimite,Produto,DescProduto,Quantidade,QuantRecebida,QuantRecebimento,UM,PrecoUnitario,ValorDesconto,TipoOrigem,NumIntOrigem,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa,Status,Observacao,ValorIPI,AliquotaIPI,AliquotaICMS FROM ItensPedCompraN WHERE PedCompra = ? AND StatusBaixa = ?", 0, tItemPedido.lNumIntDoc, tItemPedido.dtDataLimite, tItemPedido.sProduto, tItemPedido.sDescProduto, tItemPedido.dQuantidade, tItemPedido.dQuantRecebida, tItemPedido.dQuantRecebimento, tItemPedido.sUM, tItemPedido.dPrecoUnitario, tItemPedido.dValorDesconto, tItemPedido.iTipoOrigem, tItemPedido.lNumIntOrigem, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRebebForaFaixa, tItemPedido.iStatus, tItemPedido.lObservacao, tItemPedido.dValorIPI, tItemPedido.dAliquotaIPI, tItemPedido.dAliquotaICMS, objPedidoCompra.lNumIntDoc, ITENSPEDCOMPRA_STATUS_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 53156

    'Buscar primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53159
    If lErro = AD_SQL_SEM_DADOS Then gError 53160

    'Enquanto sucesso
    Do While lErro = AD_SQL_SUCESSO

        Call Transfere_Dados_ItemPC(objItemPC, tItemPedido)

        'Lock exclusive
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 53161
               
        'Exclui o item de ItensPedCompra
        lErro = Comando_ExecutarPos(alComando(4), "UPDATE ItensPedCompraN SET StatusBaixa=?", alComando(1), ITENSPEDCOMPRA_STATUS_BAIXADO)
        If lErro <> AD_SQL_SUCESSO Then gError 53162

        'Altera a Tabela ItensReqCompra
        lErro = ItemPedCompraBaixado_Altera_ItemRC(objItemPC.lNumIntDoc)
        If lErro <> SUCESSO Then gError 31150
 
        ' Buscar próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53163

    Loop

    'Exclui o item de PedidoCompra
    lErro = Comando_ExecutarPos(alComando(5), "UPDATE PedidoCompraN SET Status = ?, DataBaixa = ?", alComando(0), PEDIDOCOMPRA_STATUS_BAIXADO, objPedidoCompra.dtDataBaixa)
    If lErro <> AD_SQL_SUCESSO Then gError 53164
    
    If objPedidoCompra.dtDataEnvio <> DATA_NULA Then
    
        objFilFornFilEmp.lCodFornecedor = objPedidoCompra.lFornecedor
        objFilFornFilEmp.iCodFilial = objPedidoCompra.iFilial
        objFilFornFilEmp.iFilialEmpresa = objPedidoCompra.iFilialEmpresa
        objFilFornFilEmp.lPedidosEmAberto = objFilFornFilEmp.lPedidosEmAberto - 1

        'Grava a Estatística da Filial Fornecedor
        lErro = CF("FilialFornecedor_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_FILIALFORNFILEMP, objFilFornFilEmp)
        If lErro <> SUCESSO Then gError 89431
    
    End If
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 53165

    PedidoCompra_Baixar1 = SUCESSO

    Exit Function

Erro_PedidoCompra_Baixar1:

    PedidoCompra_Baixar1 = gErr

    Select Case gErr
        
        Case 53153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 53154, 53157, 31150, 89431

        Case 53155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", lErro, objPedidoCompra.lCodigo)

        Case 53156, 53159, 53163
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", lErro, objPedidoCompra.lCodigo)

        Case 53158
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPEDCOMPRABAIXADOS", gErr, tItemPedido.lNumIntDoc)

        Case 53160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", lErro, objItemPC.lNumIntDoc)

        Case 53161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDCOMPRA", gErr)

        Case 53162
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPEDCOMPRA", gErr, objItemPC.lNumIntDoc)

        Case 53164
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PEDIDOCOMPRA", gErr, objPedidoCompra.lCodigo)

        Case 53165
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147455)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function



Function PedidoCotacao_Baixar(ByVal objPedCotacao As ClassPedidoCotacao, alComando() As Long) As Long
'Baixa o Pedido de Cotação passado em objPedCotacao

Dim lErro As Long
Dim lNumIntItem As Long
Dim lPedidoCotacao As Long
Dim lCotacaoProduto As Long
Dim tPedidoCotacao As typePedidoCotacao
Dim lPedCompra As Long
Dim lTransacao As Long
Dim lCotacao As Long
Dim tItemCotacao As typeItemCotacao
Dim lItem As Long

On Error GoTo Erro_PedidoCotacao_Baixar
    
    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 67608
    
    tPedidoCotacao.sContato = String(STRING_CONTATO, 0)
        
    'Lê a Cotação vinculado ao Pedido de Cotação
    lErro = Comando_Executar(alComando(0), "SELECT DISTINCT Cotacao.NumIntDoc FROM Cotacao, CotacaoProduto, ItemPedCotacao, PedidoCotacao WHERE PedidoCotacao.NumIntDoc = ItemPedCotacao.PedidoCotacao AND ItemPedCotacao.CotacaoProduto = CotacaoProduto.NumIntDoc AND CotacaoProduto.Cotacao = Cotacao.NumIntDoc AND PedidoCotacao.Codigo = ? AND PedidoCotacao.FilialEmpresa = ?", lCotacao, objPedCotacao.lCodigo, objPedCotacao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 67826
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67827
    
    'Se não encontrou a Cotação, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67852
    
    'Verifica se a Cotação vinculada ao pedido de Cotação a ser baixado possui outros vínculos com outros Pedidos de Cotação
    lErro = Comando_Executar(alComando(1), "SELECT DISTINCT Cotacao.NumIntDoc FROM Cotacao, CotacaoProduto, ItemPedCotacao, PedidoCotacao WHERE PedidoCotacao.NumIntDoc = ItemPedCotacao.PedidoCotacao AND ItemPedCotacao.CotacaoProduto = CotacaoProduto.NumIntDoc AND CotacaoProduto.Cotacao = Cotacao.NumIntDoc AND NOT (PedidoCotacao.Codigo = ? AND PedidoCotacao.FilialEmpresa = ?) AND Cotacao.NumIntDoc = ?", lCotacao, objPedCotacao.lCodigo, objPedCotacao.iFilialEmpresa, lCotacao)
    If lErro <> AD_SQL_SUCESSO Then gError 67853
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67854
    
    'Se não encontrou Pedidos de Cotação vinculados a Cotação
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Baixa a Cotação
        lErro = CF("Cotacao_Baixa", lCotacao, alComando())
        If lErro <> SUCESSO Then gError 67830
        
    End If
    
    'Lê o Pedido de Cotação passado pelo código e Filial Empresa
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, FilialEmpresa, Codigo, Fornecedor, Filial, Contato, DataEmissao, Data, DataValidade, TipoFrete, Status, CondPagtoPrazo FROM PedidoCotacaoN WHERE Codigo = ? AND FilialEmpresa = ? AND StatusBaixa = ?", 0, _
    tPedidoCotacao.lNumIntDoc, tPedidoCotacao.iFilialEmpresa, tPedidoCotacao.lCodigo, tPedidoCotacao.lFornecedor, tPedidoCotacao.iFilial, tPedidoCotacao.sContato, tPedidoCotacao.dtDataEmissao, tPedidoCotacao.dtData, tPedidoCotacao.dtDataValidade, tPedidoCotacao.iTipoFrete, tPedidoCotacao.iStatus, tPedidoCotacao.iCondPagtoPrazo, objPedCotacao.lCodigo, objPedCotacao.iFilialEmpresa, PEDIDOCOTACAO_STATUS_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 67609
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And AD_SQL_SEM_DADOS Then gError 67610
    
    'Se não encontrou o Pedido de Cotação, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67611
    
    'Faz "Lock" no Pedido de Cotação
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 67612
    
    'Lê os Itens do Pedido de Cotação
    lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc, CotacaoProduto, PedidoCotacao FROM ItemPedCotacaoN WHERE PedidoCotacao = ? AND Status=?", 0, lNumIntItem, lCotacaoProduto, lPedidoCotacao, tPedidoCotacao.lNumIntDoc, ITEMPEDCOTACAO_STATUS_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 67613
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67614
       
    'Se não encontrou o item
    If lErro = AD_SQL_SEM_DADOS Then gError 67616
    
    'Enquanto encontrar Itens
    Do While lErro = AD_SQL_SUCESSO
    
        'Verifica se o ItemPedCotacao possui vínculo com Pedido de Compra não Baixado
        lErro = Comando_Executar(alComando(3), "SELECT ItensPedCompra.PedCompra FROM ItemPedCotacao, ItensCotacao, ItensPedCompra WHERE ItemPedCotacao.NumIntDoc = ItensCotacao.ItemPedCotacao AND ItensCotacao.NumIntDoc = ItensPedCompra.NumIntOrigem AND ItensPedCompra.TipoOrigem = ? AND ItemPedCotacao = ?", lPedCompra, TIPO_ORIGEM_PEDCOTACAO, lNumIntItem)
        If lErro <> AD_SQL_SUCESSO Then gError 67617
        
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SEM_DADOS Then gError 67618
        
        'Se encontrou vínculo, Erro
        If lErro = AD_SQL_SUCESSO Then gError 67619
        
        'Verifica se o ItemPedCotacao possui vínculo com Concorrência não baixada
        lErro = Comando_Executar(alComando(4), "SELECT ItemPedCotacao.PedidoCotacao FROM ItemPedCotacao, ItensCotacao, CotacaoItemConcorrencia WHERE ItemPedCotacao.NumIntDoc = ItensCotacao.ItemPedCotacao AND ItensCotacao.NumIntDoc = CotacaoItemConcorrencia.ItemCotacao AND ItemPedCotacao.NumIntDoc = ?", lPedidoCotacao, lNumIntItem)
        If lErro <> AD_SQL_SUCESSO Then gError 67620
        
        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67621
        
        'Se encontrou vínculo, Erro
        If lErro = AD_SQL_SUCESSO Then gError 67622
        
        'Lê os Itens de Cotacao ligados ao ItemPedCotacao já lido
        lErro = Comando_ExecutarPos(alComando(9), "SELECT NumIntDoc,ItemPedCotacao,CondPagto,DataReferencia,PrecoUnitario,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,AliquotaIPI,AliquotaICMS,PrazoEntrega,QuantEntrega,Observacao FROM ItensCotacaoN WHERE ItemPedCotacao=? AND Status=?", 0, tItemCotacao.lNumIntDoc, lItem, tItemCotacao.iCondPagto, tItemCotacao.dtDataReferencia, tItemCotacao.dPrecoUnitario, tItemCotacao.dOutrasDespesas, tItemCotacao.dValorFrete, tItemCotacao.dValorSeguro, tItemCotacao.dValorDesconto, tItemCotacao.dValorTotal, tItemCotacao.dValorIPI, tItemCotacao.dAliquotaIPI, tItemCotacao.dAliquotaICMS, tItemCotacao.iPrazoEntrega, tItemCotacao.dQuantEntrega, tItemCotacao.lObservacao, lNumIntItem, ITENSCOTACAO_STATUS_ATIVO)
        If lErro <> AD_SQL_SUCESSO Then gError 74923
        
        lErro = Comando_BuscarPrimeiro(alComando(9))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 74924
        
        Do While lErro = AD_SQL_SUCESSO
        
            lErro = Comando_LockExclusive(alComando(9))
            If lErro <> AD_SQL_SUCESSO Then gError 74926
        
            'Exclui Item Cotacao
            lErro = Comando_ExecutarPos(alComando(10), "UPDATE ItensCotacaoN SET Status = ?", alComando(9), ITENSCOTACAO_STATUS_BAIXADO)
            If lErro <> AD_SQL_SUCESSO Then gError 74927
        
            lErro = Comando_BuscarProximo(alComando(9))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 74925
            
            
        Loop
        
        'Faz "Lock" no item
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 67623
        
        'Exclui ItemPedCotação
        lErro = Comando_ExecutarPos(alComando(5), "UPDATE ItemPedCotacaoN SET Status=?", alComando(2), ITEMPEDCOTACAO_STATUS_BAIXADO)
        If lErro <> AD_SQL_SUCESSO Then gError 67624
        
        'Busca próximo Item do Pedido de Cotação
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67615
    
    Loop

    'Atualiza Pedido de Cotação
    lErro = Comando_ExecutarPos(alComando(7), "UPDATE PedidoCotacaoN SET DataBaixa = ?, StatusBaixa = ?", alComando(1), gdtDataHoje, PEDIDOCOTACAO_STATUS_BAIXADO)
    If lErro <> AD_SQL_SUCESSO Then gError 67626
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 67628
    
    PedidoCotacao_Baixar = SUCESSO
    
    Exit Function
    
    PedidoCotacao_Baixar = gErr

Erro_PedidoCotacao_Baixar:

    Select Case gErr
    
        Case 67608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 67609, 67610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAO", gErr, objPedCotacao.lCodigo)
            
        Case 67611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", gErr, objPedCotacao.lCodigo)
                    
        Case 67612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOCOTACAO", gErr)
                
        Case 67613, 67614, 67615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDCOTACAO", gErr, objPedCotacao.lCodigo)
                
        Case 67616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_PEDIDOCOTACAO", gErr)
        
        Case 67617, 67618
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)
        
        Case 67619
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDCOTACAO_VINCULADO_PEDCOMPRA", gErr, objPedCotacao.lCodigo, lPedCompra)
                    
        Case 67620, 67621
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOITEMCONCORRENCIA", gErr)
                    
        Case 67622
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDCOTACAO_VINCULADO_CONCORRENCIA", gErr, objPedCotacao.lCodigo)
                
        Case 67623
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMPEDCOTACAO", gErr, objPedCotacao.lCodigo)
        
        Case 67624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITEMPEDCOTACAO", gErr)
            
        Case 67625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMPEDCOTACAOBAIXADO", gErr)
                
        Case 67626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOCOTACAO", gErr, objPedCotacao.lCodigo)
                    
        Case 67627
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PEDIDOCOTACAOBAIXADO", gErr)
            
        Case 67628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 67826, 67827, 67853, 67854
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO", gErr)
        
        Case 67830
        
        Case 67852
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COTACAO_VINCULADA_PEDIDOCOT_NAO_CADASTRADA", gErr, objPedCotacao.lCodigo)
            
        Case 74923, 74924, 74925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCOTACAO", gErr)
            
        Case 74926
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSCOTACAO", gErr)
            
        Case 74927
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSCOTACAO1", gErr)
            
        Case 74928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSCOTACAOBAIXADOS", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147456)
            
    End Select
        
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function Cotacao_Baixa(ByVal lCotacao As Long, alComando() As Long) As Long
'Baixa a Cotação e seus Itens

Dim lErro As Long
Dim lCotacaoProduto As Long
Dim lItemCotacao As Long
Dim lItemPedCotacao As Long
Dim tCotacao As typeCotacao
Dim tCotacaoProduto As typeCotacaoProduto
Dim tItemCotacao As typeItemCotacao
Dim iStatus As Integer

On Error GoTo Erro_Cotacao_Baixa

    tCotacao.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)
    
    'Lê os dados da Cotação
    lErro = Comando_ExecutarPos(alComando(9), "SELECT NumIntDoc, FilialEmpresa, Codigo, Descricao, Data, TipoDestino, FornCliDestino, FilialDestino, Comprador, Status FROM CotacaoN WHERE NumIntDoc = ?", 0, _
    tCotacao.lNumIntDoc, tCotacao.iFilialEmpresa, tCotacao.lCodigo, tCotacao.sDescricao, tCotacao.dtData, tCotacao.iTipoDestino, tCotacao.lFornCliDestino, tCotacao.iFilialDestino, tCotacao.iComprador, iStatus, lCotacao)
    If lErro <> AD_SQL_SUCESSO Then gError 67831
        
    lErro = Comando_BuscarPrimeiro(alComando(9))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67832
    
    'Se não encontrou a Cotação, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67835
    
    If iStatus = COTACAO_STATUS_BAIXADO Then gError 89483
    
    'Faz "lock" em Cotação
    lErro = Comando_LockExclusive(alComando(9))
    If lErro <> AD_SQL_SUCESSO Then gError 67833
    
    tCotacaoProduto.sProduto = String(STRING_PRODUTO, 0)
    tCotacaoProduto.sUM = String(STRING_UM_SIGLA, 0)
    
    'Lê as CotaçõesProduto relacionadas a essa Cotação
    lErro = Comando_ExecutarPos(alComando(10), "SELECT NumIntDoc, Cotacao, Produto, Quantidade, UM, Fornecedor, Filial FROM CotacaoProdutoN WHERE Cotacao =? AND Status=?", 0, _
    tCotacaoProduto.lNumIntDoc, tCotacaoProduto.lCotacao, tCotacaoProduto.sProduto, tCotacaoProduto.dQuantidade, tCotacaoProduto.sUM, tCotacaoProduto.lFornecedor, tCotacaoProduto.iFilial, lCotacao, COTACAOPRODUTO_STATUS_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 67834
    
    lErro = Comando_BuscarPrimeiro(alComando(10))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67836
    
    'Enquanto encontrar Cotação Produto
    Do While lErro = AD_SQL_SUCESSO
                
        'Faz "lock" na CotacaoProduto
        lErro = Comando_LockExclusive(alComando(10))
        If lErro <> AD_SQL_SUCESSO Then gError 67838
                
        'Exclui a CotaçãoProduto
        lErro = Comando_ExecutarPos(alComando(15), "UPDATE CotacaoProdutoN SET Status=?", alComando(10), COTACAOPRODUTO_STATUS_BAIXADO)
        If lErro <> AD_SQL_SUCESSO Then gError 67848
        
        'Busca próximo CotaçãoProduto
        lErro = Comando_BuscarProximo(alComando(10))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67837
    
    Loop
    
    'Exclui a Cotação
    lErro = Comando_ExecutarPos(alComando(17), "UPDATE CotacaoN SET Status=?", alComando(9), COTACAO_STATUS_BAIXADO)
    If lErro <> AD_SQL_SUCESSO Then gError 67850
    
    Cotacao_Baixa = SUCESSO
    
    Exit Function
    
Erro_Cotacao_Baixa:

    Cotacao_Baixa = gErr
    
    Select Case gErr
            
        Case 67831, 67832
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO", gErr)
                    
        Case 67833
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_COTACAO", gErr)
        
        Case 67834, 67836, 67837
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTO", gErr)
        
        Case 67835
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COTACAO_NAO_CADASTRADA", gErr, lCotacao)
            
        Case 67838
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_COTACAOPRODUTO", gErr)
                        
        Case 67848
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COTACAOPRODUTO", gErr)
            
        Case 67849
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAOPRODUTOBAIXADO", gErr)
                
        Case 67850
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COTACAO", gErr, lCotacao)
                    
        Case 67851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAOBAIXADA", gErr)
            
        Case 89483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COTACAO_BAIXADA", gErr, lCotacao)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147457)
        
    End Select
    
    Exit Function
    
End Function

Function BaixaPedCotacao_Baixar_Pedidos(ByVal colPedCotacao As Collection) As Long
'Faz baixa dos Pedidos de Cotação passados na coleção colPedCotacao

Dim lErro As Long
Dim alComando(18) As Long
Dim iIndice As Integer
Dim objPedCotacao As ClassPedidoCotacao

On Error GoTo Erro_BaixaPedCotacao_Baixar_Pedidos

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 67606
    Next
    
    'Para cada Pedido de cotação
    For Each objPedCotacao In colPedCotacao
        
        'Baixa o Pedido de cotação
        lErro = CF("PedidoCotacao_Baixar", objPedCotacao, alComando())
        If lErro <> SUCESSO Then gError 67607
    
        'Fechamento de comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
        'Abre os Comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir
            If alComando(iIndice) = 0 Then gError 67606
        Next
    Next
    
    'Fechamento de comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    BaixaPedCotacao_Baixar_Pedidos = SUCESSO
    
    Exit Function
    
Erro_BaixaPedCotacao_Baixar_Pedidos:
    
    BaixaPedCotacao_Baixar_Pedidos = gErr
    
    Select Case gErr
            
        Case 67606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 67607
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147458)
            
    End Select
    
    'Fechamento de comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Function Comprador_Exclui(ByVal objComprador As ClassComprador) As Long
'Exclui o comprador da tabela

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim lTransacao As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Comprador_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 50100
    Next

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 50101

    'Seleciona o comprador no bd
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Comprador WHERE CodUsuario = ?", 0, iCodigo, objComprador.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 50103

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 50104

    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then Error 50105

    'Verifica se o comprador possui relacionamentos com outras tabelas
    lErro = CF("Comprador_Existe_Relacionamento", objComprador)
    If lErro <> SUCESSO Then Error 50107

    'Exclui a comprador do BD
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Comprador", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 50106

   'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 50102

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Comprador_Exclui = SUCESSO

    Exit Function

Erro_Comprador_Exclui:

    Comprador_Exclui = Err

    Select Case Err

        Case 50100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 50101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 50103, 50104
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRADOR1", Err, objComprador.sCodUsuario)

        Case 50105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_COMPRADOR", Err)

        Case 50106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMPRADOR", Err, objComprador.sCodUsuario)

        Case 50107

        Case 50102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147459)

    End Select

    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Comprador_Grava(ByVal objComprador As ClassComprador) As Long
'Grava o Comprador objComprador

Dim lErro As Long
Dim lTransacao As Long
Dim lTransacaoDic As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim sCodUsuario As String
Dim sNomeReduzido As String
Dim iAumentaQuant As Integer

On Error GoTo Erro_Comprador_Grava

    lTransacaoDic = Transacao_AbrirDic
    If lTransacaoDic = 0 Then Error 50091
    
    'Abre a transação e comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 50088

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 50089

    sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    sNomeReduzido = String(STRING_USUARIO_NOMEREDUZIDO, 0)

    'abre comando para que seja consultado o dicdados
    lComando2 = Comando_AbrirExt(GL_lConexaoDic)
    If lComando2 = 0 Then Error 50090

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 50091

    
    'Seleciona o Comprador com o código passado
    lErro = Comando_ExecutarPos(lComando, "SELECT CodUsuario, AumentaQuant FROM Comprador WHERE Codigo= ?", 0, sCodUsuario, iAumentaQuant, objComprador.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 50095

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50096

    'Se nao encontrar
    If lErro <> SUCESSO Then

        'Seleciona o usuario
        lErro = Comando_ExecutarLockado(lComando2, "SELECT CodUsuario,NomeReduzido FROM Usuarios WHERE CodUsuario=?", sCodUsuario, sNomeReduzido, objComprador.sCodUsuario)
        If lErro <> AD_SQL_SUCESSO Then Error 50138

        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50092

        objComprador.sNomeReduzido = sNomeReduzido
        
        lErro = Comando_LockShared(lComando2)
        If lErro <> AD_SQL_SUCESSO Then Error 50093

        lErro = Comando_Executar(lComando, "INSERT INTO Comprador (Codigo, CodUsuario, AumentaQuant, FilialEmpresa,Email) VALUES (?, ?, ?, ?,?)", objComprador.iCodigo, objComprador.sCodUsuario, objComprador.iAumentaQuant, giFilialEmpresa, objComprador.sEmail)
        If lErro <> AD_SQL_SUCESSO Then Error 50099

    Else

        'Se encontrar

        'Verifica se o usuario foi alterado
        If sCodUsuario <> objComprador.sCodUsuario Then Error 50121

        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 50097

        lErro = Comando_ExecutarPos(lComando1, "UPDATE Comprador SET AumentaQuant = ?, Email = ?", lComando, objComprador.iAumentaQuant, objComprador.sEmail)
        If lErro <> AD_SQL_SUCESSO Then Error 50098


    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 50094

    Call Transacao_RollbackDic
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Comprador_Grava = SUCESSO

    Exit Function

Erro_Comprador_Grava:

    Comprador_Grava = Err

    Select Case Err

        Case 50088, 50089, 50090 'Erro Abertura de comando
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 50091 'Erro abertura Transacao
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 50095, 50096 'Erro de leitura na tabela Comprador
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRADOR", Err, objComprador.iCodigo)

        Case 50097 'Erro na tentativa de fazer 'lock' na tabela Comprador
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_COMPRADOR", Err)

        Case 50098 'Erro na tentativa de atualizar a tabela Comprador.
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COMPRADOR", Err)

        Case 50092, 50138 'Erro de leitura na tabela Usuarios.
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", Err)

        Case 50093 'Erro na tentativa de fazer 'lock' na tabela Usuarios.
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_USUARIOS_DIC", Err)

        Case 50099 'Erro na tentiva de inserir um registro na tabela Comprador.
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMPRADOR", Err)

        Case 50094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 50121
            lErro = Rotina_Erro(vbOKOnly, "USUARIO_COMPRADOR_NAO_ALTERAVEL", Err, objComprador.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147460)

    End Select

    Call Transacao_Rollback
    Call Transacao_RollbackDic
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function ConfiguraCOM_GravarTrans(ByVal objConfiguraCOM As ClassConfiguraCOM) As Long

Dim lErro As Long
Dim colComprasConfig As New colComprasConfig
Dim objComprasConfig As ClassComprasConfig
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim sCodigo As String
Dim dtDataLib As Date
Dim dtDataNula As Date
Dim iTipoBloqueio As Integer
Dim iFilialEmpresa As Integer

On Error GoTo Erro_ConfiguraCOM_GravarTrans

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 49346
    Next

    'Verifica se não será feito controle de alcada
    If (objConfiguraCOM.iControleAlcada = DESMARCADO And gobjCOM.iControleAlcada = MARCADO) Then
    
        dtDataNula = DATA_NULA
        'Verifica se existe BloqueioPC com Tipo de Bloqueio = Alcada
        lErro = Comando_Executar(alComando(0), "SELECT FilialEmpresa,DataLib,TipoDeBloqueio FROM BloqueiosPC WHERE DataLib=? AND TipoDeBloqueio=? ", iFilialEmpresa, dtDataLib, iTipoBloqueio, dtDataNula, BLOQUEIO_ALCADA)
        If lErro <> AD_SQL_SUCESSO Then gError 72494
        
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 72495
        
        'Se encontrou bloqueio do tipo alcada==> erro
        If lErro = AD_SQL_SUCESSO Then gError 72496
    
    End If
    lErro = CF("ConfiguraCOM_MoverCampos_ColComprasConfig", objConfiguraCOM, colComprasConfig)
    If lErro <> SUCESSO Then gError 49348

    For Each objComprasConfig In colComprasConfig

        sCodigo = String(STRING_CONTEUDO, 0)

        'Seleciona na tabela de ComprasConfig o codigo passado como parametro
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM ComprasConfig WHERE Codigo = ? AND FilialEmpresa =? ", 0, sCodigo, objComprasConfig.sCodigo, objComprasConfig.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 49349

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 49350
        
        'Nao encontrou o registro
        If lErro = AD_SQL_SEM_DADOS Then
        
            'trata valores que podem só estar disponiveis apos a inicializacao do modulo
            Select Case objComprasConfig.sCodigo
            
                Case "NFISCAL_DIFERENTE_PED_COMPRA"
                
                Case Else
                    gError 56199
    
            End Select
            
        Else
        
            'É uma alteracao => atualiza o campo Conteudo da tabela ComprasConfig
            If lErro = AD_SQL_SUCESSO Then
    
                'Atualiza a tabela de ComprasConfig
                lErro = Comando_ExecutarPos(alComando(1), "UPDATE ComprasConfig SET Conteudo = ? ", alComando(0), objComprasConfig.sConteudo)
                If lErro <> AD_SQL_SUCESSO Then gError 49351
    
            End If
            
        End If
        
    Next

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    ConfiguraCOM_GravarTrans = SUCESSO

    Exit Function

Erro_ConfiguraCOM_GravarTrans:

    ConfiguraCOM_GravarTrans = gErr

    Select Case gErr

        Case 49346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 49348

        Case 49349, 49350
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPRASCONFIG", gErr, objComprasConfig.sCodigo)

        Case 49351
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COMPRASCONFIG", gErr, objComprasConfig.sCodigo)

        Case 56199
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_COMPRAS_CONFIG_NAO_ENCONTRADO", gErr, objComprasConfig.sCodigo, objComprasConfig.iFilialEmpresa)
            
        Case 72494, 72495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPC", gErr, BLOQUEIO_ALCADA)
            
        Case 72496
            'Erro tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147461)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Function Concorrencia_Automatica(lConcorrencia As Long) As Long
'Devolve o proximo código de concorrencia disponivel para a FilialEmpresa em questao.

Dim lErro As Long

On Error GoTo Erro_Concorrencia_Automatica

    'Obtem o proximo codigo de Concorrencia disponivel
    lErro = CF("Config_ObterAutomatico", "ComprasConfig", "NUM_PROXIMO_CODIGO_CONCORRENCIA", "Concorrencia", "Codigo", lConcorrencia)
    If lErro <> SUCESSO Then Error 63743

    Concorrencia_Automatica = SUCESSO

    Exit Function

Erro_Concorrencia_Automatica:

    Concorrencia_Automatica = gErr

    Select Case gErr

        Case 63743
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147462)

    End Select

    Exit Function

End Function

Function Concorrencia_Grava(ByVal objConcorrencia As ClassConcorrencia) As Long
'Grava concorrência

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Concorrencia_Grava

    'Abre Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 66931

    'Grava concorrência
    lErro = CF("Concorrencia_Grava_EmTrans", objConcorrencia)
    If lErro <> SUCESSO Then gError 66912

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 66932

    Concorrencia_Grava = SUCESSO

    Exit Function

Erro_Concorrencia_Grava:

    Concorrencia_Grava = gErr

    Select Case gErr

        Case 66912

        Case 66931
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 66932
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147463)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Concorrencia_Grava_EmTrans(ByVal objConcorrencia As ClassConcorrencia) As Long
'Grava a Concorrência e seus itens

Dim lErro As Long
Dim alComando(2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim iStatus As Integer

On Error GoTo Erro_Concorrencia_Grava_EmTrans

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66858
    Next

    'Procura no BD a Concorrência passada através do código e FilialEmpresa
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, Status FROM ConcorrenciaN WHERE Codigo = ? AND FilialEmpresa  = ?", 0, lNumIntDoc, iStatus, objConcorrencia.lCodigo, objConcorrencia.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 66861

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66862

    If lErro = AD_SQL_SUCESSO And iStatus = CONCORRENCIA_STATUS_BAIXADO Then gError 89481

    'Se não encontrou a concorrência passada, INSERÇÃO
    If lErro = AD_SQL_SEM_DADOS Then

        'Gera NumIntDoc para nova concorrência
        lErro = CF("Concorrencia_ObterNumIntAutomatico", lNumIntDoc)
        If lErro <> SUCESSO Then gError 66863

        'Guarda novo NumIntDoc
        objConcorrencia.lNumIntDoc = lNumIntDoc

        'Insere nova Concorrência
        lErro = Comando_Executar(alComando(1), "INSERT INTO ConcorrenciaN (NumIntDoc,FilialEmpresa,Codigo,Descricao,Data,Comprador,TaxaFinanceira,TipoDestino,FornCliDestino,FilialDestino, Status) VALUES (?,?,?,?,?,?,?,?,?,?,?)", objConcorrencia.lNumIntDoc, objConcorrencia.iFilialEmpresa, objConcorrencia.lCodigo, objConcorrencia.sDescricao, objConcorrencia.dtData, objConcorrencia.iComprador, objConcorrencia.dTaxaFinanceira, objConcorrencia.iTipoDestino, objConcorrencia.lFornCliDestino, objConcorrencia.iFilialDestino, CONCORRENCIA_STATUS_ATIVO)
        If lErro <> AD_SQL_SUCESSO Then gError 66864

        'Grava os itens da nova concorrência
        lErro = CF("ItensConcorrencia_Inclui", objConcorrencia)
        If lErro <> SUCESSO Then gError 66867

    'Se encontrou, ALTERAÇÃO
    Else

        'Guarda o NumIntDoc
        objConcorrencia.lNumIntDoc = lNumIntDoc

        'Faz "Lock" na concorrencia
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 66865

        'Exclui os itens da concorrência existente
        lErro = CF("ItensConcorrencia_Exclui", objConcorrencia)
        If lErro <> SUCESSO Then gError 66868

        'Grava os itens da concorrência
        lErro = CF("ItensConcorrencia_Inclui", objConcorrencia)
        If lErro <> SUCESSO Then gError 66869

        'Altera concorrência
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE ConcorrenciaN Set Descricao = ?, Data = ?, Comprador = ?, TaxaFinanceira = ?, TipoDestino = ?, FornCliDestino = ?, FilialDestino = ?", alComando(0), objConcorrencia.sDescricao, objConcorrencia.dtData, objConcorrencia.iComprador, objConcorrencia.dTaxaFinanceira, objConcorrencia.iTipoDestino, objConcorrencia.lFornCliDestino, objConcorrencia.iFilialDestino)
        If lErro <> AD_SQL_SUCESSO Then gError 66866

    End If


    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Concorrencia_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Concorrencia_Grava_EmTrans:

    Concorrencia_Grava_EmTrans = gErr

    Select Case gErr

        Case 66858
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66861, 66862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", gErr)

        Case 66863, 66867, 66868, 66869

        Case 66864
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONCORRENCIA", gErr, objConcorrencia.lCodigo)

        Case 66865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONCORRENCIA", gErr, objConcorrencia.lCodigo)

        Case 66866
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONCORRENCIA", gErr, objConcorrencia.lCodigo)

        Case 89481
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONCORRENCIA_BAIXADA", gErr, objConcorrencia.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147464)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensConcorrencia_Exclui(ByVal objConcorrencia As ClassConcorrencia) As Long
'Exclui itens de concorrência e tabelas associadas

Dim lErro As Long
Dim alComando(1) As Long
Dim lNumIntItem As Long
Dim iIndice As Integer

On Error GoTo Erro_ItensConcorrencia_Exclui

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66879
    Next

    'Busca no BD os Itens Relacionados a Concorrência passada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ItensConcorrencia WHERE Concorrencia = ?", 0, lNumIntItem, objConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 66880

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66881

    'Equanto encontrar
    Do While lErro <> AD_SQL_SEM_DADOS

        'Quebra vínculos do Item de Concorrência
        lErro = CF("ItensConcorrencia_Exclui_Relacionamento", lNumIntItem)
        If lErro <> SUCESSO Then gError 66883

        'Faz "Lock" no Item
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 66888

        'Exclui o Item Concorrência
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItensConcorrencia", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 66889

        'Busca próximo item de concorrência
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66882

    Loop

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensConcorrencia_Exclui = SUCESSO

    Exit Function

Erro_ItensConcorrencia_Exclui:

    ItensConcorrencia_Exclui = gErr

    Select Case gErr

        Case 66879
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66880, 66881, 66882
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCONCORRENCIA", gErr, objConcorrencia.lCodigo)

        Case 66888
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSCONCORRENCIA", gErr)

        Case 66889
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSCONCORRENCIA", gErr, lNumIntItem)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147465)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensConcorrencia_Exclui_Relacionamento(ByVal lNumIntConc As Long) As Long
'Exclui vínculos do Item de Concorrência com CotaçãoItemConcorrência, ItemRCItemConcorrência
'e QuantidadesSuplementates

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(5) As Long
Dim lNumIntDoc As Long
Dim lItemReqCompra As Long

On Error GoTo Erro_ItensConcorrencia_Exclui_Relacionamento

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66884
    Next

    'Busca Relacionamento entre o ItemConrrência e CotaçãoItemConcorrência
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM CotacaoItemConcorrencia WHERE ItemConcorrencia = ?", 0, lNumIntDoc, lNumIntConc)
    If lErro <> AD_SQL_SUCESSO Then gError 66885

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66886

    'Equanto encontrar
    Do While lErro <> AD_SQL_SEM_DADOS

        'Faz "LockExclusive" em CotaçãoItemConcorrência
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 66891

        'Quebra o vínculo Excluindo a Cotação Item Concorrência
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM CotacaoItemConcorrencia", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 66890

        'Busca Próximo Relacionamento
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66887

    Loop

    'Procura relacionamento entre o ItemConcorrência e ItemRCItemConcorrência
    lErro = Comando_ExecutarPos(alComando(2), "SELECT ItemReqCompra FROM ItemRCItemConcorrencia WHERE ItemConcorrencia = ?", 0, lItemReqCompra, lNumIntConc)
    If lErro <> AD_SQL_SUCESSO Then gError 66892

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66893

    'Enquanto encontrar
    Do While lErro <> AD_SQL_SEM_DADOS

        'Faz "lockExclusive" em ItemRCItemConcorrencia
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 66895

        'Exclui o vínculo com ItemRCItemConcorrência
        lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM ItemRCItemConcorrencia", alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 66896

        'Busca próximo relacionamento com ItemRCItemConcorrência
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66894

    Loop

    'Procura relacionamento de ItemConcorrência com QuantidadesSuplementares
    lErro = Comando_ExecutarPos(alComando(4), "SELECT ItemConcorrencia FROM QuantidadesSuplementares WHERE ItemConcorrencia = ?", 0, lNumIntDoc, lNumIntConc)
    If lErro <> AD_SQL_SUCESSO Then gError 66897

    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66898

    'Equanto existir relacionamento
    Do While lErro <> AD_SQL_SEM_DADOS

        'Exclui o vínculo com QuantidadesSuplementares
        lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM QuantidadesSuplementares", alComando(4))
        If lErro <> AD_SQL_SUCESSO Then gError 66900

        'Busca Próximo vínculo com QuantidadesSuplementares
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66899

    Loop

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensConcorrencia_Exclui_Relacionamento = SUCESSO

    Exit Function

Erro_ItensConcorrencia_Exclui_Relacionamento:

    ItensConcorrencia_Exclui_Relacionamento = gErr

    Select Case gErr

        Case 66884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66885, 66886, 66887
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOITEMCONCORRENCIA", gErr)

        Case 66890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAOITEMCONCORRENCIA", gErr, lNumIntConc)

        Case 66891
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_COTACAOITEMCONCORRENCIA", gErr)

        Case 66892, 66893, 66894
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMRCITEMCONCORRENCIA1", gErr)

        Case 66895
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMRCITEMCONCORRENCIA", gErr)

        Case 66896
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMRCITEMCONCORRENCIA", gErr, lNumIntConc)

        Case 66897, 66898, 66899
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_QUANTIDADESSUPLEMENTARES", gErr)

        Case 66900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_QUANTIDADESSUPLEMENTARES", gErr, lNumIntConc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147466)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensConcorrencia_Inclui(ByVal objConcorrencia As ClassConcorrencia) As Long
'Grava novos Itens de Concorrência

Dim lErro As Long
Dim objItemConcorrencia As ClassItemConcorrencia
Dim alComando(4) As Long
Dim lNumIntItem As Long
Dim lNumIntDoc As Long
Dim lNumIntConc As Long
Dim objCotacaoItemConc As ClassCotacaoItemConc
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim objQuantSupl As ClassQuantSuplementar
Dim iIndice As Integer
Dim dFator As Double, sUM As String
Dim objProduto As New ClassProduto
Dim dQuantidade As Double
Dim dQuantPedida As Double, dQuantReceb As Double
Dim dQuantItemRC As Double, dQuantCancel As Double
Dim bExclui As Boolean
Dim dDif As Double

On Error GoTo Erro_ItensConcorrencia_Inclui
    
    'Abertura dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66875
    Next

    bExclui = False

    'Para cada de Item de concorrência
    For Each objItemConcorrencia In objConcorrencia.colItens

        objProduto.sCodigo = objItemConcorrencia.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 23080 Then gError 76187
        If lErro <> SUCESSO Then gError 76191
                        
        'Gera número automático para item da concorrência
        lErro = CF("ItemConcorrencia_ObterNumIntAutomatico", lNumIntItem)
        If lErro <> SUCESSO Then gError 66876

        'Insere novo item de concorrência
        lErro = Comando_Executar(alComando(0), "INSERT INTO ItensConcorrenciaN (NumIntDoc, Concorrencia, Produto, Fornecedor, Filial, Quantidade, UM, DataNecessidade, Status) VALUES (?,?,?,?,?,?,?,?,?)", lNumIntItem, objConcorrencia.lNumIntDoc, objItemConcorrencia.sProduto, objItemConcorrencia.lFornecedor, objItemConcorrencia.iFilial, objItemConcorrencia.dQuantidade, objItemConcorrencia.sUM, objItemConcorrencia.dtDataNecessidade, ITENSCONCORRENCIA_STATUS_ATIVO)
        If lErro <> SUCESSO Then gError 66877

        objItemConcorrencia.lNumIntDoc = lNumIntItem

        'Para cada Item Concorrência vinculado a Cotação
        For Each objCotacaoItemConc In objItemConcorrencia.colCotacaoItemConc
            
            If objCotacaoItemConc.iSelecionada = MARCADO Then

                'Gera número automático para CotaçãoItemConc
                lErro = CF("CotacaoItemConc_ObterNumIntAutomatico", lNumIntConc)
                If lErro <> SUCESSO Then gError 66902
    
                'Insere nova CotaçãoItemConc
                lErro = Comando_Executar(alComando(1), "INSERT INTO CotacaoItemConcorrenciaN (NumIntDoc,ItemConcorrencia,ItemCotacao,ValorPresente,Escolhido,MotivoEscolha,QuantidadeComprar,DataEntrega,PrecoAjustado, Status) VALUES (?,?,?,?,?,?,?,?,?,?)", lNumIntConc, lNumIntItem, objCotacaoItemConc.lItemCotacao, objCotacaoItemConc.dValorPresente, objCotacaoItemConc.iEscolhido, objCotacaoItemConc.sMotivoEscolha, objCotacaoItemConc.dQuantidadeComprar, objCotacaoItemConc.dtDataEntrega, objCotacaoItemConc.dPrecoAjustado, COTACAOITEMCONCORRENCIA_STATUS_ATIVO)
                If lErro <> AD_SQL_SUCESSO Then gError 66903
    
                'Guarda o número interno do item de cotação
                objCotacaoItemConc.lNumIntDoc = lNumIntConc
            End If
        Next
        
        'Para cada Item de Requisição vinculado a Concorrência
        For Each objItemRCItemConc In objItemConcorrencia.colItemRCItemConcorrencia
            
            sUM = String(STRING_UM_SIGLA, 0)
            
            lErro = Comando_ExecutarPos(alComando(4), "SELECT UM,NumIntDoc,QuantPedida, Quantidade,QuantCancelada,QuantRecebida FROM ItensReqCompra WHERE NumIntDoc=?", 0, sUM, lNumIntDoc, dQuantPedida, dQuantItemRC, dQuantCancel, dQuantReceb, objItemRCItemConc.lItemReqCompra)
            If lErro <> AD_SQL_SUCESSO Then gError 62811
            
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62812

            If lErro = AD_SQL_SEM_DADOS Then gError 62813
            
            lErro = Comando_LockExclusive(alComando(4))
            If lErro <> AD_SQL_SUCESSO Then gError 62814
            
            lErro = ItemReqCompras_Verifica_Vinculo(objItemRCItemConc.lItemReqCompra, bExclui)
            If lErro <> SUCESSO Then gError 62816
                    
            lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemConcorrencia.sUM, sUM, dFator)
            If lErro <> SUCESSO Then gError 62815
            
            dQuantidade = objItemRCItemConc.dQuantidade * dFator
            
            dDif = (dQuantItemRC - dQuantidade - dQuantCancel - dQuantReceb - dQuantPedida)
            
            If dDif < 0 Then
                'Verifica se a diferença é significativa
                If Abs(dDif) > QTDE_ESTOQUE_DELTA Then gError 62819
            End If

            
            'Insere novo ItemRCItemConcorrência
            lErro = Comando_Executar(alComando(2), "INSERT INTO ItemRCItemConcorrencia (ItemReqCompra, ItemConcorrencia, Quantidade) VALUES (?,?,?)", objItemRCItemConc.lItemReqCompra, lNumIntItem, objItemRCItemConc.dQuantidade)
            If lErro <> AD_SQL_SUCESSO Then gError 66904

        Next
        
        If objItemConcorrencia.colItemRCItemConcorrencia.Count > 0 Then
            'Para cada Quantidade Suplementar vinculada ao Item Concorrência
            For Each objQuantSupl In objItemConcorrencia.colQuantSuplementar
    
                If (objQuantSupl.dQuantidade - objQuantSupl.dQuantRequisitada) > 0 Then
                    'Insere uma nova Quantidade Suplementar
                    lErro = Comando_Executar(alComando(3), "INSERT INTO QuantidadesSuplementares (ItemConcorrencia,TipoDestino,FornCliDestino,FilialDestino,Quantidade) VALUES (?,?,?,?,?)", lNumIntItem, objQuantSupl.iTipoDestino, objQuantSupl.lFornCliDestino, objQuantSupl.iFilialDestino, objQuantSupl.dQuantidade - objQuantSupl.dQuantRequisitada)
                    If lErro <> AD_SQL_SUCESSO Then gError 66905
                End If
            Next
        End If
    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensConcorrencia_Inclui = SUCESSO

    Exit Function

Erro_ItensConcorrencia_Inclui:

    ItensConcorrencia_Inclui = gErr

    Select Case gErr

        Case 62811, 62812
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)
        
        Case 62813
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENSREQCOMPRA_NAO_CADASTRADO2", gErr, objItemRCItemConc.lItemReqCompra)
        
        Case 62814
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSREQCOMPRA1", gErr, objItemRCItemConc.lItemReqCompra)
        
        Case 62815, 66876, 66902, 76187, 62816
        
        Case 62819
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_MAIOR_QUANTCOMPRARMAX", gErr, dQuantidade, dQuantItemRC - dQuantCancel - dQuantReceb - dQuantPedida)
        
        Case 66875
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMCONCONCORRENICA", gErr, lNumIntItem)

        Case 66903
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAOITEMCONCORRENCIA", gErr, lNumIntConc)

        Case 66904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMRCITEMCONCORRENCIA", gErr, lNumIntItem)

        Case 66905
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_QUANTIDADESSUPLEMENTARES", gErr, lNumIntItem)

        Case 76191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147467)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PedCompra_Concorrencia_Grava(ByVal objConcorrencia As ClassConcorrencia, ByVal colPedidoCompras As Collection) As Long
'Atualiza Concorrência e grava Pedidos de Compras

Dim lErro As Long
Dim lTransacao As Long
Dim objPedCompra As ClassPedidoCompras

On Error GoTo Erro_PedCompra_Concorrencia_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 66859

    'Grava Pedidos de Compra
    lErro = CF("PedidosCompra_Grava_Concorrencia_EmTrans", colPedidoCompras, objConcorrencia)
    If lErro <> SUCESSO Then gError 66906
    
    'Baixa a Concorrência
    lErro = CF("Concorrencia_Baixa", objConcorrencia)
    If lErro <> SUCESSO Then gError 66950
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 66860
    
    'Para cada Pedido de compras da coleção
    For Each objPedCompra In colPedidoCompras
        
        'Se o Pedido de Compras possui Itens
        If objPedCompra.colItens.Count > 0 Then

            lErro = CF("WorkFlow_Trata_Transacao", "COM", objPedCompra, gsUsuario, objPedCompra, 1002)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
    Next
            
    PedCompra_Concorrencia_Grava = SUCESSO

    Exit Function

Erro_PedCompra_Concorrencia_Grava:

    PedCompra_Concorrencia_Grava = gErr

    Select Case gErr

        Case 66906, 66950

        Case 66859
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 66860
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147468)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function PedidosCompra_Grava_Concorrencia_EmTrans(ByVal colPedidoCompras As Collection, ByVal objConcorrencia As ClassConcorrencia) As Long
'Grava Pedidos de Compras

Dim lErro As Long
Dim alComando(8) As Long
Dim objItemPC As ClassItemPedCompra
Dim objItemConcorrencia As ClassItemConcorrencia
Dim objPedidoCompra As ClassPedidoCompras
Dim lNumIntOrigin As Long
Dim iIndice As Integer
Dim lItemConc As Long
Dim lCodigo As Long
Dim objItemRCItemConc As ClassItemRCItemConcorrencia
Dim dQuantidade As Double
Dim dQuantPedida As Double
Dim dQuantRecebida As Double
Dim dQuantCancelada As Double
Dim lNumIntDoc As Long
Dim iItemConcEncontrado As Integer
Dim objCotacaoItemConc As ClassCotacaoItemConc

On Error GoTo Erro_PedidosCompra_Grava_Concorrencia_EmTrans

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66907
    Next

    'Para cada Pedido de compras da coleção
    For Each objPedidoCompra In colPedidoCompras
        
        'Se o Pedido de Compras possui Itens
        If objPedidoCompra.colItens.Count > 0 Then
        
            'Gera código automatico para o Pedido de Compras
            lErro = CF("PedidoCompra_Codigo_Automatico_EmTrans", lCodigo)
            If lErro <> SUCESSO Then gError 66854
            
            objPedidoCompra.lCodigo = lCodigo
    
            'Gera NumIntDoc para o Pedido de Compra
            lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_PC", lNumIntDoc)
            If lErro <> SUCESSO Then gError 67013
    
            objPedidoCompra.lNumIntDoc = lNumIntDoc
    
            If Len(Trim(objPedidoCompra.sTipoFrete)) = 0 Then objPedidoCompra.sTipoFrete = 0
            
            lErro = CF("PedidoCompra_Valida_Campos_Cust", objPedidoCompra)
            If lErro <> SUCESSO Then gError 66854
    
            objPedidoCompra.sUsuReg = gsUsuario
            objPedidoCompra.bInclusao = True
    
            'Insere o novo Pedido de Compras
            lErro = Comando_Executar(alComando(0), "INSERT INTO PedidoCompraN (NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,ObsEmbalagem,Observacao,DataBaixa,ValorProdutos, Status, Moeda, Taxa, DataReg, HoraReg, UsuReg) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objPedidoCompra.lNumIntDoc, objPedidoCompra.iFilialEmpresa, objPedidoCompra.lCodigo, objPedidoCompra.lFornecedor, objPedidoCompra.iFilial, objPedidoCompra.iComprador, objPedidoCompra.sContato, objPedidoCompra.iTipoDestino, objPedidoCompra.lFornCliDestino, objPedidoCompra.iFilialDestino, objPedidoCompra.dtData, objPedidoCompra.dtDataEmissao, objPedidoCompra.dtDataEnvio, objPedidoCompra.dtDataAlteracao, objPedidoCompra.iCondicaoPagto, objPedidoCompra.dOutrasDespesas, objPedidoCompra.dValorFrete, _
                objPedidoCompra.dValorSeguro, objPedidoCompra.dValorDesconto, objPedidoCompra.dValorTotal, objPedidoCompra.dValorIPI, objPedidoCompra.sTipoFrete, objPedidoCompra.iTransportadora, objPedidoCompra.iProxSeqBloqueio, objPedidoCompra.iTipoBaixa, objPedidoCompra.sMotivoBaixa, objPedidoCompra.sAlcada, objPedidoCompra.sObsEmbalagem, objPedidoCompra.lObservacao, objPedidoCompra.dtDataBaixa, objPedidoCompra.dValorProdutos, PEDIDOCOMPRA_STATUS_ATIVO, objPedidoCompra.iMoeda, objPedidoCompra.dTaxa, Date, CDbl(Time), gsUsuario)
            If lErro <> SUCESSO Then gError 67011
    
            'Para cada Item do Pedido de Compras
            For Each objItemPC In objPedidoCompra.colItens
    
                'Gera numero interno automatico
                lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROX_ITEM_PEDCOMPRA", lNumIntDoc)
                If lErro <> SUCESSO Then gError 67014
    
                objItemPC.lNumIntDoc = lNumIntDoc
    
                'Insere novo ItemPC
                lErro = Comando_Executar(alComando(1), "INSERT INTO ItensPedCompraN (NumIntDoc,PedCompra,DataLimite,Produto,DescProduto,Quantidade,QuantRecebida,QuantRecebimento,UM,PrecoUnitario,ValorDesconto,TipoOrigem,NumIntOrigem,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa,Status,Observacao,ValorIPI,AliquotaIPI,AliquotaICMS, StatusBaixa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDoc, objPedidoCompra.lNumIntDoc, objItemPC.dtDataLimite, objItemPC.sProduto, objItemPC.sDescProduto, objItemPC.dQuantidade, objItemPC.dQuantRecebida, objItemPC.dQuantRecebimento, objItemPC.sUM, objItemPC.dPrecoUnitario, objItemPC.dValorDesconto, objItemPC.iTipoOrigem, objItemPC.lNumIntOrigem, objItemPC.dPercentMaisReceb, objItemPC.dPercentMenosReceb, objItemPC.iRebebForaFaixa, objItemPC.iStatus, objItemPC.lObservacao, objItemPC.dValorIPI, objItemPC.dAliquotaIPI, objItemPC.dAliquotaICMS, ITENSPEDCOMPRA_STATUS_ATIVO)
                If lErro <> AD_SQL_SUCESSO Then gError 67012
    
    
                'Faz ligação ItemRC e itemPC e atualiza a quantidade pedida do ItemRC
                lErro = CF("ItemPC_Associa_ItemRC", objItemPC)
                If lErro <> SUCESSO Then gError 62765
            Next
    
            lErro = CF("LocalizacaoItemPC_Inclui", objPedidoCompra)
            If lErro <> SUCESSO Then gError 86128
        
        End If
        
    Next
                
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidosCompra_Grava_Concorrencia_EmTrans = SUCESSO

    Exit Function

Erro_PedidosCompra_Grava_Concorrencia_EmTrans:

    PedidosCompra_Grava_Concorrencia_EmTrans = gErr

    Select Case gErr

        Case 66907
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PEDIDOCOMPRA", gErr, objPedidoCompra.lNumIntDoc)

        Case 67012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPEDCOMPRA", gErr)

        Case 62765, 67013, 67014, 66854, 86128

        Case 67018, 67019
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)

        Case 67020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSREQCOMPRA1", gErr, objItemRCItemConc.lItemReqCompra)

        Case 67021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_MAIOR_RC", gErr, objItemRCItemConc.dQuantidade, dQuantidade - dQuantPedida - dQuantRecebida - dQuantCancelada)

        Case 67028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMRCITEMPC", gErr, objItemPC.lNumIntDoc, objItemRCItemConc.lItemReqCompra)

        Case 67029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSREQCOMPRA", gErr, objItemRCItemConc.lItemReqCompra)

        Case 67030, 67031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMRCITEMCONCORRENCIA", gErr, objItemRCItemConc.lItemReqCompra)

        Case 67032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMRCITEMCONCORRENCIA_NAO_CADASTRADO", gErr, objItemRCItemConc.lItemReqCompra, objItemConcorrencia.lNumIntDoc)

        Case 70509, 70510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTOITEMRC", gErr)
        
        Case 70511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COTACAOPRODUTOITEMRC_NAO_CADASTRADA", gErr, objItemRCItemConc.lItemReqCompra)
        
        Case 70512, 70521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAOPRODUTOITEMRC", gErr, objItemRCItemConc.lItemReqCompra)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147469)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Concorrencia_ObterNumIntAutomatico(lNumInt As Long) As Long
'gera o próximo número interno disponível para Concorrencia

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_Concorrencia_ObterNumIntAutomatico

    'Gera o proximo numero interno para ItemConcorrencia
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_CONCORRENCIA", lNumInt)
    If lErro <> SUCESSO Then gError 66878

    Concorrencia_ObterNumIntAutomatico = SUCESSO

    Exit Function

Erro_Concorrencia_ObterNumIntAutomatico:

    Concorrencia_ObterNumIntAutomatico = gErr

    Select Case gErr

        Case 66878
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147470)

    End Select

    Exit Function

End Function

Function ItemConcorrencia_ObterNumIntAutomatico(lNumInt As Long) As Long
'gera o próximo número interno disponível para ItemConcorrencia

Dim lErro As Long
Dim objItemConcorrencia As New ClassItemConcorrencia

On Error GoTo Erro_ItemConcorrencia_ObterNumIntAutomatico

    'Gera o proximo numero interno para ItemConcorrencia
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_ITEMCONCORRENCIA", lNumInt)
    If lErro <> SUCESSO Then gError 63810

    ItemConcorrencia_ObterNumIntAutomatico = SUCESSO

    Exit Function

Erro_ItemConcorrencia_ObterNumIntAutomatico:

    ItemConcorrencia_ObterNumIntAutomatico = gErr

    Select Case gErr

        Case 63810
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147471)

    End Select

    Exit Function

End Function

Function CotacaoItemConc_ObterNumIntAutomatico(lNumInt As Long) As Long
'gera o próximo número interno disponível para CotaçãoItenConc

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_CotacaoItemConc_ObterNumIntAutomatico

    'Gera o proximo numero interno para CotaçãoItenConc
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_COTACAOITEMCONC", lNumInt)
    If lErro <> SUCESSO Then gError 66901

    CotacaoItemConc_ObterNumIntAutomatico = SUCESSO

    Exit Function

Erro_CotacaoItemConc_ObterNumIntAutomatico:

    CotacaoItemConc_ObterNumIntAutomatico = gErr

    Select Case gErr

        Case 66901
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147472)

    End Select

    Exit Function

End Function

Function PedidoCompra_Codigo_Automatico(lCodigo As Long) As Long
'Gera o proximo codigo da Tabela de Pedido Compra

Dim lErro As Long

On Error GoTo Erro_PedidoCompra_Codigo_Automatico

    'Chama a rotina que gera o sequencial
    lErro = CF("Config_ObterAutomatico", "ComprasConfig", "NUM_PROXIMO_CODIGO_PC", "PedidoCompraN", "Codigo", lCodigo)
    If lErro <> SUCESSO Then Error 56114

    PedidoCompra_Codigo_Automatico = SUCESSO

    Exit Function

Erro_PedidoCompra_Codigo_Automatico:

    PedidoCompra_Codigo_Automatico = gErr

    Select Case gErr

        Case 56114
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147473)

    End Select

    Exit Function

End Function

Function Cotacao_Inclui(ByVal objCotacao As ClassCotacao, ByVal colPedidoCotacao As Collection) As Long
'Atualiza as tabelas Cotacao,CotacaoCondPagto e CotacaoProduto a partir
'da Cotacao passada como parametro

Dim lErro As Long
Dim lNumIntCotacao As Long
Dim lComando As Long

On Error GoTo Erro_Cotacao_Inclui

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 63187

    'Gera um novo numero interno para cotacao
    lErro = CF("Cotacao_Automatico", lNumIntCotacao)
    If lErro <> SUCESSO Then Error 63183

    objCotacao.lNumIntDoc = lNumIntCotacao
    
    'Inclui a nova cotacao na tabela Cotacao
    lErro = Comando_Executar(lComando, "INSERT INTO CotacaoN (NumIntDoc, FilialEmpresa,Codigo,Descricao,Data,TipoDestino,FornCliDestino,FilialDestino,Comprador, Status) VALUES (?,?,?,?,?,?,?,?,?,?)", lNumIntCotacao, objCotacao.iFilialEmpresa, objCotacao.lCodigo, objCotacao.sDescricao, objCotacao.dtData, objCotacao.iTipoDestino, objCotacao.lFornCliDestino, objCotacao.iFilialDestino, objCotacao.iComprador, COTACAO_STATUS_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then Error 63184

    'Inclui as condicoes de pagamento na tabela CotacaoCondPagto
    lErro = CF("CondicoesPagto_Inclui", objCotacao)
    If lErro <> SUCESSO Then Error 63185

    'Inclui CotacoesProduto de Cotacao na tabela de CotacaoProduto
    lErro = CF("CotacoesProduto_Inclui", objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then Error 63186

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Cotacao_Inclui = SUCESSO

    Exit Function

Erro_Cotacao_Inclui:

    Cotacao_Inclui = Err

        Select Case Err

            Case 63183, 63185, 63186

            Case 63184
                lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAO", Err)
                
            Case 63187
                lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

            Case Else
                lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147474)

        End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'''Function CotacoesProduto_Inclui(objCotacao As ClassCotacao, colItensPedCotacao As Collection) As Long
'''
'''Dim lErro As Long
'''Dim objCotacaoProduto As New ClassCotacaoProduto
'''Dim lComando As Long
'''Dim objItemPedCotacao As New ClassItemPedCotacao
'''Dim lNumCotacaoProduto As Long
'''Dim lCotacaoProduto As Long
'''
'''On Error GoTo Erro_CotacoesProduto_Inclui
'''
'''    'Abre o comando
'''    lComando = Comando_Abrir()
'''    If lComando = 0 Then Error 63213
'''
'''
'''    Set objCotacaoProduto = New ClassCotacaoProduto
'''
'''    For Each objCotacaoProduto In objCotacao.colCotacaoProduto
'''
'''        Set objItemPedCotacao = New ClassItemPedCotacao
'''
'''        lErro = CF("CotacaoProduto_Automatico",lNumCotacaoProduto)
'''        If lErro <> SUCESSO Then Error 63318
'''
'''        objCotacaoProduto.lNumIntDoc = lNumCotacaoProduto
'''        objCotacaoProduto.lCotacao = objCotacao.lNumIntDoc
'''        objItemPedCotacao.lCotacaoProduto = objCotacaoProduto.lNumIntDoc
'''
'''        'Insere as CotacoesProduto de Cotacao na tabela CotacaoProduto
'''        lErro = Comando_Executar(lComando, "INSERT INTO CotacaoProduto (NumIntDoc,Cotacao,Produto,Quantidade,UM, Fornecedor,Filial) VALUES (?,?,?,?,?,?,?)", objCotacaoProduto.lNumIntDoc, objCotacaoProduto.lCotacao, objCotacaoProduto.sProduto, objCotacaoProduto.dQuantidade, objCotacaoProduto.sUM, objCotacaoProduto.lFornecedor, objCotacaoProduto.iFilial)
'''        If lErro <> AD_SQL_SUCESSO Then Error 63214
'''
'''        colItensPedCotacao.Add objItemPedCotacao
'''
'''    Next
'''
'''    'Fecha o comando
'''    lErro = Comando_Fechar(lComando)
'''
'''    CotacoesProduto_Inclui = SUCESSO
'''
'''    Exit Function
'''
'''Erro_CotacoesProduto_Inclui:
'''
'''    CotacoesProduto_Inclui = Err
'''
'''    Select Case Err
'''
'''        Case 63213
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'''
'''        Case 63214
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAOPRODUTO", Err)
'''
'''        Case 63318
'''
'''        Case Else
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147475)
'''
'''    End Select
'''
'''    Call Comando_Fechar(lComando)
'''
'''    Exit Function
'''
'''End Function

Function CotacoesProduto_Inclui(ByVal objCotacao As ClassCotacao, ByVal colPedidoCotacao As Collection) As Long
'Grava Cotações de Produto

Dim lErro As Long
Dim objCotacaoProduto As ClassCotacaoProduto
Dim objItemPedCotacao As ClassItemPedCotacao
Dim objPedidoCotacao As ClassPedidoCotacao
Dim lComando As Long
Dim lNumCotacaoProduto As Long
Dim lCotacaoProduto As Long, objFornecedorProdutoFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_CotacoesProduto_Inclui

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 63213

    'Incluido por Wagner 08/04/2008
    For Each objCotacaoProduto In objCotacao.colCotacaoProduto
        objCotacaoProduto.lLinkProv = objCotacaoProduto.lNumIntDoc
    Next
    For Each objPedidoCotacao In colPedidoCotacao
        For Each objItemPedCotacao In objPedidoCotacao.colItens
            objItemPedCotacao.lLinkProv = objItemPedCotacao.lCotacaoProduto
        Next
    Next
    'Wagner
    
    'Para cada Cotação Produto da Cotação passada
    For Each objCotacaoProduto In objCotacao.colCotacaoProduto

        'Gera número automático para Cotação Produto
        lErro = CF("CotacaoProduto_Automatico", lNumCotacaoProduto)
        If lErro <> SUCESSO Then Error 63318
        
        'Procura por ItemPedCotacao vinculada a Cotacao Produto através do link provisório que foi criado(por numIntDoc)
        For Each objPedidoCotacao In colPedidoCotacao
            For Each objItemPedCotacao In objPedidoCotacao.colItens
                
                'Se encontrar, guarda o NumIntDoc da Cotação Produto
                If objItemPedCotacao.lLinkProv = objCotacaoProduto.lLinkProv Then
                    objItemPedCotacao.lCotacaoProduto = lNumCotacaoProduto
                    
                    'lockar fornecedorprodutoff
                    With objFornecedorProdutoFF
                        .iFilialEmpresa = objPedidoCotacao.iFilialEmpresa
                        .lFornecedor = objPedidoCotacao.lFornecedor
                        .iFilialForn = objPedidoCotacao.iFilial
                        .sProduto = objCotacaoProduto.sProduto
                    End With
                    
                    lErro = CF("FornecedorProdutoFF_Lock", objFornecedorProdutoFF)
                    If lErro <> SUCESSO And lErro <> 67872 Then Error 33332
                    If lErro <> SUCESSO Then Error 33333
                    
                    Exit For
                End If
            Next
        Next
        
        'Guarda os dados da Cotação Produto
        objCotacaoProduto.lNumIntDoc = lNumCotacaoProduto
        objCotacaoProduto.lCotacao = objCotacao.lNumIntDoc
                
        'Insere as CotacoesProduto de Cotacao na tabela CotacaoProduto
        lErro = Comando_Executar(lComando, "INSERT INTO CotacaoProdutoN (NumIntDoc,Cotacao,Produto,Quantidade,UM, Fornecedor,Filial, Status) VALUES (?,?,?,?,?,?,?,?)", objCotacaoProduto.lNumIntDoc, objCotacaoProduto.lCotacao, objCotacaoProduto.sProduto, objCotacaoProduto.dQuantidade, objCotacaoProduto.sUM, objCotacaoProduto.lFornecedor, objCotacaoProduto.iFilial, COTACAOPRODUTO_STATUS_ATIVO)
        If lErro <> AD_SQL_SUCESSO Then Error 63214
                
    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    CotacoesProduto_Inclui = SUCESSO

    Exit Function

Erro_CotacoesProduto_Inclui:

    CotacoesProduto_Inclui = Err

    Select Case Err

        Case 63213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 63214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAOPRODUTO", Err)

        Case 63318, 33332

        Case 33333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_ENCONTRADO", Err, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147476)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CondicoesPagto_Inclui(ByVal objCotacao As ClassCotacao) As Long
'Inclui na tabela CotacaoCondPagto as condicoes de pagto de Cotacao

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim lComando As Long
Dim iIndice As Integer
On Error GoTo Erro_CondicoesPagto_Inclui

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 63211

    Set objCondicaoPagto = New ClassCondicaoPagto

    For iIndice = 1 To objCotacao.colCondPagtos.Count

        objCondicaoPagto.iCodigo = objCotacao.colCondPagtos.Item(iIndice)

        'Insere as Condicoes de Pagto na tabela CotacaoCondPagto
        lErro = Comando_Executar(lComando, "INSERT INTO CotacaoCondPagto (NumIntCotacao,CondPagto) VALUES (?,?)", objCotacao.lNumIntDoc, objCondicaoPagto.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 63212

    Next

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    CondicoesPagto_Inclui = SUCESSO

    Exit Function

Erro_CondicoesPagto_Inclui:

    CondicoesPagto_Inclui = Err

    Select Case Err

        Case 63211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 63212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAOCONDPAGTO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147477)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensPedCotacao_Inclui(ByVal objCotacao As ClassCotacao, ByVal objPedidoCotacao As ClassPedidoCotacao) As Long
'Inclui os itens de Pedido de Cotacao na tabela ItemPedCotacao do BD

Dim lErro As Long
Dim lComando As Long
Dim lNumIntItemPC As Long, sObservacao As String
Dim objItemPedidoCotacao As ClassItemPedCotacao
Dim objCotacaoProduto As ClassCotacaoProduto
Dim objItemReqCompras As ClassItemReqCompras

On Error GoTo Erro_ItensPedCotacao_Inclui

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 63196
    
    'Para cada ItemPedCotacao do Pedido de Cotação passado
    For Each objItemPedidoCotacao In objPedidoCotacao.colItens
                
        'Gera o próximo número interno de ItemPedCotacao
        lErro = CF("ItensPedCotacao_Automatico", lNumIntItemPC)
        If lErro <> SUCESSO Then Error 63197
    
        sObservacao = ""
    
        For Each objCotacaoProduto In objCotacao.colCotacaoProduto
        
            If objItemPedidoCotacao.lCotacaoProduto = objCotacaoProduto.lNumIntDoc Then
        
                For Each objItemReqCompras In objCotacaoProduto.colItemReqCompras
                    
                    If Len(Trim(objItemReqCompras.sObservacao)) <> 0 Then
                    
                        If sObservacao <> "" Then sObservacao = sObservacao & " "
                        sObservacao = sObservacao & Trim(objItemReqCompras.sObservacao)
                    
                    End If
                    
                Next
    
            End If
    
        Next
        
        sObservacao = left(sObservacao, STRING_ITEMPEDCOTACAO_OBSERVACAO)
    
        'Insere os ItensPedCotacao na tabela ItemPedCotacao
        lErro = Comando_Executar(lComando, "INSERT INTO ItemPedCotacao (NumIntDoc,CotacaoProduto,PedidoCotacao, Status, Observacao) VALUES (?,?,?,?,?)", lNumIntItemPC, objItemPedidoCotacao.lCotacaoProduto, objPedidoCotacao.lNumIntDoc, ITEMPEDCOTACAO_STATUS_ATIVO, sObservacao)
        If lErro <> AD_SQL_SUCESSO Then Error 63198

    Next
            
    'Fecha comandos
    Call Comando_Fechar(lComando)

    ItensPedCotacao_Inclui = SUCESSO

    Exit Function

Erro_ItensPedCotacao_Inclui:

    ItensPedCotacao_Inclui = Err

    Select Case Err

        Case 63196
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 63197
            'Erro tratado na rotina chamada

        Case 63198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMPEDCOTACAO", Err, objItemPedidoCotacao.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147478)

    End Select
    
    'Fecha comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidosCotacao_Inclui(ByVal objCotacao As ClassCotacao, ByVal colPedidoCotacao As Collection) As Long
'Inclui os Pedidos de Cotacao na tabela PedidoCotacao

Dim lErro As Long
Dim lNumIntPedCotacao As Long
Dim lComando As Long
Dim lPedCotacao As Long
Dim objPedidoCotacao As ClassPedidoCotacao

On Error GoTo Erro_PedidosCotacao_Inclui

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 63193
        
    'Para cada Pedido de Cotação da coleção de pedidos
    For Each objPedidoCotacao In colPedidoCotacao

        'Gera novo número interno para PedidoCotacao
        lErro = CF("PedidoCotacaoNumInt_Automatico", lNumIntPedCotacao)
        If lErro <> SUCESSO Then Error 63189

        'Gera código para PedidoCotacao
        lErro = CF("PedidoCotacao_Automatico", lPedCotacao)
        If lErro <> SUCESSO Then Error 63190

        objPedidoCotacao.lCodigo = lPedCotacao
        objPedidoCotacao.lNumIntDoc = lNumIntPedCotacao
        
        'Insere um Novo Pedido de Cotação
        lErro = Comando_Executar(lComando, "INSERT INTO PedidoCotacaoN(NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Contato,DataEmissao,Data,DataValidade,TipoFrete,Status,CondPagtoPrazo, StatusBaixa)VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntPedCotacao, objPedidoCotacao.iFilialEmpresa, lPedCotacao, objPedidoCotacao.lFornecedor, objPedidoCotacao.iFilial, objPedidoCotacao.sContato, objPedidoCotacao.dtDataEmissao, objPedidoCotacao.dtData, objPedidoCotacao.dtDataValidade, objPedidoCotacao.iTipoFrete, objPedidoCotacao.iStatus, objPedidoCotacao.iCondPagtoPrazo, PEDIDOCOTACAO_STATUS_ATIVO)
        If lErro <> AD_SQL_SUCESSO Then Error 63191
         
        'Inclui os itens de pedido de cotacao na tabela ItemPedCotacao
        lErro = CF("ItensPedCotacao_Inclui", objCotacao, objPedidoCotacao)
        If lErro <> SUCESSO Then Error 63192
 
    Next
   
    'Fecha comando
    Call Comando_Fechar(lComando)

    PedidosCotacao_Inclui = SUCESSO

    Exit Function

Erro_PedidosCotacao_Inclui:

    PedidosCotacao_Inclui = Err

    Select Case Err

        Case 63189, 63190, 63192

        Case 63191
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PEDIDOCOTACAO", Err, lNumIntPedCotacao)

        Case 63193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147479)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cotacao_Automatico(lNumIntCotacao As Long) As Long
'Devolve o proximo número interno de cotacao disponivel.

Dim lErro As Long

On Error GoTo Erro_Cotacao_Automatico

    'Obtem o próximo número interno de cotacao disponivel
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMA_COTACAO", lNumIntCotacao)
    If lErro <> SUCESSO Then Error 63188

    Cotacao_Automatico = SUCESSO

    Exit Function

Erro_Cotacao_Automatico:

    Cotacao_Automatico = Err

    Select Case Err

        Case 63188
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147480)

    End Select

    Exit Function

End Function

Function CotacaoProduto_Automatico(lNumCotacaoProduto As Long) As Long
'Devolve o proximo número interno de cotacao produto disponivel.

Dim lErro As Long

On Error GoTo Erro_CotacaoProduto_Automatico

    'Obtem o próximo número interno de cotacao disponivel
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMA_COTACAOPRODUTO", lNumCotacaoProduto)
    If lErro <> SUCESSO Then Error 63318

    CotacaoProduto_Automatico = SUCESSO

    Exit Function

Erro_CotacaoProduto_Automatico:

    CotacaoProduto_Automatico = Err

    Select Case Err

        Case 63318
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147481)

    End Select

    Exit Function

End Function

Function Cotacao_Grava(ByVal objCotacao As ClassCotacao, ByVal colPedidoCotacao As Collection) As Long
'Grava o Pedido de Cotacao

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Cotacao_Grava

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 63178

    'Le e locka os objetos associados a Cotacao
    lErro = CF("Cotacao_ObjetosAssociados_Lock", objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then Error 63179

    'Atualiza Cotacao, CotacaoCondPagto e CotacaoProduto
    lErro = CF("Cotacao_Inclui", objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then Error 63180

    'Inclui os Pedidos de Cotacao na tabela PedidoCotacao
    lErro = CF("PedidosCotacao_Inclui", objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then Error 63181

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 63524

    Cotacao_Grava = SUCESSO

    Exit Function

Erro_Cotacao_Grava:

    Cotacao_Grava = Err

    Select Case Err

        Case 63178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 63179, 63180, 63181
            'Erro tratado na rotina chamada

        Case 63524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147482)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Cotacao_Automatica(lCotacao As Long) As Long
'Devolve o proximo código de cotacao disponivel para a FilialEmpresa em questao.

Dim lErro As Long

On Error GoTo Erro_Cotacao_Automatica

    'Obtem o proximo codigo de Cotacao disponivel
    lErro = CF("Config_ObterAutomatico", "ComprasConfig", "NUM_PROXIMO_CODIGO_COTACAO", "Cotacao", "Codigo", lCotacao)
    If lErro <> SUCESSO Then Error 63093

    Cotacao_Automatica = SUCESSO

    Exit Function

Erro_Cotacao_Automatica:

    Cotacao_Automatica = Err

    Select Case Err

        Case 63093
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147483)

    End Select

    Exit Function

End Function

Function PedidoCotacaoNumInt_Automatico(lNumIntPedCotacao As Long) As Long
'gera o próximo número interno disponível para PedidoCotacao

Dim lErro As Long
Dim objPedidoCotacao As New ClassPedidoCotacao

On Error GoTo Erro_PedidoCotacaoNumInt_Automatico

    'Gera o proximo numero interno para PedidoCotacao
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_PEDIDOCOTACAO", lNumIntPedCotacao)
    If lErro <> SUCESSO Then Error 63194

    PedidoCotacaoNumInt_Automatico = SUCESSO

    Exit Function

Erro_PedidoCotacaoNumInt_Automatico:

    PedidoCotacaoNumInt_Automatico = Err

    Select Case Err

        Case 63194
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147484)

    End Select

    Exit Function

End Function

Function PedidoCotacao_Automatico(lPedCotacao As Long) As Long
'Gera o próximo código disponível para PedidoCotacao

Dim lErro As Long

On Error GoTo Erro_PedidoCotacao_Automatico

    'Obtem o proximo codigo disponivel para PedidoCotacao
    lErro = CF("Config_ObterAutomatico_EmTrans", "ComprasConfig", "PROXIMO_CODIGO_PEDIDOCOTACAO", "PedidoCotacao", "Codigo", lPedCotacao)
    If lErro <> SUCESSO Then Error 63195

    PedidoCotacao_Automatico = SUCESSO

    Exit Function

Erro_PedidoCotacao_Automatico:

    PedidoCotacao_Automatico = Err

    Select Case Err

        Case 63195
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147485)

    End Select

    Exit Function

End Function

Function ItensPedCotacao_Automatico(lNumIntItemPC As Long) As Long
'Gera o próximo NumInt de ItemPedCotacao

Dim lErro As Long

On Error GoTo Erro_ItensPedCotacao_Automatico

    'Obtem o próximo NumInt de ItemPedCotacao
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_ITEM_PEDCOTACAO", lNumIntItemPC)
    If lErro <> SUCESSO Then Error 63199

    ItensPedCotacao_Automatico = SUCESSO

    Exit Function

Erro_ItensPedCotacao_Automatico:

    ItensPedCotacao_Automatico = Err

    Select Case Err

        Case 63199
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147486)

    End Select

    Exit Function

End Function

Function CotacaoProdutoItemRC_Inclui(ByVal objCotacao As ClassCotacao) As Long

Dim lErro As Long
Dim lComando As Long
Dim objItemReqCompras As New ClassItemReqCompras
Dim iIndice As Integer
Dim objCotacaoProduto As ClassCotacaoProduto

On Error GoTo Erro_CotacaoProdutoItemRC_Inclui

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 63526

    For Each objCotacaoProduto In objCotacao.colCotacaoProduto

        For Each objItemReqCompras In objCotacaoProduto.colItemReqCompras

            'Inclui na tabela CotacaoProdutoItemRC os itens armazenados em colItemReqCompras
            lErro = Comando_Executar(lComando, "INSERT INTO CotacaoProdutoItemRC (CotacaoProduto,ItemReqCompra,QuantidadeCotar) VALUES (?,?,?)", objCotacaoProduto.lNumIntDoc, objItemReqCompras.lNumIntDoc, objItemReqCompras.dQuantCotar)
            If lErro <> AD_SQL_SUCESSO Then Error 63527

        Next
    Next

    'Fecha o  comando
    Comando_Fechar (lComando)

    CotacaoProdutoItemRC_Inclui = SUCESSO

    Exit Function

Erro_CotacaoProdutoItemRC_Inclui:

    CotacaoProdutoItemRC_Inclui = Err

    Select Case Err

        Case 63526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 63527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAOPRODUTOITEMRC", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147487)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Cotacao_Grava_Pedidos(ByVal objCotacao As ClassCotacao, ByVal colPedidoCotacao As Collection) As Long
'Grava Pedidos de Cotacao gerados

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Cotacao_Grava_Pedidos

     'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 63523

    lErro = CF("Cotacao_Grava_Pedidos_EmTrans", objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then gError 63522

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 63525

    Cotacao_Grava_Pedidos = SUCESSO

    Exit Function

Erro_Cotacao_Grava_Pedidos:

    Cotacao_Grava_Pedidos = gErr

    Select Case gErr

        Case 63523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 63522

        Case 63525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147488)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function GeracaoReqPedVenda_Grava(ByVal objReqCompras As ClassRequisicaoCompras, ByVal colPedidoVenda As Collection) As Long
'Busca no BD os Pedidos de Venda de acordo com os parametros selecionados no TabSelecao

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 3) As Long
Dim lCodigo As Long
Dim objPedVenda As New ClassPedidoDeVenda
Dim iIndice As Integer
Dim lReqCompra As Long
Dim objItemPV As New ClassItemPedido

On Error GoTo Erro_GeracaoReqPedVenda_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 68159
    Next

    'Abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 68160

    'Obtem o próximo código para Requisicao de Compras
    lErro = CF("RequisicaoCompras_Codigo_Automatico_EmTrans", lCodigo)
    If lErro <> SUCESSO Then gError 68161

    objReqCompras.lCodigo = lCodigo
    objReqCompras.iFilialEmpresa = giFilialEmpresa

    'Grava a Requisicao de Compra
    lErro = CF("RequisicaoCompra_Grava_EmTrans", objReqCompras)
    If lErro <> SUCESSO Then gError 68162
    
'    For Each objPedVenda In colPedidoVenda
'
'        objPedVenda.iFilialEmpresa = giFilialEmpresa
'
'        'Busca no BD o Pedido de Venda com código e filialempresa passados como parametros
'        lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM PedidosDeVenda WHERE FilialEmpresa=? AND Codigo=?", 0, lCodigo, objPedVenda.iFilialEmpresa, objPedVenda.lCodigo)
'        If lErro <> AD_SQL_SUCESSO Then gError 68273
'
'        lErro = Comando_BuscarPrimeiro(alComando(0))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68301
'
'        If lErro = AD_SQL_SEM_DADOS Then gError 68302
'
'        'Faz lock no Pedido de Venda lido
'        lErro = Comando_LockExclusive(alComando(0))
'        If lErro <> SUCESSO Then gError 68163
'
'''''        'Altera o campo ReqCompra de PedidosDeVenda
'''''        lErro = Comando_ExecutarPos(alComando(1), "UPDATE PedidosDeVenda SET ReqCompra=?", alComando(0), objReqCompras.lNumIntDoc)
'''''        If lErro <> AD_SQL_SUCESSO Then gError 68257
'
'        For Each objItemPV In objPedVenda.colItensPedido
'
'            If objItemPV.dQuantidade > 0 Then
'
'                'Busca no Bd o Item do Pedido de Venda cujo código e produto são informados
'                lErro = Comando_ExecutarPos(alComando(2), "SELECT CodPedido FROM ItensPedidoDeVenda WHERE CodPedido=? AND Produto=?", 0, lCodigo, objItemPV.lCodPedido, objItemPV.sProduto)
'                If lErro <> AD_SQL_SUCESSO Then gError 68274
'
'                lErro = Comando_BuscarPrimeiro(alComando(2))
'                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 68275
'
'                'Se não encontrou o Item de Pedido de Venda ==> erro
'                If lErro = AD_SQL_SEM_DADOS Then gError 68276
'
'                'Altera o campo QuantSC na tabela ItensPedidoDeVenda
'                lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensPedidoDeVenda SET QuantSC=QuantSC + ? ", alComando(2), objItemPV.dQuantSC)
'                If lErro <> AD_SQL_SUCESSO Then gError 68277
'
'            End If
'
'        Next
'
'    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 68164

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    GeracaoReqPedVenda_Grava = SUCESSO

    Exit Function

Erro_GeracaoReqPedVenda_Grava:

    GeracaoReqPedVenda_Grava = gErr

    Select Case gErr

        Case 68159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 68160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 68161, 68162, 68163
            'Erros tratados nas rotinas chamadas

        Case 68164
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case 68257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDVENDA", gErr, objPedVenda.lCodigo)

        Case 68273, 68301
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", gErr)

        Case 68274, 68275
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case 68276
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_NAO_ENCONTRADO", gErr, objItemPV.lNumIntDoc)

        Case 68277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", gErr)

        Case 68302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedVenda.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147489)

    End Select

    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RequisicaoCompra_Grava_EmTrans(ByVal objRequisicaoCompra As ClassRequisicaoCompras, Optional ByVal bGeracao As Boolean = False) As Long
'Grava a Requisição de Compra

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objItemReqCompra As ClassItemReqCompras
Dim lNumIntItem As Long
Dim dtDataEnvio As Date
Dim lRequisitante As Long
Dim iStatus As Integer
'Inserido por Wagner
Dim objItemOP As New ClassItemOP
Dim objOP As New ClassOrdemDeProducao
Dim sOPCodigo As String
Dim lNumIntDocItemOP As Long
Dim tReqCompra As typeRequisicaoCompras

On Error GoTo Erro_RequisicaoCompra_Grava_EmTrans

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 61994
    Next
    
    'Faz "lock" nas tabelas relacionadas a Requisição Compras
    lErro = CF("RequisicaoCompra_Lock_Gravacao", objRequisicaoCompra)
    If lErro <> SUCESSO Then gError 65502

    'Faz "lock" nas tabelas relacionadas aos itens da Requisição de compras
    lErro = CF("ItensReqCompra_Lock_Gravacao", objRequisicaoCompra)
    If lErro <> SUCESSO Then gError 65503

    'Busca em RequisicaoCompra uma requisição com o mesmo código e FilialEmpresa
    sOPCodigo = String(STRING_OPCODIGO, 0)
    tReqCompra.sUsuReg = String(STRING_USUARIO_CODIGO, 0)
    tReqCompra.sUsuRegAprov = String(STRING_USUARIO_CODIGO, 0)
    tReqCompra.sUsuRegBaixa = String(STRING_USUARIO_CODIGO, 0)
    tReqCompra.sUsuRegEnvio = String(STRING_USUARIO_CODIGO, 0)
    
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, DataEnvio, Requisitante, Status, OPCodigo, NumIntDocItemOP, UsuReg, UsuRegAprov, UsuRegBaixa, UsuRegEnvio FROM RequisicaoCompraN WHERE Codigo = ? AND FilialEmpresa = ?", 0, lNumIntDoc, dtDataEnvio, lRequisitante, iStatus, sOPCodigo, lNumIntDocItemOP, tReqCompra.sUsuReg, tReqCompra.sUsuRegAprov, tReqCompra.sUsuRegBaixa, tReqCompra.sUsuRegEnvio, objRequisicaoCompra.lCodigo, objRequisicaoCompra.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 61999

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65500
   
    'se a requisição estiver baixada
    If lErro = AD_SQL_SUCESSO And iStatus = REQUISICAOCOMPRA_STATUS_BAIXADO Then gError 61998

    'Se a Requisição de Compras já foi enviada, Erro
    If lErro = AD_SQL_SUCESSO And dtDataEnvio <> DATA_NULA Then gError 66607

    'Se não encontrou, INSERÇÃO
    If lErro = AD_SQL_SEM_DADOS Then

        If (Not bGeracao) And objRequisicaoCompra.lRequisitante = REQUISITANTE_AUTOMATICO_CODIGO Then Error 62681
        
        'Gera número interno para a próxima Requisição de Compra
        lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_RC", lNumIntDoc)
        If lErro <> SUCESSO Then gError 65504

        'Guarda NumIntDoc em objRequisicaoCompra
        objRequisicaoCompra.lNumIntDoc = lNumIntDoc
        
        objRequisicaoCompra.bInclusao = True
        objRequisicaoCompra.sUsuReg = gsUsuario

        'Grava a nova Requisição de Compras no BD
        If objRequisicaoCompra.dtDataEnvio <> DATA_NULA Then
        
            objRequisicaoCompra.bEnvio = True
            objRequisicaoCompra.sUsuRegEnvio = gsUsuario
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO RequisicaoCompraN (NumIntDoc, FilialEmpresa, Codigo, Data, DataEnvio, DataLimite,DataBaixa, Urgente, Requisitante, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao, OPCodigo, Digitador, Status, NumIntDocItemOP, DataReg, HoraReg, UsuReg, DataRegEnvio, HoraRegEnvio, UsuRegEnvio) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objRequisicaoCompra.lNumIntDoc, objRequisicaoCompra.iFilialEmpresa, objRequisicaoCompra.lCodigo, objRequisicaoCompra.dtData, objRequisicaoCompra.dtDataEnvio, objRequisicaoCompra.dtDataLimite, objRequisicaoCompra.dtDataBaixa, objRequisicaoCompra.lUrgente, objRequisicaoCompra.lRequisitante, objRequisicaoCompra.sCcl, objRequisicaoCompra.iFilialCompra, objRequisicaoCompra.iTipoDestino, objRequisicaoCompra.lFornCliDestino, objRequisicaoCompra.iFilialDestino, objRequisicaoCompra.lObservacao, objRequisicaoCompra.iTipoTributacao, objRequisicaoCompra.sOPCodigo, objRequisicaoCompra.sDigitador, REQUISICAOCOMPRA_STATUS_ATIVO, objRequisicaoCompra.lNumIntDocItemOP, Date, CDbl(Time), gsUsuario, Date, CDbl(Time), gsUsuario)
        
        Else
            lErro = Comando_Executar(alComando(1), "INSERT INTO RequisicaoCompraN (NumIntDoc, FilialEmpresa, Codigo, Data, DataEnvio, DataLimite,DataBaixa, Urgente, Requisitante, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao, OPCodigo, Digitador, Status, NumIntDocItemOP, DataReg, HoraReg, UsuReg) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objRequisicaoCompra.lNumIntDoc, objRequisicaoCompra.iFilialEmpresa, objRequisicaoCompra.lCodigo, objRequisicaoCompra.dtData, objRequisicaoCompra.dtDataEnvio, objRequisicaoCompra.dtDataLimite, objRequisicaoCompra.dtDataBaixa, objRequisicaoCompra.lUrgente, objRequisicaoCompra.lRequisitante, objRequisicaoCompra.sCcl, objRequisicaoCompra.iFilialCompra, objRequisicaoCompra.iTipoDestino, objRequisicaoCompra.lFornCliDestino, objRequisicaoCompra.iFilialDestino, objRequisicaoCompra.lObservacao, objRequisicaoCompra.iTipoTributacao, objRequisicaoCompra.sOPCodigo, objRequisicaoCompra.sDigitador, REQUISICAOCOMPRA_STATUS_ATIVO, objRequisicaoCompra.lNumIntDocItemOP, Date, CDbl(Time), gsUsuario)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 65505

        'Grava os Itens da Requisição no BD
        lErro = CF("ItensReqCompra_Inclui", objRequisicaoCompra)
        If lErro <> SUCESSO Then gError 65506

    'Encontrou, ALTERAÇÃO
    Else

        objRequisicaoCompra.sUsuReg = tReqCompra.sUsuReg
        objRequisicaoCompra.sUsuRegAprov = tReqCompra.sUsuRegAprov
        objRequisicaoCompra.sUsuRegBaixa = tReqCompra.sUsuRegBaixa
        objRequisicaoCompra.sUsuRegEnvio = tReqCompra.sUsuRegEnvio
        
        'Se a RC vai ser enviada para compras
        If objRequisicaoCompra.dtDataEnvio <> DATA_NULA Then
            
            'Se está associada a uma OP
            If Len(Trim(sOPCodigo)) > 0 Then
                
                'Se é um item Para cada OP/RC
                If lNumIntDocItemOP <> 0 Then
                    objItemOP.lNumIntDoc = lNumIntDocItemOP
                    
                    'Lê a o item OP
                    lErro = CF("ItemOP_Le_NumIntDoc", objItemOP)
                    If lErro <> SUCESSO And lErro <> 33226 Then gError 131273
                
                    'Se o item estiver bloqueado => Erro
                    If objItemOP.iSituacao = ITEMOP_SITUACAO_DESAB Then gError 131274
            
                Else
                
                    objOP.sCodigo = sOPCodigo
                    objOP.iFilialEmpresa = objRequisicaoCompra.iFilialEmpresa
                    
                    'Lê a OP
                    lErro = CF("OrdemDeProducao_Le_ComItens", objOP)
                    If lErro <> SUCESSO And lErro <> 21960 Then gError 131275
                    
                    For Each objItemOP In objOP.colItens
                    
                        If objItemOP.iSituacao = ITEMOP_SITUACAO_DESAB Then gError 131276
                    
                    Next
                
                End If
            
            End If
        End If
        
        'Faz LockExclusive na Requisição
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 66608

        objRequisicaoCompra.lNumIntDoc = lNumIntDoc

        If lRequisitante <> REQUISITANTE_AUTOMATICO_CODIGO And objRequisicaoCompra.lRequisitante = REQUISITANTE_AUTOMATICO_CODIGO Then Error 62682

        'Atualiza os itens da Requisição de Compras
        lErro = CF("ItensReqCompra_Exclui", objRequisicaoCompra)
        If lErro <> SUCESSO Then gError 66609

        lErro = CF("ItensReqCompra_Inclui", objRequisicaoCompra)
        If lErro <> SUCESSO Then gError 66610

        'Atualiza os Dados da Requisição de Compras
        If objRequisicaoCompra.dtDataEnvio <> DATA_NULA And dtDataEnvio = DATA_NULA Then
        
            objRequisicaoCompra.bEnvio = True
            objRequisicaoCompra.sUsuRegEnvio = gsUsuario
        
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE RequisicaoCompraN SET FilialEmpresa = ?, Data = ?, DataEnvio = ?, DataLimite = ?, Urgente = ?, Requisitante = ?, Digitador = ?, Ccl = ?, FilialCompra = ?, TipoDestino = ?, FornCliDestino = ?, FilialDestino = ?, Observacao = ?, TipoTributacao = ?, OPCodigo = ?, DataRegEnvio = ?, HoraRegEnvio = ?, UsuRegEnvio = ?", alComando(0), objRequisicaoCompra.iFilialEmpresa, objRequisicaoCompra.dtData, objRequisicaoCompra.dtDataEnvio, objRequisicaoCompra.dtDataLimite, objRequisicaoCompra.lUrgente, objRequisicaoCompra.lRequisitante, objRequisicaoCompra.sDigitador, objRequisicaoCompra.sCcl, objRequisicaoCompra.iFilialCompra, objRequisicaoCompra.iTipoDestino, objRequisicaoCompra.lFornCliDestino, objRequisicaoCompra.iFilialDestino, objRequisicaoCompra.lObservacao, objRequisicaoCompra.iTipoTributacao, objRequisicaoCompra.sOPCodigo, Date, CDbl(Time), gsUsuario)
        Else
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE RequisicaoCompraN SET FilialEmpresa = ?, Data = ?, DataEnvio = ?, DataLimite = ?, Urgente = ?, Requisitante = ?, Digitador = ?, Ccl = ?, FilialCompra = ?, TipoDestino = ?, FornCliDestino = ?, FilialDestino = ?, Observacao = ?, TipoTributacao = ?, OPCodigo = ?", alComando(0), objRequisicaoCompra.iFilialEmpresa, objRequisicaoCompra.dtData, objRequisicaoCompra.dtDataEnvio, objRequisicaoCompra.dtDataLimite, objRequisicaoCompra.lUrgente, objRequisicaoCompra.lRequisitante, objRequisicaoCompra.sDigitador, objRequisicaoCompra.sCcl, objRequisicaoCompra.iFilialCompra, objRequisicaoCompra.iTipoDestino, objRequisicaoCompra.lFornCliDestino, objRequisicaoCompra.iFilialDestino, objRequisicaoCompra.lObservacao, objRequisicaoCompra.iTipoTributacao, objRequisicaoCompra.sOPCodigo)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 66616

    End If

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RequisicaoCompra_Grava_EmTrans = SUCESSO

    Exit Function

Erro_RequisicaoCompra_Grava_EmTrans:

    RequisicaoCompra_Grava_EmTrans = gErr

    Select Case gErr

        Case 61994
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 61995
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 61996, 61997
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAO_COMPRA_BAIXADA", gErr, objRequisicaoCompra.lCodigo)

        Case 61998
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_COMPRA_BAIXADA", gErr, objRequisicaoCompra.lCodigo)

        Case 61999, 65500
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr, objRequisicaoCompra.lCodigo)

        Case 62681, 62682
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_REQUISITANTE_AUTOMATICO", Err)
        
        Case 65502, 65503, 65504, 65506, 66609, 66610, 131273, 131275

        Case 65505
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REQUISICAO_COMPRA", gErr, objRequisicaoCompra.lCodigo)

        Case 65507
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 66607
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_ENVIADA", gErr, objRequisicaoCompra.lCodigo)

        Case 66608
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISICAOCOMPRA", gErr, objRequisicaoCompra.lCodigo)

        Case 66616
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_REQUISICAOCOMPRA", gErr, objRequisicaoCompra.lCodigo)

        Case 131274, 131276
            Call Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIO_ITEMOP", gErr, objItemOP.iItem, objItemOP.sCodigo, objItemOP.sJustificativaBloqueio)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147490)

    End Select

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensReqCompra_Inclui(ByVal objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Insere os Itens da Requisição Compras no BD

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim objItemReqCompra As ClassItemReqCompras
Dim iIndice As Integer
Dim lNumIntItem As Long

On Error GoTo Erro_ItensReqCompra_Inclui

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError (65531)
    Next

    'Para cada item da Requisição de Compras
    For Each objItemReqCompra In objRequisicaoCompra.colItens

        'Gera numero interno automatico para o item
        lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_ITEM_RC", lNumIntItem)
        If lErro <> SUCESSO Then gError (65533)

        objItemReqCompra.lNumIntDoc = lNumIntItem

        'Grava o Item da Requisição de Compra no BD
        lErro = Comando_Executar(alComando(0), "INSERT INTO ItensReqCompraN (NumIntDoc, ReqCompra, Produto, DescProduto, Quantidade, UM, Ccl, Almoxarifado, ContaContabil, TipoTributacao, Observacao, Fornecedor, Filial, Exclusivo, StatusBaixa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objItemReqCompra.lNumIntDoc, objRequisicaoCompra.lNumIntDoc, objItemReqCompra.sProduto, objItemReqCompra.sDescProduto, objItemReqCompra.dQuantidade, objItemReqCompra.sUM, objItemReqCompra.sCcl, objItemReqCompra.iAlmoxarifado, objItemReqCompra.sContaContabil, objItemReqCompra.iTipoTributacao, objItemReqCompra.lObservacao, objItemReqCompra.lFornecedor, objItemReqCompra.iFilial, objItemReqCompra.iExclusivo, ITENSREQCOMPRA_STATUS_ATIVO)
        If lErro <> AD_SQL_SUCESSO Then gError (65532)

        lErro = ItemPVItemRC_Inclui(objItemReqCompra)
        If lErro <> SUCESSO Then gError 86181

    Next

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensReqCompra_Inclui = SUCESSO

    Exit Function

Erro_ItensReqCompra_Inclui:

    ItensReqCompra_Inclui = gErr

    Select Case gErr

        Case 65531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65532
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMREQCOMPRA", gErr, objItemReqCompra.lNumIntDoc)

        Case 65533, 86181

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147491)

    End Select

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensReqCompra_Exclui(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Exclui os Itens da Requisição de Compras

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntItem As Long
Dim objItemReqCompra As New ClassItemReqCompras

On Error GoTo Erro_ItensReqCompra_Exclui

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66611
    Next

    'Lê o item da Requisição
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ItensReqCompra WHERE ReqCompra = ?", 0, lNumIntItem, objRequisicaoCompras.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 66612

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66613

    'Para cada item de Requisição
    Do While lErro <> AD_SQL_SEM_DADOS

        'Exclui o Item do banco de dados
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItensReqCompra", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 66615

        objItemReqCompra.lNumIntDoc = lNumIntItem

        lErro = ItemPVItemRC_Exclui(objItemReqCompra)
        If lErro <> SUCESSO Then gError 178636

        'Busca Próximo Item
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66614


    Loop

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensReqCompra_Exclui = SUCESSO

    Exit Function

Erro_ItensReqCompra_Exclui:

    ItensReqCompra_Exclui = gErr

    Select Case gErr

        Case 66611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66612, 66613, 66614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)

        Case 66615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMREQCOMPRA", gErr, lNumIntItem)

        Case 178636

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147492)

    End Select

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Observacao_Automatico(lNumInt As Long) As Long
'Gera o proximo numero interno da Tabela de Observacao

Dim lErro As Long

On Error GoTo Erro_Observacao_Automatico

    'Chama a rotina que gera o sequencial
    lErro = CF("Config_ObterAutomatico_EmTrans", "ComprasConfig", "NUM_PROX_OBSERVACAO", "Observacao", "NumInt", lNumInt)
    If lErro <> SUCESSO Then Error 56238

    Observacao_Automatico = SUCESSO

    Exit Function

Erro_Observacao_Automatico:

    Observacao_Automatico = Err

    Select Case Err

        Case 56238
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147493)

    End Select

    Exit Function

End Function

Function RequisicaoCompras_Codigo_Automatico_EmTrans(lCodigo As Long) As Long
'Gera código automático para a próxima Requisição de Compras

Dim lErro As Long
Dim sCodigo As String, sTabelaConfig As String, sTabelaObjeto As String
Dim sCampo As String

On Error GoTo Erro_RequisicaoCompras_Codigo_Automatico_EmTrans

    sCodigo = "NUM_PROXIMO_CODIGO_RC"
    sTabelaConfig = "ComprasConfig"
    sTabelaObjeto = "RequisicaoCompraN"
    sCampo = "Codigo"

    'Chama Config_ObterAutomatico
    lErro = CF("Config_ObterAutomatico_EmTrans", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lCodigo)
    If lErro <> SUCESSO Then gError 68219

    RequisicaoCompras_Codigo_Automatico_EmTrans = SUCESSO

    Exit Function

Erro_RequisicaoCompras_Codigo_Automatico_EmTrans:

    RequisicaoCompras_Codigo_Automatico_EmTrans = gErr

    Select Case gErr

        Case 68219 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147494)

    End Select

    Exit Function

End Function

Function PedidoCompra_Codigo_Automatico_EmTrans(lCodigo As Long) As Long
'Gera o proximo codigo da Tabela de Pedido Compra

Dim lErro As Long

On Error GoTo Erro_PedidoCompra_Codigo_Automatico_EmTrans

    'Chama a rotina que gera o sequencial
    lErro = CF("Config_ObterAutomatico_EmTrans", "ComprasConfig", "NUM_PROXIMO_CODIGO_PC", "PedidoCompraN", "Codigo", lCodigo)
    If lErro <> SUCESSO Then Error 56114

    PedidoCompra_Codigo_Automatico_EmTrans = SUCESSO

    Exit Function

Erro_PedidoCompra_Codigo_Automatico_EmTrans:

    PedidoCompra_Codigo_Automatico_EmTrans = Err

    Select Case Err

        Case 56114
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147495)

    End Select

    Exit Function

End Function

Function Concorrencia_Baixa(ByVal objConcorrencia As ClassConcorrencia) As Long
'Baixa a Concorrência passada por objConcorrencia

Dim lErro As Long
Dim alComando(11) As Long
Dim iIndice As Integer
Dim tConcorrencia As typeConcorrencia
Dim tItemConcorrencia As typeItemConcorrencia
Dim tCotacaoItemConc As typeCotacaoItemConcorrencia
Dim lItemConcorrencia As Long
Dim tQuantSupl As typeQuantSuplementar
Dim iStatus As Integer

On Error GoTo Erro_Concorrencia_Baixa

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 67813
    Next
        
    tConcorrencia.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)
    
    'Lê a Concorrência passada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, FilialEmpresa, Codigo, Descricao, Data, Comprador, TaxaFinanceira, TipoDestino, FornCliDestino, FilialDestino, Status FROM ConcorrenciaN WHERE NumIntDoc = ?", 0, _
    tConcorrencia.lNumIntDoc, tConcorrencia.iFilialEmpresa, tConcorrencia.lCodigo, tConcorrencia.sDescricao, tConcorrencia.dtData, tConcorrencia.iComprador, tConcorrencia.dTaxaFinanceira, tConcorrencia.iTipoDestino, tConcorrencia.lFornCliDestino, tConcorrencia.iFilialDestino, iStatus, objConcorrencia.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 67814
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67815
        
    'Se não encontrou a Concorrência, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67816
    
    If iStatus = CONCORRENCIA_STATUS_BAIXADO Then gError 89482
    
    'Faz "lock" na Concorrência
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 67817
    
    tItemConcorrencia.sProduto = String(STRING_PRODUTO, 0)
    tItemConcorrencia.sUM = String(STRING_UM_SIGLA, 0)
    
    'Lê os Itens da Concorrência
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, Produto, Fornecedor, Filial, Quantidade, UM, DataNecessidade FROM ItensConcorrenciaN WHERE Concorrencia = ? AND Status = ?", 0, _
    tItemConcorrencia.lNumIntDoc, tItemConcorrencia.sProduto, tItemConcorrencia.lFornecedor, tItemConcorrencia.iFilial, tItemConcorrencia.dQuantidade, tItemConcorrencia.sUM, tItemConcorrencia.dtDataNecessidade, objConcorrencia.lNumIntDoc, ITENSCONCORRENCIA_STATUS_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 67818
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67819
    
    'Se não encontrar Itens, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67823
    
    'Equanto encontrar
    Do While lErro = AD_SQL_SUCESSO
    
        'Faz "Lock" no Item da Concorrência
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 67820
        
        tCotacaoItemConc.sMotivoEscolha = String(STRING_MOTIVO_ESCOLHA, 0)
        
        'Lê CotacaoItemConcorrencia
        lErro = Comando_ExecutarPos(alComando(6), "SELECT NumIntDoc,ItemConcorrencia,ItemCotacao,ValorPresente,Escolhido,MotivoEscolha,QuantidadeComprar,DataEntrega,PrecoAjustado FROM CotacaoItemConcorrenciaN WHERE ItemConcorrencia=? AND Status = ?", 0, _
        tCotacaoItemConc.lNumIntDoc, lItemConcorrencia, tCotacaoItemConc.lItemCotacao, tCotacaoItemConc.dValorPresente, tCotacaoItemConc.iEscolhido, tCotacaoItemConc.sMotivoEscolha, tCotacaoItemConc.dQuantidadeComprar, tCotacaoItemConc.dtDataEntrega, _
        tCotacaoItemConc.dPrecoAjustado, tItemConcorrencia.lNumIntDoc, COTACAOITEMCONCORRENCIA_STATUS_ATIVO)
        If lErro <> AD_SQL_SUCESSO Then gError 74862

        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 74863
        
        Do While lErro = AD_SQL_SUCESSO
            
            'Exclui CotacaoItemConcorrencia
            lErro = Comando_ExecutarPos(alComando(7), "UPDATE CotacaoItemConcorrenciaN SET Status = ?", alComando(6), COTACAOITEMCONCORRENCIA_STATUS_BAIXADO)
            If lErro <> AD_SQL_SUCESSO Then gError 74864
            
            lErro = Comando_BuscarProximo(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 74863
                
        Loop
        
        lErro = Comando_ExecutarPos(alComando(9), "SELECT  TipoDestino, FornCliDestino, FilialDestino, Quantidade FROM QuantidadesSuplementares WHERE ItemConcorrencia = ?", 0, tQuantSupl.iTipoDestino, tQuantSupl.lFornCliDestino, tQuantSupl.iFilialDestino, tQuantSupl.dQuantidade, tItemConcorrencia.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 86112
        
        lErro = Comando_BuscarPrimeiro(alComando(9))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86113
        
        Do While lErro = AD_SQL_SUCESSO
        
            lErro = Comando_Executar(alComando(10), "INSERT INTO QuantidadesSuplementaresBaixa (ItemConcorrencia,TipoDestino,FornCliDestino,FilialDestino,Quantidade) VALUES (?,?,?,?,?) ", tItemConcorrencia.lNumIntDoc, tQuantSupl.iTipoDestino, tQuantSupl.lFornCliDestino, tQuantSupl.iFilialDestino, tQuantSupl.dQuantidade)
            If lErro <> AD_SQL_SUCESSO Then gError 86114
            
            lErro = Comando_ExecutarPos(alComando(11), "DELETE FROM QuantidadesSuplementares", alComando(9))
            If lErro <> AD_SQL_SUCESSO Then gError 86115
            
            lErro = Comando_BuscarProximo(alComando(9))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86116
            
        Loop
    
        'Exclui o Item da Concorrência
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensConcorrenciaN SET Status = ?", alComando(1), ITENSCONCORRENCIA_STATUS_BAIXADO)
        If lErro <> AD_SQL_SUCESSO Then gError 67821
        
        'Busca próximo item concorrência
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67837
        
    Loop
    
    'Atualiza a Concorrência como baixada
    lErro = Comando_ExecutarPos(alComando(4), "UPDATE ConcorrenciaN SET Status = ?", alComando(0), CONCORRENCIA_STATUS_BAIXADO)
    If lErro <> AD_SQL_SUCESSO Then gError 67824
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Concorrencia_Baixa = SUCESSO

    Exit Function

Erro_Concorrencia_Baixa:
    
    Concorrencia_Baixa = gErr
    
    Select Case gErr
    
        Case 67813
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                
        Case 67814, 67815
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", gErr)
            
        Case 67816
            Call Rotina_Erro(vbOKOnly, "ERRO_CONCORRENCIA_NAO_CADASTRADA", gErr, objConcorrencia.lCodigo)
            
        Case 67817
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONCORRENCIA", gErr, objConcorrencia.lCodigo)
            
        Case 67818, 67819, 67837
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCONCORRENCIA", gErr, objConcorrencia.lCodigo)
            
        Case 67820
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSCONCORRENCIA", gErr)
        
        Case 67821
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSCONCORRENCIA", gErr, tItemConcorrencia.lNumIntDoc)
            
        Case 67822
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMCONCONCORRENICA", gErr, tItemConcorrencia.lNumIntDoc)
            
        Case 67823
            Call Rotina_Erro(vbOKOnly, "ERRO_CONCORRENCIA_AUSENCIA_ITENS", gErr, objConcorrencia.lCodigo)

        Case 67824
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONCORRENCIA", gErr, objConcorrencia.lCodigo)

        Case 67825
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONCORRENCIABAIXADA", gErr, objConcorrencia.lCodigo)

        Case 74862, 74863
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOITEMCONCORRENCIA", gErr)

        Case 74864
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COTACAOITEMCONCORRENCA", gErr)

        Case 74865
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAOITEMCONCORRENCIA", gErr)

        Case 86112, 86113, 86116
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_QUANTIDADESSUPLEMENTARES", gErr)

        Case 86114
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_QUANTIDADESSUPLEMENTARESBAIXA", gErr)
        
        Case 86115
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_QUANTIDADESSUPLEMENTARES", gErr, tItemConcorrencia.lNumIntDoc)
        
        Case 89482
            Call Rotina_Erro(vbOKOnly, "ERRO_CONCORRENCIA_BAIXADA1", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147496)
    
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function RequisicaoCompras_Codigo_Automatico(lCodigo As Long) As Long
'Gera código automático para a próxima Requisição de Compras

Dim lErro As Long
Dim sCodigo As String, sTabelaConfig As String, sTabelaObjeto As String
Dim sCampo As String

On Error GoTo Erro_RequisicaoCompras_Codigo_Automatico

    sCodigo = "NUM_PROXIMO_CODIGO_RC"
    sTabelaConfig = "ComprasConfig"
    sTabelaObjeto = "RequisicaoCompra"
    sCampo = "Codigo"

    'Chama Config_ObterAutomatico
    lErro = CF("Config_ObterAutomatico", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lCodigo)
    If lErro <> SUCESSO Then Error 61840

    RequisicaoCompras_Codigo_Automatico = SUCESSO
    
    Exit Function

Erro_RequisicaoCompras_Codigo_Automatico:

    RequisicaoCompras_Codigo_Automatico = Err
    
    Select Case Err

        Case 61840 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147497)

    End Select

    Exit Function

End Function

Function GeracaoRequisicaoPtoPedido_Grava(ByVal objRequisicao As ClassRequisicaoCompras) As Long
'Grava a Requisição de Compra

Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_GeracaoRequisicaoPtoPedido_Grava
    
    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 67414
    
    lErro = CF("RequisicaoCompra_Grava_EmTrans", objRequisicao, True)
    If lErro <> SUCESSO Then gError 67415
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 67416
   
    GeracaoRequisicaoPtoPedido_Grava = SUCESSO

    Exit Function
    
Erro_GeracaoRequisicaoPtoPedido_Grava:

    GeracaoRequisicaoPtoPedido_Grava = gErr

    Select Case gErr
    
        Case 67414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 67415
               
        Case 67416
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147498)
    
    End Select
    
    Call Transacao_Rollback

    Exit Function

End Function


Function BloqueiosPC_Libera(ByVal colBloqueiosPC As Collection) As Long
'libera os bloqueios passados na colecao colBloqueioPC

Dim alComando(0 To 4) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim objBloqueioPC As ClassBloqueioPC

On Error GoTo Erro_BloqueiosPC_Libera

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 49230
    Next

    For Each objBloqueioPC In colBloqueiosPC

        'processa cada liberação de bloqueio. Mesmo que haja erro, prossegue para o proximo bloqueio
         lErro = BloqueiosPC_Libera_Trans(alComando, objBloqueioPC)

        'fecha os comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then Error 36646
        Next

    Next

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    BloqueiosPC_Libera = SUCESSO

    Exit Function

Erro_BloqueiosPC_Libera:

    BloqueiosPC_Libera = Err

    Select Case Err

        Case 49230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147499)

    End Select

    'fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function BloqueiosPC_Libera_Trans(alComando() As Long, ByVal objBloqueioPC As ClassBloqueioPC) As Long
'tenta liberar o bloqueio passado como parametro

Dim lTransacao As Long
Dim lErro As Long
Dim iIndice As Integer
Dim iLiberaBloqueio As Integer
Dim dtData As Date
Dim lPedCompras As Long
Dim iFilialEmpresa As Integer

On Error GoTo Erro_BloqueiosPC_Libera_Trans

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 49231

    dtData = gdtDataHoje

     iLiberaBloqueio = True

    'Busca em BloqueiosPC selecionados no grid
    lErro = Comando_ExecutarPos(alComando(0), "SELECT PedCompras FROM BloqueiosPC WHERE FilialEmpresa = ? AND PedCompras = ? AND Sequencial = ?", 0, lPedCompras, objBloqueioPC.iFilialEmpresa, objBloqueioPC.lPedCompras, objBloqueioPC.iSequencial)
    If lErro <> AD_SQL_SUCESSO Then Error 49244

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49245

    'Se nao existe nenhum BloqueioPC
    If lErro = AD_SQL_SEM_DADOS Then Error 49246

    'Lock de BloqueiosPC
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 49247

    'Seleciona o Pedido de Compra do BloqueioPC
    lErro = Comando_ExecutarPos(alComando(1), "SELECT FilialEmpresa FROM  PedidoCompra WHERE FilialEmpresa = ? and Codigo = ?", 0, iFilialEmpresa, objBloqueioPC.iFilialEmpresa, objBloqueioPC.lPedCompras)
    If lErro <> AD_SQL_SUCESSO Then Error 49278

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49279

    'Se nao encontrar o Pedido de Compra
    If lErro = AD_SQL_SEM_DADOS Then Error 49280

    'Lock de Pedido de Compra
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 49281

    'Verifica se o bloqueio é do tipo alcada
    If objBloqueioPC.iTipoBloqueio = BLOQUEIO_ALCADA Then

        lErro = CF("Alcada_Avalia", gsUsuario, objBloqueioPC.dValorPedido, dtData)
        If lErro <> SUCESSO And lErro <> 49257 Then

            iLiberaBloqueio = False
            'Se usuario nao possui alcada
            If lErro = 49258 Then Error 49277

        Else

            lErro = CF("ValorPCLiberado_Grava", objBloqueioPC.dValorPedido, objBloqueioPC.sCodUsuario, dtData)
            If lErro <> SUCESSO Then Error 49243
        End If
    End If


    If iLiberaBloqueio = True Then

        'Atualiza a tabela BloqueiosPC
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE BloqueiosPC SET DataLib = ?,Data = ?,CodUsuarioLib=?", alComando(0), objBloqueioPC.dtDataLib, objBloqueioPC.dtData, objBloqueioPC.sCodUsuarioLib)
        If lErro <> AD_SQL_SUCESSO Then Error 49249

        'Atualiza a tabela de Pedido de Compra
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE PedidoCompra SET Alcada = ?", alComando(1), objBloqueioPC.sCodUsuarioLib)
        If lErro <> AD_SQL_SUCESSO Then Error 49282

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 49248

    BloqueiosPC_Libera_Trans = SUCESSO

    Exit Function

Erro_BloqueiosPC_Libera_Trans:

    BloqueiosPC_Libera_Trans = Err

    Select Case Err


        Case 49231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 49243

        Case 49244, 49245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPC1", Err, objBloqueioPC.lPedCompras)

        Case 49246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIOPC_INEXISTENTE", Err, objBloqueioPC, lPedCompras)
            
        Case 49247
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_BLOQUEIOSPC", Err, objBloqueioPC.lPedCompras)

        Case 49248
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 49249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_BLOQUEIOSPC", Err, objBloqueioPC.lPedCompras)

        Case 49277

        Case 49278, 49279
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", Err, objBloqueioPC.lPedCompras)

        Case 49280
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", Err, objBloqueioPC.lPedCompras)
            
        Case 49281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOCOMPRA", Err, objBloqueioPC.lPedCompras)

        Case 49282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOCOMPRA", Err, objBloqueioPC.lPedCompras)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147500)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function ValorPCLiberado_Grava(ByVal dValor As Double, ByVal sCodUsuario As String, dtData As Date) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim iAno As Integer

On Error GoTo Erro_ValorPCLiberado_Grava

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 49285
    Next

    'Busca usuario na tabela ValorPCLiberado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Ano FROM ValorPCLiberado WHERE CodUsuario = ? AND Ano = ? ", 0, iAno, sCodUsuario, Year(dtData))
    If lErro <> AD_SQL_SUCESSO Then Error 49260

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49261

    'Liberacao não cadastrada ==> insere a liberacao
    If lErro = AD_SQL_SEM_DADOS Then

        lErro = Comando_Executar(alComando(1), "INSERT INTO ValorPCLiberado (CodUsuario, Ano, Liberado" + CStr(Month(dtData)) + ") VALUES (?,?,?)", sCodUsuario, Year(dtData), dValor)
        If lErro <> AD_SQL_SUCESSO Then Error 49262

    Else

        'Liberacao cadastrada ==> atualiza liberacao
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 49263

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE ValorPCLiberado SET Liberado" + CStr(Month(dtData)) + " = Liberado" + CStr(Month(dtData)) + " + ?", alComando(0), dValor)
        If lErro <> AD_SQL_SUCESSO Then Error 49264

    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    ValorPCLiberado_Grava = SUCESSO

    Exit Function

Erro_ValorPCLiberado_Grava:

    ValorPCLiberado_Grava = Err

    Select Case Err

        Case 49260, 49261
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VALORPCLIBERADO", Err)

        Case 49262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_VALORPCLIBERADO", Err, sCodUsuario)

        Case 49263
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_VALORPCLIBERADO", Err, sCodUsuario)

        Case 49264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_VALORPCLIBERADO", Err, sCodUsuario)

        Case 49285
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 147501)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoCompra_Atualiza_DataEmissao(ByVal objPedidoCompra As ClassPedidoCompras, Optional ByVal bAtualizaDataEnvio As Boolean = False) As Long
'Atualiza data de emissao no BD para a data atual

Dim lErro As Long
Dim lComando(0 To 2) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lCodigo As Long
Dim iFilialEmpresa As Long
Dim dtDataRegAprov As Date

On Error GoTo Erro_PedidoCompra_Atualiza_DataEmissao

    'Abre os  comandos
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 56156
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 56157

    'OK você tem o numintdoc do pedido para usar como filtro
    'Busca na tabela de PedidoCompra, o pedido de compra passado como parametro
    lErro = Comando_ExecutarPos(lComando(0), "SELECT NumIntDoc,Codigo,FilialEmpresa,DataRegAprov  FROM PedidoCompraN WHERE  Codigo = ? and FilialEmpresa =? AND Status = ?", 0, lNumIntDoc, lCodigo, iFilialEmpresa, dtDataRegAprov, objPedidoCompra.lCodigo, objPedidoCompra.iFilialEmpresa, PEDIDOCOMPRA_STATUS_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then Error 56158

    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56159

    'se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then Error 56348

    'Atualiza Pedido de Compra
    If bAtualizaDataEnvio And gobjCOM.iEmailAlteraDataEnvioPC = MARCADO Then
    
        objPedidoCompra.bEnvio = True
        objPedidoCompra.sUsuRegEnvio = True
        
        If dtDataRegAprov = DATA_NULA Then
            If gobjCOM.iPedCompraBloqEnvioSemAprov = MARCADO Then gError 213170
        End If
    
        lErro = Comando_ExecutarPos(lComando(1), "UPDATE PedidoCompraN SET DataEmissao = ?, DataEnvio = ?, DataRegEnvio = ?, HoraRegEnvio = ?, UsuRegEnvio = ?", lComando(0), gdtDataHoje, gdtDataHoje, Date, CDbl(Time), gsUsuario) 'OK não usar os controles da tela nas rotinas que irão subir.
    Else
        lErro = Comando_ExecutarPos(lComando(1), "UPDATE PedidoCompraN SET DataEmissao = ?", lComando(0), gdtDataHoje) 'OK não usar os controles da tela nas rotinas que irão subir.
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 56160

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 56161
    
    lErro = CF("WorkFlow_Trata_Transacao", "COM", objPedidoCompra, gsUsuario, objPedidoCompra, 1002)
    If lErro <> SUCESSO Then gError 56348

    'Fecha os comandos
    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next

    PedidoCompra_Atualiza_DataEmissao = SUCESSO

    Exit Function

Erro_PedidoCompra_Atualiza_DataEmissao:

    PedidoCompra_Atualiza_DataEmissao = Err

    Select Case Err

        Case 56156
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56157
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 56158, 56159
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", Err, objPedidoCompra.lCodigo)

        Case 56160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOCOMPRA", Err, objPedidoCompra.lCodigo)

        Case 56161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 56348

        Case 213170
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_APROVADO", gErr, objPedidoCompra.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147502)

    End Select

     Call Transacao_Rollback

    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoCompraGerado_Grava(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Grava Pedido de Compra no BD

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 3) As Long
Dim lTransacao As Long
Dim iFilialEmpresa As Integer
Dim lCodigo As Long
Dim lNumIntDoc As Long
Dim objPedidoCompra_BD As New ClassPedidoCompras
Dim iValorAlterado As Integer
Dim objItemPC As New ClassItemPedCompra
Dim objComprador As New ClassComprador
Dim objFilFornFilEmp As New ClassFilFornFilEmp
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_PedidoCompraGerado_Grava

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 53362
    Next

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 53363

    lErro = Comando_Executar(alComando(0), "SELECT FilialEmpresa,Codigo FROM PedidoCompraBaixado WHERE FilialEmpresa = ? AND Codigo = ?", iFilialEmpresa, lCodigo, objPedidoCompra.iFilialEmpresa, objPedidoCompra.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 53364

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53365

    If lErro = AD_SQL_SUCESSO Then gError 53366

    objPedidoCompra_BD.lCodigo = objPedidoCompra.lCodigo
    objPedidoCompra_BD.iFilialEmpresa = objPedidoCompra.iFilialEmpresa
    
    lErro = CF("PedidoCompra_Valida_Campos_Cust", objPedidoCompra)
    If lErro <> SUCESSO Then gError 53367

    lErro = CF("PedidoCompra_LeMod_Gravacao", objPedidoCompra)
    If lErro <> SUCESSO Then gError 53367

    lErro = CF("PedidoCompra_LeMod", objPedidoCompra_BD, alComando(1))
    If lErro <> SUCESSO And lErro <> 62646 Then gError 53368

    'Nao encontrou => erro
    If lErro <> SUCESSO Then gError 53369

    'Verifica se DataEnvio foi preenchida
    If objPedidoCompra_BD.dtDataEnvio <> DATA_NULA Then gError 53370
    
    objPedidoCompra.sUsuReg = objPedidoCompra_BD.sUsuReg

    'verifica se a DataEnvio da tela foi preenchida
    If objPedidoCompra.dtDataEnvio <> DATA_NULA Then
        
        objPedidoCompra.bEnvio = True
        objPedidoCompra.sUsuRegEnvio = gsUsuario
        
        If (Len(Trim(objPedidoCompra_BD.sAlcada)) <> 0 And (objPedidoCompra.dValorTotal > objPedidoCompra_BD.dValorTotal)) Or Len(Trim(objPedidoCompra_BD.sAlcada)) = 0 Then

            lErro = CF("BloqueioAlcada_Critica", objPedidoCompra, iValorAlterado)
            If lErro <> SUCESSO Then gError 53371

        End If

        If objPedidoCompra.colBloqueiosPC.Count > 0 Then
            Call Rotina_Aviso(vbOKOnly, "AVISO_PEDIDOCOMPRA_BLOQUEADO", objPedidoCompra.lCodigo)
            objPedidoCompra.dtDataEnvio = DATA_NULA
        Else
            lErro = VerificaLiberacao_BloqueioPC(objPedidoCompra) 'William
            If lErro <> SUCESSO Then gError 31149
        End If
        
        objFilFornFilEmp.lCodFornecedor = objPedidoCompra.lFornecedor
        objFilFornFilEmp.iCodFilial = objPedidoCompra.iFilial
        objFilFornFilEmp.iFilialEmpresa = objPedidoCompra.iFilialEmpresa
        objFilFornFilEmp.lPedidosEmAberto = objFilFornFilEmp.lPedidosEmAberto + 1

        'Grava a Estatística da Filial Fornecedor
        lErro = CF("FilialFornecedor_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_FILIALFORNFILEMP, objFilFornFilEmp)
        If lErro <> SUCESSO Then gError 89430
        
    End If

    objPedidoCompra.dtDataAlteracao = gdtDataHoje

    objPedidoCompra.lNumIntDoc = objPedidoCompra_BD.lNumIntDoc

    lErro = CF("ItensPedCompra_Lock", objPedidoCompra_BD)
    If lErro <> SUCESSO And lErro <> 56332 Then gError 53372

    Set objItemPC = objPedidoCompra_BD.colItens.Item(1)

    'Verifica se NumIntOrigem esta preenchido
    If objItemPC.lNumIntOrigem = 0 Then gError 53373

    lErro = CF("ItensPedCompra_AlteraBD", objPedidoCompra)
    If lErro <> SUCESSO Then gError 53374

    lErro = CF("LocalizacaoItemPC_Altera", objPedidoCompra)
    If lErro <> SUCESSO Then gError 53375

    lErro = CF("BloqueiosPC_Grava", objPedidoCompra)
    If lErro <> SUCESSO Then gError 53376

    'Altera o Pedido de Compra
    If objPedidoCompra.bEnvio Then
        If Len(Trim(objPedidoCompra_BD.sUsuRegAprov)) = 0 Then
            If gobjCOM.iPedCompraBloqEnvioSemAprov = MARCADO Then gError 213170
        End If
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE PedidoCompraN SET CotacaoMoeda=?, Embalagem=?, Taxa=?, Fornecedor=?, Filial=?, Comprador=?, Contato=?, TipoDestino=?, FornCliDestino=?, FilialDestino=?, DataEmissao=?, DataEnvio=?, DataAlteracao=?, CondicaoPagto=?, OutrasDespesas=?, ValorFrete=?, ValorSeguro=?, ValorDesconto=?, ValorTotal=?, ValorIPI=?, TipoFrete=?, Transportadora=?, Alcada=?, ObsEmbalagem=?, Observacao=?, DataRefFluxo=?, TabelaPreco=?, DataRegEnvio = ?, HoraRegEnvio = ?, UsuRegEnvio = ?", _
            alComando(1), objPedidoCompra.dCotacaoMoeda, objPedidoCompra.iEmbalagem, objPedidoCompra.dTaxa, objPedidoCompra.lFornecedor, objPedidoCompra.iFilial, objPedidoCompra.iComprador, objPedidoCompra.sContato, objPedidoCompra.iTipoDestino, objPedidoCompra.lFornCliDestino, objPedidoCompra.iFilialDestino, objPedidoCompra.dtDataEmissao, objPedidoCompra.dtDataEnvio, objPedidoCompra.dtDataAlteracao, objPedidoCompra.iCondicaoPagto, objPedidoCompra.dOutrasDespesas, objPedidoCompra.dValorFrete, objPedidoCompra.dValorSeguro, objPedidoCompra.dValorDesconto, objPedidoCompra.dValorTotal, objPedidoCompra.dValorIPI, objPedidoCompra.sTipoFrete, objPedidoCompra.iTransportadora, objPedidoCompra.sAlcada, objPedidoCompra.sObsEmbalagem, objPedidoCompra.lObservacao, objPedidoCompra.dtDataRefFluxo, objPedidoCompra.iTabelaPreco, Date, CDbl(Time), gsUsuario)
    Else
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE PedidoCompraN SET CotacaoMoeda=?, Embalagem=?, Taxa=?, Fornecedor=?, Filial=?, Comprador=?, Contato=?, TipoDestino=?, FornCliDestino=?, FilialDestino=?, DataEmissao=?, DataEnvio=?, DataAlteracao=?, CondicaoPagto=?, OutrasDespesas=?, ValorFrete=?, ValorSeguro=?, ValorDesconto=?, ValorTotal=?, ValorIPI=?, TipoFrete=?, Transportadora=?, Alcada=?, ObsEmbalagem=?, Observacao=?, DataRefFluxo=?, TabelaPreco=?", _
            alComando(1), objPedidoCompra.dCotacaoMoeda, objPedidoCompra.iEmbalagem, objPedidoCompra.dTaxa, objPedidoCompra.lFornecedor, objPedidoCompra.iFilial, objPedidoCompra.iComprador, objPedidoCompra.sContato, objPedidoCompra.iTipoDestino, objPedidoCompra.lFornCliDestino, objPedidoCompra.iFilialDestino, objPedidoCompra.dtDataEmissao, objPedidoCompra.dtDataEnvio, objPedidoCompra.dtDataAlteracao, objPedidoCompra.iCondicaoPagto, objPedidoCompra.dOutrasDespesas, objPedidoCompra.dValorFrete, objPedidoCompra.dValorSeguro, objPedidoCompra.dValorDesconto, objPedidoCompra.dValorTotal, objPedidoCompra.dValorIPI, objPedidoCompra.sTipoFrete, objPedidoCompra.iTransportadora, objPedidoCompra.sAlcada, objPedidoCompra.sObsEmbalagem, objPedidoCompra.lObservacao, objPedidoCompra.dtDataRefFluxo, objPedidoCompra.iTabelaPreco)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 53377

'Por Leo

    lErro = NotasPedCompra_Exclui(objPedidoCompra_BD)
    If lErro <> AD_SQL_SUCESSO Then gError 103355
    
    objPedidoCompra.lNumIntDoc = objPedidoCompra_BD.lNumIntDoc
    
    lErro = NotasPedCompra_Insere(objPedidoCompra)
    If lErro <> AD_SQL_SUCESSO Then gError 103356
        
'Leo até aqui

    '#####################################################
    'Inserido por Wagner
    'Verifica se DataEnvio foi preenchida
    If objPedidoCompra.dtDataEnvio <> DATA_NULA Then
        'Verifica se o PC está ligado a uma OP bloqueada, SE SIM => ERRO
        lErro = CF("PedidoCompra_Verifica_BloqueioOP", objPedidoCompra)
        If lErro <> SUCESSO Then gError 131657
    End If
    '#####################################################

        
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 53378

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    lErro = CF("WorkFlow_Trata_Transacao", "COM", objPedidoCompra, gsUsuario, objPedidoCompra, 1002)
    If lErro <> SUCESSO Then gError 131657

    PedidoCompraGerado_Grava = SUCESSO

    Exit Function

Erro_PedidoCompraGerado_Grava:

    PedidoCompraGerado_Grava = gErr

    Select Case gErr

        Case 31149
            'Erro já tratado na rotina chamadora

        Case 53362
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 53363
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 53364, 53365
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA_BAIXADO", gErr, objPedidoCompra.lCodigo)

        Case 53366
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_BAIXADO", gErr, objPedidoCompra.lCodigo)

        Case 53367, 53368, 53371, 53372, 53374, 53375, 53376, 89430, 103355, 103356
            'Erros tratados nas rotinas chamadas

        Case 53369
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case 53370
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_ENVIADO", gErr, objPedidoCompra.lCodigo)

        Case 53373
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_GERADO", gErr, objPedidoCompra.lCodigo)

        Case 53377
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOCOMPRA", gErr, objPedidoCompra.lCodigo)

        Case 53378
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 86161
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO1", gErr, objPedidoCompra.lCodigo)

        Case 131657

        Case 213170
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_APROVADO", gErr, objPedidoCompra_BD.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147503)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoCompra_LeMod_Gravacao(ByVal objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim lFornecedor As Double
Dim iFilial As Integer
Dim lNumInt As Long
Dim iCondicaoPagto As Integer
Dim alComando(0 To 4) As Long
Dim iIndice As Integer
Dim objItemPC As New ClassItemPedCompra
Dim objProduto As New ClassProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objLocalizacao As New ClassLocalizacaoItemPC
Dim iAlmoxarifado As Integer
Dim sCcl As String
Dim sConta As String
Dim lNumObservacao As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_PedidoCompra_LeMod_Gravacao

     'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then Error 56241
    Next


    'Le e locka FilialFornecedor
    lErro = CF("FilialFornecedor_Lock", objPedidoCompra.lFornecedor, objPedidoCompra.iFilial)
    If lErro <> SUCESSO And lErro <> 43050 Then Error 56217
    If lErro = 43050 Then Error 56447

    'Le e Locka Condicao de Pagamento
    lErro = CF("CondicaoPagto_Lock", alComando(1), objPedidoCompra.iCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 16361 Then Error 56444

    If objPedidoCompra.iTipoDestino = TIPO_DESTINO_FORNECEDOR Then

        lErro = CF("FilialFornecedor_Lock", objPedidoCompra.lFornecedor, objPedidoCompra.iFilialDestino)
        If lErro <> SUCESSO And lErro <> 42870 Then Error 56225

        'Se nao encontrou FilialFornecedor
       If lErro = 42870 Then Error 56233

    End If


    'Verifica se observacao esta preenchida
    If Len(Trim(objPedidoCompra.sObservacao)) > 0 Then

        'Busca na tabela de Observacao a observacao do Pedido
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumInt FROM Observacao WHERE Observacao = ?", 0, lNumInt, objPedidoCompra.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then Error 56234

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56235

        'Se nao encontrou
        If lErro = AD_SQL_SEM_DADOS Then

            'gera codigo automatico
            lErro = CF("Observacao_Automatico", lNumInt)
            If lErro <> SUCESSO Then Error 56240

            'Insere observacao na tabela de Observacao
            lErro = Comando_Executar(alComando(2), "INSERT INTO Observacao(Observacao,NumInt)VALUES (?,?)", objPedidoCompra.sObservacao, lNumInt)
            If lErro <> AD_SQL_SUCESSO Then Error 56239

            objPedidoCompra.lObservacao = lNumInt

        Else

            lErro = Comando_LockShared(alComando(2))
            If lErro <> AD_SQL_SUCESSO Then Error 56236

            objPedidoCompra.lObservacao = lNumInt

        End If

    End If

    'Lock shared nos Produtos e Almoxarifados
    For Each objItemPC In objPedidoCompra.colItens

        objProduto.sCodigo = objItemPC.sProduto
        lErro = CF("Produto_Lock", objProduto)
        If lErro <> SUCESSO Then Error 56300

        If Len(Trim(objItemPC.sObservacao)) > 0 Then

            'Le a observacao na tabela de Observacao
            lErro = Comando_ExecutarPos(alComando(2), "SELECT NumInt FROM Observacao WHERE Observacao = ?", 0, lNumInt, objItemPC.sObservacao)
            If lErro <> AD_SQL_SUCESSO Then Error 56242

            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56243

            'Se nao encontrou
            If lErro = AD_SQL_SEM_DADOS Then

                'gera codigo automatico
                lErro = CF("Observacao_Automatico", lNumInt)
                If lErro <> SUCESSO Then Error 56244

                'Insere observacao na tabela de Observacao
                lErro = Comando_Executar(alComando(2), "INSERT INTO Observacao(Observacao,NumInt)VALUES (?,?)", objItemPC.sObservacao, lNumInt)
                If lErro <> AD_SQL_SUCESSO Then Error 56245

                objItemPC.lObservacao = lNumInt

            Else

                lErro = Comando_LockShared(alComando(2))
                If lErro <> AD_SQL_SUCESSO Then Error 56246

                objItemPC.lObservacao = lNumInt

            End If

        End If
        For Each objLocalizacao In objItemPC.colLocalizacao

            If objLocalizacao.iAlmoxarifado <> 0 Then

                lErro = CF("Almoxarifado_Lock", objLocalizacao.iAlmoxarifado, alComando(3))
                If lErro <> SUCESSO And lErro <> 34552 Then Error 56301
                'Se nao encontrou => erro
                If lErro = 34552 Then Error 56445

            End If

            If Len(Trim(objLocalizacao.sCcl)) > 0 Then

                sCcl = String(STRING_CCL, 0)

                lErro = Comando_ExecutarPos(alComando(0), "SELECT Ccl FROM Ccl WHERE Ccl = ? ", 0, sCcl, objLocalizacao.sCcl)
                If lErro <> AD_SQL_SUCESSO Then Error 56302

                lErro = Comando_BuscarPrimeiro(alComando(0))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56303
                If lErro = AD_SQL_SEM_DADOS Then Error 56304

                lErro = Comando_LockShared(alComando(0))
                If lErro <> AD_SQL_SUCESSO Then Error 56305

            End If

            If Len(Trim(objLocalizacao.sContaContabil)) > 0 Then

                lErro = CF("Conta_Analitica_Lock", objLocalizacao.sContaContabil, MODULO_COMPRAS)
                If lErro <> SUCESSO Then Error 56306

            End If

        Next

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoCompra_LeMod_Gravacao = SUCESSO

    Exit Function

Erro_PedidoCompra_LeMod_Gravacao:

    PedidoCompra_LeMod_Gravacao = Err

    Select Case Err

        Case 56300, 56301, 56306

        Case 56302, 56303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", Err, objLocalizacao.sCcl)

        Case 56304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objLocalizacao.sCcl)

        Case 56305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", Err, objLocalizacao.sCcl)

        Case 56217, 56218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES", Err)

        Case 56219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISFORNECEDORES", Err, lFornecedor)

        Case 56223, 56224, 56225

        Case 56231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", Err, objPedidoCompra.iFilialDestino)

        Case 56233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objPedidoCompra.lFornCliDestino, objPedidoCompra.iFilialDestino)

        Case 56234, 56235, 56242, 56243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OBSERVACAO", Err, lNumInt)

        Case 56236, 56246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_OBSERVACAO", Err)

        Case 56239, 56245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OBSERVACAO", Err, lNumInt)

        Case 56240, 56244
            'Erro tratado na rotina chamada

        Case 56241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56445
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", Err, objLocalizacao.iAlmoxarifado)

        Case 56447
           lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_INEXISTENTE", Err, objPedidoCompra.iFilialEmpresa, objPedidoCompra.lFornecedor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147504)

    End Select

     For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function BloqueioAlcada_Critica(ByVal objPedidoCompra As ClassPedidoCompras, ByVal iValorAlterado As Integer) As Long
'Inclui no BD um bloqueio de Alcada, se o usuário não tiver alçada

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim dLiberado As Double
Dim objAlcada As New ClassAlcada
Dim objItemPC As ClassItemPedCompra
Dim tBloqueioPC As typeBloqueioPC
Dim objBloqueiosPC As New ClassBloqueioPC
Dim iExisteBloqAlcada As Integer
Dim iIndice As Integer
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda
Dim dValorReal As Double, dValorTotalPC As Double

On Error GoTo Erro_BloqueioAlcada_Critica

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 51215
    Next
    
    'Se a moeda nao for R$ ...
    If objPedidoCompra.iMoeda <> MOEDA_REAL Then
        
        'Se o pedido tiver a taxa preenchida ...
        If objPedidoCompra.dCotacaoMoeda <> 0 Then
            
            objPedidoCompra.dCotacaoMoeda = objPedidoCompra.dTaxa
                    
        Else
        
            'Carrega objCotacao
            objCotacao.dtData = gdtDataAtual
            
            'Preeche com a Moeda selecionada
            objCotacao.iMoeda = objPedidoCompra.iMoeda
            objCotacaoAnterior.iMoeda = objPedidoCompra.iMoeda
        
            'Chama função de leitura
            lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
            If lErro <> SUCESSO Then gError 108948
            
            'Se nao existe cotacao para a data informada => Busca a última.
            If objCotacao.dValor <> 0 Then
                objPedidoCompra.dCotacaoMoeda = objCotacao.dValor
                
            'Verifica se existe cotacao anterior ...
            ElseIf objCotacaoAnterior.dValor <> 0 Then
                objPedidoCompra.dCotacaoMoeda = objCotacaoAnterior.dValor
                
            'Se nao existir => Fator de conversao = 1 para nao alterar o valor
            Else
                objPedidoCompra.dCotacaoMoeda = 1
            End If
    
        End If
        
        'Converte para Real
        'dValorReal = objPedidoCompra.dValorTotal / objPedidoCompra.dCotacaoMoeda 'Pedido de 120 dólares com cotação de 3 reais não são 40 reais e sim 360 reais
        dValorReal = objPedidoCompra.dValorTotal * objPedidoCompra.dCotacaoMoeda
    
    End If
    
    tBloqueioPC.sCodUsuario = String(STRING_USUARIO, 0)
    tBloqueioPC.sResponsavel = String(STRING_BLOQUEIOSPC_RESPONSAVEL, 0)
    tBloqueioPC.sCodUsuarioLib = String(STRING_USUARIO, 0)

    lErro = Comando_ExecutarPos(alComando(0), "SELECT Sequencial,CodUsuario,Responsavel,Data,CodUsuarioLib FROM BloqueiosPC WHERE FilialEmpresa =? AND PedCompras =? AND DataLib = ? AND TipoDeBloqueio = ?", 0, tBloqueioPC.iSequencial, tBloqueioPC.sCodUsuario, tBloqueioPC.sResponsavel, tBloqueioPC.dtData, tBloqueioPC.sCodUsuarioLib, objPedidoCompra.iFilialEmpresa, objPedidoCompra.lCodigo, DATA_NULA, BLOQUEIO_ALCADA)
    If lErro <> SUCESSO Then gError 51217

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 51218

    iExisteBloqAlcada = False

    If lErro = AD_SQL_SUCESSO Then
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 51219

        iExisteBloqAlcada = True
    End If

    If iValorAlterado <> NAO_ALTERADO Or Len(Trim(objPedidoCompra.sAlcada)) = 0 Or gobjCOM.iControleAlcada = 1 Then

        objAlcada.sCodUsuario = gsUsuario

        lErro = CF("Alcada_Le", objAlcada)
        If lErro <> SUCESSO And lErro <> 49208 Then gError 51220
        
        If lErro <> SUCESSO Then gError 51221 'Não encontrou

        'Busca valor liberado do mes em questao
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Liberado" & CStr(Month(objPedidoCompra.dtDataEnvio)) & " FROM ValorPCLiberado WHERE CodUsuario = ? And  Ano = ?", 0, dLiberado, gsUsuario, Year(objPedidoCompra.dtDataEnvio))
        If lErro <> AD_SQL_SUCESSO Then gError 51222

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 51223

        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 51352

        'Se o valor foi convertido => Moeda <> R$, logo o Valor do Pedido = Valor Convertido ...
        If dValorReal <> 0 Then
            dValorTotalPC = dValorReal
        Else
            dValorTotalPC = objPedidoCompra.dValorTotal
        End If
        
        'Verifica se o valor ultrapassa o limite de operacao
        If (objAlcada.dLimiteOperacao <= dValorTotalPC) Or (objAlcada.dLimiteMensal <= dValorTotalPC) Or (objAlcada.dLimiteMensal < (dValorTotalPC + dLiberado)) Then

            If iExisteBloqAlcada = True Then

                lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM BloqueiosPC", alComando(0))
                If lErro <> AD_SQL_SUCESSO Then gError 51224

            End If

            objBloqueiosPC.dtData = gdtDataHoje
            objBloqueiosPC.dtDataEmissao = DATA_NULA
            objBloqueiosPC.dtDataLib = DATA_NULA
            objBloqueiosPC.iTipoBloqueio = BLOQUEIO_ALCADA
            objBloqueiosPC.sCodUsuario = gsUsuario
            objBloqueiosPC.sResponsavel = BLOQUEIO_AUTO_RESP
            objPedidoCompra.dtDataEnvio = DATA_NULA

            objPedidoCompra.colBloqueiosPC.Add objBloqueiosPC

        Else
            objPedidoCompra.sAlcada = gsUsuario

            If iExisteBloqAlcada Then
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE BloqueiosPC SET DataLib = ? , CodUsuarioLib = ?", alComando(0), gdtDataAtual, gsUsuario)
                If lErro <> SUCESSO Then gError 51225

                lErro = Comando_ExecutarPos(alComando(3), "UPDATE ValorPCLiberado SET Liberado" & CStr(Month(objPedidoCompra.dtDataEnvio)) & " = Liberado" & CStr(Month(objPedidoCompra.dtDataEnvio)) & " +  ?", alComando(1), dValorTotalPC)
                If lErro <> AD_SQL_SUCESSO Then gError 51353

            End If
        End If

    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    BloqueioAlcada_Critica = SUCESSO

    Exit Function

Erro_BloqueioAlcada_Critica:

    BloqueioAlcada_Critica = gErr
        
    Select Case gErr

        Case 51215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 51217, 51218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPC", gErr, BLOQUEIO_ALCADA)

        Case 51219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_BLOQUEIOSPC", gErr, objPedidoCompra.lCodigo)

        Case 51220, 108948

        Case 51221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALCADA_NAO_CADASTRADA", gErr, gsUsuario)

        Case 51222, 51223
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VALORPCLIBERADO", gErr)

        Case 51224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BLOQUEIOSPC", gErr)

        Case 51225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_BLOQUEIOSPC", gErr, objPedidoCompra.lCodigo)

        Case 51352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_VALORPCLIBERADO", gErr)

        Case 51353
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_VALORPCLIBERADO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147505)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensPedCompra_Altera(ByVal objPedidoCompra_BD As ClassPedidoCompras, ByVal objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim objItemPC As New ClassItemPedCompra
Dim iIndice As Integer

On Error GoTo Erro_ItensPedCompra_Altera

    lErro = CF("ItensPedCompra_Lock", objPedidoCompra_BD)
    If lErro <> SUCESSO And lErro <> 56332 Then Error 56251

    Set objItemPC = objPedidoCompra_BD.colItens.Item(1)

    'Verifica se NumIntOrigem esta preenchido
    If objItemPC.lNumIntOrigem > 0 Then Error 56335

'    lErro = CF("ItensPedCompra_AlteraBD", objPedidoCompra)
'    If lErro <> SUCESSO Then Error 56252
'
'    lErro = CF("LocalizacaoItemPC_Altera", objPedidoCompra)
'    If lErro <> SUCESSO Then Error 56253

    For Each objItemPC In objPedidoCompra.colItens
        For iIndice = objPedidoCompra_BD.colItens.Count To 1 Step -1
            If objItemPC.lNumIntDoc = objPedidoCompra_BD.colItens(iIndice).lNumIntDoc Then
                objPedidoCompra_BD.colItens.Remove iIndice
                Exit For
            End If
        Next
    Next

    'Exclui de ItensPedCompra os itens do PedidoCompra
    lErro = CF("ItensPedCompra_Exclui", objPedidoCompra_BD)
    If lErro <> SUCESSO Then Error 56254
    
    lErro = CF("ItensPedCompra_AlteraBD", objPedidoCompra)
    If lErro <> SUCESSO Then Error 56252

    lErro = CF("LocalizacaoItemPC_Altera", objPedidoCompra)
    If lErro <> SUCESSO Then Error 56253

    ItensPedCompra_Altera = SUCESSO
    
    Exit Function
    
Erro_ItensPedCompra_Altera:

    ItensPedCompra_Altera = Err
    
    Select Case Err
        
        Case 56335
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_GERADO", Err, objPedidoCompra.lCodigo)
        
        Case 56251 To 56254
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147506)
            
    End Select
    
    Exit Function
    
End Function

Function LocalizacaoItemPC_Altera(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Atualiza a tabela de LocalizacaoItensPC

Dim lErro As Long

On Error GoTo Erro_LocalizacaoItemPC_Altera

    lErro = CF("LocalizacaoItemPC_Exclui", objPedidoCompra)
    If lErro <> SUCESSO Then Error 56326

    lErro = CF("LocalizacaoItemPC_Inclui", objPedidoCompra)
    If lErro <> SUCESSO Then Error 56327

    LocalizacaoItemPC_Altera = SUCESSO

    Exit Function

Erro_LocalizacaoItemPC_Altera:

    LocalizacaoItemPC_Altera = Err

    Select Case Err

        Case 56326, 56327

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147507)

    End Select

    Exit Function

End Function

Function LocalizacaoItemPC_Inclui(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Insere registro na tabela de LocalizacaoItemPC

Dim lErro As Long
Dim lComando As Long
Dim objLocalizacaoItemPC As New ClassLocalizacaoItemPC
Dim objItemPC As New ClassItemPedCompra

On Error GoTo Erro_LocalizacaoItemPC_Inclui

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56327

    For Each objItemPC In objPedidoCompra.colItens

        For Each objLocalizacaoItemPC In objItemPC.colLocalizacao

            'Insere localizacoes de itens na tabela de LocalizacaoItensPC
            lErro = Comando_Executar(lComando, "INSERT INTO LocalizacaoItensPC(ItemPC,Almoxarifado,Ccl,PedCompra,Quantidade,ContaContabil) VALUES(?,?,?,?,?,?)", objItemPC.lNumIntDoc, objLocalizacaoItemPC.iAlmoxarifado, objLocalizacaoItemPC.sCcl, objPedidoCompra.lNumIntDoc, objLocalizacaoItemPC.dQuantidade, objLocalizacaoItemPC.sContaContabil)
            If lErro <> AD_SQL_SUCESSO Then Error 56328

        Next

    Next

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    LocalizacaoItemPC_Inclui = SUCESSO

    Exit Function

Erro_LocalizacaoItemPC_Inclui:

    LocalizacaoItemPC_Inclui = Err

    Select Case Err

        Case 56327
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPEDCOMPRA", Err, objPedidoCompra.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147508)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function LocalizacaoItemPC_Exclui(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Exclui registro da tabela de LocalizacaoItemPC

Dim lErro As Long
Dim lComando(0 To 2) As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_LocalizacaoItemPC_Exclui

    'Abrir comando
    For iIndice = LBound(lComando) To UBound(lComando)

        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 56324

    Next

    lErro = Comando_ExecutarPos(lComando(0), "SELECT PedCompra FROM LocalizacaoItensPC WHERE PedCompra = ?", 0, lCodigo, objPedidoCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 56365

    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56366

     Do While lErro = AD_SQL_SUCESSO

        lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM LocalizacaoItensPC", lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 56325

        lErro = Comando_BuscarProximo(lComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56367

    Loop

    'Fecha o comando
    For iIndice = LBound(lComando) To UBound(lComando)

        Call Comando_Fechar(lComando(iIndice))
    Next

    LocalizacaoItemPC_Exclui = SUCESSO

    Exit Function

Erro_LocalizacaoItemPC_Exclui:

    LocalizacaoItemPC_Exclui = Err

    Select Case Err

        Case 56324
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPEDCOMPRA", Err, objPedidoCompra.lNumIntDoc)

        Case 56365, 56366, 56367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOCALIZACAOITENSPC", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147509)

    End Select

    For iIndice = UBound(lComando) To LBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next


    Exit Function

End Function

Function BloqueiosPC_Grava(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Grava o BloqueioPC

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim objBloqueiosPC As ClassBloqueioPC
Dim iSequencial As Integer

On Error GoTo Erro_BloqueiosPC_Grava

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56313

     For Each objBloqueiosPC In objPedidoCompra.colBloqueiosPC

        'Gera o sequencial automático para o Bloqueio
        lErro = CF("BloqueioPC_Automatico", objPedidoCompra.lCodigo, iSequencial)
        If lErro <> SUCESSO Then Error 56412

        objBloqueiosPC.iSequencial = iSequencial

        lErro = Comando_Executar(lComando, "INSERT INTO BloqueiosPC (FilialEmpresa,PedCompras,Sequencial,TipoDeBloqueio,CodUsuario,Responsavel,Data,CodUsuarioLib,DataLib) VALUES(?,?,?,?,?,?,?,?,?)", objPedidoCompra.iFilialEmpresa, objPedidoCompra.lCodigo, objBloqueiosPC.iSequencial, objBloqueiosPC.iTipoBloqueio, objBloqueiosPC.sCodUsuario, objBloqueiosPC.sResponsavel, objBloqueiosPC.dtData, objBloqueiosPC.sCodUsuarioLib, objBloqueiosPC.dtDataLib)
        If lErro <> AD_SQL_SUCESSO Then Error 56321

     Next

    BloqueiosPC_Grava = SUCESSO

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    Exit Function

Erro_BloqueiosPC_Grava:

    BloqueiosPC_Grava = Err

    Select Case Err

        Case 56313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BLOQUEIOPC", Err, objPedidoCompra.lCodigo)

        Case 56412

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147510)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCompra_Baixar(ByVal objPedidoCompra As ClassPedidoCompras) As Long

Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_PedidoCompra_Baixar

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 68395
    Next

    lErro = CF("PedidoCompra_Baixar1", objPedidoCompra, alComando())
    If lErro <> SUCESSO Then gError 68396
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoCompra_Baixar = SUCESSO

    Exit Function

Erro_PedidoCompra_Baixar:

    PedidoCompra_Baixar = gErr

    Select Case gErr

        Case 68395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 68396
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147511)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function BloqueioPC_Automatico(ByVal lCodigo As Long, iProxSeqBloqueio As Integer) As Long
'retorna o número da próximo Sequencial de Bloqueio disponivel
'OBS: ESTA FUNÇÃO DEVE SER CHAMADA DE DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim iFim_de_Arquivo As Integer
Dim iSequencial_Automatico As Integer
Dim iConteudo As Integer
Dim iConteudo2 As Integer

On Error GoTo Erro_BloqueioPC_Automatico

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56406

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 56407

    'Pesquisa o número do Bloqueio gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT ProxSeqBloqueio FROM PedidoCompra WHERE Codigo = ? AND FilialEmpresa = ?", 0, iConteudo, lCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 56408

    'Lê o registro com o número do Sequencial automático
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 56409

    'Loca o registro que guarda o numero do Sequencial gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 56410

    iSequencial_Automatico = iConteudo

    'Valor de retorno da funcao
    iProxSeqBloqueio = iSequencial_Automatico

    iSequencial_Automatico = iSequencial_Automatico + 1

    iConteudo2 = iSequencial_Automatico

    'Atualiza o numero do Sequencial gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE PedidoCompra SET ProxSeqBloqueio = ?", lComando, iConteudo2)
    If lErro <> AD_SQL_SUCESSO Then Error 56411

    'Fecha comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    BloqueioPC_Automatico = SUCESSO

    Exit Function

Erro_BloqueioPC_Automatico:

    BloqueioPC_Automatico = Err

    Select Case Err

    Case 56406, 56407
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

    Case 56408, 56409
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", Err)

    Case 56410
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOCOMPRA", Err, lCodigo)

    Case 56411
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOCOMPRA", Err, lCodigo)

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147512)

    End Select

    'Fecha comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function PedidoCompra_Exclui(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Exclui o Pedido de Compra passado como parametro

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim lTotal As Long
Dim iIndice As Integer
Dim lCodigo As Long
Dim objPedidoCompra_BD As New ClassPedidoCompras
Dim tPedCompra As typePedidoCompras
Dim lComando(0 To 2) As Long

On Error GoTo Erro_PedidoCompra_Exclui

    'Abre os comandos
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then gError 63621
    Next
       
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 56145

    objPedidoCompra_BD.lCodigo = objPedidoCompra.lCodigo
    objPedidoCompra_BD.iFilialEmpresa = objPedidoCompra.iFilialEmpresa

    'Le e locka o Pedido de Compra
    '** Muda nome para PedidoCompra_LeLock. Não passa alComando.
    'Ok
    lErro = CF("PedidoCompra_LeLock", objPedidoCompra_BD)
    If lErro <> SUCESSO And lErro <> 56259 Then gError 56146

    'Se nao encontrou => gerro
    If lErro = 56259 Then gError 56147

    'Verifica se data de envio foi preenchida
    If objPedidoCompra_BD.dtDataEnvio <> DATA_NULA Then gError 56148

    'Le e locka os itens do Pedido de Compra
    lErro = CF("ItensPedCompra_Lock", objPedidoCompra)
    If lErro <> SUCESSO And lErro <> 56332 Then gError 56149
    If lErro = 56332 Then gError 56334

    'Exclui BloqueiosPC associados a este Pedido de Compra
    lErro = CF("BloqueiosPC_Exclui", objPedidoCompra)
    If lErro <> SUCESSO Then gError 56150

    'Exclui as Notas associadas a este Pedido de Compra
    lErro = CF("NotasPedCompra_Exclui", objPedidoCompra)
    If lErro <> SUCESSO Then gError 103348

    'exclui localizacoes relacionadas aos itens do Pedido de Compra
    lErro = CF("LocalizacaoItemPC_Exclui", objPedidoCompra)
    If lErro <> SUCESSO Then gError 56151

    'Exclui itens do Pedido de Compra
    lErro = CF("ItensPedCompra_Exclui", objPedidoCompra)
    If lErro <> SUCESSO Then gError 56152
        
    tPedCompra.sContato = String(STRING_CONTATO, 0)
    tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
    tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
    tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)

    'Le o pedido de compra passado como parametro
    lErro = Comando_ExecutarPos(lComando(0), "SELECT NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,Observacao FROM PedidoCompra WHERE Codigo = ? AND FilialEmpresa=?", 0, tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, _
        tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.lObservacao, objPedidoCompra_BD.lCodigo, objPedidoCompra_BD.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 56257

    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 56258
    
    'Exclui Pedido de Compra da tabela de PedidoCompra
    lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM PedidoCompra", lComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 56264

    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 56154

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoCompra_Exclui = SUCESSO

    Exit Function

Erro_PedidoCompra_Exclui:

    PedidoCompra_Exclui = gErr

    Select Case gErr

        Case 56144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 56145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 56147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case 56148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_ENVIADO", gErr, objPedidoCompra.lCodigo)

        Case 56146, 56149, 56150, 56151, 56152, 103348

        Case 56154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 56264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PEDIDOCOMPRA", gErr, objPedidoCompra.lCodigo)

        Case 56334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", gErr, objPedidoCompra.lNumIntDoc)

        Case 63621
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147513)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function BloqueiosPC_Exclui(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Exclui todos os bloqueios ligados ao Pedido de Compra passado como parametro

Dim lErro As Long
Dim lComando(0 To 2) As Long
Dim iFilialEmpresa As Integer
Dim lPedCompras As Long
Dim iIndice As Integer
Dim iSequencial As Integer
Dim lTransacao As Long

On Error GoTo Erro_BloqueiosPC_Exclui

    'Abre o comando
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 56261
    Next

    'Busca BloqueioPC ligado ao Pedido de Compra de objPedidoCompra
    lErro = Comando_ExecutarPos(lComando(0), "SELECT FilialEmpresa,PedCompras,Sequencial FROM BloqueiosPC WHERE FilialEmpresa = ? AND PedCompras=? ", 0, iFilialEmpresa, lPedCompras, iSequencial, objPedidoCompra.iFilialEmpresa, objPedidoCompra.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 56262

    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56263

    Do While lErro = AD_SQL_SUCESSO

        'Exclui BloqueioPC ligado ao Pedido de Compra
        lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM BloqueiosPC", lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 56346

        lErro = Comando_BuscarProximo(lComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56347

    Loop

    'Fecha o comando
    For iIndice = LBound(lComando) To UBound(lComando)
         Call Comando_Fechar(lComando(iIndice))
    Next

    BloqueiosPC_Exclui = SUCESSO

    Exit Function

Erro_BloqueiosPC_Exclui:

    BloqueiosPC_Exclui = Err

    Select Case Err

        Case 56261
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56262, 56263, 56347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPC1", Err, objPedidoCompra.lCodigo)

        Case 56346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BLOQUEIOSPC", Err, objPedidoCompra.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147514)

    End Select

    For iIndice = LBound(lComando) To UBound(lComando)
         Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoCompra_Grava(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Grava Pedido de Compra no BD

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_PedidoCompra_Grava

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 56082

    lErro = CF("PedidoCompra_Grava_EmTrans", objPedidoCompra)
    If lErro <> AD_SQL_SUCESSO Then gError 56085
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 56255

    lErro = CF("WorkFlow_Trata_Transacao", "COM", objPedidoCompra, gsUsuario, objPedidoCompra, 1002)
    If lErro <> SUCESSO Then gError 56085

    PedidoCompra_Grava = SUCESSO

    Exit Function

Erro_PedidoCompra_Grava:

    PedidoCompra_Grava = gErr

    Select Case gErr

        Case 56082
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 56085

        Case 56255
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147515)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function ItensPedCompra_AlteraBD(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Atualiza Item de Pedido de Compra na tabela de ItensPedCompra

Dim alComando(0 To 4) As Long
Dim lErro As Long
Dim objItemPC As ClassItemPedCompra
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim sProduto As String
Dim lCodigo As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim dQuant As Double
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_ItensPedCompra_AlteraBD

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 56375
    Next

    objFornecedorProdutoFF.lFornecedor = objPedidoCompra.lFornecedor
    objFornecedorProdutoFF.iFilialForn = objPedidoCompra.iFilial
    objFornecedorProdutoFF.iFilialEmpresa = objPedidoCompra.iFilialEmpresa

    'Para cada item da coleção
    For Each objItemPC In objPedidoCompra.colItens

        objFornecedorProdutoFF.sProduto = objItemPC.sProduto

        'Procura FornecedorFilialProduto relacionado ao ItemPC
        lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
        If lErro <> SUCESSO And lErro <> 54217 Then gError 70345

        'Se não encontrou, erro
        If lErro = 54217 Then gError 70346

        'Se for uma inclusão
        If objItemPC.lNumIntDoc = 0 Then

            'Se o Lote Minimo de FornecedorProdutoFF estiver Preenchido, Verifica se a quantidade a comprar do item é maior do que o lote minimo
            If objFornecedorProdutoFF.dLoteMinimo > 0 And objFornecedorProdutoFF.dLoteMinimo > objItemPC.dQuantidade Then
                vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_LOTEMINIMO_MAIOR_QUANTCOMPRAR", objItemPC.sProduto, objItemPC.dQuantidade, objFornecedorProdutoFF.dLoteMinimo)
                If vbMsgBox = vbNo Then gError 70347
            End If

            'Obtém o número interno automático do item
            lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROX_ITEM_PEDCOMPRA", lNumIntDoc)
            If lErro <> SUCESSO Then gError 56376

            objItemPC.lNumIntDoc = lNumIntDoc
            'Insere o item no BD
            lErro = Comando_Executar(alComando(0), "INSERT INTO ItensPedCompraN (NumIntDoc,PedCompra,DataLimite,Produto,DescProduto,Quantidade,QuantRecebida,QuantRecebimento,UM,PrecoUnitario,ValorDesconto,TipoOrigem,NumIntOrigem,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa,Status,Observacao,ValorIPI,AliquotaIPI,AliquotaICMS, StatusBaixa,DeliveryDate,TempoTransito) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDoc, objPedidoCompra.lNumIntDoc, objItemPC.dtDataLimite, objItemPC.sProduto, objItemPC.sDescProduto, objItemPC.dQuantidade, objItemPC.dQuantRecebida, objItemPC.dQuantRecebimento, objItemPC.sUM, objItemPC.dPrecoUnitario, objItemPC.dValorDesconto, objItemPC.iTipoOrigem, objItemPC.lNumIntOrigem, objItemPC.dPercentMaisReceb, objItemPC.dPercentMenosReceb, objItemPC.iRebebForaFaixa, objItemPC.iStatus, objItemPC.lObservacao, objItemPC.dValorIPI, objItemPC.dAliquotaIPI, objItemPC.dAliquotaICMS, ITENSPEDCOMPRA_STATUS_ATIVO, objItemPC.dtDeliveryDate, objItemPC.iTempoTransito)
            If lErro <> AD_SQL_SUCESSO Then gError 56377

            '#############################################
            'Inserido por Wagner 17/07/2006
            lErro = CF("PCItens_Grava_Customizado", objItemPC)
            If lErro <> SUCESSO Then gError 181227
            '#############################################

        'Se for uma alteração
        Else

            'Localiza o item no Banco de Dados
            lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc,PedCompra,Quantidade FROM ItensPedCompra WHERE NumIntDoc = ? ", 0, lCodigo, lNumIntDoc, dQuant, objItemPC.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 56392

            If dQuant <> objItemPC.dQuantidade Then
                'Se o Lote Minimo de FornecedorProdutoFF estiver Preenchido, Verifica se a quantidade a comprar do item é maior do que o lote minimo
                If objFornecedorProdutoFF.dLoteMinimo > 0 And objFornecedorProdutoFF.dLoteMinimo > objItemPC.dQuantidade Then
                    vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_LOTEMINIMO_MAIOR_QUANTCOMPRAR", objItemPC.sProduto, objItemPC.dQuantidade, objFornecedorProdutoFF.dLoteMinimo)
                    If vbMsgBox = vbNo Then gError 70347
                End If
            End If

            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56393
            If lErro <> AD_SQL_SUCESSO Then gError 56443

            'Faz a atualização no Banco de Dados
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensPedCompra SET DataLimite=?,Produto =?,DescProduto=?, Quantidade = ?, QuantRecebida = ?, QuantRecebimento = ?,UM=?,PrecoUnitario=?,ValorDesconto=?,PercentMaisReceb=?,PercentMenosReceb=?,RecebForaFaixa=?,Status=?,Observacao=?,ValorIPI=?,AliquotaIPI=?,AliquotaICMS=?,DeliveryDate=?,TempoTransito=?", alComando(0), objItemPC.dtDataLimite, objItemPC.sProduto, objItemPC.sDescProduto, objItemPC.dQuantidade, objItemPC.dQuantRecebida, objItemPC.dQuantRecebimento, objItemPC.sUM, objItemPC.dPrecoUnitario, objItemPC.dValorDesconto, objItemPC.dPercentMaisReceb, objItemPC.dPercentMenosReceb, objItemPC.iRebebForaFaixa, objItemPC.iStatus, objItemPC.lObservacao, objItemPC.dValorIPI, objItemPC.dAliquotaIPI, objItemPC.dAliquotaICMS, objItemPC.dtDeliveryDate, objItemPC.iTempoTransito)
            If lErro <> AD_SQL_SUCESSO Then gError 56449

            '#############################################
            'Inserido por Wagner 17/07/2006
            lErro = CF("PCItens_Grava_Customizado", objItemPC)
            If lErro <> SUCESSO Then gError 181227
            '#############################################

        End If

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensPedCompra_AlteraBD = SUCESSO

    Exit Function

Erro_ItensPedCompra_AlteraBD:

    ItensPedCompra_AlteraBD = gErr

    Select Case gErr

        Case 56376, 70345, 70347, 181227

        Case 56443
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", gErr, objPedidoCompra.lNumIntDoc)

        Case 56375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 56377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPEDCOMPRA", gErr)

        Case 56449
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDCOMPRA", gErr, objItemPC.lNumIntDoc)

        Case 56292, 56393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr, objPedidoCompra.lNumIntDoc)

        Case 70346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147516)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensPedCompra_Exclui(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Exclui de ItensPedCompra os itens do PedidoCompra passados como parametro

Dim lErro As Long
Dim lComando(0 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objItemPC As New ClassItemPedCompra

On Error GoTo Erro_ItensPedCompra_Exclui

    'Abre o comando
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 56322
    Next

    For Each objItemPC In objPedidoCompra.colItens

        lErro = Comando_ExecutarPos(lComando(0), "SELECT NumIntDoc FROM ItensPedCompra WHERE NumIntdoc =?", 0, lNumIntDoc, objItemPC.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 56368

        lErro = Comando_BuscarPrimeiro(lComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56369
        If lErro <> AD_SQL_SUCESSO Then Error 51339

        'Exclui o Pedido de Compra em ItensPedCompra
        lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM ItensPedCompra", lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 56323

    Next

    'Fecha o comando
    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next

    ItensPedCompra_Exclui = SUCESSO

    Exit Function

Erro_ItensPedCompra_Exclui:

    ItensPedCompra_Exclui = Err

    Select Case Err

        Case 56322
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPEDCOMPRA", Err, objPedidoCompra.lNumIntDoc)

        Case 56368, 56369, 56370
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", Err, objPedidoCompra.lNumIntDoc)

        Case 51339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_NAO_CADASTRADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147517)

    End Select

    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function BloqueiosPC_Insere(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Insere o BloqueioPC

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim objBloqueiosPC As ClassBloqueioPC
Dim iSequencial As Integer
Dim iTipoDeBloqueio As Integer
Dim colTipoDeBloqueioPC As New Collection

On Error GoTo Erro_BloqueiosPC_Insere

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 56313

    
    For Each objBloqueiosPC In objPedidoCompra.colBloqueiosPC
    
        lErro = Comando_Executar(lComando, "SELECT TipoDeBloqueio FROM BloqueiosPC WHERE PedCompras = ? AND FilialEmpresa = ? AND TipoDeBloqueio = ? AND DataLib = ?", iTipoDeBloqueio, objPedidoCompra.lCodigo, objPedidoCompra.iFilialEmpresa, objBloqueiosPC.iTipoBloqueio, DATA_NULA)
        If lErro <> AD_SQL_SUCESSO Then gError 84613

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84614
        
        If lErro = AD_SQL_SUCESSO Then gError 84615
        
    Next
    
    For Each objBloqueiosPC In objPedidoCompra.colBloqueiosPC

        'Gera o sequencial automático para o Bloqueio
        lErro = CF("BloqueioPC_Automatico", objPedidoCompra.lCodigo, iSequencial)
        If lErro <> SUCESSO Then gError 56412

        objBloqueiosPC.iSequencial = iSequencial
        
        lErro = Comando_Executar(lComando, "INSERT INTO BloqueiosPC (FilialEmpresa,PedCompras,Sequencial,TipoDeBloqueio,CodUsuario,Responsavel,Data,CodUsuarioLib,DataLib) VALUES(?,?,?,?,?,?,?,?,?)", objPedidoCompra.iFilialEmpresa, objPedidoCompra.lCodigo, objBloqueiosPC.iSequencial, objBloqueiosPC.iTipoBloqueio, objBloqueiosPC.sCodUsuario, objBloqueiosPC.sResponsavel, objBloqueiosPC.dtData, objBloqueiosPC.sCodUsuarioLib, objBloqueiosPC.dtDataLib)
        If lErro <> AD_SQL_SUCESSO Then gError 56321

     Next

    BloqueiosPC_Insere = SUCESSO

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    Exit Function

Erro_BloqueiosPC_Insere:

    BloqueiosPC_Insere = gErr

    Select Case gErr

        Case 56313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 56321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_BLOQUEIOPC", gErr, objPedidoCompra.lCodigo)

        Case 84613, 84614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPC", gErr)
        
        Case 84615
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_EXISTE", gErr, objBloqueiosPC.iTipoBloqueio)
        
        Case 56412

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147518)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensPedCompra_Insere(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Insere itens do PedidoCompra na tabela ItensPedCompra

Dim lErro As Long
Dim objItemPC As ClassItemPedCompra
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim lComando As Long

On Error GoTo Erro_ItensPedCompra_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 56309

    'Para cada Item do PedidoCompra
    For iIndice = 1 To objPedidoCompra.colItens.Count

        Set objItemPC = objPedidoCompra.colItens(iIndice)

        'Gera numero interno automatico
        lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROX_ITEM_PEDCOMPRA", lNumIntDoc)
        If lErro <> SUCESSO Then gError 56312

        objItemPC.lNumIntDoc = lNumIntDoc

        'Insere no BD em ItensPedCompra
        lErro = Comando_Executar(lComando, "INSERT INTO ItensPedCompraN (NumIntDoc, PedCompra,DataLimite,Produto,DescProduto, Quantidade, QuantRecebida,QuantRecebimento,UM,PrecoUnitario, ValorDesconto,TipoOrigem,NumIntOrigem,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa,Status,Observacao, ValorIPI,AliquotaIPI,AliquotaICMS, StatusBaixa,DeliveryDate,TempoTransito) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objItemPC.lNumIntDoc, objPedidoCompra.lNumIntDoc, objItemPC.dtDataLimite, objItemPC.sProduto, objItemPC.sDescProduto, objItemPC.dQuantidade, objItemPC.dQuantRecebida, objItemPC.dQuantRecebimento, objItemPC.sUM, objItemPC.dPrecoUnitario, objItemPC.dValorDesconto, objItemPC.iTipoOrigem, objItemPC.lNumIntOrigem, objItemPC.dPercentMaisReceb, objItemPC.dPercentMenosReceb, objItemPC.iRebebForaFaixa, objItemPC.iStatus, objItemPC.lObservacao, objItemPC.dValorIPI, objItemPC.dAliquotaIPI, objItemPC.dAliquotaICMS, ITENSPEDCOMPRA_STATUS_ATIVO, objItemPC.dtDeliveryDate, objItemPC.iTempoTransito)
        If lErro <> AD_SQL_SUCESSO Then gError 56310

        '#############################################
        'Inserido por Wagner 17/07/2006
        lErro = CF("PCItens_Grava_Customizado", objItemPC)
        If lErro <> SUCESSO Then gError 181226
        '#############################################
    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensPedCompra_Insere = SUCESSO

    Exit Function

Erro_ItensPedCompra_Insere:

    ItensPedCompra_Insere = gErr

    Select Case gErr

        Case 56309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 56310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSPEDCOMPRA", Err, objItemPC.lNumIntDoc)

        Case 56312, 181226
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147519)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function LocalizacaoItensPC_Insere(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Insere dados de distribuicao na tabela de LocalizacaoItensPC

Dim lErro As Long
Dim lComando As Long
Dim objLocalizacao As New ClassLocalizacaoItemPC
Dim objItemPC As New ClassItemPedCompra
Dim iIndice As Integer
Dim iItem As Integer

On Error GoTo Erro_LocalizacaoItensPC_Insere

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 56311

    For Each objItemPC In objPedidoCompra.colItens

        For Each objLocalizacao In objItemPC.colLocalizacao

            lErro = Comando_Executar(lComando, "INSERT INTO LocalizacaoItensPC (ItemPC,Almoxarifado,Ccl,PedCompra,Quantidade,ContaContabil) VALUES(?,?,?,?,?,?)", objItemPC.lNumIntDoc, objLocalizacao.iAlmoxarifado, objLocalizacao.sCcl, objPedidoCompra.lNumIntDoc, objLocalizacao.dQuantidade, objLocalizacao.sContaContabil)
            If lErro <> AD_SQL_SUCESSO Then Error 56320

        Next
    Next

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    LocalizacaoItensPC_Insere = SUCESSO

    Exit Function

Erro_LocalizacaoItensPC_Insere:

    LocalizacaoItensPC_Insere = Err

    Select Case Err

        Case 56311
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 56320
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOCALIZACAOITENSPC", Err, objPedidoCompra.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147520)
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCompra_Insere(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Insere PedidoCompra na tabela de PedidoCompra

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_PedidoCompra_Insere

    'Abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then Error 63618
    
    'Gera NumIntDoc para o Pedido de Compra
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_PC", lNumIntDoc)
    If lErro <> SUCESSO Then Error 56202

    objPedidoCompra.lNumIntDoc = lNumIntDoc
'leo

    If Len(Trim(objPedidoCompra.sTipoFrete)) = 0 Then objPedidoCompra.sTipoFrete = 0
    
    objPedidoCompra.sUsuReg = gsUsuario
    objPedidoCompra.bInclusao = True

    If objPedidoCompra.bEnvio Then
        lErro = Comando_Executar(lComando, "INSERT INTO PedidoCompraN(CotacaoMoeda, Embalagem, Taxa, Moeda, " & _
        "NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino," & _
        "FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete," & _
        "ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio," & _
        "TipoBaixa,MotivoBaixa,Alcada,ObsEmbalagem,Observacao,DataBaixa,ValorProdutos, Status, CodigoPV, DataRefFluxo, TabelaPreco, DataReg, HoraReg, UsuReg, DataRegEnvio, HoraRegEnvio, UsuRegEnvio) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objPedidoCompra.dCotacaoMoeda, objPedidoCompra.iEmbalagem, objPedidoCompra.dTaxa, objPedidoCompra.iMoeda, _
        objPedidoCompra.lNumIntDoc, objPedidoCompra.iFilialEmpresa, objPedidoCompra.lCodigo, _
        objPedidoCompra.lFornecedor, objPedidoCompra.iFilial, objPedidoCompra.iComprador, _
        objPedidoCompra.sContato, objPedidoCompra.iTipoDestino, objPedidoCompra.lFornCliDestino, _
        objPedidoCompra.iFilialDestino, objPedidoCompra.dtData, objPedidoCompra.dtDataEmissao, _
        objPedidoCompra.dtDataEnvio, objPedidoCompra.dtDataAlteracao, objPedidoCompra.iCondicaoPagto, _
        objPedidoCompra.dOutrasDespesas, objPedidoCompra.dValorFrete, objPedidoCompra.dValorSeguro, _
        objPedidoCompra.dValorDesconto, objPedidoCompra.dValorTotal, objPedidoCompra.dValorIPI, _
        objPedidoCompra.sTipoFrete, objPedidoCompra.iTransportadora, objPedidoCompra.iProxSeqBloqueio, _
        objPedidoCompra.iTipoBaixa, objPedidoCompra.sMotivoBaixa, objPedidoCompra.sAlcada, _
        objPedidoCompra.sObsEmbalagem, objPedidoCompra.lObservacao, objPedidoCompra.dtDataBaixa, objPedidoCompra.dValorProdutos, _
        PEDIDOCOMPRA_STATUS_ATIVO, objPedidoCompra.lCodigoPV, objPedidoCompra.dtDataRefFluxo, objPedidoCompra.iTabelaPreco, Date, CDbl(Time), gsUsuario, Date, CDbl(Time), gsUsuario)
    Else
        lErro = Comando_Executar(lComando, "INSERT INTO PedidoCompraN(CotacaoMoeda, Embalagem, Taxa, Moeda, " & _
        "NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino," & _
        "FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete," & _
        "ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio," & _
        "TipoBaixa,MotivoBaixa,Alcada,ObsEmbalagem,Observacao,DataBaixa,ValorProdutos, Status, CodigoPV, DataRefFluxo, TabelaPreco, DataReg, HoraReg, UsuReg) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        objPedidoCompra.dCotacaoMoeda, objPedidoCompra.iEmbalagem, objPedidoCompra.dTaxa, objPedidoCompra.iMoeda, _
        objPedidoCompra.lNumIntDoc, objPedidoCompra.iFilialEmpresa, objPedidoCompra.lCodigo, _
        objPedidoCompra.lFornecedor, objPedidoCompra.iFilial, objPedidoCompra.iComprador, _
        objPedidoCompra.sContato, objPedidoCompra.iTipoDestino, objPedidoCompra.lFornCliDestino, _
        objPedidoCompra.iFilialDestino, objPedidoCompra.dtData, objPedidoCompra.dtDataEmissao, _
        objPedidoCompra.dtDataEnvio, objPedidoCompra.dtDataAlteracao, objPedidoCompra.iCondicaoPagto, _
        objPedidoCompra.dOutrasDespesas, objPedidoCompra.dValorFrete, objPedidoCompra.dValorSeguro, _
        objPedidoCompra.dValorDesconto, objPedidoCompra.dValorTotal, objPedidoCompra.dValorIPI, _
        objPedidoCompra.sTipoFrete, objPedidoCompra.iTransportadora, objPedidoCompra.iProxSeqBloqueio, _
        objPedidoCompra.iTipoBaixa, objPedidoCompra.sMotivoBaixa, objPedidoCompra.sAlcada, _
        objPedidoCompra.sObsEmbalagem, objPedidoCompra.lObservacao, objPedidoCompra.dtDataBaixa, objPedidoCompra.dValorProdutos, _
        PEDIDOCOMPRA_STATUS_ATIVO, objPedidoCompra.lCodigoPV, objPedidoCompra.dtDataRefFluxo, objPedidoCompra.iTabelaPreco, Date, CDbl(Time), gsUsuario)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 56308

    'Insere itens de pedido de compra na tabela ItensPedCompra
    lErro = CF("ItensPedCompra_Insere", objPedidoCompra)
    If lErro <> SUCESSO Then Error 56203

    'insere dados de distribuicao em LocalizacaoItensPC
    lErro = CF("LocalizacaoItensPC_Insere", objPedidoCompra)
    If lErro <> SUCESSO Then Error 56204

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    PedidoCompra_Insere = SUCESSO

    Exit Function

Erro_PedidoCompra_Insere:

    PedidoCompra_Insere = Err

    Select Case Err

        Case 56308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PEDIDOCOMPRA", Err, objPedidoCompra.lNumIntDoc)

        Case 63618
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 56203, 56204
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147521)
    
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PedidoCompra_Lock_Gravacao(ByVal objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim lFornecedor As Double
Dim iFilial As Integer
Dim lNumInt As Long
Dim iCondicaoPagto As Integer
Dim alComando(0 To 6) As Long
Dim iIndice As Integer
Dim objItemPC As New ClassItemPedCompra
Dim objProduto As New ClassProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objLocalizacao As New ClassLocalizacaoItemPC
Dim iAlmoxarifado As Integer
Dim sCcl As String
Dim sConta As String
Dim lNumObservacao As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodEmb As Integer
Dim iCodMoeda As Integer

On Error GoTo Erro_PedidoCompra_Lock_Gravacao

     'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 56241
    Next

    'Le e locka FilialFornecedor
    lErro = CF("FilialFornecedor_Lock", objPedidoCompra.lFornecedor, objPedidoCompra.iFilial)
    If lErro <> SUCESSO And lErro <> 43050 Then gError 56217
    If lErro = 43050 Then gError 56447

    'Le e Locka Condicao de Pagamento
    lErro = CF("CondicaoPagto_Lock", alComando(1), objPedidoCompra.iCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 16361 Then gError 56444

    'Leo
    
    If objPedidoCompra.iEmbalagem <> 0 Then
    
        lErro = Comando_ExecutarPos(alComando(5), "SELECT Codigo FROM Embalagens WHERE Codigo = ?", 0, iCodEmb, objPedidoCompra.iEmbalagem)
        If lErro <> AD_SQL_SUCESSO Then gError 103381
        
        lErro = Comando_BuscarPrimeiro(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103382
        
        lErro = Comando_LockShared(alComando(5))
        If lErro <> AD_SQL_SUCESSO Then gError 103383
        
    End If
    
    lErro = Comando_ExecutarPos(alComando(6), "SELECT Codigo FROM Moedas WHERE Codigo = ?", 0, iCodMoeda, objPedidoCompra.iMoeda)
    If lErro <> AD_SQL_SUCESSO Then gError 103385
    
    lErro = Comando_BuscarPrimeiro(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103386
    
    lErro = Comando_LockShared(alComando(6))
    If lErro <> AD_SQL_SUCESSO Then gError 103387
           
    'leo
    
    If objPedidoCompra.iTipoDestino = TIPO_DESTINO_FORNECEDOR Then

        lErro = CF("FilialFornecedor_Lock", objPedidoCompra.lFornCliDestino, objPedidoCompra.iFilialDestino)
        If lErro <> SUCESSO And lErro <> 42870 Then gError 56225

        'Se nao encontrou FilialFornecedor
       If lErro = 42870 Then gError 56233

    End If


    'Verifica se observacao esta preenchida
    If Len(Trim(objPedidoCompra.sObservacao)) > 0 Then

        'Busca na tabela de Observacao a observacao do Pedido
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumInt FROM Observacao WHERE Observacao = ?", 0, lNumInt, objPedidoCompra.sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 56234

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 56235

        'Se nao encontrou
        If lErro = AD_SQL_SEM_DADOS Then

            'gera codigo automatico
            lErro = CF("Observacao_Automatico", lNumInt)
            If lErro <> SUCESSO Then gError 56240

            'Insere observacao na tabela de Observacao
            lErro = Comando_Executar(alComando(2), "INSERT INTO Observacao(Observacao,NumInt)VALUES (?,?)", objPedidoCompra.sObservacao, lNumInt)
            If lErro <> AD_SQL_SUCESSO Then gError 56239

            objPedidoCompra.lObservacao = lNumInt

        Else

            lErro = Comando_LockShared(alComando(2))
            If lErro <> AD_SQL_SUCESSO Then gError 56236

            objPedidoCompra.lObservacao = lNumInt

        End If

    End If

    'Lock shared nos Produtos e Almoxarifados
    For Each objItemPC In objPedidoCompra.colItens
        
        objProduto.sCodigo = objItemPC.sProduto
        lErro = CF("Produto_Lock", objProduto)
        If lErro <> SUCESSO Then gError 56300

        If Len(Trim(objItemPC.sObservacao)) > 0 Then

            'Le a observacao na tabela de Observacao
            lErro = Comando_ExecutarPos(alComando(2), "SELECT NumInt FROM Observacao WHERE Observacao = ?", 0, lNumInt, objItemPC.sObservacao)
            If lErro <> AD_SQL_SUCESSO Then gError 56242

            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 56243

            'Se nao encontrou
            If lErro = AD_SQL_SEM_DADOS Then

                'gera codigo automatico
                lErro = CF("Observacao_Automatico", lNumInt)
                If lErro <> SUCESSO Then gError 56244

                'Insere observacao na tabela de Observacao
                lErro = Comando_Executar(alComando(2), "INSERT INTO Observacao(Observacao,NumInt)VALUES (?,?)", objItemPC.sObservacao, lNumInt)
                If lErro <> AD_SQL_SUCESSO Then gError 56245

                objItemPC.lObservacao = lNumInt

            Else

                lErro = Comando_LockShared(alComando(2))
                If lErro <> AD_SQL_SUCESSO Then gError 56246

                objItemPC.lObservacao = lNumInt

            End If

        End If
        
        For Each objLocalizacao In objItemPC.colLocalizacao

            If objLocalizacao.iAlmoxarifado <> 0 Then

                lErro = CF("Almoxarifado_Lock", objLocalizacao.iAlmoxarifado, alComando(3))
                If lErro <> SUCESSO And lErro <> 34552 Then gError 56301
                'Se nao encontrou => gErro
                If lErro = 34552 Then gError 56445

            End If

            If Len(Trim(objLocalizacao.sCcl)) > 0 Then

                sCcl = String(STRING_CCL, 0)

                lErro = Comando_ExecutarPos(alComando(0), "SELECT Ccl FROM Ccl WHERE Ccl = ? ", 0, sCcl, objLocalizacao.sCcl)
                If lErro <> AD_SQL_SUCESSO Then gError 56302

                lErro = Comando_BuscarPrimeiro(alComando(0))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 56303
                If lErro = AD_SQL_SEM_DADOS Then gError 56304

                lErro = Comando_LockShared(alComando(0))
                If lErro <> AD_SQL_SUCESSO Then gError 56305

            End If

            If Len(Trim(objLocalizacao.sContaContabil)) > 0 Then

                lErro = CF("Conta_Analitica_Lock", objLocalizacao.sContaContabil, MODULO_COMPRAS)
                If lErro <> SUCESSO Then gError 56306

            End If

        Next

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoCompra_Lock_Gravacao = SUCESSO

    Exit Function

Erro_PedidoCompra_Lock_Gravacao:

    PedidoCompra_Lock_Gravacao = gErr

    Select Case gErr

        Case 103383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EMBALAGEM", gErr, objPedidoCompra.iEmbalagem)
            
        Case 103387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MOEDAS", gErr)
            
        Case 103381, 103382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr)
            
        Case 103385, 103386
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOEDAS", gErr)
        
        Case 56300, 56301, 56306

        Case 56302, 56303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CCL", gErr, objLocalizacao.sCcl)

        Case 56304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, objLocalizacao.sCcl)

        Case 56305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CCL", gErr, objLocalizacao.sCcl)

        Case 56217, 56218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES", gErr)

        Case 56219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISFORNECEDORES", gErr, lFornecedor)

        Case 56223, 56225

        Case 56231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objPedidoCompra.iFilialDestino)

        Case 56233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objPedidoCompra.lFornCliDestino, objPedidoCompra.iFilialDestino)

        Case 56234, 56235, 56242, 56243
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OBSERVACAO", gErr, lNumInt)

        Case 56236, 56246
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_OBSERVACAO", gErr)

        Case 56239, 56245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OBSERVACAO", gErr, lNumInt)

        Case 56240, 56244
            'Erro tratado na rotina chamada

        Case 56241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 56445
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objLocalizacao.iAlmoxarifado)

        Case 56447
           lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_INEXISTENTE", gErr, objPedidoCompra.iFilialEmpresa, objPedidoCompra.lFornecedor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147522)

    End Select

     For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoCotacao_Atualiza(ByVal objPedidoCotacao As ClassPedidoCotacao) As Long

Dim alComando(0 To 5) As Long
Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim objItemPedCotacao As New ClassItemPedCotacao
Dim objItensCotacao As New ClassItemCotacao
Dim sObservacao As String
Dim lNumInt As Long
Dim lCodigo As Long
Dim tPedidoCotacao As typePedidoCotacao

On Error GoTo Erro_PedidoCotacao_Atualiza

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 53789
    Next

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 53713

    tPedidoCotacao.sContato = String(STRING_CONTATO, 0)

    'Busca em PedidoCotacao o pedido passado
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Contato, NumIntDoc FROM PedidoCotacao WHERE Codigo = ? AND FilialEmpresa=?", 0, tPedidoCotacao.sContato, tPedidoCotacao.lNumIntDoc, objPedidoCotacao.lCodigo, objPedidoCotacao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 53714

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 53715
    
    If lErro = AD_SQL_SEM_DADOS Then Error 53716

    'Passa o NumIntDoc lido para o obj.
    objPedidoCotacao.lNumIntDoc = tPedidoCotacao.lNumIntDoc

    'Lock do PedidoCotacao
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 53717

    'Se a condição de pagamento estiver preenchida, lê e faz lock
    If objPedidoCotacao.iCondPagtoPrazo <> 0 Then

        lErro = CF("CondicaoPagto_Lock", alComando(1), objPedidoCotacao.iCondPagtoPrazo)
        If lErro <> SUCESSO Then Error 53718

    End If

    'Para cada item, verifica se a Observação está preenchida
    For Each objItemPedCotacao In objPedidoCotacao.colItens

        'Verifica se existem itens dentro do objItemPedCotacao.colItensCotacao.
        If objItemPedCotacao.colItensCotacao.Count > 0 Then

            For Each objItensCotacao In objItemPedCotacao.colItensCotacao

                If Len(Trim(objItensCotacao.sObservacao)) > 0 Then
    
                    'Busca na tabela Observacao uma observação que seja igual a informada.
                    lErro = Comando_Executar(alComando(2), "SELECT NumInt from Observacao WHERE Observacao=?", lNumInt, objItensCotacao.sObservacao)
                    If lErro <> AD_SQL_SUCESSO Then Error 53719
    
                    lErro = Comando_BuscarPrimeiro(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 53720
    
                    'Se encontrar a observação, armazena em objItensCotacao.lObservacao o NumInt da observação.
                    If lErro = SUCESSO Then
    
                        objItensCotacao.lObservacao = lNumInt
                        
                    Else
    
                        'Se a Observação não existir no BD, vamos criá-la
                        'Gera um número automático para a observação
                        lErro = CF("Observacao_Automatico", lCodigo)
                        If lErro <> SUCESSO Then Error 53721
    
                        'Insere a observação na tabela de observações
                        lErro = Comando_Executar(alComando(3), "INSERT INTO Observacao (NumInt, Observacao) VALUES (?,?)", lCodigo, objItensCotacao.sObservacao)
                        If lErro <> AD_SQL_SUCESSO Then Error 53722
    
                        'Vamos armazenar em objItensCotacao.lObservacao o código gerado para a nova observação.
                        objItensCotacao.lObservacao = lCodigo
    
                    End If
                    
                End If
            Next
        End If

    Next
            
    'Altera em PedidoCotacao os campos DataValidade, Contato, TipoFrete ,CondicaoPagamento e Status
    lErro = Comando_ExecutarPos(alComando(4), "UPDATE PedidoCotacao SET DataValidade= ? , Contato= ? , TipoFrete = ?, CondPagtoPrazo=?, Status=?", alComando(0), objPedidoCotacao.dtDataValidade, objPedidoCotacao.sContato, objPedidoCotacao.iTipoFrete, objPedidoCotacao.iCondPagtoPrazo, objPedidoCotacao.iStatus)
    If lErro <> AD_SQL_SUCESSO Then Error 53723

    'Altera os dados do item em ItensCotacao
    lErro = CF("ItensCotacao_Atualiza", objPedidoCotacao)
    If lErro <> SUCESSO Then Error 53724

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 53725

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoCotacao_Atualiza = SUCESSO

    Exit Function

Erro_PedidoCotacao_Atualiza:

    PedidoCotacao_Atualiza = Err

    Select Case Err

        Case 53789
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 53713
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 53714, 53715
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAO", Err, objPedidoCotacao.lCodigo)

        Case 53716
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", Err, objPedidoCotacao.lCodigo)

        Case 53717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDO_COTACAO", Err, objPedidoCotacao.lCodigo)

        Case 53718, 53721, 53724 'Erros tratados nas rotinas chamadas.

        Case 53719, 53720
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OBSERVACAO", Err)

        Case 53722
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OBSERVACAO", Err, objItensCotacao.sObservacao)

        Case 53723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOCOTACAO", Err, objPedidoCotacao.lCodigo)

        Case 53725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147523)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoCotacao_Exclui(ByVal objPedidoCotacao As ClassPedidoCotacao) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 25) As Long
Dim lTransacao As Long
Dim sProduto As String
Dim objCotacao As New ClassCotacao
Dim dQuantCotar As Double
Dim lNumIntCotacao As Long
Dim lCotacaoProduto As Long
Dim lItemConcorrencia As Long
Dim lNumIntDocItemPed As Long
Dim lNumIntDocPedCotacao As Long
Dim lNumIntDocItemCotacao As Long, lNumIntItemFormPreco As Long

On Error GoTo Erro_PedidoCotacao_Exclui

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 53745
    Next

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 53746

    sProduto = String(STRING_PRODUTO, 0)

    'Lê no BD o Pedido de Cotação a ser excluído
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PedidoCotacao WHERE Codigo = ? AND FilialEmpresa=?", 0, lNumIntDocPedCotacao, objPedidoCotacao.lCodigo, objPedidoCotacao.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 53747

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53748
    If lErro = AD_SQL_SEM_DADOS Then gError 53749

    'Lock do PedidoCotacao
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 53750

    'Lê os Itens do Pedido de Cotação
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, CotacaoProduto FROM ItemPedCotacao WHERE ItemPedCotacao.PedidoCotacao=?", 0, lNumIntDocItemPed, lCotacaoProduto, lNumIntDocPedCotacao)
    If lErro <> AD_SQL_SUCESSO Then gError 53751

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53752

    Do While lErro = AD_SQL_SUCESSO

        'Lock do ItemPedCotacao
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 53754

        sProduto = String(STRING_PRODUTO, 0)

        'Busca em ItensPedCompra algum vínculo entre um item de pedido de compra e o item de pedido de cotação
        lErro = Comando_Executar(alComando(3), "SELECT Produto FROM ItensPedCompra WHERE TipoOrigem = ? AND NumIntOrigem=?", sProduto, TIPO_ITEMPEDCOTACAO, lNumIntDocItemPed)
        If lErro <> AD_SQL_SUCESSO Then gError 53759

        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53760
        
        If lErro <> AD_SQL_SEM_DADOS Then gError 53761

        'Lê a tabela ItensCotacao
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc FROM ItensCotacao WHERE ItemPedCotacao=?", 0, lNumIntDocItemCotacao, lNumIntDocItemPed)
        If lErro <> AD_SQL_SUCESSO Then gError 53755

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53756

        Do While lErro = AD_SQL_SUCESSO

            'Lock de ItensCotacao
            lErro = Comando_LockExclusive(alComando(2))
            If lErro <> AD_SQL_SUCESSO Then gError 53758

            'Busca em CotacaoItemConcorrencia algum vínculo entre uma concorrência e o item de cotacao associado ao pedido.
            lErro = Comando_Executar(alComando(3), "SELECT ItemConcorrencia FROM CotacaoItemConcorrencia WHERE CotacaoItemConcorrencia.ItemCotacao=?", lItemConcorrencia, lNumIntDocItemCotacao)
            If lErro <> AD_SQL_SUCESSO Then gError 53762

            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53763
            If lErro <> AD_SQL_SEM_DADOS Then gError 53764 'Se tirver, erro.

            'Exclui o item de ItensCotacao
            lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM ItensCotacao", alComando(2))
            If lErro <> AD_SQL_SUCESSO Then gError 53765

            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53766

        Loop

        'Exclui os registros de ItemPedCotacao
        lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM ItemPedCotacao", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 53767

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53768

    Loop

    'Exclui o pedido de cotação
    lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM PedidoCotacao", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 53769
    
    'Verifica se existe outro pedido de cotação associado a cotação do Pedido de Cotacao em questão
    lErro = Testa_Pedido_Cotacao(lCotacaoProduto, objPedidoCotacao, objCotacao)
    If lErro <> SUCESSO And lErro <> 74922 Then gError 74904
        
    'Se nao encontrou outro pedido de cotação associado a cotação do Pedido de Cotacao em questão
    If lErro = 74922 Then
        
        'Verifica se existe algum Pedido de Cotacao Baixado ligado a Cotacao
        lErro = Comando_Executar(alComando(7), "SELECT ItemPedCotacaoBaixado.CotacaoProduto FROM PedidoCotacaoBaixado,ItemPedCotacaoBaixado,CotacaoProduto,Cotacao WHERE Cotacao.NumIntDoc=? AND ItemPedCotacaoBaixado.CotacaoProduto=CotacaoProduto.NumIntDoc AND ItemPedCotacaoBaixado.PedidoCotacao=PedidoCotacaoBaixado.NumIntDoc AND CotacaoProduto.Cotacao=Cotacao.NumIntDoc", lCotacaoProduto, objCotacao.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 74905

        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 74906

        'Se nao encontrou Pedido de Cotacao Baixado ligado a Cotacao ==> Exclui a Cotacao e CotacaoProduto
        If lErro = AD_SQL_SEM_DADOS Then
        
            'Lê a Cotacao
            lErro = Comando_ExecutarPos(alComando(8), "SELECT NumIntDoc FROM Cotacao WHERE NumIntDoc=?", 0, lNumIntCotacao, objCotacao.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 74929

            lErro = Comando_BuscarPrimeiro(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 74930

            'Lê a CotacaoProduto
            lErro = Comando_ExecutarPos(alComando(10), "SELECT NumIntDoc FROM CotacaoProduto WHERE Cotacao=?", 0, lCotacaoProduto, objCotacao.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 74909

            lErro = Comando_BuscarPrimeiro(alComando(10))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 74910

            Do While lErro = AD_SQL_SUCESSO
                
                'Lê os relacionamentos entre a Cotação Produto e Itens de Requisição
                lErro = Comando_ExecutarPos(alComando(21), "SELECT QuantidadeCotar FROM CotacaoProdutoItemRC WHERE CotacaoProduto = ? ", 0, dQuantCotar, lCotacaoProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 86108
                
                lErro = Comando_BuscarPrimeiro(alComando(21))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86109
                
                'Enquanto houver um relacionamento
                Do While lErro = AD_SQL_SUCESSO
                    'Exclui a ligação
                    lErro = Comando_ExecutarPos(alComando(22), "DELETE FROM CotacaoProdutoItemRC", alComando(21))
                    If lErro <> AD_SQL_SUCESSO Then gError 86110
                    
                    'BUsca o próximo
                    lErro = Comando_BuscarProximo(alComando(21))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86111
                Loop
                
                'Lê os relacionamentos entre a Cotação Produto e Itens FormPreco
                lErro = Comando_ExecutarPos(alComando(23), "SELECT NumIntItemFormPreco FROM CotacaoProdutoItemFP WHERE CotacaoProduto = ? ", 0, lNumIntItemFormPreco, lCotacaoProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 86108
                
                lErro = Comando_BuscarPrimeiro(alComando(23))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86109
                
                'Enquanto houver um relacionamento
                Do While lErro = AD_SQL_SUCESSO
                    'Exclui a ligação
                    lErro = Comando_ExecutarPos(alComando(24), "DELETE FROM CotacaoProdutoItemFP", alComando(23))
                    If lErro <> AD_SQL_SUCESSO Then gError 86110
                    
                    lErro = Comando_Executar(alComando(25), "UPDATE ItensFormPreco SET CustoUnit = 0, PrecoUnit = 0, PrecoTotal = 0 WHERE NumIntDoc = ?", lNumIntItemFormPreco)
                    If lErro <> AD_SQL_SUCESSO Then gError 184378
                    
                    'BUsca o próximo
                    lErro = Comando_BuscarProximo(alComando(23))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86111
                Loop
                
                'Exclui as Cotacoes Produto ligadas ao Pedido de Cotacao
                lErro = Comando_ExecutarPos(alComando(11), "DELETE FROM CotacaoProduto", alComando(10))
                If lErro <> AD_SQL_SUCESSO Then gError 74911

                lErro = Comando_BuscarProximo(alComando(10))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 74912

            Loop

            'Exclui a Cotacao ligada ao Pedido de Cotacao
            lErro = Comando_ExecutarPos(alComando(9), "DELETE FROM Cotacao", alComando(8))
            If lErro <> AD_SQL_SUCESSO Then gError 74913
    
        'Se encontrou Pedido de Cotacao Baixado ==> Baixa a Cotacao e CotacaoProduto
        ElseIf lErro = AD_SQL_SUCESSO Then
    
            'Baixa a Cotacao e CotacaoProduto
            lErro = CF("Cotacao_Baixa", objCotacao.lNumIntDoc, alComando())
            If lErro <> SUCESSO Then gError 74914
        
        End If
    
    End If
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 53770

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoCotacao_Exclui = SUCESSO

    Exit Function

Erro_PedidoCotacao_Exclui:

    PedidoCotacao_Exclui = gErr

    Select Case gErr

        Case 53745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 53746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 53747, 53748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOTACAO_NAO_ENCONTRADO", gErr, objPedidoCotacao.lCodigo)

        Case 53750
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDO_COTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53751, 53752, 53768
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDCOTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOTACAO_NAO_ENCONTRADO", gErr, objPedidoCotacao.lCodigo)

        Case 53754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMPEDCOTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53755, 53756, 53766
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCOTACAO", gErr)

        Case 53757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENSCOTACAO_NAO_ENCONTRADOS", gErr, objPedidoCotacao.lCodigo)

        Case 53758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSCOTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53759, 53760
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)

        Case 53761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PEDCOTACAO_VINCULADO_ITEM_PEDCOMPRA", gErr, objPedidoCotacao.lCodigo)

        Case 53762, 53763
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOITEMCONCORRENCIA", gErr)

        Case 53764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PEDCOTACAO_VINCULADO_CONCORRENCIA", gErr, objPedidoCotacao.lCodigo)

        Case 53765
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSCOTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMPEDCOTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PEDIDO_COTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53770
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 74904, 74914, 74931
        
        Case 74905, 74906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAOBAIXADO", gErr)
            
        Case 74909, 74910, 74912
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTO", gErr)
            
        Case 74911
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAOPRODUTO", gErr)
        
        Case 74913
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAO", gErr)
            
        Case 74929, 74930
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAO", gErr)
        
        Case 86108, 86109, 86111
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTOITEMRC", gErr)
        
        Case 86110, 184378
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAOPRODUTOITEMRC1", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147524)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensCotacao_Atualiza(ByVal objPedidoCotacao As ClassPedidoCotacao) As Long

Dim alComando(0 To 10) As Long
Dim lErro As Long
Dim objItemPedCotacao As ClassItemPedCotacao
Dim objItensCotacao As ClassItemCotacao
Dim lNumIntDoc As Long, lNumIntCotacaoProdutoN As Long, dQtdeCotacaoProd As Double
Dim tItensCotacao As typeItemCotacao
Dim tItemPedCotacao As typeItemPedCotacao
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim bAchou As Boolean
Dim lItemConcorrencia As Long
Dim lTransacao As Long
Dim lFornecedor As Long, iFilialForn As Integer
Dim sProduto As String

On Error GoTo Erro_ItensCotacao_Atualiza

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 53728
    Next

    iIndice1 = 0

    For Each objItemPedCotacao In objPedidoCotacao.colItens

        iIndice1 = iIndice1 + 1

        'Busca no BD ItemPedCotacao.NumIntDoc
        If objItemPedCotacao.iExclusivo = MARCADO Then
            lFornecedor = objPedidoCotacao.lFornecedor
            iFilialForn = objPedidoCotacao.iFilial
        Else
            lFornecedor = 0
            iFilialForn = 0
        End If
        
        lErro = Comando_Executar(alComando(0), "SELECT ItemPedCotacao.NumIntDoc, CotacaoProduto.NumIntDoc, CotacaoProduto.Quantidade FROM ItemPedCotacao, CotacaoProduto WHERE CotacaoProduto.Produto=? AND ItemPedCotacao.PedidoCotacao=? and ItemPedCotacao.CotacaoProduto=CotacaoProduto.NumIntDoc AND CotacaoProduto.Fornecedor = ? AND CotacaoProduto.Filial = ?", tItemPedCotacao.lNumIntDoc, lNumIntCotacaoProdutoN, dQtdeCotacaoProd, objItemPedCotacao.sProduto, objPedidoCotacao.lNumIntDoc, lFornecedor, iFilialForn)
        If lErro <> AD_SQL_SUCESSO Then gError 53733

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53734
        
        'se não encontrar a cotaçãoproduto associada ao item pedido de cotação ==> erro
        If lErro = AD_SQL_SEM_DADOS Then gError 53735

        'se a qtde cotada foi alterada
        If Abs(dQtdeCotacaoProd - objItemPedCotacao.dQuantidade) > QTDE_ESTOQUE_DELTA Then
                    
            lErro = Comando_ExecutarPos(alComando(9), "SELECT Quantidade FROM CotacaoProdutoN WHERE NumIntDoc = ?", 0, dQtdeCotacaoProd, lNumIntCotacaoProdutoN)
            If lErro <> AD_SQL_SUCESSO Then gError 106741
    
            lErro = Comando_BuscarPrimeiro(alComando(9))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106742
            If lErro = AD_SQL_SEM_DADOS Then gError 106743
            
            lErro = Comando_LockExclusive(alComando(9))
            If lErro <> AD_SQL_SUCESSO Then gError 106744
            
            lErro = Comando_ExecutarPos(alComando(10), "UPDATE CotacaoProdutoN SET Quantidade = ?", alComando(9), objItemPedCotacao.dQuantidade)
            If lErro <> AD_SQL_SUCESSO Then gError 106745
            
        End If
        
        sProduto = String(STRING_PRODUTO, 0)

        'Busca em ItensPedCompra algum vínculo entre um item de pedido de compra e o item de pedido de cotação
        lErro = Comando_Executar(alComando(7), "SELECT Produto FROM ItensPedCompra WHERE TipoOrigem = ? AND NumIntOrigem=?", sProduto, TIPO_ITEMPEDCOTACAO, tItemPedCotacao.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 83805

        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83806
        
        'se já houver um pedido de compra vinculado ao item pedido de cotação ==> erro
        If lErro = AD_SQL_SUCESSO Then gError 83807

        'Busca em CotacaoItemConcorrencia algum vínculo entre uma concorrência e o item de pedido de cotacao.
        lErro = Comando_Executar(alComando(8), "SELECT ItemConcorrencia FROM CotacaoItemConcorrencia, ItensCotacao WHERE CotacaoItemConcorrencia.ItemCotacao=ItensCotacao.NumIntDoc AND ItensCotacao.ItemPedCotacao=?", lItemConcorrencia, tItemPedCotacao.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 83808

        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83809
        
        'se já houver uma concorrencia vinculada ao item pedido de cotação ==> erro
        If lErro = AD_SQL_SUCESSO Then gError 83810

        'Armazena em objItemPedCotacao.lNumIntDoc  o NumIntDoc lido.
        objItemPedCotacao.lNumIntDoc = tItemPedCotacao.lNumIntDoc

        'Busca os itens de cotação para ItemPedCotacao com o número interno encontrado.
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, CondPagto FROM ItensCotacao WHERE ItensCotacao.ItemPedCotacao=?", 0, tItensCotacao.lNumIntDoc, tItensCotacao.iCondPagto, objItemPedCotacao.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 53736

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53737

        Do While lErro = AD_SQL_SUCESSO

            iIndice = 0

            bAchou = False

            For iIndice = 1 To objItemPedCotacao.colItensCotacao.Count

                Set objItensCotacao = objItemPedCotacao.colItensCotacao(iIndice)

                If (tItensCotacao.iCondPagto = COD_A_VISTA And objItensCotacao.iCondPagto = tItensCotacao.iCondPagto) Or (tItensCotacao.iCondPagto <> COD_A_VISTA And objItensCotacao.iCondPagto <> COD_A_VISTA) Then

                        bAchou = True

                        'Altera o item de cotação com os dados do objItenscotacao.
                        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensCotacao SET CondPagto=?, PrecoUnitario=?, OutrasDespesas=?, ValorFrete=?, ValorSeguro=?, ValorDesconto=?, ValorTotal=?, ValorIPI=?, AliquotaIPI=?, AliquotaICMS=?, PrazoEntrega=?, QuantEntrega=?, Observacao=?, Moeda=?, Taxa=?", alComando(1), objItensCotacao.iCondPagto, objItensCotacao.dPrecoUnitario, objItensCotacao.dOutrasDespesas, objItensCotacao.dValorFrete, objItensCotacao.dValorSeguro, objItensCotacao.dValorDesconto, objItensCotacao.dValorTotal, objItensCotacao.dValorIPI, objItensCotacao.dAliquotaIPI, objItensCotacao.dAliquotaICMS, objItensCotacao.iPrazoEntrega, objItensCotacao.dQuantEntrega, objItensCotacao.lObservacao, objItensCotacao.iMoeda, objItensCotacao.dTaxa)
                        If lErro <> AD_SQL_SUCESSO Then gError 53739

                        'Exclui o item de cotação da coleção de itens do pedido
                        objItemPedCotacao.colItensCotacao.Remove (iIndice)

                        Exit For

                End If

            Next

            If Not bAchou Then

                'Exclui o item de cotação
                lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM ItensCotacao", alComando(1))
                If lErro <> AD_SQL_SUCESSO Then gError 53740

            End If

            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 53741

        Loop

        For Each objItensCotacao In objItemPedCotacao.colItensCotacao

            'Gera um numero interno para o próximo item de cotação
            lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_ITEMCOTACAO", lNumIntDoc)
            If lErro <> SUCESSO Then gError 53742

            'Insere na tabela de itens de cotação.
            lErro = Comando_Executar(alComando(5), "INSERT INTO ItensCotacaoN (NumIntDoc, ItemPedCotacao, CondPagto, DataReferencia, PrecoUnitario, OutrasDespesas, ValorFrete, ValorSeguro, ValorDesconto, ValorTotal, ValorIPI, AliquotaIPI, AliquotaICMS, PrazoEntrega, QuantEntrega, Observacao, Status, Moeda, Taxa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", lNumIntDoc, objItemPedCotacao.lNumIntDoc, objItensCotacao.iCondPagto, objItensCotacao.dtDataReferencia, objItensCotacao.dPrecoUnitario, objItensCotacao.dOutrasDespesas, objItensCotacao.dValorFrete, objItensCotacao.dValorSeguro, objItensCotacao.dValorDesconto, objItensCotacao.dValorTotal, objItensCotacao.dValorIPI, objItensCotacao.dAliquotaIPI, objItensCotacao.dAliquotaICMS, objItensCotacao.iPrazoEntrega, objItensCotacao.dQuantEntrega, objItensCotacao.lObservacao, ITENSCOTACAO_STATUS_ATIVO, objItensCotacao.iMoeda, objItensCotacao.dTaxa)
            If lErro <> AD_SQL_SUCESSO Then gError 53743

        Next

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensCotacao_Atualiza = SUCESSO

    Exit Function

Erro_ItensCotacao_Atualiza:

    ItensCotacao_Atualiza = gErr

    Select Case gErr

        Case 53728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 53732
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 53733, 53734, 106741, 106742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPEDCOTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53735, 106743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOTACAO_NAO_ENCONTRADO", gErr)

        Case 53736, 53737, 53741
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSCOTACAO", gErr)

        Case 53739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSCOTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSCOTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53742 'Erro tratado na rotina chamada.

        Case 53743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSCOTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 53744
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 53729, 53730
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOITEMCONCORRENCIA", gErr)

        Case 53731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOTACAO_VINCULADO_CONCORRENCIA1", gErr, objPedidoCotacao.iCondPagtoPrazo, iIndice1)

        Case 83805, 83806
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)

        Case 83807
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PEDCOTACAO_VINCULADO_ITEM_PEDCOMPRA", gErr, objPedidoCotacao.lCodigo)

        Case 83808, 83809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOITEMCONCORRENCIA", gErr)

        Case 83810
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_PEDCOTACAO_VINCULADO_CONCORRENCIA", gErr, objPedidoCotacao.lCodigo)

        Case 106744
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIO_COTACAOPRODUTO", gErr, objPedidoCotacao.lCodigo)
        
        Case 106745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COTACAOPRODUTO", gErr, objPedidoCotacao.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147525)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ReqCompraEnv_BaixarEmTrans(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Baixa a Requisição passada por objRequisiçãoCompras
'CHAMADA EM TRANSACAO

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 8) As Long
Dim tItemReqCompra As typeItemReqCompra
Dim tRequisicaoCompras As typeRequisicaoCompras
Dim lNumIntDoc As Long
Dim lItemConcorrencia As Long
Dim dQuantidadeCotar As Double
Dim iStatus As Integer

On Error GoTo Erro_ReqCompraEnv_BaixarEmTrans

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 54199
    Next

    tRequisicaoCompras.sCcl = String(STRING_CCL, 0)
    tRequisicaoCompras.sDigitador = String(STRING_USUARIO_CODIGO, 0)
    tRequisicaoCompras.sOPCodigo = String(STRING_OPCODIGO, 0)
    tRequisicaoCompras.sUsuReg = String(STRING_USUARIO_CODIGO, 0)
    tRequisicaoCompras.sUsuRegAprov = String(STRING_USUARIO_CODIGO, 0)
    tRequisicaoCompras.sUsuRegBaixa = String(STRING_USUARIO_CODIGO, 0)
    tRequisicaoCompras.sUsuRegEnvio = String(STRING_USUARIO_CODIGO, 0)

    'Lê em Requisição Compras todos os dados da Requisição vinda em objRequisiçãoCompras
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc, FilialEmpresa, Codigo, Data, DataEnvio, DataLimite, Urgente, Requisitante, Digitador, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao, OPCodigo, Status, UsuReg, UsuRegAprov, UsuRegBaixa, UsuRegEnvio FROM RequisicaoCompraN WHERE FilialEmpresa = ? And Codigo = ?", 0, _
        tRequisicaoCompras.lNumIntDoc, tRequisicaoCompras.iFilialEmpresa, tRequisicaoCompras.lCodigo, tRequisicaoCompras.dtData, tRequisicaoCompras.dtDataEnvio, tRequisicaoCompras.dtDataLimite, tRequisicaoCompras.lUrgente, tRequisicaoCompras.lRequisitante, tRequisicaoCompras.sDigitador, tRequisicaoCompras.sCcl, tRequisicaoCompras.iFilialCompra, tRequisicaoCompras.iTipoDestino, tRequisicaoCompras.lFornCliDestino, tRequisicaoCompras.iFilialDestino, tRequisicaoCompras.lObservacao, tRequisicaoCompras.iTipoTributacao, tRequisicaoCompras.sOPCodigo, iStatus, tRequisicaoCompras.sUsuReg, tRequisicaoCompras.sUsuRegAprov, tRequisicaoCompras.sUsuRegBaixa, tRequisicaoCompras.sUsuRegEnvio, objRequisicaoCompras.iFilialEmpresa, objRequisicaoCompras.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 54208

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 54209

    'Se não encontrou --> Erro
    If lErro <> AD_SQL_SUCESSO Then gError 54231

    'se a requisicao estiver baixada ==> erro
    If iStatus = REQUISICAOCOMPRA_STATUS_BAIXADO Then gError 89484

    'Faz "LockExclusive" em RequisiçãoCompra
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 54232

    tRequisicaoCompras.dtDataBaixa = gdtDataHoje
    
    objRequisicaoCompras.sUsuReg = tRequisicaoCompras.sUsuReg
    objRequisicaoCompras.sUsuRegAprov = tRequisicaoCompras.sUsuRegAprov
    objRequisicaoCompras.sUsuRegBaixa = tRequisicaoCompras.sUsuRegBaixa
    objRequisicaoCompras.sUsuRegEnvio = tRequisicaoCompras.sUsuRegEnvio

    tItemReqCompra.sProduto = String(STRING_PRODUTO, 0)
    tItemReqCompra.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemReqCompra.sCcl = String(STRING_CCL, 0)
    tItemReqCompra.sContaContabil = String(STRING_CONTA, 0)
    tItemReqCompra.sUM = String(STRING_UM_SIGLA, 0)

    objRequisicaoCompras.lNumIntDoc = tRequisicaoCompras.lNumIntDoc

    iIndice = 0

    'Lê o Item da Requisição Compras
    lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc, ReqCompra, Produto, DescProduto, Status, Quantidade, QuantPedida, QuantRecebida, QuantCancelada, UM, Ccl, Almoxarifado, ContaContabil, TipoTributacao, Observacao, Fornecedor, Filial, Exclusivo FROM ItensReqCompraN WHERE ReqCompra = ? AND StatusBaixa = ?", 0, _
        tItemReqCompra.lNumIntDoc, tItemReqCompra.lReqCompra, tItemReqCompra.sProduto, tItemReqCompra.sDescProduto, tItemReqCompra.iStatus, tItemReqCompra.dQuantidade, tItemReqCompra.dQuantPedida, tItemReqCompra.dQuantRecebida, tItemReqCompra.dQuantCancelada, tItemReqCompra.sUM, tItemReqCompra.sCcl, tItemReqCompra.iAlmoxarifado, tItemReqCompra.sContaContabil, tItemReqCompra.iTipoTributacao, tItemReqCompra.lObservacao, tItemReqCompra.lFornecedor, tItemReqCompra.iFilial, tItemReqCompra.iExclusivo, objRequisicaoCompras.lNumIntDoc, ITENSREQCOMPRA_STATUS_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 54233

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 54234

    'Para cada Item de Requisição de Compras
    Do While lErro <> AD_SQL_SEM_DADOS

        iIndice = iIndice + 1

        'Fazer "LockExclusive" em Item Requisição Compras
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 54254

        'Se o Item de Requisição ainda possui vínculo com o Item de um Pedido de Compras, Erro
        If tItemReqCompra.dQuantPedida > QTDE_ESTOQUE_DELTA2 Then gError 54271

        'Pesquisa em ItemRCItemConcorrencia vínculo entre o item da requisição e um item de alguma Concorrencia não baixado
        lErro = Comando_Executar(alComando(8), "SELECT ItemConcorrencia FROM ItemRCItemConcorrencia,ItensConcorrencia WHERE ItemReqCompra = ? AND ItensConcorrencia.NumIntDoc = ItemRCItemConcorrencia.ItemConcorrencia ", lItemConcorrencia, tItemReqCompra.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 54275

        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 54276

        'Se encontrou --> Erro
        If lErro = AD_SQL_SUCESSO Then gError 54310

        'Busca em CotaçãoProdutoItemRC as Cotações para o item
        lErro = Comando_ExecutarPos(alComando(3), "SELECT QuantidadeCotar FROM CotacaoProdutoItemRC WHERE ItemReqCompra = ?", 0, dQuantidadeCotar, tItemReqCompra.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 54327

        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 54328

        'Para cada registro encontrado
        Do While lErro = AD_SQL_SUCESSO

            'Exclui o Registro da tabela CotaçãoProdutoItemRC
            lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM CotacaoProdutoItemRC", alComando(3))
            If lErro <> AD_SQL_SUCESSO Then gError 54330

            'Busca a Próxima CotaçãoProdutoItemRC
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 54329

        Loop
        
        'Atualiza a Quantidade Cancelada para fazer a Baixa
        'OK ** Código não está certo.  = tItemReqCompra.dQuantidade - tItemReqCompra.dQuantRecebida
        tItemReqCompra.dQuantCancelada = tItemReqCompra.dQuantidade - tItemReqCompra.dQuantRecebida
        
        'Exclui o item Requisição de Compra da tabela ItemRequisiçãoCompra
        lErro = Comando_ExecutarPos(alComando(6), "UPDATE ItensReqCompraN SET QuantCancelada = ?, StatusBaixa = ?", alComando(2), tItemReqCompra.dQuantCancelada, ITENSREQCOMPRA_STATUS_BAIXADO)
        If lErro <> AD_SQL_SUCESSO Then gError 54332

        'Busca o próximo Item de Requisição de Compras
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 54402

    Loop
    
    objRequisicaoCompras.sUsuRegBaixa = gsUsuario
    objRequisicaoCompras.bBaixa = True

    'Altera o status da Requisição de Compras
    lErro = Comando_ExecutarPos(alComando(7), "UPDATE RequisicaoCompraN SET Status = ?, DataBaixa = ?, DataRegBaixa = ?, HoraRegBaixa = ?, UsuRegBaixa = ?", alComando(0), REQUISICAOCOMPRA_STATUS_BAIXADO, tRequisicaoCompras.dtDataBaixa, Date, CDbl(Time), gsUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 54410

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ReqCompraEnv_BaixarEmTrans = SUCESSO

    Exit Function

Erro_ReqCompraEnv_BaixarEmTrans:

    ReqCompraEnv_BaixarEmTrans = gErr

    Select Case gErr

        Case 54199
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 54207
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 54208, 54209
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr, objRequisicaoCompras.lCodigo)

        Case 54231
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objRequisicaoCompras.lCodigo)
            
        Case 54232
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISICAOCOMPRA", gErr, objRequisicaoCompras.lCodigo)

        Case 54233, 54234, 54402
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr, objRequisicaoCompras.lCodigo)

        Case 54254
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSREQCOMPRA", gErr, objRequisicaoCompras.lCodigo)

        Case 54267
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REQUISICAOCOMPRABAIXADA", gErr, objRequisicaoCompras.lCodigo)

        Case 54271
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXA_ITEMRC_VINCULADO_ITEMPC_NAO_BAIXADO", gErr, objRequisicaoCompras.lCodigo, iIndice)

        Case 54275, 54276
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMRCITEMCONCORRENCIA", gErr, tItemReqCompra.lNumIntDoc)

        Case 54310
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXA_ITEMRC_VINCULADO_ITEMCONCORRENCIA_NAO_BAIXADO", gErr, objRequisicaoCompras.lCodigo, lItemConcorrencia)

        Case 54327, 54328, 54329
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTOITEMRC", gErr, tItemReqCompra.lNumIntDoc)

        Case 54330
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAOPRODUTOITEMRC", gErr, tItemReqCompra.lNumIntDoc)

        Case 54331
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITENSREQCOMPRABAIXADOS", gErr, tItemReqCompra.lNumIntDoc)

        Case 54332
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSREQCOMPRAS", gErr, objRequisicaoCompras.lCodigo)

        Case 54410
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REQUISICAOCOMPRA", gErr, objRequisicaoCompras.lNumIntDoc)

        Case 89484
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_BAIXADA", gErr, objRequisicaoCompras.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147526)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RequisicaoCompraEnv_Grava(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Grava a Requisição passada pelo objRequisicaoCompras na tabela RequisiçãoCompras

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 7) As Long
Dim lTransacao As Long
Dim dQuantidade As Double
Dim objItemReqCompras As New ClassItemReqCompras
Dim tItemReqCompra As typeItemReqCompra
Dim dQuantMaximaCancelar As Double
Dim dQuantTotalConcorrencia As Double
Dim dQuantCancelada As Double
Dim lNumIntDoc As Long
Dim iTotalItensFechados As Integer
Dim iStatus As Integer

On Error GoTo Erro_RequisicaoCompraEnv_Grava

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then Error 54339
    Next

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 54340

    'Pesquisa no BD a requisição passada por objRequisicaoCompras na tabela RequisiçãoCompra
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RequisicaoCompra WHERE Codigo = ? AND FilialEmpresa = ?", 0, lNumIntDoc, objRequisicaoCompras.lCodigo, objRequisicaoCompras.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 54341

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54342

    'Se não encontrou --> Erro
    If lErro <> AD_SQL_SUCESSO Then Error 54343

    'Faz "LockExclusive" em RequisiçãoCompra
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 54344

    'Para cada Item de Requisição de Compras
    For Each objItemReqCompras In objRequisicaoCompras.colItens

        tItemReqCompra.sProduto = String(STRING_PRODUTO, 0)
        tItemReqCompra.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        tItemReqCompra.sUM = String(STRING_UM_SIGLA, 0)
        tItemReqCompra.sCcl = String(STRING_CCL, 0)
        tItemReqCompra.sContaContabil = String(STRING_CONTA, 0)

        'Lê o Item da Requisição Compras
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, ReqCompra, Produto, DescProduto, Status, Quantidade, QuantPedida, QuantRecebida, QuantCancelada, UM, Ccl, Almoxarifado, ContaContabil, TipoTributacao, Observacao, Fornecedor, Filial, Exclusivo FROM ItensReqCompra WHERE ReqCompra = ? And NumIntDoc = ?", 0, _
            tItemReqCompra.lNumIntDoc, tItemReqCompra.lReqCompra, tItemReqCompra.sProduto, tItemReqCompra.sDescProduto, tItemReqCompra.iStatus, tItemReqCompra.dQuantidade, tItemReqCompra.dQuantPedida, tItemReqCompra.dQuantRecebida, tItemReqCompra.dQuantCancelada, tItemReqCompra.sUM, tItemReqCompra.sCcl, tItemReqCompra.iAlmoxarifado, tItemReqCompra.sContaContabil, tItemReqCompra.iTipoTributacao, tItemReqCompra.lObservacao, tItemReqCompra.lFornecedor, tItemReqCompra.iFilial, tItemReqCompra.iExclusivo, lNumIntDoc, objItemReqCompras.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 54345

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54346

        If lErro = AD_SQL_SEM_DADOS Then Error 54426
    
        'Faz "LockExclusive" em Item Requisição Compras
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 54347

        'Se Quantidade Cancelada desse item foi alterada
        If objItemReqCompras.dQuantCancelada <> tItemReqCompra.dQuantCancelada Then

            'Verifica se o item está vinculado a concorrência gerada não baixada.
            lErro = Comando_Executar(alComando(2), "SELECT ItemRCItemConcorrencia.Quantidade FROM ItensConcorrencia, ItemRCItemConcorrencia WHERE ItensConcorrencia.NumIntDoc = ItemRCItemConcorrencia.ItemConcorrencia AND ItemRCItemConcorrencia.ItemReqCompra = ?", dQuantidade, objItemReqCompras.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then Error 54348

            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54349

            'Se estiver vinculada, para cada item
            Do While lErro = AD_SQL_SUCESSO
                'Acumula a quantidade
                dQuantTotalConcorrencia = dQuantTotalConcorrencia + dQuantidade

                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54425
                
            Loop

            'Preenche o objItemReqCompras
            objItemReqCompras.dQuantidade = tItemReqCompra.dQuantidade
            objItemReqCompras.dQuantPedida = tItemReqCompra.dQuantPedida

            dQuantMaximaCancelar = objItemReqCompras.dQuantidade - objItemReqCompras.dQuantPedida - objItemReqCompras.dQuantRecebida - dQuantTotalConcorrencia

            'Se a Quantida a cancelar for maior que a quantidade máxima a Cancelar --> Erro
            If objItemReqCompras.dQuantCancelada > 0 And objItemReqCompras.dQuantCancelada > dQuantMaximaCancelar Then Error 54350
                
            'Verifica se o item está vinculado a alguma Cotação
            lErro = Comando_ExecutarPos(alComando(4), "SELECT QuantidadeCotar FROM CotacaoProdutoItemRC WHERE ItemReqCompra = ?", 0, dQuantidade, objItemReqCompras.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then Error 54353

            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54354

            'Se estiver, quebra o vínculo
            Do While lErro = AD_SQL_SUCESSO

                'Exclui a Cotação Produto Item Requisição Compras
                lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM CotacaoProdutoItemRC", alComando(4))
                If lErro <> AD_SQL_SUCESSO Then Error 54355

                'Busca a próxima Cotação Produto Item Requisição Compras
                lErro = Comando_BuscarProximo(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54356

            Loop

            If objItemReqCompras.dQuantidade - objItemReqCompras.dQuantRecebida - objItemReqCompras.dQuantCancelada = 0 Then
                iStatus = ITEM_REQ_RECEBIDO
                iTotalItensFechados = iTotalItensFechados + 1
            End If

            'Altera na tabela de Itens Requisição Compras o valor da quantidade cancelada
            lErro = Comando_ExecutarPos(alComando(6), "UPDATE ItensReqCompra SET QuantCancelada = ?, Status = ?", alComando(1), objItemReqCompras.dQuantCancelada, iStatus)
            If lErro <> AD_SQL_SUCESSO Then Error 54357

        ElseIf objItemReqCompras.dQuantidade - objItemReqCompras.dQuantRecebida - objItemReqCompras.dQuantCancelada = 0 Then
            iTotalItensFechados = iTotalItensFechados + 1
        End If

    Next

    'Se todos os Itens Foram Recebidos, baixa a Requisição
    If iTotalItensFechados = objRequisicaoCompras.colItens.Count Then
        
        lErro = CF("ReqCompraEnv_BaixarEmTrans", objRequisicaoCompras)
        If lErro <> SUCESSO Then gError 66098
        
    End If
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 56255

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RequisicaoCompraEnv_Grava = SUCESSO
    
    Exit Function

Erro_RequisicaoCompraEnv_Grava:

    RequisicaoCompraEnv_Grava = gErr

    Select Case gErr

        Case 54339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 54341, 54342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", Err, objRequisicaoCompras.lCodigo)

        Case 54343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", Err, objRequisicaoCompras.lCodigo)

        Case 54344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISICAOCOMPRA", Err, objRequisicaoCompras.lCodigo)

        Case 54345, 54346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", Err, objRequisicaoCompras.lCodigo)

        Case 54347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSREQCOMPRA", Err, objRequisicaoCompras.lCodigo)

        Case 54348, 54349, 54425
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMRCITEMCONCORRENCIA1", Err)
        
        Case 54350
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTCANCELAR_SUPERIOR_QUANTDISPONIVEL_CANCELAR", Err, objItemReqCompras.sProduto, dQuantMaximaCancelar)

        Case 54353, 54354, 54356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTOITEMRC", Err, objItemReqCompras.lNumIntDoc)

        Case 54355
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COTACAOPRODUTOITEMRC", Err, objItemReqCompras.lNumIntDoc)

        Case 54357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSREQCOMPRA", Err, objItemReqCompras.lNumIntDoc)

        Case 54426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENSREQCOMPRA_NAO_CADASTRADO", Err, tItemReqCompra.lNumIntDoc, objRequisicaoCompras.lCodigo)
        
        Case 66098
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147527)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RequisicaoCompraEnv_Baixar(ByVal objRequisicaoCompras As ClassRequisicaoCompras) As Long

Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_RequisicaoCompraEnv_Baixar

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 54204

    'Baixa a Requisição
    lErro = CF("ReqCompraEnv_BaixarEmTrans", objRequisicaoCompras)
    If lErro <> SUCESSO Then gError 66099
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 54207

    RequisicaoCompraEnv_Baixar = SUCESSO

    Exit Function

Erro_RequisicaoCompraEnv_Baixar:

    RequisicaoCompraEnv_Baixar = gErr

    Select Case gErr

        Case 54204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
    
        Case 66099
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147528)
            
    End Select
    
    Call Transacao_Rollback

    Exit Function

End Function

Function RequisicaoCompra_Exclui(ByVal objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Exclui a Requisição de Compras a partir do NumIntDoc passado

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RequisicaoCompra_Exclui

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError (65539)
        
    lErro = CF("RequisicaoCompra_Exclui_EmTrans", objRequisicaoCompra)
    If lErro <> SUCESSO Then gError 62564
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError (65552)

    RequisicaoCompra_Exclui = SUCESSO
    
    Exit Function
    
Erro_RequisicaoCompra_Exclui:
    
    RequisicaoCompra_Exclui = gErr

    Select Case gErr
        
        Case 62564
        
        Case 65539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                                
        Case 65552
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147529)
    
    End Select
    
    Call Transacao_Rollback

    Exit Function
    
End Function

Function RequisicaoCompra_Exclui_EmTrans(ByVal objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Exclui a Requisição de Compras a partir do NumIntDoc passado

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumIntItem As Long
Dim objItemReqCompra As ClassItemReqCompras

On Error GoTo Erro_RequisicaoCompra_Exclui_EmTrans
    
    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError (65538)
    Next

    'Tenta ler a Requisição Modelo passada por objRequisicaoCompra
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RequisicaoCompra WHERE NumIntDoc = ?", 0, lNumIntDoc, objRequisicaoCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError (65540)
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError (65541)
    
    'Se não encontrou a Requisição de Compras, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError (65543)
    
    'Faz LockExclusive na Requisição de Compras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError (65544)
    
'    lErro = RequisicaoCompra_Desvincula_PedidoVenda(objRequisicaoCompra)
'    If lErro <> SUCESSO Then gError 86129
        
    'Lê e Faz "lock" do item
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM ItensReqCompra WHERE ReqCompra = ?", 0, lNumIntItem, objRequisicaoCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError (65545)
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError (65546)
    
    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError (65547)
        
    Do While lErro = AD_SQL_SUCESSO
    
        Set objItemReqCompra = New ClassItemReqCompras
        
        'Guarda o NumIntDoc do Item
        objItemReqCompra.lNumIntDoc = lNumIntItem
        
        'Adiciona o item na coleção
        objRequisicaoCompra.colItens.Add objItemReqCompra
        
        'Faz "LockExclusive" nos item da Requisição
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError (65548)

        'Busca próximo item
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError (65549)
    
    Loop
    
    'Exclui todos os Itens da Requisição Modelo
    lErro = CF("ItensReqCompra_Exclui", objRequisicaoCompra)
    If lErro <> SUCESSO Then gError (65550)
    
    'Exclui a Requisição de Compras
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM RequisicaoCompra", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError (65551)
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    RequisicaoCompra_Exclui_EmTrans = SUCESSO
    
    Exit Function
    
Erro_RequisicaoCompra_Exclui_EmTrans:
    
    RequisicaoCompra_Exclui_EmTrans = gErr

    Select Case gErr
        
        Case 65538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 65540, 65541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr, objRequisicaoCompra.lCodigo)
                
        Case 65543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objRequisicaoCompra.lCodigo)
            
        Case 65544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISICAOCOMPRA", gErr, objRequisicaoCompra.lCodigo)
        
        Case 65545, 65546, 65549
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)
        
        Case 65547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENSREQCOMPRA_NAO_CADASTRADO1", gErr, objRequisicaoCompra.lCodigo)
             
        Case 65548
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMREQCOMPRA", gErr, objItemReqCompra.lNumIntDoc)
        
        Case 65550, 86129
        
        Case 65551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REQUISICAOCOMPRA", gErr, objRequisicaoCompra.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147530)
    
    End Select

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function


'????========== Código PARA COMPILAR

Function PedidoCompra_Avalia_EstoqueMaximo(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Verifica se quantidade ordenada + disponivel > quantidade máxima de estoque

Dim lErro As Long
Dim lComando As Long, lComando1 As Long
Dim sComando_Sql As String
Dim dEstoqueMaximo As Double
Dim dTotal As Double
Dim objItemPC As ClassItemPedCompra
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_PedidoCompra_Avalia_EstoqueMaximo

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62651
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 62652
    
    For Each objItemPC In objPedidoCompra.colItens

        sComando_Sql = "SELECT SUM (EstoqueProduto.QuantDispNossa + EstoqueProduto.QuantConsig3 + EstoqueProduto.QuantReservada + EstoqueProduto.QuantReservadaConsig) FROM EstoqueProduto, Almoxarifado WHERE (EstoqueProduto.Produto = ?) AND (Almoxarifado.FilialEmpresa = ?) AND (Almoxarifado.Codigo = EstoqueProduto.Almoxarifado)"
        lErro = Comando_Executar(lComando, sComando_Sql, dTotal, objItemPC.sProduto, objPedidoCompra.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 62656
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62658

        If lErro = AD_SQL_SEM_DADOS Then Error 62659
        
       'Lê de ProdutosFilial a Quantidade máxima de estoque desse produto
        sComando_Sql = "SELECT EstoqueMaximo FROM ProdutosFilial WHERE FilialEmpresa = ? AND Produto = ?"
        lErro = Comando_Executar(lComando1, sComando_Sql, dEstoqueMaximo, objPedidoCompra.iFilialEmpresa, objItemPC.sProduto)
        If lErro <> AD_SQL_SUCESSO Then Error 62653
    
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62654
    
        If lErro = AD_SQL_SUCESSO Then
                
            If dEstoqueMaximo > 0 Then
                If objItemPC.dQuantidade + dTotal > dEstoqueMaximo Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_QUANTIDADE_PEDIDA_ESTOQUEMAXIMO", objItemPC.sProduto)
                    If vbMsgRes = vbNo Then Error 62655
                End If
            End If
            
        End If
    
    Next
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    PedidoCompra_Avalia_EstoqueMaximo = SUCESSO

    Exit Function

Erro_PedidoCompra_Avalia_EstoqueMaximo:

    PedidoCompra_Avalia_EstoqueMaximo = Err

    Select Case Err

        Case 62651, 62652
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 62653, 62654, 62656, 62658
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err)
        
        Case 62655
            
        Case 62659
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DETERMINACAO_QUANTMAX", Err, objPedidoCompra.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147531)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function
    
End Function

Private Function Testa_Pedido_Cotacao(ByVal lCotacaoProduto As Long, ByVal objPedidoCotacao As ClassPedidoCotacao, ByVal objCotacao As ClassCotacao) As Long
'Verifica se existe algum pedido de cotação diferente do pedido passado como parametro associado a cotação do pedido

Dim lErro As Long
Dim lCotacao As Long
Dim lComando As Long
Dim lPedidoCotacao As Long

On Error GoTo Erro_Testa_Pedido_Cotacao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 74916

    'Pesquisa na tabela CotacaoProduto o NumIntDoc da Cotação associada.
    lErro = Comando_Executar(lComando, "SELECT Cotacao FROM CotacaoProduto WHERE NumIntDoc = ? ", lCotacao, lCotacaoProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 74917

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 74921
    
    'se não encontrou cotação produto no BD ==> erro
    If lErro <> AD_SQL_SUCESSO Then gError 74918

    objCotacao.lNumIntDoc = lCotacao
                 
    'Verifica se existe Pedido de Cotacao ligado a Cotacao informada
    lErro = Comando_Executar(lComando, "SELECT PedidoCotacao.Codigo FROM PedidoCotacao,ItemPedCotacao,CotacaoProduto WHERE ItemPedCotacao.CotacaoProduto=CotacaoProduto.NumIntDoc AND ItemPedCotacao.PedidoCotacao=PedidoCotacao.NumIntDoc AND PedidoCotacao.Codigo<>? AND PedidoCotacao.FilialEmpresa = ? AND CotacaoProduto.Cotacao= ?", lPedidoCotacao, objPedidoCotacao.lCodigo, objPedidoCotacao.iFilialEmpresa, objCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 83843
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83844
    
    If lErro = AD_SQL_SEM_DADOS Then gError 74922
    
    Call Comando_Fechar(lComando)
    
    Testa_Pedido_Cotacao = SUCESSO
    
    Exit Function
    
Erro_Testa_Pedido_Cotacao:

    Testa_Pedido_Cotacao = gErr
    
    Select Case gErr
    
        Case 74916
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 74917, 74921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTO", gErr)
            
        Case 74918
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COTACAOPRODUTO_INEXISTENTE", gErr, lCotacaoProduto)
        
        Case 74919, 74922, 83843, 83844
        
        Case 74920
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COTACAOPRODUTO_INEXISTENTE1", gErr, objCotacao.lNumIntDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147532)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function


Function RequisicaoCompra_Grava(ByVal objRequisicaoCompra As ClassRequisicaoCompras, Optional ByVal bGeracao As Boolean = False) As Long
'Grava a Requisição de Compra

Dim lTransacao As Long
Dim lErro As Long

On Error GoTo Erro_RequisicaoCompra_Grava
    
    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError (61995)
    
    lErro = CF("RequisicaoCompra_Grava_EmTrans", objRequisicaoCompra, bGeracao)
    If lErro <> SUCESSO Then gError (62561)
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError (65507)
   
    lErro = CF("WorkFlow_Trata_Transacao", "COM", objRequisicaoCompra.objTela, gsUsuario, objRequisicaoCompra)
    If lErro <> SUCESSO Then gError 178684
   
    RequisicaoCompra_Grava = SUCESSO

    Exit Function
    
Erro_RequisicaoCompra_Grava:

    RequisicaoCompra_Grava = gErr

    Select Case gErr
    
        Case 61995
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 62561, 178684
               
        Case 65507
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
                                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147533)
    
    End Select
    
    Call Transacao_Rollback

    Exit Function
    
End Function

Function RequisicaoModelo_Exclui(ByVal objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Exclui a Requisição Modelo a partir do NumIntDoc passado

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lNumIntItem As Long
Dim objItemReqModelo As ClassItemReqModelo

On Error GoTo Erro_RequisicaoModelo_Exclui
    
    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 61687
    Next

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 61688
    
    'Tenta ler a Requisição Modelo passada por objRequisicaoModelo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RequisicaoModelo WHERE NumIntDoc = ?", 0, lNumIntDoc, objRequisicaoModelo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 61691
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61692
    
    'Se não encontrou a Requisição Modelo, Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 61693
    
    'Faz LockExclusive na Requisição Modelo
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 61690
            
    'Exclui todos os Itens da Requisição Modelo
    lErro = CF("ItensReqModelo_Exclui", objRequisicaoModelo)
    If lErro <> SUCESSO Then Error 61698
    
    'Exclui a Requisição Modelo
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM RequisicaoModelo ", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 61699
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 61689

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    RequisicaoModelo_Exclui = SUCESSO
    
    Exit Function
    
Erro_RequisicaoModelo_Exclui:

    
    RequisicaoModelo_Exclui = Err

    Select Case Err
        
        Case 61687
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 61688
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 61689
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 61690
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISICAOMODELO", Err, objRequisicaoModelo.lCodigo)
                
        Case 61691, 61692
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOMODELO", Err, objRequisicaoModelo.lNumIntDoc)
                
        Case 61693
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOMODELO_NAO_CADASTRADA1", Err, objRequisicaoModelo.lNumIntDoc)
            
        Case 61694, 61695, 61811
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQMODELO", Err, objRequisicaoModelo.lCodigo)
        
        Case 61696
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMREQMODELO_NAO_CADASTRADO", Err, objRequisicaoModelo.lCodigo)
             
        Case 61697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMREQMODELO", Err, objItemReqModelo.lNumIntDoc)
        
        Case 61698
        
        Case 61699
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REQUISICAOMODELO", Err, objRequisicaoModelo.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147534)
    
    End Select
    
    Call Transacao_Rollback

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function ItensReqModelo_Exclui(ByVal objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Exclui os Itens da Requisição Modelo

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lNumIntItem As Long

On Error GoTo Erro_ItensReqModelo_Exclui
    
    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 61769
    Next
        
    'Lê e Faz "lock" do item
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ItensReqModelo WHERE ReqModelo = ?", 0, lNumIntItem, objRequisicaoModelo.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 61694
        
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61695
            
    Do While lErro = AD_SQL_SUCESSO
        
        'Faz "LockExclusive" no item da Requisição
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 61697

        'Exclui o Item do banco de dados
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ItensReqModelo", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 61772

        'Busca próximo item
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61811
    
    Loop
                
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ItensReqModelo_Exclui = SUCESSO
    
    Exit Function


    
Erro_ItensReqModelo_Exclui:


    
    ItensReqModelo_Exclui = Err

    Select Case Err
        
        Case 61769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 61694, 61695, 61811
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQMODELO", Err, objRequisicaoModelo.lCodigo)
                     
        Case 61697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEMREQMODELO", Err, lNumIntItem)
        
        Case 61772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMREQMODELO", lNumIntItem)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147535)
    
    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next


    
    Exit Function

End Function

Function RequisicaoModelo_Grava(ByVal objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Grava a Requisição Modelo

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim objItemReqModelo As ClassItemReqModelo
Dim lNumIntItem As Long

On Error GoTo Erro_RequisicaoModelo_Grava
    
    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 61667
    Next

    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 61668

    'Faz "lock" nas tabelas relacionadas a Requisição Modelo
    lErro = CF("RequisicaoModelo_Lock_Gravacao", objRequisicaoModelo)
    If lErro <> SUCESSO Then Error 61670
    
    'Faz "lock" nas tabelas relacionadas ao item da Requisição Modelo
    lErro = CF("ItensReqModelo_Lock_Gravacao", objRequisicaoModelo)
    If lErro <> SUCESSO Then Error 61671
    
    'Tenta ler a Requisição Modelo passada por objRequisicaoModelo
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM RequisicaoModelo WHERE FilialEmpresa = ? AND Codigo = ?", 0, lNumIntDoc, objRequisicaoModelo.iFilialEmpresa, objRequisicaoModelo.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 61672
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61673

    'Se não encontrou, INSERÇÃO
    If lErro = AD_SQL_SEM_DADOS Then
        
        'Gera número interno para a próxima Requisição Modelo
        lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_RC_MODELO", lNumIntDoc)
        If lErro <> SUCESSO Then Error 61674

        'Guarda NumIntDoc em objRequisicaoModelo
        objRequisicaoModelo.lNumIntDoc = lNumIntDoc
            
        'Grava a nova Requisição Modelo no BD
        lErro = Comando_Executar(alComando(1), "INSERT INTO RequisicaoModelo (NumIntDoc, FilialEmpresa, Codigo, Descricao, Urgente, Requisitante, Ccl, FilialCompra, TipoDestino, FornCliDestino, FilialDestino, Observacao, TipoTributacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", objRequisicaoModelo.lNumIntDoc, objRequisicaoModelo.iFilialEmpresa, objRequisicaoModelo.lCodigo, objRequisicaoModelo.sDescricao, objRequisicaoModelo.iUrgente, objRequisicaoModelo.lRequisitante, objRequisicaoModelo.sCcl, objRequisicaoModelo.iFilialCompra, objRequisicaoModelo.iTipoDestido, objRequisicaoModelo.lFornCliDestino, objRequisicaoModelo.iFilialDestino, objRequisicaoModelo.lObservacao, objRequisicaoModelo.iTipoTributacao)
        If lErro <> AD_SQL_SUCESSO Then Error 61675
        
        'Grava os Itens da Requisição no BD
        lErro = CF("ItensReqModelo_Inclui", objRequisicaoModelo)
        If lErro <> SUCESSO Then Error 61676
            
    'Encontrou, ALTERAÇÃO
    Else
    
        'Faz LockExclusive na Requisição
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 61677
      
        objRequisicaoModelo.lNumIntDoc = lNumIntDoc
                
        'Atualiza os itens da Requisição Modelo
        lErro = CF("ItensReqModelo_Exclui", objRequisicaoModelo)
        If lErro <> SUCESSO Then Error 61773

        lErro = CF("ItensReqModelo_Inclui", objRequisicaoModelo)
        If lErro <> SUCESSO Then Error 61774
                               
        'Atualiza os Dados da Requisição Modelo
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE RequisicaoModelo SET FilialEmpresa = ?, Descricao = ?, Urgente = ?, Requisitante = ?, Ccl = ?, FilialCompra = ?, TipoDestino = ?, FornCliDestino = ?, FilialDestino = ?, Observacao = ?, TipoTributacao = ?", alComando(0), objRequisicaoModelo.iFilialEmpresa, objRequisicaoModelo.sDescricao, objRequisicaoModelo.iUrgente, objRequisicaoModelo.lRequisitante, objRequisicaoModelo.sCcl, objRequisicaoModelo.iFilialCompra, objRequisicaoModelo.iTipoDestido, objRequisicaoModelo.lFornCliDestino, objRequisicaoModelo.iFilialDestino, objRequisicaoModelo.lObservacao, objRequisicaoModelo.iTipoTributacao)
        If lErro <> AD_SQL_SUCESSO Then Error 61682

    End If
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then Error 61669

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RequisicaoModelo_Grava = SUCESSO

    Exit Function
    
Erro_RequisicaoModelo_Grava:

    RequisicaoModelo_Grava = Err

    Select Case Err
    
        Case 61667
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 61668
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 61669
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 61670, 61671, 61674, 61676, 61773, 61774
        
        Case 61672, 61673
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOMODELO", Err, objRequisicaoModelo.lCodigo)
        
        Case 61675
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REQUISICAOMODELO", Err, objRequisicaoModelo.lCodigo)
        
        Case 61677
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISICAOMODELO", Err, objRequisicaoModelo.lCodigo)
                        
        Case 61682
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_REQUISICAOMODELO", Err, objRequisicaoModelo.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147536)
    
    End Select
    
    Call Transacao_Rollback

    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

'====== mais codigo que havia sumido Jones 29/08/00

Function RequisicaoModelo_Lock_Gravacao(ByVal objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Faz Lock nas tabelas Associadas a Requisição Modelo

Dim lErro As Long
Dim lObservacao As Long

On Error GoTo Erro_RequisicaoModelo_Lock_Gravacao

    'Se a FilialCompra tiver sido preenchida
    If objRequisicaoModelo.iFilialCompra <> 0 Then
    
        'Le e locka FilialCompra
        lErro = CF("FilialEmpresa_Lock", objRequisicaoModelo.iFilialCompra)
        If lErro <> SUCESSO And lErro <> 61725 Then Error 61721
        If lErro = 61725 Then Error 61722
        
    End If
    
    'Se o Ccl está preenchido
    If Len(Trim(objRequisicaoModelo.sCcl)) > 0 Then
        
        'Lê o locka Ccl
        lErro = CF("Ccl_Lock", objRequisicaoModelo.sCcl)
        If lErro <> SUCESSO And lErro <> 61731 Then Error 61732
        If lErro = 61731 Then Error 61733
    
    End If
    
    'Se o Requisitante estiver preenchido
    If objRequisicaoModelo.lRequisitante <> 0 Then
        
        'Lê e Faz Lock do Requisitante
        lErro = CF("Requisitante_Lock", objRequisicaoModelo.lRequisitante)
        If lErro <> SUCESSO And lErro <> 61737 Then Error 61740
        If lErro = 61737 Then Error 61741
    
    End If
    
    'Trata o Lock do Tipo Destino
    If objRequisicaoModelo.iTipoDestido = TIPO_DESTINO_EMPRESA Then
        
        lErro = CF("FilialEmpresa_Lock", objRequisicaoModelo.iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 61725 Then Error 61742
        
        'Se não encontrou FilialEmpresa, Erro
        If lErro = 61725 Then Error 61743

    ElseIf objRequisicaoModelo.iTipoDestido = TIPO_DESTINO_FORNECEDOR Then

        lErro = CF("FilialFornecedor_Lock", objRequisicaoModelo.lFornCliDestino, objRequisicaoModelo.iFilialDestino)
        If lErro <> SUCESSO And lErro <> 42870 Then Error 61746

        'Se nao encontrou FilialFornecedor, Erro
        If lErro = 42870 Then Error 61747

    End If

    'Verifica se observacao esta preenchida
    If Len(Trim(objRequisicaoModelo.sObservacao)) > 0 Then

        'Faz Lock e insere nova Observação
        lErro = CF("Observacao_Lock", objRequisicaoModelo.sObservacao, lObservacao)
        If lErro <> SUCESSO Then Error 61753
        
        objRequisicaoModelo.lObservacao = lObservacao
            
    End If
    
    RequisicaoModelo_Lock_Gravacao = SUCESSO
    
    Exit Function
    
Erro_RequisicaoModelo_Lock_Gravacao:

    RequisicaoModelo_Lock_Gravacao = Err
    
    Select Case Err
                
        Case 61721, 61732, 61740, 61742, 61746, 61753
        
        Case 61722
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", Err, objRequisicaoModelo.iFilialCompra)
        
        Case 61733
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, objRequisicaoModelo.sCcl)
        
        Case 61741
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", Err, objRequisicaoModelo.lRequisitante)
        
        Case 61743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", Err, objRequisicaoModelo.iFilialEmpresa)
                    
        Case 61747
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objRequisicaoModelo.lFornCliDestino, objRequisicaoModelo.iFilialDestino)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147537)
    
    End Select

    Exit Function
    
End Function

Function ItensReqModelo_Lock_Gravacao(ByVal objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Faz lock nas tabelas associados ao item da Requisição Modelo

Dim lErro As Long
Dim objItemReqModelo As ClassItemReqModelo
Dim objProduto As New ClassProduto
Dim lComando As Long
Dim lObservacao As Long
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_ItensReqModelo_Lock_Gravacao
    
    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 61582
    
    'Para cada item da Requisição Modelo
    For Each objItemReqModelo In objRequisicaoModelo.colItensReqModelo
                        
        If objItemReqModelo.lFornecedor > 0 Then
        
            'Lê e faz lock do FornecedorProdutoFF
            objFornecedorProdutoFF.lFornecedor = objItemReqModelo.lFornecedor
            objFornecedorProdutoFF.iFilialForn = objItemReqModelo.iFilial
            objFornecedorProdutoFF.sProduto = objItemReqModelo.sProduto
            objFornecedorProdutoFF.iFilialEmpresa = objRequisicaoModelo.iFilialCompra
            
            lErro = CF("FornecedorProdutoFF_Lock", objFornecedorProdutoFF)
            If lErro <> SUCESSO And lErro <> 67872 Then gError 67876
            
            'Se o FornecedorProdutoFF não está cadastrado, erro
            If lErro = 67872 Then gError 67877
        
        End If
        
        'Lê e Faz Lock do Produto
        objProduto.sCodigo = objItemReqModelo.sProduto
        lErro = CF("Produto_Lock", objProduto)
        If lErro <> SUCESSO Then gError 61754
        
        'Se o Almoxarifado estiver preenchido
        If objItemReqModelo.iAlmoxarifado <> 0 Then
        
            'Lê e Faz Lock do Almoxarifado
            lErro = CF("Almoxarifado_Lock", objItemReqModelo.iAlmoxarifado, lComando)
            If lErro <> SUCESSO Then gError 61756
        
        End If
        
        'Se o Ccl estiver preenchido
        If Len(Trim(objItemReqModelo.sCcl)) > 0 Then
        
            'Lê e Faz lock do Ccl
            lErro = CF("Ccl_Lock", objItemReqModelo.sCcl)
            If lErro <> SUCESSO And lErro <> 61731 Then gError 61757
            If lErro = 61731 Then gError 61758
        
        End If
        
        'Se a Conta Contábil estiver preenchida
        If Len(Trim(objItemReqModelo.sContaContabil)) > 0 Then
            
            'Lê e Faz Lock da Conta Contábil
            lErro = CF("ContaContabil_Lock", objItemReqModelo.sContaContabil)
            If lErro <> SUCESSO And lErro <> 61762 Then gError 61764
            If lErro = 61762 Then gError 61765
            
        End If
        
        'Se Fornecedor e Filial estiverem Preenchidos
        If objItemReqModelo.lFornecedor <> 0 And objItemReqModelo.iFilial <> 0 Then
        
            'Lê e Faz Lock em FilialFornecedor
            lErro = CF("FilialFornecedor_Lock", objItemReqModelo.lFornecedor, objItemReqModelo.iFilial)
            If lErro <> SUCESSO And lErro <> 42870 Then gError 61766
            If lErro = 42870 Then gError 61767
        End If
        
        'Se a Observação estiver Preenchida
        If Len(Trim(objItemReqModelo.sObservacao)) > 0 Then
    
            'Lê e Faz Lock em Observação
            lErro = CF("Observacao_Lock", objItemReqModelo.sObservacao, lObservacao)
            If lErro <> SUCESSO Then gError 61768
                        
            objItemReqModelo.lObservacao = lObservacao
        
        End If
        
    Next
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    ItensReqModelo_Lock_Gravacao = SUCESSO
    
    Exit Function
    
Erro_ItensReqModelo_Lock_Gravacao:

    ItensReqModelo_Lock_Gravacao = gErr

    Select Case gErr
    
        Case 61582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 61754, 61756, 61757, 61764, 61766, 61768, 67876
                 
        Case 61758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, objItemReqModelo.sCcl)
        
        Case 61765
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, objItemReqModelo.sContaContabil)
        
        Case 61767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objItemReqModelo.lFornecedor, objItemReqModelo.iFilial)
        
        Case 67878
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_ENCONTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147538)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function ItensReqModelo_Inclui(ByVal objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Insere os Itens da Requisição Modelo no BD

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim objItemReqModelo As ClassItemReqModelo
Dim iIndice As Integer
Dim lNumIntItem As Long

On Error GoTo Erro_ItensReqModelo_Inclui

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 61509
    Next
    
    'Para cada item da Requisição Modelo
    For Each objItemReqModelo In objRequisicaoModelo.colItensReqModelo
        
        'Gera numero interno automatico para o item
        lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_ITEM_RC_MODELO", lNumIntItem)
        If lErro <> SUCESSO Then Error 61809

        objItemReqModelo.lNumIntDoc = lNumIntItem
    
        'Grava o Item da Requisição Modelo no BD
        lErro = Comando_Executar(alComando(0), "INSERT INTO ItensReqModelo (NumIntDoc, ReqModelo, Produto, DescProduto, Quantidade, UM, Ccl, Almoxarifado, ContaContabil, TipoTributacao, Observacao, Fornecedor, Filial, Exclusivo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objItemReqModelo.lNumIntDoc, objRequisicaoModelo.lNumIntDoc, objItemReqModelo.sProduto, objItemReqModelo.sDescProduto, objItemReqModelo.dQuantidade, objItemReqModelo.sUM, objItemReqModelo.sCcl, objItemReqModelo.iAlmoxarifado, objItemReqModelo.sContaContabil, objItemReqModelo.iTipoTributacao, objItemReqModelo.lObservacao, objItemReqModelo.lFornecedor, objItemReqModelo.iFilial, objItemReqModelo.iExclusivo)
        If lErro <> AD_SQL_SUCESSO Then Error 61664

    Next
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItensReqModelo_Inclui = SUCESSO
    
    Exit Function
    
Erro_ItensReqModelo_Inclui:

    ItensReqModelo_Inclui = Err

    Select Case Err
            
        Case 61509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 61664
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMREQMODELO", Err, objItemReqModelo.lNumIntDoc)
            
        Case 61809
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147539)
    
    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function RequisicaoModelo_Codigo_Automatico(lCodigo As Long) As Long
'Gera código automático para a próxima Requisição Modelo

Dim lErro As Long
Dim sCodigo As String, sTabelaConfig As String, sTabelaObjeto As String
Dim sCampo As String

On Error GoTo Erro_RequisicaoModelo_Codigo_Automatico

    sCodigo = "NUM_PROXIMO_CODIGO_RC_MODELO"
    sTabelaConfig = "ComprasConfig"
    sTabelaObjeto = "RequisicaoModelo"
    sCampo = "Codigo"

    'Chama Config_ObterAutomatico
    lErro = CF("Config_ObterAutomatico", sTabelaConfig, sCodigo, sTabelaObjeto, sCampo, lCodigo)
    If lErro <> SUCESSO Then Error 61580

    RequisicaoModelo_Codigo_Automatico = SUCESSO
    
    Exit Function

Erro_RequisicaoModelo_Codigo_Automatico:

    RequisicaoModelo_Codigo_Automatico = Err
    
    Select Case Err

        Case 61580 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147540)

    End Select

    Exit Function

End Function

Function Requisitante_Grava(ByVal objRequisitante As ClassRequisitante) As Long
'Grava o Requisitante no BD

Dim lTransacao As Long
Dim alComando(0 To 2) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim sCcl As String
Dim lCodigo As Long
Dim sRequisitante As String

On Error GoTo Erro_Requisitante_Grava

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 49068
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 49069

    'Procura por Requisitante com mesmo nome reduzido
     lErro = Comando_Executar(alComando(0), "SELECT Codigo FROM Requisitante WHERE Codigo <> ? AND NomeReduzido = ?", lCodigo, objRequisitante.lCodigo, objRequisitante.sNomeReduzido)
     If lErro <> AD_SQL_SUCESSO Then gError 49070

    'Tenta ler o código deste Requisitante
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 49074
    
    'Se existir outro Requisitante com o mesmo Nome Reduzido, erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 49075

    'Procura por Requisitante com mesmo nome
    lErro = Comando_Executar(alComando(0), "SELECT Codigo FROM Requisitante WHERE Codigo <> ? AND Nome = ?", lCodigo, objRequisitante.lCodigo, objRequisitante.sNome)
    If lErro <> AD_SQL_SUCESSO Then gError 49288

    'Tenta ler o código deste Requisitante
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 49289
    
    'Se existir outro Requisitante com o mesmo Nome, erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 49290

    'se o ccl foi informado
    If Len(Trim(objRequisitante.sCcl)) > 0 Then

'Ok?CATHARINE
        lErro = CF("Ccl_Lock", objRequisitante.sCcl)
        If lErro <> SUCESSO Then gError 49076

    End If

    sRequisitante = String(STRING_REQUISITANTE_NOMERED, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NomeReduzido FROM Requisitante WHERE Codigo = ?", 0, sRequisitante, objRequisitante.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 49085

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 49086

    'Encontrou o Requisitante
    If lErro = AD_SQL_SUCESSO Then
                
        'Lock do Requisitante
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 49080

        'Atualiza os dados do Requisitante
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Requisitante SET Nome = ?,NomeReduzido = ?,Ccl = ?, Email = ?,CodUsuario=?", alComando(1), objRequisitante.sNome, objRequisitante.sNomeReduzido, objRequisitante.sCcl, objRequisitante.sEmail, objRequisitante.sCodUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 49087

    Else
        'Insere o Requisitante
        lErro = Comando_Executar(alComando(2), "INSERT INTO Requisitante(Codigo,Nome,NomeReduzido,Ccl, Email,CodUsuario) VALUES (?,?,?,?,?,?)", objRequisitante.lCodigo, objRequisitante.sNome, objRequisitante.sNomeReduzido, objRequisitante.sCcl, objRequisitante.sEmail, objRequisitante.sCodUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 49088

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 49089

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Requisitante_Grava = SUCESSO

    Exit Function

Erro_Requisitante_Grava:

    Requisitante_Grava = gErr

    Select Case gErr

        Case 49068
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 49069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 49070, 49074, 49085, 49086, 49288, 49289
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISITANTE", gErr, objRequisitante.lCodigo)

        Case 49075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NOMERED_DUPLICADO", gErr, objRequisitante.sNomeReduzido)

        Case 49076
        
        Case 49080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISITANTE", gErr, objRequisitante.lCodigo)

        Case 49087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_REQUISITANTE", gErr, objRequisitante.lCodigo)

        Case 49088
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_REQUISITANTE", gErr, objRequisitante.lCodigo)

        Case 49089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 49290
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NOME_DUPLICADO", gErr, objRequisitante.sNome)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147541)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Requisitante_Exclui(ByVal objRequisitante As ClassRequisitante) As Long
'Exclui o Requisitante

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim lTransacao As Long
Dim lTotal As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_Requisitante_Exclui

   'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 49062
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 49063

    'Seleciona o Requisitante a ser excluido
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Requisitante WHERE Codigo =?", 0, lCodigo, objRequisitante.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49064

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49090
    If lErro = AD_SQL_SEM_DADOS Then Error 49091

    'Lock do Requisitante
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 49065

    lErro = CF("Requisitante_Valida_Exclusao", objRequisitante, alComando)
    If lErro <> SUCESSO Then Error 51351
        
    'Exclui o Requisitante
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Requisitante", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 49066

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 49067

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Requisitante_Exclui = SUCESSO

    Exit Function

Erro_Requisitante_Exclui:

    Requisitante_Exclui = Err

    Select Case Err

        Case 49062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49063
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 49064, 49090
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISITANTE", Err, objRequisitante.lCodigo)

        Case 49065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISITANTE", Err, objRequisitante.lCodigo)

        Case 49066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_REQUISITANTE", Err, objRequisitante.lCodigo)

        Case 49067
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 49091
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISITANTE_NAO_CADASTRADO", Err, objRequisitante.lCodigo)
            
        Case 51351
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147542)

    End Select

    Call Transacao_Rollback
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Requisitante_Automatico(lCodigo As Long) As Long
'Gera o proximo codigo da Tabela de Requisitante

Dim lErro As Long

On Error GoTo Erro_Requisitante_Automatico

    'Chama a rotina que gera o sequencial
    lErro = CF("Config_ObterAutomatico", "ComprasConfig", "NUM_PROXIMO_REQUISITANTE", "Requisitante", "Codigo", lCodigo)
    If lErro <> SUCESSO Then Error 49071

    Requisitante_Automatico = SUCESSO

    Exit Function

Erro_Requisitante_Automatico:

    Requisitante_Automatico = Err

    Select Case Err

        Case 49071

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 147543)

    End Select

    Exit Function

End Function

Function TipoBloqueioPC_Grava(ByVal objTipoDeBloqueioPC As ClassTipoBloqueioPC) As Long
'Grava no BD

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim sNomeReduzido As String
Dim sDescricao As String
Dim iCodigo As Integer
Dim iIndice As Integer
Dim lTransacao As Long

On Error GoTo Erro_TipoBloqueioPC_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
         alComando(iIndice) = Comando_Abrir()
         If alComando(iIndice) = 0 Then Error 49120
    Next

    'Inicia a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 49121

    sNomeReduzido = String(STRING_TIPODEBLOQUEIOPC_NOME_REDUZIDO, 0)

    'Procura por um TipoDeBloqueioPC com o mesmo Nome Reduzido
    lErro = Comando_Executar(alComando(0), "SELECT NomeReduzido FROM TiposDeBloqueioPC WHERE Codigo <>? and NomeReduzido = ?", sNomeReduzido, objTipoDeBloqueioPC.iCodigo, objTipoDeBloqueioPC.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then Error 49122

    'Tenta ler o código deste Tipo de BloqueioPC
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49123

    'Encontrou Tipo de Bloqueio PC com o mesmo Nome Reduzido
    If lErro <> AD_SQL_SEM_DADOS Then Error 49124

    sDescricao = String(STRING_TIPODEBLOQUEIOPC_DESCRICAO, 0)

    'Procura por um TipoDeBloqueioPC com mesma Descricao
    lErro = Comando_Executar(alComando(0), "SELECT Descricao FROM TiposDeBloqueioPC WHERE Codigo <>? and Descricao = ?", sDescricao, objTipoDeBloqueioPC.iCodigo, objTipoDeBloqueioPC.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then Error 49125

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49139

    'Encontrou Tipo de Bloqueio PC com a mesma Descricao
    If lErro <> AD_SQL_SEM_DADOS Then Error 49126

    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM TiposDeBloqueioPC WHERE Codigo = ?", 0, iCodigo, objTipoDeBloqueioPC.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49127

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49148

    If lErro = AD_SQL_SUCESSO Then

        'Lock do Tipo de Bloqueio PC
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 49211

        'atualiza os dados do Tipo de Bloqueio PC
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TiposDeBloqueioPC SET NomeReduzido = ?,Descricao= ?", alComando(0), objTipoDeBloqueioPC.sNomeReduzido, objTipoDeBloqueioPC.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then Error 49128

    Else

        'insere o Tipo de Bloqueio PC
        lErro = Comando_Executar(alComando(1), "INSERT INTO TiposDeBloqueioPC(Codigo,NomeReduzido,Descricao) VALUES (?,?,?)", objTipoDeBloqueioPC.iCodigo, objTipoDeBloqueioPC.sNomeReduzido, objTipoDeBloqueioPC.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then Error 49129

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 49130

    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        Call Comando_Fechar(alComando(iIndice))

    Next

    TipoBloqueioPC_Grava = SUCESSO

    Exit Function

Erro_TipoBloqueioPC_Grava:

   TipoBloqueioPC_Grava = Err

    Select Case Err

        Case 49120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 49122, 49123, 49125, 49127, 49139, 49148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODEBLOQUEIOPC", Err, objTipoDeBloqueioPC.iCodigo)

        Case 49124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_MESMO_NOME", Err, objTipoDeBloqueioPC.sNomeReduzido)

        Case 49126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_MESMA_DESCRICAO", Err, objTipoDeBloqueioPC.sDescricao)

        Case 49128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TIPODEBLOQUEIOPC", Err, objTipoDeBloqueioPC.iCodigo)

        Case 49129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPODEBLOQUEIOPC", Err, objTipoDeBloqueioPC.iCodigo)

        Case 49130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 49211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPODEBLOQUEIOPC", Err, objTipoDeBloqueioPC.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147544)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TipoBloqueioPC_Exclui(ByVal objTipoDeBloqueioPC As ClassTipoBloqueioPC) As Long
'Exclui do BD

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(0 To 5) As Long
Dim iIndice As Integer
Dim iCodigo As Integer

On Error GoTo Erro_TipoBloqueioPC_Exclui

    'Abre os comandos
     For iIndice = LBound(alComando) To UBound(alComando)
         alComando(iIndice) = Comando_Abrir()
         If alComando(iIndice) = 0 Then Error 49131
    Next

    'Inicia transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 49132

    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM TiposDeBloqueioPC WHERE codigo = ?", 0, iCodigo, objTipoDeBloqueioPC.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49133

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49149
    If lErro = AD_SQL_SEM_DADOS Then Error 49150

    'Lock do Tipo de Bloqueio PC
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 49134

    'Verifica se o Tipo de Bloqueio PC esta relacionado com algum BloqueioPC
    lErro = Comando_ExecutarPos(alComando(1), "SELECT TipoDeBloqueio FROM BloqueiosPC WHERE TipoDeBloqueio =?", 0, iCodigo, objTipoDeBloqueioPC.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49135

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49151

    'O Tipo de  Bloqueio PC esta relacionado a um Bloqueio PC
    If lErro = AD_SQL_SUCESSO Then Error 49136

    'Exclui Tipo de BloqueioPC
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TiposDeBloqueioPC", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 49137

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 49138

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TipoBloqueioPC_Exclui = SUCESSO

    Exit Function

Erro_TipoBloqueioPC_Exclui:

    TipoBloqueioPC_Exclui = Err

    Select Case Err

        Case 49131
             lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 49133, 49149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODEBLOQUEIOPC", Err, objTipoDeBloqueioPC.iCodigo)

        Case 49134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPODEBLOQUEIOPC", Err, objTipoDeBloqueioPC.iCodigo)

        Case 49135, 49151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPC", Err, objTipoDeBloqueioPC.iCodigo)

        Case 49136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_VINCULADO_BLOQUEIOSPC", Err, objTipoDeBloqueioPC.iCodigo)

        Case 49137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPODEBLOQUEIOPC", Err, objTipoDeBloqueioPC.iCodigo)

        Case 49138
             lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 49150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIOPC_NAO_CADASTRADO", Err, objTipoDeBloqueioPC.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147545)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ConfiguraCOM_Gravar(ByVal objConfiguraCOM As ClassConfiguraCOM) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ConfiguraCOM_Gravar

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 49347

    lErro = CF("ConfiguraCOM_GravarTrans", objConfiguraCOM)
    If lErro <> SUCESSO Then gError 74933
    
     'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 49352

    ConfiguraCOM_Gravar = SUCESSO

    Exit Function

Erro_ConfiguraCOM_Gravar:

    ConfiguraCOM_Gravar = gErr

    Select Case gErr

        Case 49347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 49352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 74933
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147546)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function PedidoCompraBaixado_Atualiza_DataEmissao(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Atualiza data de emissao no BD, em PedidoCompraBaixado, para a data atual

Dim lErro As Long
Dim lComando(0 To 2) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lCodigo As Long
Dim iFilialEmpresa As Long

On Error GoTo Erro_PedidoCompraBaixado_Atualiza_DataEmissao

    'Abre os  comandos
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then gError 76065
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 76067

    'OK você tem o numintdoc do pedido para usar como filtro
    'Busca na tabela de PedidoCompraBaixado, o pedido de compra passado como parametro
    lErro = Comando_ExecutarPos(lComando(0), "SELECT NumIntDoc,Codigo,FilialEmpresa  FROM PedidoCompraN WHERE  Codigo = ? and FilialEmpresa =? AND Status = ?", 0, lNumIntDoc, lCodigo, iFilialEmpresa, objPedidoCompra.lCodigo, objPedidoCompra.iFilialEmpresa, PEDIDOCOMPRA_STATUS_BAIXADO)
    If lErro <> AD_SQL_SUCESSO Then gError 76068

    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76069

    'se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 76070

    'Atualiza Pedido de CompraBaixado
    lErro = Comando_ExecutarPos(lComando(1), "UPDATE PedidoCompraN SET DataEmissao = ?", lComando(0), Date) 'OK não usar os controles da tela nas rotinas que irão subir.
    If lErro <> AD_SQL_SUCESSO Then gError 76071

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 76072

    'Fecha os comandos
    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next

    PedidoCompraBaixado_Atualiza_DataEmissao = SUCESSO

    Exit Function

Erro_PedidoCompraBaixado_Atualiza_DataEmissao:

    PedidoCompraBaixado_Atualiza_DataEmissao = gErr

    Select Case gErr

        Case 76065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 76067
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 76068, 76069
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr, objPedidoCompra.lCodigo)

        Case 76071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOCOMPRA", gErr, objPedidoCompra.lCodigo)

        Case 76072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 76070

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147547)

    End Select

     Call Transacao_Rollback

    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoCotacao_Atualiza_DataEmissao(ByVal objPedidoCotacao As ClassPedidoCotacao) As Long
'Atualiza data de emissao no BD para a data atual

Dim lErro As Long
Dim lComando(0 To 2) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lCodigo As Long
Dim iFilialEmpresa As Long

On Error GoTo Erro_PedidoCotacao_Atualiza_DataEmissao

    'Abre os  comandos
    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then gError 76127
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 76128

    'Busca na tabela de PedidoCotacao, o pedido de cotacao passado como parametro
    lErro = Comando_ExecutarPos(lComando(0), "SELECT NumIntDoc,Codigo,FilialEmpresa  FROM PedidoCotacao WHERE  NumIntDoc=?", 0, lNumIntDoc, lCodigo, iFilialEmpresa, objPedidoCotacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 76129

    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76130

    'se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 76131

    'Atualiza Pedido de Cotacao
    lErro = Comando_ExecutarPos(lComando(1), "UPDATE PedidoCotacao SET DataEmissao = ?", lComando(0), gdtDataHoje)
    If lErro <> AD_SQL_SUCESSO Then gError 76132

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 76133

    'Fecha os comandos
    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next

    PedidoCotacao_Atualiza_DataEmissao = SUCESSO

    Exit Function

Erro_PedidoCotacao_Atualiza_DataEmissao:

    PedidoCotacao_Atualiza_DataEmissao = Err

    Select Case gErr

        Case 76127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 76128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 76129, 76130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAO1", gErr)

        Case 76131
        
        Case 76132
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOCOTACAO", gErr, objPedidoCotacao.lCodigo)

        Case 76133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147548)

    End Select

     Call Transacao_Rollback

    For iIndice = LBound(lComando) To UBound(lComando)
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function PedidoCotacao_Grava_PedidoCompra(ByVal colPedidoCompra As Collection) As Long
'Grava os Pedidos de Compra gerados a partir do Pedido de Cotacao

Dim lErro As Long
Dim lTransacao As Long
Dim objPedidoCompra As ClassPedidoCompras
Dim objItemPC As New ClassItemPedCompra
Dim lCodigo As Long

On Error GoTo Erro_PedidoCotacao_Grava_PedidoCompra

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 76159
    
    'Para cada Pedido de Compra gerado
    For Each objPedidoCompra In colPedidoCompra
        
        lErro = CF("PedidoCompra_Codigo_Automatico_EmTrans", lCodigo)
        If lErro <> SUCESSO Then gError 76246
        
        objPedidoCompra.lCodigo = lCodigo
        
        'Grava o Pedido de Compra
        lErro = CF("PedidoCompra_Grava_EmTrans", objPedidoCompra)
        If lErro <> SUCESSO Then gError 76160
        
        For Each objItemPC In objPedidoCompra.colItens
            
            'Faz o link entre ItemRC e ItemPC para cada Item do Pedido de Compra gravado
            'e desfaz o link na tabela CotacaoProdutoItemRC
            lErro = CF("ItemPC_Associa_ItemRC", objItemPC)
            If lErro <> SUCESSO Then gError 76161
                        
        Next
        
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 76162
    
    'Para cada Pedido de compras da coleção
    For Each objPedidoCompra In colPedidoCompra

        lErro = CF("WorkFlow_Trata_Transacao", "COM", objPedidoCompra, gsUsuario, objPedidoCompra, 1002)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next

    PedidoCotacao_Grava_PedidoCompra = SUCESSO
    
    Exit Function
    
Erro_PedidoCotacao_Grava_PedidoCompra:

    PedidoCotacao_Grava_PedidoCompra = gErr
    
    Select Case gErr
    
        Case 76159
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 76160, 76161, 76193, 76246
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147549)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function
        
End Function

Function PedidoCompra_Grava_EmTrans(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Grava Pedido de Compra no BD

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 13) As Long
Dim lTransacao As Long
Dim iFilialEmpresa As Integer
Dim lCodigo As Long
Dim objPedidoCompra_BD As New ClassPedidoCompras
Dim iValorAlterado As Integer
Dim tPedCompra As typePedidoCompras
Dim objItemPC As ClassItemPedCompra
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim vbMsgBox As VbMsgBoxResult
Dim objFilFornFilEmp As New ClassFilFornFilEmp
Dim objNFiscal As ClassNFiscal, dtDataRegAprov As Date

On Error GoTo Erro_PedidoCompra_Grava_EmTrans

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 76163
    Next

    lErro = PedidoCompra_Valida_CodigoPV(objPedidoCompra)
    If lErro <> SUCESSO Then gError 178826

    'Pesquisa existencia de Pedido Baixado com mesmo código
    lErro = Comando_Executar(alComando(0), "SELECT FilialEmpresa, Codigo FROM PedidoCompraBaixado WHERE FilialEmpresa = ? AND Codigo = ?", iFilialEmpresa, lCodigo, objPedidoCompra.iFilialEmpresa, objPedidoCompra.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 76164

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76165

    'se o pedido de compra estiver baixado ==> erro
    If lErro = AD_SQL_SUCESSO Then gError 76166

    objPedidoCompra_BD.lCodigo = objPedidoCompra.lCodigo
    objPedidoCompra_BD.iFilialEmpresa = objPedidoCompra.iFilialEmpresa
        
    'Guarda dados do Pedido de Compras
    objFornecedorProdutoFF.iFilialEmpresa = objPedidoCompra.iFilialEmpresa
    objFornecedorProdutoFF.lFornecedor = objPedidoCompra.lFornecedor
    objFornecedorProdutoFF.iFilialForn = objPedidoCompra.iFilial
    
    lErro = CF("PedidoCompra_Valida_Campos_Cust", objPedidoCompra)
    If lErro <> SUCESSO Then gError 210943
    
    lErro = CF("PedidoCompra_Lock_Gravacao", objPedidoCompra)
    If lErro <> SUCESSO Then gError 76167

    lErro = CF("PedidoCompra_LeLock", objPedidoCompra_BD)
    If lErro <> SUCESSO And lErro <> 56259 Then gError 76168

    objPedidoCompra.sUsuReg = objPedidoCompra_BD.sUsuReg
    
    'Nao encontrou => é uma insercao
    If lErro = 56259 Then

        'Para cada Item de Pedido de Compra
        For Each objItemPC In objPedidoCompra.colItens
            
            objFornecedorProdutoFF.sProduto = objItemPC.sProduto
            
            'Procura FornecedorFilialProduto relacionado ao ItemPC
            lErro = CF("FornecedorProdutoFF_Le", objFornecedorProdutoFF)
            If lErro <> SUCESSO And lErro <> 54217 Then gError 76169
            
            'Se não encontrou, erro
            If lErro = 54217 Then gError 76170
            
            'Se o Lote Minimo de FornecedorProdutoFF estiver Preenchido, Verifica se a quantidade a comprar do item é maior do que o lote minimo
            If objFornecedorProdutoFF.dLoteMinimo > 0 And objFornecedorProdutoFF.dLoteMinimo > objItemPC.dQuantidade Then
                vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_LOTEMINIMO_MAIOR_QUANTCOMPRAR", objItemPC.sProduto, objItemPC.dQuantidade, objFornecedorProdutoFF.dLoteMinimo)
                If vbMsgBox = vbNo Then gError 76171
            End If
        
        Next
        
        'Verifica se DataEnvio foi preenchida
        If objPedidoCompra.dtDataEnvio <> DATA_NULA Then
        
            If gobjCOM.iPedCompraBloqEnvioSemAprov = MARCADO Then gError 213170
        
            'Verifica se há chance de ultrapassar o ESTOQUE MÁXIMO
            'Se hover, dá AVISO e pergunta se deseja seguir.
            lErro = PedidoCompra_Avalia_EstoqueMaximo(objPedidoCompra)
            If lErro <> SUCESSO Then gError 76173
        
            If gobjCOM.iControleAlcada <> 0 Then

                iValorAlterado = NAO_ALTERADO
                lErro = CF("BloqueioAlcada_Critica", objPedidoCompra, iValorAlterado)
                If lErro <> SUCESSO Then gError 76174
            End If
                

            objFilFornFilEmp.lCodFornecedor = objPedidoCompra.lFornecedor
            objFilFornFilEmp.iCodFilial = objPedidoCompra.iFilial
            objFilFornFilEmp.iFilialEmpresa = objPedidoCompra.iFilialEmpresa
            objFilFornFilEmp.lPedidosEmAberto = objFilFornFilEmp.lPedidosEmAberto + 1
    
            'Grava a Estatística da Filial Fornecedor
            lErro = CF("FilialFornecedor_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_FILIALFORNFILEMP, objFilFornFilEmp)
            If lErro <> SUCESSO Then gError 89428
            
            objPedidoCompra.bEnvio = True
            objPedidoCompra.sUsuRegEnvio = gsUsuario
        
        End If
        
        objPedidoCompra.bInclusao = True
        objPedidoCompra.sUsuReg = gsUsuario
        
        If objPedidoCompra.colBloqueiosPC.Count > 0 Then
            Call Rotina_Aviso(vbOKOnly, "AVISO_PEDIDOCOMPRA_BLOQUEADO", objPedidoCompra.lCodigo)
            objPedidoCompra.dtDataEnvio = DATA_NULA
            objPedidoCompra.bEnvio = False
        End If
        
        lErro = CF("PedidoCompra_Insere", objPedidoCompra)
        If lErro <> SUCESSO Then gError 76172

        
        lErro = CF("BloqueiosPC_Insere", objPedidoCompra)
        If lErro <> SUCESSO Then gError 76175
        
    Else 'é uma alteracao

        objPedidoCompra.dtDataAlteracao = gdtDataHoje

        objPedidoCompra.lNumIntDoc = objPedidoCompra_BD.lNumIntDoc

        If objPedidoCompra_BD.dtDataEnvio <> DATA_NULA Then gError 76176

        lErro = CF("ItensPedCompra_Altera", objPedidoCompra_BD, objPedidoCompra)
        If lErro <> SUCESSO Then gError 76177
        
        tPedCompra.sContato = String(STRING_CONTATO, 0)
        tPedCompra.sTipoFrete = String(STRING_TIPOFRETE, 0)
        tPedCompra.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
        tPedCompra.sAlcada = String(STRING_ALCADA_CODUSUARIO, 0)
    
        'Le o pedido de compra passado como parametro
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc,FilialEmpresa,Codigo,Fornecedor,Filial,Comprador,Contato,TipoDestino,FornCliDestino,FilialDestino,Data,DataEmissao,DataEnvio,DataAlteracao,CondicaoPagto,OutrasDespesas,ValorFrete,ValorSeguro,ValorDesconto,ValorTotal,ValorIPI,TipoFrete,Transportadora,ProxSeqBloqueio,TipoBaixa,MotivoBaixa,Alcada,Observacao, DataRegAprov FROM PedidoCompra WHERE Codigo = ? AND FilialEmpresa=?", 0, tPedCompra.lNumIntDoc, tPedCompra.iFilialEmpresa, tPedCompra.lCodigo, tPedCompra.lFornecedor, tPedCompra.iFilial, tPedCompra.iComprador, tPedCompra.sContato, tPedCompra.iTipoDestino, tPedCompra.lFornCliDestino, tPedCompra.iFilialDestino, tPedCompra.dtData, tPedCompra.dtDataEmissao, tPedCompra.dtDataEnvio, tPedCompra.dtDataAlteracao, tPedCompra.iCondicaoPagto, tPedCompra.dOutrasDespesas, tPedCompra.dValorFrete, tPedCompra.dValorSeguro, tPedCompra.dValorDesconto, tPedCompra.dValorTotal, tPedCompra.dValorIPI, _
            tPedCompra.sTipoFrete, tPedCompra.iTransportadora, tPedCompra.iProxSeqBloqueio, tPedCompra.iTipoBaixa, tPedCompra.sMotivoBaixa, tPedCompra.sAlcada, tPedCompra.lObservacao, dtDataRegAprov, objPedidoCompra_BD.lCodigo, objPedidoCompra_BD.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 76178
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76179
            
        If objPedidoCompra.dtDataEnvio <> DATA_NULA Then

            If dtDataRegAprov = DATA_NULA Then
                If gobjCOM.iPedCompraBloqEnvioSemAprov = MARCADO Then gError 213170
            End If
            
            If (Len(Trim(objPedidoCompra_BD.sAlcada)) <> 0 And (objPedidoCompra.dValorTotal > objPedidoCompra_BD.dValorTotal)) Or Len(Trim(objPedidoCompra_BD.sAlcada)) = 0 And gobjCOM.iControleAlcada <> 0 Then

                iValorAlterado = ALTERADO
                lErro = CF("BloqueioAlcada_Critica", objPedidoCompra, iValorAlterado)
                If lErro <> SUCESSO Then gError 76180

            End If
        
            If objPedidoCompra.colBloqueiosPC.Count > 0 Then
                Call Rotina_Aviso(vbOKOnly, "AVISO_PEDIDOCOMPRA_BLOQUEADO", objPedidoCompra.lCodigo)
                objPedidoCompra.dtDataEnvio = DATA_NULA
            Else
                lErro = VerificaLiberacao_BloqueioPC(objPedidoCompra) 'William
                If lErro <> SUCESSO Then gError 87547
            End If
            
            objFilFornFilEmp.lCodFornecedor = objPedidoCompra.lFornecedor
            objFilFornFilEmp.iCodFilial = objPedidoCompra.iFilial
            objFilFornFilEmp.iFilialEmpresa = objPedidoCompra.iFilialEmpresa
            objFilFornFilEmp.lPedidosEmAberto = objFilFornFilEmp.lPedidosEmAberto + 1
    
            'Grava a Estatística da Filial Fornecedor
            lErro = CF("FilialFornecedor_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_FILIALFORNFILEMP, objFilFornFilEmp)
            If lErro <> SUCESSO Then gError 89429
            
            objPedidoCompra.bEnvio = True
            objPedidoCompra.sUsuRegEnvio = gsUsuario
            
        End If
        
        lErro = CF("BloqueiosPC_Insere", objPedidoCompra)
        If lErro <> SUCESSO Then gError 76181
        
        lErro = NotasPedCompra_Exclui(objPedidoCompra)
        If lErro <> SUCESSO Then gError 103341
'Leo
        'Altera o Pedido de Compra
        If objPedidoCompra.bEnvio Then
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE PedidoCompra SET CotacaoMoeda=?, Embalagem=?, Taxa=?, Moeda=?, Fornecedor=?, Filial=?, Comprador=?, Contato=?, TipoDestino=?, FornCliDestino=?, FilialDestino=?, DataEmissao=?, DataEnvio=?, DataAlteracao=?, CondicaoPagto=?, OutrasDespesas=?,ValorProdutos = ?, ValorFrete=?, ValorSeguro=?, ValorDesconto=?, ValorTotal=?, ValorIPI=?, TipoFrete=?, Transportadora=?, Alcada=?, ObsEmbalagem=?, Observacao=?, CodigoPV=?, Data = ?, DataRefFluxo = ?, TabelaPreco = ?, DataRegEnvio = ?, HoraRegEnvio = ?, UsuRegEnvio = ?", _
                alComando(1), objPedidoCompra.dCotacaoMoeda, objPedidoCompra.iEmbalagem, objPedidoCompra.dTaxa, objPedidoCompra.iMoeda, objPedidoCompra.lFornecedor, objPedidoCompra.iFilial, objPedidoCompra.iComprador, objPedidoCompra.sContato, objPedidoCompra.iTipoDestino, objPedidoCompra.lFornCliDestino, objPedidoCompra.iFilialDestino, objPedidoCompra.dtDataEmissao, objPedidoCompra.dtDataEnvio, objPedidoCompra.dtDataAlteracao, objPedidoCompra.iCondicaoPagto, objPedidoCompra.dOutrasDespesas, objPedidoCompra.dValorProdutos, objPedidoCompra.dValorFrete, objPedidoCompra.dValorSeguro, objPedidoCompra.dValorDesconto, objPedidoCompra.dValorTotal, objPedidoCompra.dValorIPI, objPedidoCompra.sTipoFrete, objPedidoCompra.iTransportadora, objPedidoCompra.sAlcada, objPedidoCompra.sObsEmbalagem, objPedidoCompra.lObservacao, objPedidoCompra.lCodigoPV, objPedidoCompra.dtData, objPedidoCompra.dtDataRefFluxo, objPedidoCompra.iTabelaPreco, Date, CDbl(Time), gsUsuario)
            If lErro <> AD_SQL_SUCESSO Then gError 76182
        Else
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE PedidoCompra SET CotacaoMoeda=?, Embalagem=?, Taxa=?, Moeda=?, Fornecedor=?, Filial=?, Comprador=?, Contato=?, TipoDestino=?, FornCliDestino=?, FilialDestino=?, DataEmissao=?, DataEnvio=?, DataAlteracao=?, CondicaoPagto=?, OutrasDespesas=?,ValorProdutos = ?, ValorFrete=?, ValorSeguro=?, ValorDesconto=?, ValorTotal=?, ValorIPI=?, TipoFrete=?, Transportadora=?, Alcada=?, ObsEmbalagem=?, Observacao=?, CodigoPV=?, Data = ?, DataRefFluxo = ?, TabelaPreco = ?", _
                alComando(1), objPedidoCompra.dCotacaoMoeda, objPedidoCompra.iEmbalagem, objPedidoCompra.dTaxa, objPedidoCompra.iMoeda, objPedidoCompra.lFornecedor, objPedidoCompra.iFilial, objPedidoCompra.iComprador, objPedidoCompra.sContato, objPedidoCompra.iTipoDestino, objPedidoCompra.lFornCliDestino, objPedidoCompra.iFilialDestino, objPedidoCompra.dtDataEmissao, objPedidoCompra.dtDataEnvio, objPedidoCompra.dtDataAlteracao, objPedidoCompra.iCondicaoPagto, objPedidoCompra.dOutrasDespesas, objPedidoCompra.dValorProdutos, objPedidoCompra.dValorFrete, objPedidoCompra.dValorSeguro, objPedidoCompra.dValorDesconto, objPedidoCompra.dValorTotal, objPedidoCompra.dValorIPI, objPedidoCompra.sTipoFrete, objPedidoCompra.iTransportadora, objPedidoCompra.sAlcada, objPedidoCompra.sObsEmbalagem, objPedidoCompra.lObservacao, objPedidoCompra.lCodigoPV, objPedidoCompra.dtData, objPedidoCompra.dtDataRefFluxo, objPedidoCompra.iTabelaPreco)
            If lErro <> AD_SQL_SUCESSO Then gError 76182
        End If
        
        lErro = ItensPCEntrega_Exclui(alComando(11), objPedidoCompra.lNumIntDoc, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    lErro = NotasPedCompra_Insere(objPedidoCompra)
    If lErro <> SUCESSO Then gError 103340
    
    objPedidoCompra.objInfoAdic.lNumIntDoc = objPedidoCompra.lNumIntDoc
    objPedidoCompra.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_PC
    lErro = CF("InfoAdicionais_Grava", objPedidoCompra.objInfoAdic)
    If lErro <> SUCESSO Then gError 103340
    
    lErro = CF("ProjetoInfo_Grava", objPedidoCompra.objProjetoInfo, PRJ_CR_TIPO_PC, objPedidoCompra.lNumIntDoc)
    If lErro <> SUCESSO Then gError 103340
    
    For Each objItemPC In objPedidoCompra.colItens
        lErro = ItensPCEntrega_Insere(alComando(12), alComando(13), objItemPC)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Next

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PedidoCompra_Grava_EmTrans = SUCESSO

    Exit Function

Erro_PedidoCompra_Grava_EmTrans:

    PedidoCompra_Grava_EmTrans = gErr

    Select Case gErr

        Case 76163
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 76167, 76168, 76169, 76171, 76172, 76173, 76174, 76175, 76177, 76180, 76181, 87547, 89428, 89429, 103340, 103341, 178826, 210943

        Case 76164, 76165
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA_BAIXADO", gErr, objPedidoCompra.lCodigo)

        Case 76166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_BAIXADO", gErr, objPedidoCompra.lCodigo)

        Case 76170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_NAO_CADASTRADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
                            
        Case 76176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_ENVIADO", gErr, objPedidoCompra_BD.lCodigo)

        Case 76182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDIDOCOMPRA", gErr, objPedidoCompra.lCodigo)

        Case 76178, 76179
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
            
        Case 131647
        
        Case 213170
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_APROVADO", gErr, objPedidoCompra_BD.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147550)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'??? parametro opcional nao utilizado. Jones 2002
Function ItemPC_Associa_ItemRC(ByVal objItemPC As ClassItemPedCompra, Optional ByVal lPedCotacao As Long) As Long
'Faz o link entre ItemRC e ItemPC e atualiza QuantPedida de ItemRC e
'desfaz o link na tabela CotacaoProdutoItemRC

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 5) As Long
Dim objItemRC As New ClassItemReqCompras
Dim objItemRCItemPC As New ClassItemRCItemPC
Dim dQuantidade  As Double
Dim lNumIntDoc As Long
Dim dQuantPedida As Double
Dim lItemReqCompra As Long
Dim lCotacao As Long
Dim dQuantCancel As Double
Dim dQuantReceb As Double
Dim dQuantItemRC As Double

On Error GoTo Erro_ItemPC_Associa_ItemRC

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 76183
    Next
    
    For Each objItemRC In objItemPC.colItemReqCompras
       
        objItemRCItemPC.lItemRC = objItemRC.lNumIntDoc
        objItemRCItemPC.lItemPC = objItemPC.lNumIntDoc
        objItemRCItemPC.dQuantidade = objItemRC.dQuantComprar
        
        'Faz o Link entre ItemPedCompra e ItemReqCompra
        lErro = Comando_Executar(alComando(1), "INSERT INTO ItemRCItemPC (ItemPC,ItemRC,Quantidade,QuantRecebida) VALUES (?,?,?,?)", objItemRCItemPC.lItemPC, objItemRCItemPC.lItemRC, objItemRCItemPC.dQuantidade, objItemRCItemPC.dQuantRecebida)
        If lErro <> AD_SQL_SUCESSO Then gError 76184
                
        dQuantidade = objItemRC.dQuantComprar
                
        'Busca o ItemReqCompra no BD
        lErro = Comando_ExecutarPos(alComando(5), "SELECT NumIntDoc,QuantPedida, Quantidade,QuantCancelada,QuantRecebida FROM ItensReqCompra WHERE NumIntDoc=?", 0, lNumIntDoc, dQuantPedida, dQuantItemRC, dQuantCancel, dQuantReceb, objItemRC.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 76188
        
        lErro = Comando_BuscarPrimeiro(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76189
        
        'Se não encontrou o ItemReqCompra ==>erro
        If lErro = AD_SQL_SEM_DADOS Then gError 76192
        
        lErro = Comando_LockExclusive(alComando(5))
        If lErro <> AD_SQL_SUCESSO Then Error 76194
               
        If (dQuantItemRC - dQuantidade - dQuantCancel - dQuantReceb - dQuantPedida) < 0 Then gError 62819
        
        'Atualiza QuantPedida do ItemReqCompra
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensReqCompra SET QuantPedida = QuantPedida  + ?", alComando(5), objItemRC.dQuantComprar)
        If lErro <> AD_SQL_SUCESSO Then gError 76190
            
        'Desvincula as cotações dos itens de requisições envolvidos
        'na geração de pedidos de compra.
           
        'Busca na tabela CotacaoProdutoItemRC
        lErro = Comando_ExecutarPos(alComando(3), "SELECT CotacaoProduto,ItemReqCompra FROM CotacaoProdutoItemRC WHERE ItemReqCompra=?", 0, lCotacao, lItemReqCompra, objItemRC.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 76197
    
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76198
        
        Do While lErro = AD_SQL_SUCESSO
        
            'Desfaz o link de CotacaoProdutoItemRC
            lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM CotacaoProdutoItemRC", alComando(3))
            If lErro <> AD_SQL_SUCESSO Then gError 76200
        
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62820
            
        Loop
    Next
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItemPC_Associa_ItemRC = SUCESSO
    
    Exit Function
    
Erro_ItemPC_Associa_ItemRC:

    ItemPC_Associa_ItemRC = gErr

    Select Case gErr
    
        Case 62819
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTCOMPRAR_MAIOR_QUANTCOMPRARMAX", gErr, dQuantidade, dQuantItemRC - dQuantCancel - dQuantReceb - dQuantPedida)
        
        Case 76183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMRCITEMPC", gErr, objItemRCItemPC.lItemPC, objItemRCItemPC.lItemRC)
            
        Case 76186, 76187
        
           
        Case 76188, 76189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr)
            
        Case 76190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSREQCOMPRA", gErr, objItemRC.lNumIntDoc)
            
        Case 76192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENSREQCOMPRA_NAO_CADASTRADO2", gErr, objItemRC.lNumIntDoc)
                    
         Case 76194
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSREQCOMPRA1", gErr, objItemRC.lNumIntDoc)
        
        Case 76196, 76199
            
        Case 76197, 76198, 62820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACAOPRODUTOITEMRC", gErr)
            
        Case 76200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COTACAOPRODUTOITEMRC", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147551)
            
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
     
    Exit Function
    
End Function

Function Concorrencia_Exclui(ByVal objConcorrencia As ClassConcorrencia) As Long
'Exclui a Concorrência passada

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Concorrencia_Exclui
            
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 66918
    
    lErro = Concorrencia_Exclui_EmTrans(objConcorrencia)
    If lErro <> SUCESSO Then gError 66920
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 66919
           
    Concorrencia_Exclui = SUCESSO
    
    Exit Function

Erro_Concorrencia_Exclui:

    Concorrencia_Exclui = gErr
    
    Select Case gErr
    
        Case 66918
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 66919
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
    
        Case 66920
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147552)

    End Select
        
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function Concorrencia_Exclui_EmTrans(ByVal objConcorrencia As ClassConcorrencia) As Long
'Exclui a Concorrência passada

Dim lErro As Long
Dim alComando(3) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_Concorrencia_Exclui_EmTrans

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66917
    Next
            
    'Busca no BD a Concorrência Passada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM Concorrencia WHERE Codigo = ? AND FilialEmpresa = ?", 0, lNumIntDoc, objConcorrencia.lCodigo, objConcorrencia.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 66920
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66921
    
    'Se não encontrou a Concorrência, Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66922
    
    'Faz "lockExclusive" na Concorrência
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 66923
    
    'Guarda o NumIntDoc da Concorrência
    objConcorrencia.lNumIntDoc = lNumIntDoc
    
    'Exclui os Itens da Concorrência
    lErro = CF("ItensConcorrencia_Exclui", objConcorrencia)
    If lErro <> SUCESSO Then gError 66924
    
    'Exclui a Concorrência
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Concorrencia", alComando(0))
    If lErro <> SUCESSO Then gError 66925
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Concorrencia_Exclui_EmTrans = SUCESSO
    
    Exit Function

Erro_Concorrencia_Exclui_EmTrans:

    Concorrencia_Exclui_EmTrans = gErr
    
    Select Case gErr
    
        Case 66917
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                       
        Case 66920, 66921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", gErr)
            
        Case 66922
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONCORRENCIA_NAO_CADASTRADA", gErr, objConcorrencia.lCodigo)
                    
        Case 66923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONCORRECIA", gErr, objConcorrencia.lCodigo)
                
        Case 66924
        
        Case 66925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONCORRENCIA", gErr, objConcorrencia.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147553)

    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function ItemReqCompras_Verifica_Vinculo(ByVal lItemReqCompra As Long, bExclui As Boolean) As Long
'Verifica para o Item RC passado se ele está vinculado opr outra concorrência não baixada
'Se tiver e o usuário escolher continuar nessa gravação a concorrência anterior será excluída

Dim lErro As Long
Dim lCodigo As Long
Dim lComando As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_ItemReqCompras_Verifica_Vinculo

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 62742
    
    'Verifica se exite concorrência não baixada que utiliza o item RC
    lErro = Comando_Executar(lComando, "SELECT Concorrencia.Codigo FROM Concorrencia, ItensConcorrencia, ItemRCItemConcorrencia WHERE Concorrencia.NumIntDoc = ItensConcorrencia.Concorrencia AND ItensConcorrencia.NumIntDoc = ItemRCItemConcorrencia.ItemConcorrencia AND ItemReqCompra = ?", lCodigo, lItemReqCompra)
    If lErro <> AD_SQL_SUCESSO Then gError 62743
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62769
    
    'Se encontrar a concorrência
    If lErro = AD_SQL_SUCESSO Then
        'Se a pergunta ainda não foi feita
        If Not bExclui Then
            'Pergunta se deseja prosseguir na gravação implicando na exclusão da concorrência anterior
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ITEMRC_PARTICIPA_CONCORRENCIA", lCodigo)
            If vbMsgRes = vbYes Then
                'Inicaliza flag de exclusão
                bExclui = True
            Else
                gError 62797
            End If
        End If
        'Se o usuário quer excluir a concorrência anterior
        If bExclui Then
            
            objConcorrencia.lCodigo = lCodigo
            objConcorrencia.iFilialEmpresa = giFilialEmpresa
            'Exclui a concorrência antiga do BD
            lErro = Concorrencia_Exclui_EmTrans(objConcorrencia)
            If lErro <> SUCESSO Then gError 62770
                        
        End If
    End If

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ItemReqCompras_Verifica_Vinculo = SUCESSO
    
    Exit Function
    
Erro_ItemReqCompras_Verifica_Vinculo:

    ItemReqCompras_Verifica_Vinculo = gErr
    
    Select Case gErr
    
        Case 62742
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 62743, 62769
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", gErr)
        
        Case 62770, 62797
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147554)
            
    End Select
            
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RequisicaoCompra_Desvincula_PedidoVenda(ByVal objReqCompra As ClassRequisicaoCompras) As Long
'Quebra o vínculo entre a Requisição de Compra com Numero Interno passado e os PV que a geraram

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim dQuantidade As Double
Dim objItemRC As ClassItemReqCompras
Dim lNumIntItemPV As Long
Dim dQuantSC As Double

On Error GoTo Erro_RequisicaoCompra_Desvincula_PedidoVenda

    'Abre os comandos
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 86131
    Next
    
    For Each objItemRC In objReqCompra.colItens
        
        'Busca um relacionamento com Pedido de Venda
        lErro = Comando_ExecutarPos(alComando(0), "SELECT ItemPV, Quantidade FROM ItemPVItemRC WHERE ItemRC = ?", 0, lNumIntItemPV, dQuantidade, objItemRC.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 86132
        
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86133
    
        'Enquanto encontrar ligação
        Do While lErro <> AD_SQL_SEM_DADOS
        
            'Lê o Item de Pedido de venda relacionado
            lErro = Comando_ExecutarPos(alComando(1), "SELECT QuantSC FROM ItensPedidoDeVenda WHERE NumIntDoc = ?", 0, dQuantSC, lNumIntItemPV)
            If lErro <> AD_SQL_SUCESSO Then gError 86134
            
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86139
            
            'Se não encontrou na tabela de não baixados
            If lErro = AD_SQL_SEM_DADOS Then
                
                'BUsca nos itens baixados o Item relacionado
                lErro = Comando_ExecutarPos(alComando(1), "SELECT QuantSC FROM ItensPedidoDeVendaBaixados WHERE NumIntDoc = ?", 0, dQuantSC, lNumIntItemPV)
                If lErro <> AD_SQL_SUCESSO Then gError 86140
                
                lErro = Comando_BuscarPrimeiro(alComando(1))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86145
                If lErro = AD_SQL_SUCESSO Then
                
                    dQuantSC = dQuantSC - dQuantidade
                    If dQuantSC < 0 Then dQuantSC = 0
                    
                    lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensPedidoDeVendaBaixados SET QuantSC = ?", alComando(1), dQuantSC)
                    If lErro <> AD_SQL_SUCESSO Then gError 86141
                End If
            Else
                
                dQuantSC = dQuantSC - dQuantidade
                If dQuantSC < 0 Then dQuantSC = 0
                
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE ItensPedidoDeVenda set QuantSC = ?", dQuantSC)
                If lErro <> AD_SQL_SUCESSO Then gError 86135
                
            End If
               
            lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM ItemPVItemRC ", alComando(0))
            If lErro <> AD_SQL_SUCESSO Then gError 86136
               
        Loop
    
    Next
            
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RequisicaoCompra_Desvincula_PedidoVenda = SUCESSO
    
    Exit Function
    
Erro_RequisicaoCompra_Desvincula_PedidoVenda:

    RequisicaoCompra_Desvincula_PedidoVenda = gErr

    Select Case gErr

        Case 86130, 86132, 86133, 86135, 86136, 86137
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", gErr)

        Case 86131
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 86134, 86138
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PEDVENDA", gErr, objReqCompra.lCodigo)

        Case 86139, 86140, 86141, 86142, 86143, 86144
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)
        
        Case 86145, 86146
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147555)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Sub Transfere_Dados_ItemPC(ByVal objItemPC As ClassItemPedCompra, tItemPedido As typeItemPedCompra)

    With tItemPedido
    
        objItemPC.dAliquotaICMS = .dAliquotaICMS
        objItemPC.dAliquotaIPI = .dAliquotaIPI
        objItemPC.dPercentMaisReceb = .dPercentMaisReceb
        objItemPC.dPercentMenosReceb = .dPercentMenosReceb
        objItemPC.dPrecoUnitario = .dPrecoUnitario
        objItemPC.dQuantidade = .dQuantidade
        objItemPC.dQuantRecebida = .dQuantRecebida
        objItemPC.dQuantRecebimento = .dQuantRecebimento
        objItemPC.dtDataLimite = .dtDataLimite
        objItemPC.dValorDesconto = .dValorDesconto
        objItemPC.dValorIPI = .dValorIPI
        objItemPC.iRebebForaFaixa = .iRebebForaFaixa
        objItemPC.iStatus = .iStatus
        objItemPC.iTipoOrigem = .iTipoOrigem
        objItemPC.lNumIntDoc = .lNumIntDoc
        objItemPC.lNumIntOrigem = .lNumIntOrigem
        objItemPC.lObservacao = .lObservacao
        objItemPC.sDescProduto = .sDescProduto
        objItemPC.sObservacao = .sObservacao
        objItemPC.sProduto = .sProduto
        objItemPC.sUM = .sUM
        
    End With
    
    Exit Sub

End Sub

'William - 11/04/2001
Private Function VerificaLiberacao_BloqueioPC(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Função que verifica se existem bloqueios que impeçam o envio do Pedido de Compras

Dim lErro As Long
Dim lComando As Long
Dim dtDataLib As Date

On Error GoTo Erro_VerificaLiberacao_BloqueioPC

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87542

    'Faz leitura na tabela Bloqueios PC
        '==> chaves FilialEmpresa, PedCompras, TipoBloqueio
    lErro = Comando_Executar(lComando, "SELECT DataLib FROM BloqueiosPC WHERE FilialEmpresa = ? AND PedCompras = ? AND TipoDeBloqueio <> ? AND DataLib = ?", dtDataLib, objPedidoCompra.iFilialEmpresa, objPedidoCompra.lCodigo, BLOQUEIO_ALCADA, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 87543

    'Posiciona ponteiro no primeiro registro que satisfaz a condição
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87544
    
    If lErro = AD_SQL_SUCESSO Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_PEDIDOCOMPRA_BLOQUEADO", objPedidoCompra.lCodigo)
        objPedidoCompra.dtDataEnvio = DATA_NULA
    End If
    
    'gError 87546
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    VerificaLiberacao_BloqueioPC = SUCESSO

    Exit Function

Erro_VerificaLiberacao_BloqueioPC:

    VerificaLiberacao_BloqueioPC = gErr
    
    Select Case gErr
    
        Case 87542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 87543, 87544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BLOQUEIOSPC1", gErr, objPedidoCompra.lCodigo)
            
        Case 87546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BLOQUEADO", gErr, objPedidoCompra.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147556)
            
    End Select

    'Fecha comando ==> saída por erro
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function ItemPVItemRC_Inclui(ByVal objItemReqCompra As ClassItemReqCompras) As Long

Dim lErro As Long
Dim objItemPV As ClassItemPedido
Dim alComando(1 To 3) As Long
Dim lCodigo As Long
Dim iIndice As Integer

On Error GoTo Erro_ItemPVItemRC_Inclui

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 86179
    Next

    For Each objItemPV In objItemReqCompra.colItemPV
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO ItemPVItemRC (ItemPV,ItemRC,QUantidade) VALUES (?,?,?)", objItemPV.lNumIntDoc, objItemReqCompra.lNumIntDoc, objItemPV.dQuantSC)
        If lErro <> AD_SQL_SUCESSO Then gError 86180
    
        'Busca no Bd o Item do Pedido de Venda cujo código e produto são informados
        lErro = Comando_ExecutarPos(alComando(2), "SELECT CodPedido FROM ItensPedidoDeVenda WHERE NumIntDoc=?", 0, lCodigo, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 178620

        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178621

        'Se não encontrou o Item de Pedido de Venda ==> erro
        If lErro = AD_SQL_SEM_DADOS Then gError 178622

        'Altera o campo QuantSC na tabela ItensPedidoDeVenda
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensPedidoDeVenda SET QuantSC=QuantSC + ? ", alComando(2), objItemPV.dQuantSC)
        If lErro <> AD_SQL_SUCESSO Then gError 178623

    Next

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItemPVItemRC_Inclui = SUCESSO
    
    Exit Function
    
Erro_ItemPVItemRC_Inclui:

    ItemPVItemRC_Inclui = gErr
    
    Select Case gErr
    
        Case 86179
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 86180
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMPVITEMRC", gErr)
        
        Case 178620, 178621
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case 178622
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPV_NAO_ENCONTRADO", gErr, objItemPV.lNumIntDoc)

        Case 178623
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147557)
    
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function
    
End Function

Function NotasPC_Exclui(objNotasPC As ClassNotaPC) As Long
'Exclui a NotaPC da tabela

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim lTransacao As Long
Dim lCodigo As Long
Dim iIndice As Integer

On Error GoTo Erro_NotasPC_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103320
    Next

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 103321

    'Seleciona o comprador no bd
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM NotasPC WHERE Codigo = ?", 0, lCodigo, objNotasPC.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103322

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 103323

    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 103324

    'Exclui a Nota do BD
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM NotasPC", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 103325

   'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 103326

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NotasPC_Exclui = SUCESSO

    Exit Function

Erro_NotasPC_Exclui:

    NotasPC_Exclui = gErr

    Select Case gErr

        Case 103320
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 103322, 103323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTASPC", gErr)

        Case 103324
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NOTASPC", gErr)

        Case 103325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NOTASPC", gErr, objNotasPC.lCodigo)

        Case 103326
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147558)

    End Select

    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NotasPC_Grava(ByVal objNotasPC As ClassNotaPC) As Long
'Grava uma NotasPC

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 2) As Long
Dim lTransacao As Long
Dim lCodigo As Long

On Error GoTo Erro_NotasPC_Grava

    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 103312
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 103313

    'verifica se a Nota está cadastrada
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM NotasPC WHERE Codigo = ?", 0, lCodigo, objNotasPC.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103314

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103315

    'se a nota estiver cadastrada ==> é uma alteração
    If lErro = AD_SQL_SUCESSO Then

        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 103316

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE NotasPC SET Nota = ?", alComando(0), objNotasPC.sNota)
        If lErro <> AD_SQL_SUCESSO Then gError 103317

    Else

        lErro = Comando_Executar(alComando(0), "INSERT INTO NotasPC(Codigo, Nota) VALUES (?,?)", objNotasPC.lCodigo, objNotasPC.sNota)
        If lErro <> AD_SQL_SUCESSO Then gError 103318

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 103319

    'Fecha os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NotasPC_Grava = SUCESSO

    Exit Function

Erro_NotasPC_Grava:

    NotasPC_Grava = gErr

    Select Case gErr

        Case 103312
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 103314, 103315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTASPC", gErr, objNotasPC.lCodigo)

        Case 103316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ALCADA", gErr, objNotasPC.lCodigo)

        Case 103317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NOTASPC", gErr, objNotasPC.lCodigo)

        Case 103318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NOTASPC", gErr, objNotasPC.lCodigo)

        Case 103319
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147559)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NotasPedCompra_Insere(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Insere as Notas de um pedido de compras

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim sNotas As String

On Error GoTo Erro_NotasPedCompra_Insere

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103332

    For iIndice = 1 To objPedidoCompra.colNotasPedCompras.Count
        
        sNotas = objPedidoCompra.colNotasPedCompras.Item(iIndice)
        
        lErro = Comando_Executar(lComando, "INSERT INTO NotasPedCompras(NumIntPedido, Sequencial, NotaPC) VALUES (?,?,?)", _
        objPedidoCompra.lNumIntDoc, iIndice, sNotas)
        
        If lErro <> AD_SQL_SUCESSO Then gError 103333
        
    Next
    
    NotasPedCompra_Insere = SUCESSO

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    Exit Function

Erro_NotasPedCompra_Insere:

    NotasPedCompra_Insere = gErr

    Select Case gErr

        Case 103332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NOTASPEDCOMPRA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147560)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NotasPedCompra_Exclui(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Exclui as Notas de um pedido de compras

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim iSequencial As Integer

On Error GoTo Erro_NotasPedCompra_Exclui

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103334

    'Abre o comando
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 103335

    lErro = Comando_ExecutarPos(lComando, "SELECT Sequencial FROM NotasPedCompras WHERE NumIntPedido = ?", _
    0, iSequencial, objPedidoCompra.lNumIntDoc)
    
    If lErro <> AD_SQL_SUCESSO Then gError 103336
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103337
            
    Do While lErro = AD_SQL_SUCESSO
            
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM NotasPedCompras", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 103338
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103339
        
    Loop
    
    NotasPedCompra_Exclui = SUCESSO

    'Fecha o comando
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando2)
    Exit Function

Erro_NotasPedCompra_Exclui:

    NotasPedCompra_Exclui = gErr

    Select Case gErr

        Case 103334, 103335
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103338
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NOTASPEDCOMPRA", gErr)

        Case 103336, 103337, 103339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTASPEDCOMPRA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147561)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function



Function MapaCotacao_ObterNumIntAutomatico(lNumInt As Long) As Long
'gera o próximo número interno disponível para MapaCotacao

Dim lErro As Long
Dim objConcorrencia As New ClassConcorrencia

On Error GoTo Erro_MapaCotacao_ObterNumIntAutomatico

    'Gera o proximo numero interno para ItemConcorrencia
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_NUMINTDOC_MAPACOTACAO", lNumInt)
    If lErro <> SUCESSO Then gError 114592

    MapaCotacao_ObterNumIntAutomatico = SUCESSO

    Exit Function

Erro_MapaCotacao_ObterNumIntAutomatico:

    MapaCotacao_ObterNumIntAutomatico = gErr

    Select Case gErr

        Case 114592
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147562)

    End Select

End Function

Function MapaCotacaoItem_ObterNumIntAutomatico(lNumInt As Long) As Long
'gera o próximo número interno disponível para Item de Mapa de Cotacao

Dim lErro As Long
Dim objItemConcorrencia As New ClassItemConcorrencia

On Error GoTo Erro_MapaCotacaoItem_ObterNumIntAutomatico

    'Gera o proximo numero interno para ItemConcorrencia
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROXIMO_NUMINTDOC_ITEMMAPACOTACAO", lNumInt)
    If lErro <> SUCESSO Then gError 114593

    MapaCotacaoItem_ObterNumIntAutomatico = SUCESSO

    Exit Function

Erro_MapaCotacaoItem_ObterNumIntAutomatico:

    MapaCotacaoItem_ObterNumIntAutomatico = gErr

    Select Case gErr

        Case 114593
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147563)

    End Select

End Function

Function MapaCotacao_Exclui(ByVal objMapaCotacao As ClassMapaCotacao) As Long
'Exclui o mapa de cotacao e seus itens

Dim lErro As Long
Dim iIndice As Integer
Dim lCodigo As Long
Dim lTransacao As Long
Dim alComando(3) As Long
Dim objItemMapaCotacao As New ClassMapaCotacaoItemCotacao

On Error GoTo Erro_MapaCotacao_Exclui

    'Abertura da transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 114595

    'Abertura dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 114596
    Next

    'Busca os itens para exclusao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodMapaCotacao FROM MapaCotacaoItemCotacao WHERE CodMapaCotacao = ? ", 0, lCodigo, objMapaCotacao.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 114599

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 114610

    Do While lErro = AD_SQL_SUCESSO
        
        'Exclui do BD
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM MapaCotacaoItemCotacao", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 114612
        
        'Busca o próximo item a ser excluído
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 114613
            
    Loop

    'Busca o Mapa de Cotacao passado por parametro
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Codigo FROM MapaCotacao WHERE Codigo = ?", 0, lCodigo, objMapaCotacao.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 114614
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then gError 114610

    'Se nao encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 114615
    
    'Exclui do BD
    lErro = Comando_ExecutarPos(alComando(3), "DELETE FROM MapaCotacao", alComando(2))
    If lErro <> AD_SQL_SUCESSO Then gError 114616

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 114600

    MapaCotacao_Exclui = SUCESSO

    Exit Function

Erro_MapaCotacao_Exclui:

    MapaCotacao_Exclui = gErr

    Select Case gErr

        Case 114595
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 114596
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 114599, 114610, 114613
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAPACOTACAOITEMCOTACAO", gErr)

        Case 114614
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MAPACOTACAO", gErr)
            
        Case 114615
            Call Rotina_Erro(vbOKOnly, "ERRO_MAPACOTACAO_INEXISTENTE", gErr, objMapaCotacao.lCodigo)
            
        Case 114616, 114612
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MAPACOTACAO", gErr)
            
        Case 114600
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147564)

    End Select

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

End Function

Function MapaCotacao_Grava(ByVal objMapaCotacao As ClassMapaCotacao) As Long
'Grava o mapa de cotacao e seus itens

Dim lErro As Long
Dim sSelect As String
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long
Dim alComando(1) As Long
Dim objItemMapaCotacao As New ClassMapaCotacaoItemCotacao

On Error GoTo Erro_MapaCotacao_Grava

    'Abertura da transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 114595

    'Abertura dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 114596
    Next

    'Gera NumIntDoc para novo Mapa de Cotação
    lErro = CF("MapaCotacao_ObterNumIntAutomatico", lNumIntDoc)
    If lErro <> SUCESSO Then gError 114597

    'Guarda novo NumIntDoc
    objMapaCotacao.lNumIntDoc = lNumIntDoc

    sSelect = "INSERT INTO MapaCotacao (NumIntDoc, Codigo, FilialEmpresa, Data, TaxaFinanceira) VALUES (?,?,?,?,?)"

    'Grava o Mapa de Cotacao passado por parametro
    With objMapaCotacao
        lErro = Comando_Executar(alComando(0), sSelect, .lNumIntDoc, .lCodigo, .iFilialEmpresa, .dtData, .dTaxaFinanceira)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 114598

    sSelect = "INSERT INTO MapaCotacaoItemCotacao (NumIntDoc, CodMapaCotacao, FilialEmpresa, NumIntItemCotacao) VALUES (?,?,?,?)"

    'Grava os itens do mapa de cotacao
    For Each objItemMapaCotacao In objMapaCotacao.ColItensMapaCotacao

        'Gera NumIntDoc para novo Mapa de Cotação
        lErro = CF("MapaCotacaoItem_ObterNumIntAutomatico", lNumIntDoc)
        If lErro <> SUCESSO Then gError 114601

        objItemMapaCotacao.lNumIntDoc = lNumIntDoc

        With objItemMapaCotacao
            lErro = Comando_Executar(alComando(0), sSelect, .lNumIntDoc, .lCodMapaCotacao, .iFilialEmpresa, .lNumIntItemCotacao)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 114599

    Next

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 114600

    MapaCotacao_Grava = SUCESSO

    Exit Function

Erro_MapaCotacao_Grava:

    MapaCotacao_Grava = gErr

    Select Case gErr

        Case 114595
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 114596
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 114597, 114601

        Case 114598
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MAPACOTACAO", gErr)

        Case 114599
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ITEMMAPACOTACAO", gErr)

        Case 114600
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147565)

    End Select

    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

End Function

'Incluído por Luiz Nogueira em 03/03/04
Public Function RelABCCompras_Grava(ByVal colItensRelABCCompras As Collection) As Long
'Grava no BD os dados que serão usados para impressão do Relatório ABC de Compras

Dim lErro As Long
Dim alComando As Long
Dim iIndice As Integer
Dim objRelABCCompras As ClassRelABCCompras
Dim lComando As Long
Dim lTransacao As Long
Dim lNumIntRel As Long

On Error GoTo Erro_RelABCCompras_Grava

    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 127029
    
    'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 127030

    'Obtem o Próximo Numero Interno do Relatório
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROX_REL_ABCCOMPRAS", lNumIntRel, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError 127031

    'Para cada item a ser gravado
    For Each objRelABCCompras In colItensRelABCCompras

        'Guarda no obj o número interno do relatório. É importante, pois a função chamadora irá precisar desse número
        'para filtrar o relatório
        objRelABCCompras.lNumIntRel = lNumIntRel
        
        'Insere os dados no BD
        lErro = Comando_Executar(lComando, "INSERT INTO RelABCCompras (NumIntRel, Produto, Ranking, Quantidade, Valor, PercParticipacao, ItemCategoria) VALUES ( ?, ?, ?, ?, ?, ?, ?)", lNumIntRel, objRelABCCompras.sProduto, objRelABCCompras.lRanking, objRelABCCompras.dQuantidade, objRelABCCompras.dValor, objRelABCCompras.dPercParticipacao, objRelABCCompras.sItemCategoria)
        If lErro <> AD_SQL_SUCESSO Then gError 127032

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 127033

    RelABCCompras_Grava = SUCESSO

    Exit Function

Erro_RelABCCompras_Grava:

    RelABCCompras_Grava = gErr

    Select Case gErr

        Case 127029
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 127030
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 127031
        
        Case 127032
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELABCCOMPRAS", gErr)

        Case 127033
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147566)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    'Faz rollback da transação
    Call Transacao_Rollback

End Function

'Incluído por Luiz Nogueira em 03/03/04
Public Function RelABCFornecedores_Grava(ByVal colItensRelABCFornecedores As Collection) As Long
'Grava no BD os dados que serão usados para impressão do Relatório ABC de Compras

Dim lErro As Long
Dim alComando As Long
Dim iIndice As Integer
Dim objRelABCFornecedores As ClassRelABCFornecedores
Dim lComando As Long
Dim lTransacao As Long
Dim lNumIntRel As Long

On Error GoTo Erro_RelABCFornecedores_Grava

    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 127064
    
    'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 127065

    'Obtem o Próximo Numero Interno do Relatório
    lErro = CF("Config_ObterNumInt", "ComprasConfig", "NUM_PROX_REL_ABCFORNECEDORES", lNumIntRel, 1, EMPRESA_TODA)
    If lErro <> SUCESSO Then gError 127066

    'Para cada item a ser gravado
    For Each objRelABCFornecedores In colItensRelABCFornecedores

        'Guarda no obj o número interno do relatório. É importante, pois a função chamadora irá precisar desse número
        'para filtrar o relatório
        objRelABCFornecedores.lNumIntRel = lNumIntRel
        
        'Insere os dados no BD
        lErro = Comando_Executar(lComando, "INSERT INTO RelABCFornecedores (NumIntRel, Fornecedor, FilialFornecedor, Ranking, Valor, PercParticipacao, ItemCategoria) VALUES ( ?, ?, ?, ?, ?, ?, ?)", lNumIntRel, objRelABCFornecedores.lFornecedor, objRelABCFornecedores.iFilialForn, objRelABCFornecedores.lRanking, objRelABCFornecedores.dValor, objRelABCFornecedores.dPercParticipacao, objRelABCFornecedores.sItemCategoria)
        If lErro <> AD_SQL_SUCESSO Then gError 127067

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 127068

    RelABCFornecedores_Grava = SUCESSO

    Exit Function

Erro_RelABCFornecedores_Grava:

    RelABCFornecedores_Grava = gErr

    Select Case gErr

        Case 127064
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 127065
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 127066
        
        Case 127067
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELABCCOMPRAS", gErr)

        Case 127068
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147567)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    'Faz rollback da transação
    Call Transacao_Rollback

End Function

'#########################################################
'Inserido por Wagner
Function PedidoCompra_Verifica_BloqueioOP(ByVal objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim objItemOP As ClassItemOP
Dim objItemPC As ClassItemPedCompra
Dim colRC As New Collection
Dim objRC As ClassRequisicaoCompras
Dim objOP As ClassOrdemDeProducao

On Error GoTo Erro_PedidoCompra_Verifica_BloqueioOP

    For Each objItemPC In objPedidoCompra.colItens
    
        lErro = CF("ItemRCItemPC_Le2", objItemPC.lNumIntDoc, colRC)
        If lErro <> SUCESSO And lErro <> 131651 Then gError 131653
        
        If lErro = SUCESSO Then
        
            For Each objRC In colRC
            
                'Se está associada a uma OP
                If Len(Trim(objRC.sOPCodigo)) > 0 Then
            
                    Set objItemOP = New ClassItemOP
                    
                    'Se é um item Para cada OP/RC
                    If objRC.lNumIntDocItemOP <> 0 Then
                        objItemOP.lNumIntDoc = objRC.lNumIntDocItemOP
            
                        'Lê a o item OP
                        lErro = CF("ItemOP_Le_NumIntDoc", objItemOP)
                        If lErro <> SUCESSO And lErro <> 33226 Then gError 131654
            
                        If lErro = SUCESSO Then
                                        'Se o item estiver bloqueado => Erro
                            If objItemOP.iSituacao = ITEMOP_SITUACAO_DESAB Then gError 131655
                        End If

                    Else
            
                        Set objOP = New ClassOrdemDeProducao
           
                        objOP.sCodigo = objRC.sOPCodigo
                        objOP.iFilialEmpresa = objRC.iFilialEmpresa
            
                        'Lê a OP
                        lErro = CF("OrdemDeProducao_Le_ComItens", objOP)
                        If lErro <> SUCESSO And lErro <> 21960 Then gError 131656
            
                        For Each objItemOP In objOP.colItens
            
                            If objItemOP.iSituacao = ITEMOP_SITUACAO_DESAB Then gError 131657
            
                        Next
            
                    End If
            
                End If
            
            Next
        
        End If
        
    Next
                
    PedidoCompra_Verifica_BloqueioOP = SUCESSO
    
    Exit Function
                
Erro_PedidoCompra_Verifica_BloqueioOP:

    PedidoCompra_Verifica_BloqueioOP = gErr

    Select Case gErr
    
        Case 131653, 131654, 131656
    
        Case 131655, 131657
            Call Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIO_ITEMOP_PC", gErr, objItemOP.iItem, objItemOP.sCodigo, objItemOP.sJustificativaBloqueio)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147568)

    End Select
    
    Exit Function
    
End Function
'#########################################################

'########################################################################
'Inserido por Wagner 08/12/05
Function PedidoCompra_Cancelar_Baixa(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Refaz os Pedidos de Compras já Baixados e os ItensPC

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(3) As Long
Dim tPedCompra As typePedidoCompras
Dim tItemPedido As typeItemPedCompra
Dim iStatus As Integer
Dim iStatusPC As Integer
Dim lTransacao As Long

On Error GoTo Erro_PedidoCompra_Cancelar_Baixa
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 141352
    Next
                                                           
    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 141353
    
    'Verifica se o Pedido de Compras relacionado ao ItemPC está baixado
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, Status FROM PedidoCompraN WHERE NumIntDoc = ? ", 0, tPedCompra.lNumIntDoc, iStatusPC, objPedidoCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 141354
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141355
                    
    If iStatusPC <> PEDIDOCOMPRA_STATUS_BAIXADO Then gError 141350
            
    'Se o Pedido de Compras está baixado
    If lErro = AD_SQL_SUCESSO Then
            
        'Lê os Itens desse Pedido
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc,Quantidade,QuantRecebida,QuantRecebimento FROM ItensPedCompraN WHERE PedCompra = ? AND StatusBaixa = ?", 0, _
                                                    tItemPedido.lNumIntDoc, tItemPedido.dQuantidade, tItemPedido.dQuantRecebida, tItemPedido.dQuantRecebimento, objPedidoCompra.lNumIntDoc, ITENSPEDCOMPRA_STATUS_BAIXADO)
        If lErro <> AD_SQL_SUCESSO Then gError 141356
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141357
        
        'Equanto encontrar ItensPC
        Do While lErro = AD_SQL_SUCESSO
                        
            If tItemPedido.dQuantidade > tItemPedido.dQuantRecebida Then
                iStatus = ITEM_PED_COMPRAS_ABERTO
            Else
                iStatus = ITEM_PED_COMPRAS_RECEBIDO
            End If
            
            'Atualiza o ItemPCBaixado
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensPedCompraN SET Status = ?, StatusBaixa=?", alComando(2), iStatus, ITENSPEDCOMPRA_STATUS_ATIVO)
            If lErro <> AD_SQL_SUCESSO Then gError 141358
            
            'Busca próximo ItemPC
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141359
                    
        Loop
                                 
        'Atualiza o Pedido de Compras para ativo
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE PedidoCompraN SET Status = ?, DataBaixa= ? ", alComando(1), PEDIDOCOMPRA_STATUS_ATIVO, DATA_NULA)
        If lErro <> AD_SQL_SUCESSO Then gError 141360
        
    End If
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 141361
    
    PedidoCompra_Cancelar_Baixa = SUCESSO
    
    Exit Function

Erro_PedidoCompra_Cancelar_Baixa:

    PedidoCompra_Cancelar_Baixa = gErr

    Select Case gErr
    
        Case 141350
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_BAIXADO", gErr, objPedidoCompra.lCodigo)
        
        Case 141352
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 141353
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 141354, 141355
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRABAIXADO", gErr)
                
        Case 141356, 141357, 141359
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)
                
        Case 141358
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDCOMPRA", gErr, tItemPedido.lNumIntDoc)
            
        Case 141360
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PEDIDOCOMPRABAIXADO", gErr, tItemPedido.lNumIntDoc)
                            
        Case 141361
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 141870)
    
    End Select
    
    'Fechamento dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Faz rollback da transação
    Call Transacao_Rollback
    
    Exit Function

End Function
'########################################################################

'################################################
'Inserido por Wagner 17/07/2006
Function PCItens_Grava_Customizado(ByVal objItemPC As ClassItemPedCompra) As Long
    PCItens_Grava_Customizado = SUCESSO
End Function
'################################################

Private Function ItemPVItemRC_Exclui(ByVal objItemReqCompra As ClassItemReqCompras) As Long

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim lCodigo As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long
Dim dQuant As Double
Dim dQuantSC As Double

On Error GoTo Erro_ItemPVItemRC_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 178613
    Next

    lErro = Comando_Executar(alComando(1), "SELECT ItemPV, Quantidade FROM ItemPVItemRC WHERE ItemRC =?", lNumIntDoc, dQuant, objItemReqCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 178631
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178625
        
    Do While lErro = SUCESSO
        
        'Busca no Bd o Item do Pedido de Venda cujo código e produto são informados
        lErro = Comando_ExecutarPos(alComando(2), "SELECT QuantSC FROM ItensPedidoDeVenda WHERE NumIntDoc=?", 0, dQuantSC, lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 178626
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178627
    
        'Se não encontrou o Item de Pedido de Venda
        If lErro = AD_SQL_SEM_DADOS Then
    
            'BUsca nos itens baixados o Item relacionado
            lErro = Comando_ExecutarPos(alComando(2), "SELECT QuantSC FROM ItensPedidoDeVendaBaixados WHERE NumIntDoc = ?", 0, dQuantSC, lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 178632
            
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178633
            
            If lErro = AD_SQL_SUCESSO Then
            
                dQuantSC = dQuantSC - dQuant
                If dQuantSC < 0 Then dQuantSC = 0
                
                lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensPedidoDeVendaBaixados SET QuantSC = ?", alComando(2), dQuantSC)
                If lErro <> AD_SQL_SUCESSO Then gError 178634
            End If
    
    
        Else
    
            dQuantSC = dQuantSC - dQuant
            If dQuantSC < 0 Then dQuantSC = 0
    
            'Altera o campo QuantSC na tabela ItensPedidoDeVenda
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE ItensPedidoDeVenda SET QuantSC=QuantSC - ? ", alComando(2), dQuantSC)
            If lErro <> AD_SQL_SUCESSO Then gError 178629
        
        End If
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178630
        
    Loop
    
    lErro = Comando_Executar(alComando(4), "DELETE FROM ItemPVItemRC WHERE ItemRC = ?", objItemReqCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 178614
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItemPVItemRC_Exclui = SUCESSO
    
    Exit Function
    
Erro_ItemPVItemRC_Exclui:

    ItemPVItemRC_Exclui = gErr
    
    Select Case gErr
    
        Case 178613
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 178626, 178627, 178632, 178633
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case 178629, 178634
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ITENSPEDIDODEVENDA", gErr)
        
        Case 178625, 178630, 178631
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMPVITEMRC", gErr)
        
        Case 178614
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITEMPVITEMRC", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178615)
    
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function
    
End Function

Private Function PedidoCompra_Valida_CodigoPV(ByVal objPedidoCompra As ClassPedidoCompras) As Long

Dim lErro As Long
Dim objPV As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoCompra_Valida_CodigoPV

    If objPedidoCompra.lCodigoPV > 0 Then

        objPV.iFilialEmpresa = objPedidoCompra.iFilialEmpresa
        objPV.lCodigo = objPedidoCompra.lCodigoPV
        
        'Lê o Pedido de Venda a partir do Código passado.
        lErro = CF("PedidoDeVenda_Le", objPV)
        If lErro <> SUCESSO And lErro <> 26509 Then gError 178827
    
        If lErro <> SUCESSO Then
        
            lErro = CF("PedidoVendaBaixado_Le", objPV)
            If lErro <> SUCESSO And lErro <> 46135 Then gError 178828
        
            If lErro <> SUCESSO Then gError 178829
        
        End If

    End If
    
    PedidoCompra_Valida_CodigoPV = SUCESSO
    
    Exit Function
    
Erro_PedidoCompra_Valida_CodigoPV:

    PedidoCompra_Valida_CodigoPV = gErr
    
    Select Case gErr
    
        Case 178827, 178828
    
        Case 178829
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178830)
    
    End Select
    
    Exit Function
    
End Function

Public Function Concorrencia_Grava_Baixa(ByVal objConcorrencia As ClassConcorrencia) As Long
'Atualiza Concorrência e grava Pedidos de Compras

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Concorrencia_Grava_Baixa

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 189331
    
    'Baixa a Concorrência
    lErro = CF("Concorrencia_Baixa", objConcorrencia)
    If lErro <> SUCESSO Then gError 189332
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 189333

    Concorrencia_Grava_Baixa = SUCESSO

    Exit Function

Erro_Concorrencia_Grava_Baixa:

    Concorrencia_Grava_Baixa = gErr

    Select Case gErr

        Case 189331
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 189332
        
        Case 189333
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function CotacaoProdutoItemFP_Inclui(ByVal objCotacao As ClassCotacao) As Long

Dim lErro As Long
Dim lComando As Long
Dim objItemFormPreco As New ClassItemFormPreco
Dim iIndice As Integer
Dim objCotacaoProduto As ClassCotacaoProduto

On Error GoTo Erro_CotacaoProdutoItemFP_Inclui

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 63526

    For Each objCotacaoProduto In objCotacao.colCotacaoProduto

        For Each objItemFormPreco In objCotacaoProduto.colItemFormPreco

            'Inclui na tabela CotacaoProdutoItemRC os itens armazenados em colItemReqCompras
            lErro = Comando_Executar(lComando, "INSERT INTO CotacaoProdutoItemFP (CotacaoProduto,NumIntItemFormPreco,QuantidadeCotar) VALUES (?,?,?)", objCotacaoProduto.lNumIntDoc, objItemFormPreco.lNumIntDoc, objItemFormPreco.dQtde)
            If lErro <> AD_SQL_SUCESSO Then Error 63527

        Next
    Next

    'Fecha o  comando
    Comando_Fechar (lComando)

    CotacaoProdutoItemFP_Inclui = SUCESSO

    Exit Function

Erro_CotacaoProdutoItemFP_Inclui:

    CotacaoProdutoItemFP_Inclui = Err

    Select Case Err

        Case 63526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 63527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COTACAOPRODUTOITEMFP", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147487)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Cotacao_Grava_Pedidos_EmTrans(ByVal objCotacao As ClassCotacao, ByVal colPedidoCotacao As Collection) As Long
'Grava Pedidos de Cotacao gerados

Dim lErro As Long
Dim colItensPedCotacao As New Collection

On Error GoTo Erro_Cotacao_Grava_Pedidos_EmTrans

    lErro = CF("Cotacao_Le_Codigo", objCotacao)
    If lErro <> SUCESSO And lErro <> 89163 Then gError 89165

    'se o codigo já tiver sido cadastrado
    If lErro = SUCESSO Then gError 89166

    'Le e faz lock dos objetos associados a Cotacao
    lErro = CF("Cotacao_ObjetosAssociados_Lock", objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then gError 63519

    'Inclui Cotacao
    lErro = CF("Cotacao_Inclui", objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then gError 63520

    lErro = CF("CotacaoProdutoItemRC_Inclui", objCotacao)
    If lErro <> SUCESSO Then gError 63521

    lErro = CF("CotacaoProdutoItemFP_Inclui", objCotacao)
    If lErro <> SUCESSO Then gError 63521

    'Inclui Pedido de Cotacao
    lErro = CF("PedidosCotacao_Inclui", objCotacao, colPedidoCotacao)
    If lErro <> SUCESSO Then gError 63522

    Cotacao_Grava_Pedidos_EmTrans = SUCESSO

    Exit Function

Erro_Cotacao_Grava_Pedidos_EmTrans:

    Cotacao_Grava_Pedidos_EmTrans = gErr

    Select Case gErr

        Case 63519 To 63522, 89165

        Case 89166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COTACAO_CADASTRADA", gErr, objCotacao.iFilialEmpresa, objCotacao.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147488)

    End Select

    Exit Function

End Function

Function Concorrencia_Grava_OrcVenda(ByVal objConcorrencia As ClassConcorrencia, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Grava concorrência

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Concorrencia_Grava_OrcVenda

    'Abre Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 66931

    'Grava concorrência
    lErro = CF("Concorrencia_Grava_EmTrans", objConcorrencia)
    If lErro <> SUCESSO Then gError 66912

    lErro = CF("Concorrencia_Atualiza_OrcVenda", objConcorrencia, objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 66912
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 66932

    Concorrencia_Grava_OrcVenda = SUCESSO

    Exit Function

Erro_Concorrencia_Grava_OrcVenda:

    Concorrencia_Grava_OrcVenda = gErr

    Select Case gErr

        Case 66912

        Case 66931
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 66932
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 147463)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Public Function Concorrencia_Atualiza_OrcVenda(ByVal objConcorrencia As ClassConcorrencia, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long, iIndice As Integer, alComando(1 To 33) As Long, iStatusComercial As Integer
Dim lNumIntItemOV As Long, iPCSituacao As Integer, iPCSituacaoNova As Integer, dFPQtde As Double, bFaltaCotacaoItemOV As Boolean
Dim lNumIntItemFP As Long, iFPSituacao As Integer, dPrecoAjustado As Double, dFPPercentMargem As Double
Dim dPrecoCalcItemOV As Double, dFPPrecoTotal As Double, bFaltaCotacaoOV As Boolean, dIPIAliquota As Double, dQuantidade As Double
Dim dCustoUnit As Double, dPrecoUnit As Double

On Error GoTo Erro_Concorrencia_Atualiza_OrcVenda

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 178613
    Next

    'bloquear o orcamento de venda
    lErro = Comando_ExecutarPos(alComando(1), "SELECT StatusComercial FROM OrcamentoVenda WHERE FilialEmpresa = ? AND Codigo = ?", 0, iStatusComercial, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 184391
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184392
    If lErro <> AD_SQL_SUCESSO Then gError 184393
    
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 184394
    
    'se o orcamento não está em fase de avaliacao de custos para calculo de precos, vou ignorar
    If iStatusComercial = OV_STATUS_COMERCIAL_EM_COTACAO Or iStatusComercial = OV_STATUS_COMERCIAL_COTADO Then
    
        bFaltaCotacaoOV = False
        
        'ler itens ov
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntDoc, PCSituacao, Quantidade FROM ItensOrcamentoVenda WHERE FilialEmpresa = ? AND CodOrcamento = ? AND PCSituacao IN (?,?,?)", 0, lNumIntItemOV, iPCSituacao, dQuantidade, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, ITEMOV_PCSITUACAO_COTAR, ITEMOV_PCSITUACAO_EM_COTACAO, ITEMOV_PCSITUACAO_COTADO)
        If lErro <> AD_SQL_SUCESSO Then gError 184395
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184396
        If lErro <> AD_SQL_SUCESSO Then gError 184397
        
        'para cada itemov
        Do While lErro = AD_SQL_SUCESSO
        
            iPCSituacaoNova = iPCSituacao
            bFaltaCotacaoItemOV = False
            dPrecoCalcItemOV = 0
            
            'buscar os itensformpreco
            lErro = Comando_ExecutarPos(alComando(3), "SELECT NumIntDoc, Situacao, Qtde, PercentMargem, PrecoTotal FROM ItensFormPreco WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ?", 0, lNumIntItemFP, iFPSituacao, dFPQtde, dFPPercentMargem, dFPPrecoTotal, FPORIGEM_ITEMOV, lNumIntItemOV)
            If lErro <> AD_SQL_SUCESSO Then gError 184398
        
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184399
        
            'para cada itemfp
            Do While lErro = AD_SQL_SUCESSO
            
                If iFPSituacao = FPSITUACAO_EM_COTACAO Or iFPSituacao = FPSITUACAO_COTADO Then
                
                    'buscar preco escolhido em concorrencia
                    lErro = Comando_Executar(alComando(4), "SELECT CotacaoItemConcorrenciaN.PrecoAjustado, ItensCotacaoN.AliquotaIPI FROM CotacaoProdutoN, ItensCotacaoN, ItemPedCotacaoN, CotacaoProdutoItemFP, CotacaoItemConcorrenciaN, ItensConcorrenciaN, ConcorrenciaN " & _
                        " WHERE CotacaoProdutoN.NumIntDoc = ItemPedCotacaoN.CotacaoProduto AND ItensCotacaoN.ItemPedCotacao = ItemPedCotacaoN.NumIntDoc AND " & _
                        " CotacaoProdutoItemFP.CotacaoProduto = CotacaoProdutoN.NumIntDoc AND CotacaoProdutoItemFP.NumIntItemFormPreco = ? AND CotacaoItemConcorrenciaN.Escolhido <> 0 AND " & _
                        " CotacaoItemConcorrenciaN.ItemCotacao = ItensCotacaoN.NumIntDoc AND CotacaoItemConcorrenciaN.ItemConcorrencia = ItensConcorrenciaN.NumIntDoc AND ItensConcorrenciaN.Concorrencia = ConcorrenciaN.NumIntDoc", _
                        dPrecoAjustado, dIPIAliquota, lNumIntItemFP)
                    If lErro <> AD_SQL_SUCESSO Then gError 184400
    
                    lErro = Comando_BuscarProximo(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184401
                    
                    'se o custo ainda nao está definido
                    If lErro <> AD_SQL_SUCESSO Then
                        bFaltaCotacaoItemOV = True
                        bFaltaCotacaoOV = True
                        If iFPSituacao = FPSITUACAO_COTADO Then
                            lErro = Comando_ExecutarPos(alComando(5), "UPDATE ItensFormPreco SET Situacao = ?, CustoUnit = ?, PrecoUnit = ?, PrecoTotal = ?", alComando(3), FPSITUACAO_EM_COTACAO, 0, 0, 0)
                            If lErro <> AD_SQL_SUCESSO Then gError 184402
                            iPCSituacaoNova = ITEMOV_PCSITUACAO_EM_COTACAO
                        End If
                    Else
                        If iFPSituacao = FPSITUACAO_EM_COTACAO Then
                            dCustoUnit = Arredonda_Moeda(dPrecoAjustado * (1 + dIPIAliquota))
                            dPrecoUnit = Arredonda_Moeda(dCustoUnit * (1 + dFPPercentMargem))
                            dFPPrecoTotal = Arredonda_Moeda(dPrecoUnit * dFPQtde)
                            lErro = Comando_ExecutarPos(alComando(5), "UPDATE ItensFormPreco SET Situacao = ?, CustoUnit = ?, PrecoUnit = ?, PrecoTotal = ?", alComando(3), FPSITUACAO_COTADO, dCustoUnit, dPrecoUnit, dFPPrecoTotal)
                            If lErro <> AD_SQL_SUCESSO Then gError 184403
                        End If
                        dPrecoCalcItemOV = dPrecoCalcItemOV + dFPPrecoTotal
                    End If
                
                Else
                
                    dPrecoCalcItemOV = dPrecoCalcItemOV + dFPPrecoTotal
                    
                End If
                
                'buscar proximo itemfp
                lErro = Comando_BuscarProximo(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184404
            
            Loop
            
            If dQuantidade <> 0 Then
            
                dPrecoCalcItemOV = Arredonda_Moeda(dPrecoCalcItemOV / dQuantidade)
                
            End If
            
            'atualizar itemov
            lErro = Comando_ExecutarPos(alComando(6), "UPDATE ItensOrcamentoVenda SET PCPrecoUnitCalc=?, PCSituacao=? ", alComando(2), dPrecoCalcItemOV, IIf(bFaltaCotacaoItemOV, ITEMOV_PCSITUACAO_EM_COTACAO, ITEMOV_PCSITUACAO_COTADO))
            If lErro <> AD_SQL_SUCESSO Then gError 184405
            
            'buscar proximo itemov
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184406
        
        Loop
    
        If (bFaltaCotacaoOV And iStatusComercial = OV_STATUS_COMERCIAL_COTADO) Or _
            (bFaltaCotacaoOV = False And iStatusComercial = OV_STATUS_COMERCIAL_EM_COTACAO) Then
        
            lErro = Comando_ExecutarPos(alComando(7), "UPDATE OrcamentoVenda SET StatusComercial = ?", alComando(1), IIf(bFaltaCotacaoOV, OV_STATUS_COMERCIAL_EM_COTACAO, OV_STATUS_COMERCIAL_COTADO))
            If lErro <> AD_SQL_SUCESSO Then gError 184407
        
        End If
        
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Concorrencia_Atualiza_OrcVenda = SUCESSO
    
    Exit Function
    
Erro_Concorrencia_Atualiza_OrcVenda:

    Concorrencia_Atualiza_OrcVenda = gErr

    Select Case gErr

        Case 178613
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184390)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function ReqCompraEnvio_Grava(ByVal objReqCompraEnvio As ClassReqCompraEnvio) As Long
'Grava a Requisição de Compra

Dim lTransacao As Long
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objReqCompra As ClassRequisicaoCompras
Dim dtDataEnvio As Date, iStatus As Integer

On Error GoTo Erro_ReqCompraEnvio_Grava
    
    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211124

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211125
    Next
    
    For Each objReqCompra In objReqCompraEnvio.colRequisicao
    
        If objReqCompra.iSelecionado = MARCADO Then

            lErro = Comando_ExecutarPos(alComando(0), "SELECT DataEnvio, Status FROM RequisicaoCompraN WHERE NumIntDoc = ?", 0, dtDataEnvio, iStatus, objReqCompra.lNumIntDoc)
            If lErro <> SUCESSO Then gError 211126
        
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211127
            
            If lErro <> AD_SQL_SUCESSO Then gError 211128
           
            'se a requisição estiver baixada
            If lErro = AD_SQL_SUCESSO And iStatus = REQUISICAOCOMPRA_STATUS_BAIXADO Then gError 211129
        
            'Se a Requisição de Compras já foi enviada, Erro
            If lErro = AD_SQL_SUCESSO And dtDataEnvio <> DATA_NULA Then gError 211130
            
            'Faz LockExclusive na Requisição
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> AD_SQL_SUCESSO Then gError 211131
            
            objReqCompra.bEnvio = True
            objReqCompra.sUsuRegEnvio = gsUsuario
    
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE RequisicaoCompraN SET DataEnvio = ?, DataRegEnvio = ?, HoraRegEnvio = ?, UsuRegEnvio = ? ", alComando(0), gdtDataAtual, Date, CDbl(Time), gsUsuario)
            If lErro <> AD_SQL_SUCESSO Then gError 211132
        
        End If
        
    Next
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 211133
    
    For Each objReqCompra In objReqCompraEnvio.colRequisicao
        
        If objReqCompra.iSelecionado = MARCADO Then
        
            lErro = CF("WorkFlow_Trata_Transacao", "COM", objReqCompra, gsUsuario, objReqCompra, 1004)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End If

    Next
   
    ReqCompraEnvio_Grava = SUCESSO

    Exit Function
    
Erro_ReqCompraEnvio_Grava:

    ReqCompraEnvio_Grava = gErr

    Select Case gErr
    
        Case 211124
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 211125
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211126, 211127
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr)

        Case 211128
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objReqCompra.lCodigo)

        Case 211129
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_COMPRA_BAIXADA", gErr, objReqCompra.lCodigo)

        Case 211130
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_ENVIADA", gErr, objReqCompra.lCodigo)

        Case 211131
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISICAOCOMPRA", gErr, objReqCompra.lCodigo)

        Case 211132
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_REQUISICAOCOMPRA", gErr, objReqCompra.lCodigo)
               
        Case 211133
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM
                                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211134)
    
    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback

    Exit Function
    
End Function

Function ReqCompraAprova_Grava(ByVal objReqCompraEnvio As ClassReqCompraEnvio) As Long
'Grava a Requisição de Compra

Dim lTransacao As Long
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objReqCompra As ClassRequisicaoCompras
Dim dtDataRegAprov As Date, iStatus As Integer

On Error GoTo Erro_ReqCompraAprova_Grava
    
    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211124

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211125
    Next
    
    For Each objReqCompra In objReqCompraEnvio.colRequisicao
    
        If objReqCompra.iSelecionado = MARCADO Then

            lErro = Comando_ExecutarPos(alComando(0), "SELECT DataRegAprov, Status FROM RequisicaoCompraN WHERE NumIntDoc = ?", 0, dtDataRegAprov, iStatus, objReqCompra.lNumIntDoc)
            If lErro <> SUCESSO Then gError 211126
        
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211127
            
            If lErro <> AD_SQL_SUCESSO Then gError 211128
           
            'se a requisição estiver baixada
            If lErro = AD_SQL_SUCESSO And iStatus = REQUISICAOCOMPRA_STATUS_BAIXADO Then gError 211129
        
            'Se a Requisição de Compras já foi enviada, Erro
            If lErro = AD_SQL_SUCESSO And dtDataRegAprov <> DATA_NULA Then gError 211130
            
            'Faz LockExclusive na Requisição
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> AD_SQL_SUCESSO Then gError 211131
            
            objReqCompra.bAprovacao = True
            objReqCompra.sUsuRegAprov = gsUsuario
    
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE RequisicaoCompraN SET DataRegAprov = ?, HoraRegAprov = ?, UsuRegAprov = ? ", alComando(0), Date, CDbl(Time), gsUsuario)
            If lErro <> AD_SQL_SUCESSO Then gError 211132
        
        End If
        
    Next
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 211133
    
    For Each objReqCompra In objReqCompraEnvio.colRequisicao
        
        If objReqCompra.iSelecionado = MARCADO Then
        
            lErro = CF("WorkFlow_Trata_Transacao", "COM", objReqCompra, gsUsuario, objReqCompra, 1004)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End If

    Next
   
    ReqCompraAprova_Grava = SUCESSO

    Exit Function
    
Erro_ReqCompraAprova_Grava:

    ReqCompraAprova_Grava = gErr

    Select Case gErr
    
        Case 211124
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 211125
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211126, 211127
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr)

        Case 211128
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objReqCompra.lCodigo)

        Case 211129
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_COMPRA_BAIXADA", gErr, objReqCompra.lCodigo)

        Case 211130
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_APROVADA", gErr, objReqCompra.lCodigo)

        Case 211131
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISICAOCOMPRA", gErr, objReqCompra.lCodigo)

        Case 211132
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_REQUISICAOCOMPRA", gErr, objReqCompra.lCodigo)
               
        Case 211133
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM
                                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211134)
    
    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback

    Exit Function
    
End Function

Public Function NFiscal_Trata_Eventos_WorkFlow_Compras(ByVal objNF As ClassNFiscal) As Long

Dim lErro As Long
Dim vReqCompra As Variant
Dim objReqCompra As ClassRequisicaoCompras

On Error GoTo Erro_NFiscal_Trata_Eventos_WorkFlow_Compras

    For Each vReqCompra In objNF.colReqCompraBaixadas
    
        Set objReqCompra = New ClassRequisicaoCompras
        
        objReqCompra.lNumIntDoc = vReqCompra
        
        lErro = CF("RequisicaoComprasTodas_Le", objReqCompra)
        If lErro <> SUCESSO And lErro <> 66696 Then gError ERRO_SEM_MENSAGEM
        
        objReqCompra.bBaixa = True
        objReqCompra.sUsuRegBaixa = gsUsuario
    
        lErro = CF("WorkFlow_Trata_Transacao", "COM", objReqCompra, gsUsuario, objReqCompra, 1004)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Next

    NFiscal_Trata_Eventos_WorkFlow_Compras = SUCESSO

    Exit Function

Erro_NFiscal_Trata_Eventos_WorkFlow_Compras:

    NFiscal_Trata_Eventos_WorkFlow_Compras = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211918)

    End Select
    
    Exit Function
    
End Function

Function PedCompraAprova_Grava(ByVal objPedCompraEnvio As ClassPedCompraEnvio) As Long

Dim lTransacao As Long
Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim objPedCompra As ClassPedidoCompras
Dim dtDataEnvio As Date, iStatus As Integer

On Error GoTo Erro_PedCompraAprova_Grava
    
    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 211124

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211125
    Next
    
    For Each objPedCompra In objPedCompraEnvio.colPedidos
    
        If objPedCompra.iSelecionado = MARCADO Then

            lErro = Comando_ExecutarPos(alComando(0), "SELECT DataEnvio, Status FROM PedidoCompra WHERE NumIntDoc = ?", 0, dtDataEnvio, iStatus, objPedCompra.lNumIntDoc)
            If lErro <> SUCESSO Then gError 211126
        
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211127
            
            If lErro <> AD_SQL_SUCESSO Then gError 211128
           
            'se o pedido estiver baixado
            If lErro = AD_SQL_SUCESSO And iStatus = PEDIDOCOMPRA_STATUS_BAIXADO Then gError 211129
        
            'Se o pedido de Compras já foi enviado, Erro
            If lErro = AD_SQL_SUCESSO And dtDataEnvio <> DATA_NULA Then gError 211130
            
            'Faz LockExclusive na Requisição
            lErro = Comando_LockExclusive(alComando(0))
            If lErro <> AD_SQL_SUCESSO Then gError 211131
            
            objPedCompra.bAprovacao = True
            objPedCompra.sUsuRegAprov = gsUsuario
    
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE PedidoCompra SET DataRegAprov = ?, HoraRegAprov = ?, UsuRegAprov = ? ", alComando(0), Date, CDbl(Time), gsUsuario)
            If lErro <> AD_SQL_SUCESSO Then gError 211132
        
        End If
        
    Next
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 211133
    
    For Each objPedCompra In objPedCompraEnvio.colPedidos
        
        If objPedCompra.iSelecionado = MARCADO Then
        
            lErro = CF("WorkFlow_Trata_Transacao", "COM", objPedCompra, gsUsuario, objPedCompra, 1002)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End If

    Next
   
    PedCompraAprova_Grava = SUCESSO

    Exit Function
    
Erro_PedCompraAprova_Grava:

    PedCompraAprova_Grava = gErr

    Select Case gErr
    
        Case 211124
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 211125
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211126, 211127
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr)

        Case 211128
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, objPedCompra.lCodigo)

        Case 211129
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAO_COMPRA_BAIXADA", gErr, objPedCompra.lCodigo)

        Case 211130
            Call Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_ENVIADA", gErr, objPedCompra.lCodigo)

        Case 211131
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISICAOCOMPRA", gErr, objPedCompra.lCodigo)

        Case 211132
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_REQUISICAOCOMPRA", gErr, objPedCompra.lCodigo)
               
        Case 211133
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM
                                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211134)
    
    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback

    Exit Function
    
End Function

Function ItensPCEntrega_Grava(ByVal objItemPC As ClassItemPedCompra) As Long

Dim lTransacao As Long
Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_ItensPCEntrega_Grava
    
    'Abre a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 213191

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213192
    Next
    
    If objItemPC.lNumIntDoc = 0 Then gError 213193
    
    lErro = ItensPCEntrega_Exclui(alComando(0), 0, objItemPC.lNumIntDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = ItensPCEntrega_Insere(alComando(1), alComando(2), objItemPC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 213194
   
    ItensPCEntrega_Grava = SUCESSO

    Exit Function
    
Erro_ItensPCEntrega_Grava:

    ItensPCEntrega_Grava = gErr

    Select Case gErr
    
        Case 213191
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 213192
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213193
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", gErr, objItemPC.lNumIntDoc)

        Case 213194
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case ERRO_SEM_MENSAGEM
                                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213195)
    
    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Transacao_Rollback

    Exit Function
    
End Function

Private Function ItensPCEntrega_Insere(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objItemPC As ClassItemPedCompra) As Long
'exclui as  datas de entrega do pedido e em seguinda grava as novas

Dim lErro As Long, dQtde As Double
Dim objDataEntrega As ClassDataEntrega
Dim dQtdePC As Double, lNumIntPC As Long

On Error GoTo Erro_ItensPCEntrega_Insere

    If objItemPC.colDataEntrega.Count > 0 Then
    
        For Each objDataEntrega In objItemPC.colDataEntrega
            dQtde = dQtde + objDataEntrega.dQuantidade
        Next
        
        lErro = Comando_Executar(lComando1, "SELECT PedCompra, Quantidade FROM ItensPedCompraN WHERE NumIntDoc = ? ", lNumIntPC, dQtdePC, objItemPC.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 213196
    
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213197
        
        If lErro <> AD_SQL_SUCESSO Then gError 213198
        
        If Abs(dQtde - dQtdePC) > QTDE_ESTOQUE_DELTA Then gError 213199
    
        For Each objDataEntrega In objItemPC.colDataEntrega
    
            'Insere o bloqueio na tabela de Bloqueios de Pedido de Venda
            lErro = Comando_Executar(lComando2, "INSERT INTO ItensPCEntrega(NumIntPC, NumIntItemPC, DataEntrega, Quantidade) VALUES (?,?,?,?)", lNumIntPC, objItemPC.lNumIntDoc, objDataEntrega.dtDataEntrega, objDataEntrega.dQuantidade)
            If lErro <> AD_SQL_SUCESSO Then gError 213200
    
        Next
        
    End If

    ItensPCEntrega_Insere = SUCESSO

    Exit Function

Erro_ItensPCEntrega_Insere:

    ItensPCEntrega_Insere = gErr

    Select Case gErr

        Case 213196, 213197
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)

        Case 213198
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", gErr, objItemPC.lNumIntDoc)

        Case 213199
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTSOLICITADA_DIFERE1", gErr, objItemPC.lNumIntDoc, objItemPC.sProduto, dQtdePC, dQtde)

        Case 213200
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_ITENSPCENTREGA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213201)

    End Select

    Exit Function

End Function

Private Function ItensPCEntrega_Exclui(ByVal lComando As Long, ByVal lNumIntPC As Long, ByVal lNumIntItemPC As Long) As Long
'exclui as  datas de entrega do pedido

Dim lErro As Long

On Error GoTo Erro_ItensPCEntrega_Exclui

    If lNumIntItemPC = 0 Then
        lErro = Comando_Executar(lComando, "DELETE FROM ItensPCEntrega WHERE NumIntPC = ?", lNumIntPC)
    Else
        lErro = Comando_Executar(lComando, "DELETE FROM ItensPCEntrega WHERE NumIntItemPC = ?", lNumIntItemPC)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 213202

    ItensPCEntrega_Exclui = SUCESSO

    Exit Function

Erro_ItensPCEntrega_Exclui:

    ItensPCEntrega_Exclui = gErr

    Select Case gErr

        Case 213202
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ITENSPCENTREGA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213203)

    End Select

    Exit Function

End Function

Public Function Config_Altera_Obj_COM(ByVal sNomeProperty As String, ByVal sConteudoNovo As String) As Long

Dim lErro As Long

On Error GoTo Erro_Config_Altera_Obj_COM

    Call CallByName(gobjCOM, sNomeProperty, VbLet, sConteudoNovo)

    Config_Altera_Obj_COM = SUCESSO

    Exit Function

Erro_Config_Altera_Obj_COM:

    Config_Altera_Obj_COM = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_ATRIBUICAO_VALOR_OBJ_GLOBAL", gErr, "gobjCOM", sNomeProperty, sConteudoNovo, Err.Description)
    
    End Select
    
    Exit Function
    
End Function
