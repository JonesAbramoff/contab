VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaFuncoes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Rotina_ReqComprasBaixar_Batch(sNomeArqParam As String, colReqComprasInfo As Collection) As Long
'calcula o custo médio de produção para mes/ano passados e valora os movimentos de estoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_ReqComprasBaixar_Batch
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_ROTINA_REQCOMPRASBAIXAR_BATCH", X)
    If lErro <> SUCESSO Then gError 63433
    
    
    Rotina_ReqComprasBaixar_Batch = X.Rotina_ReqComprasBaixar_Batch(sNomeArqParam, colReqComprasInfo)

    Exit Function
    
Erro_Rotina_ReqComprasBaixar_Batch:
    
    Rotina_ReqComprasBaixar_Batch = gErr
    
    Select Case gErr

        Case 63433
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145044)
        
    End Select
    
    Exit Function

End Function

Function Requisitante_Le(objRequisitante As ClassRequisitante) As Long
'Lê o Requisistante partir da chave em objRequisistante (Código) e retorna os dados em objRequisitante

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Requisitante_Le
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISITANTE_LE", X)
    If lErro <> SUCESSO Then gError 51140
    
    Requisitante_Le = X.Requisitante_Le(objRequisitante)
    
    Exit Function
    
Erro_Requisitante_Le:

    Requisitante_Le = gErr
    
    Select Case gErr

        Case 51140
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145045)
        
    End Select
    
    Exit Function
    
End Function

Function Requisitante_Le_NomeReduzido(objRequisitante As ClassRequisitante) As Long
'Lê o Requisistante partir da chave em objRequisistante (Código) e retorna os dados em objRequisitante

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Requisitante_Le_NomeReduzido
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISITANTE_LE_NOMEREDUZIDO", X)
    If lErro <> SUCESSO Then gError 51141
    
    Requisitante_Le_NomeReduzido = X.Requisitante_Le_NomeReduzido(objRequisitante)
    
    Exit Function
    
Erro_Requisitante_Le_NomeReduzido:

    Requisitante_Le_NomeReduzido = gErr
    
    Select Case gErr

        Case 51141
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145046)
        
    End Select
    
    Exit Function
    
End Function

Function LiberacaoDeBloqueiosPC_ObterBloqueios(objLiberacaoBloqueiosPC As ClassLiberacaoBloqueiosPC, colTiposBloqueio As Collection) As Long
'Preenche colBloqueioLiberacaoInfo em objLiberacaoBloqueiosPC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LiberacaoDeBloqueiosPC_ObterBloqueios
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_LIBERACAODEBLOQPC_OBTERBLOQUEIOS", X)
    If lErro <> SUCESSO Then gError 51175
    
    LiberacaoDeBloqueiosPC_ObterBloqueios = X.LiberacaoDeBloqueiosPC_ObterBloqueios(objLiberacaoBloqueiosPC, colTiposBloqueio)

    Exit Function
    
Erro_LiberacaoDeBloqueiosPC_ObterBloqueios:
    
    LiberacaoDeBloqueiosPC_ObterBloqueios = gErr
    
    Select Case gErr

        Case 51175
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145047)
        
    End Select
    
    Exit Function

End Function


Function Alcada_Le_Todas(colAlcada As Collection) As Long
'Guarda em colAlcada todas as alçadas cadastradas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Alcada_Le_Todas
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_ALCADA_LE_TODAS", X)
    If lErro <> SUCESSO Then gError 54695
    
    Alcada_Le_Todas = X.Alcada_Le_Todas(colAlcada)

    Exit Function
    
Erro_Alcada_Le_Todas:
    
    Alcada_Le_Todas = gErr
    
    Select Case gErr

        Case 54695
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145048)
        
    End Select
    
    Exit Function

End Function

Function Alcada_Le(objAlcada As ClassAlcada) As Long
'Lê a alcada a partir do código do usuário.
'Devolve os dados em objAlcada.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Alcada_Le
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_ALCADA_LE", X)
    If lErro <> SUCESSO Then gError 54696
    
    Alcada_Le = X.Alcada_Le(objAlcada)

    Exit Function
    
Erro_Alcada_Le:
    
    Alcada_Le = gErr
    
    Select Case gErr

        Case 54696
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145049)
        
    End Select
    
    Exit Function

End Function

Function Alcada_Exclui(objAlcada As ClassAlcada) As Long
'Exclui uma alçada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Alcada_Exclui
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_ALCADA_EXCLUI", X)
    If lErro <> SUCESSO Then gError 54697
    
    Alcada_Exclui = X.Alcada_Exclui(objAlcada)

    Exit Function
    
Erro_Alcada_Exclui:
    
    Alcada_Exclui = gErr
    
    Select Case gErr

        Case 54697
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145050)
        
    End Select
    
    Exit Function

End Function

Function Alcada_Grava(objAlcada As ClassAlcada) As Long
'Grava uma alçada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Alcada_Grava
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_ALCADA_GRAVA", X)
    If lErro <> SUCESSO Then gError 54698
    
    Alcada_Grava = X.Alcada_Grava(objAlcada)

    Exit Function
    
Erro_Alcada_Grava:
    
    Alcada_Grava = gErr
    
    Select Case gErr

        Case 54698
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145051)
        
    End Select
    
    Exit Function

End Function

Function BloqueiosPC_Libera(colBloqueiosPC As Collection) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPC_Libera
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_BLOQUEIOSPC_LIBERA", X)
    If lErro <> SUCESSO Then gError 54699
    
    BloqueiosPC_Libera = X.BloqueiosPC_Libera(colBloqueiosPC)

    Exit Function
    
Erro_BloqueiosPC_Libera:
    
    BloqueiosPC_Libera = gErr
    
    Select Case gErr

        Case 54699
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145052)
        
    End Select
    
    Exit Function

End Function

Function Alcada_Avalia(sCodUsuario As String, dValor As Double, dtData As Date) As Long
'Verifica se o usuario possui alcada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Alcada_Avalia
    
    
    lErro = Rotina_ObterObjetoFuncao("COM_ALCADA_AVALIA", X)
    If lErro <> SUCESSO Then gError 54700
    
    Alcada_Avalia = X.Alcada_Avalia(sCodUsuario, dValor, dtData)

    Exit Function
    
Erro_Alcada_Avalia:
    
    Alcada_Avalia = gErr
    
    Select Case gErr

        Case 54700
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145053)
        
    End Select
    
    Exit Function

End Function

Function ValorPCLiberado_Grava(dValor As Double, sCodUsuario As String, dtData) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ValorPCLiberado_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_VALORPCLIBERADO_GRAVA", X)
    If lErro <> SUCESSO Then gError 54701
    
    ValorPCLiberado_Grava = X.ValorPCLiberado_Grava(dValor, sCodUsuario, dtData)

    Exit Function
    
Erro_ValorPCLiberado_Grava:
    
    ValorPCLiberado_Grava = gErr
    
    Select Case gErr

        Case 54701
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145054)
        
    End Select
    
    Exit Function

End Function

Function Rotina_CalculaPtoPedido(sNomeArqParam As String) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_CalculaPtoPedido
    
    lErro = Rotina_ObterObjetoFuncao("COM_ROTINA_CALCULA_PTOPEDIDO", X)
    If lErro <> SUCESSO Then gError 69055
    
    Rotina_CalculaPtoPedido = X.Rotina_CalculaPtoPedido(sNomeArqParam)

    Exit Function
    
Erro_Rotina_CalculaPtoPedido:
    
    Rotina_CalculaPtoPedido = gErr
    
    Select Case gErr

        Case 69055
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145055)
        
    End Select
    
    Exit Function

End Function

Function Usuario_Le_Todos(colUsuarios As Collection) As Long
'le todos os usuários contidos na tabela de Usuario e coloca os dados em colUsuarios

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Usuario_Le_Todos
    
    lErro = Rotina_ObterObjetoFuncao("COM_USUARIO_LE_TODOS", X)
    If lErro <> SUCESSO Then gError 65490
    
    Usuario_Le_Todos = X.Usuario_Le_Todos(colUsuarios)

    Exit Function
    
Erro_Usuario_Le_Todos:
    
    Usuario_Le_Todos = gErr
    
    Select Case gErr

        Case 65490
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145056)
        
    End Select
    
    Exit Function

End Function

Function Usuario_Le(objUsuario As ClassUsuario) As Long
'le os dados do usuario passado como parametro e devolve os dados em objUsuario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Usuario_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_USUARIO_LE", X)
    If lErro <> SUCESSO Then gError 65491
    
    Usuario_Le = X.Usuario_Le(objUsuario)

    Exit Function
    
Erro_Usuario_Le:
    
    Usuario_Le = gErr
    
    Select Case gErr

        Case 65491
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145057)
        
    End Select
    
    Exit Function

End Function

Function Usuario_Le_NomeRed(objUsuario As ClassUsuario) As Long
'Le o usuário a partir do nome reduzido passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Usuario_Le_NomeRed
    
    lErro = Rotina_ObterObjetoFuncao("COM_USUARIO_LE_NOMERED", X)
    If lErro <> SUCESSO Then gError 65492
    
    Usuario_Le_NomeRed = X.Usuario_Le_NomeRed(objUsuario)

    Exit Function
    
Erro_Usuario_Le_NomeRed:
    
    Usuario_Le_NomeRed = gErr
    
    Select Case gErr

        Case 65492
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145058)
        
    End Select
    
    Exit Function

End Function

Function BaixaReqCompras_ObterRequisicoes(objBaixaReqCompras As ClassBaixaReqCompras) As Long
'Obtem as Requisicoes de Compras a Baixar de acordo
'com a seleção determinada na tela

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BaixaReqCompras_ObterRequisicoes
    
    lErro = Rotina_ObterObjetoFuncao("COM_BAIXAREQCOMPRAS_OBTERREQUISICOES", X)
    If lErro <> SUCESSO Then gError 65493
    
    BaixaReqCompras_ObterRequisicoes = X.BaixaReqCompras_ObterRequisicoes(objBaixaReqCompras)

    Exit Function
    
Erro_BaixaReqCompras_ObterRequisicoes:
    
    BaixaReqCompras_ObterRequisicoes = gErr
    
    Select Case gErr

        Case 65493
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145059)
        
    End Select
    
    Exit Function

End Function

Function BaixaPedCompras_ObterPedidos(objBaixaPedCompra As ClassBaixaPedCompra) As Long
'Preenche a colecao em objBaixaPedCompra com os pedidos que podem ser faturados e atendem
'aos criterios passados em objBaixaPedCompra com restricoes de No. de pedido, data de emissao,...

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BaixaPedCompras_ObterPedidos
    
    lErro = Rotina_ObterObjetoFuncao("COM_BAIXAPEDCOMPRAS_OBTERPEDIDOS", X)
    If lErro <> SUCESSO Then gError 65494
    
    BaixaPedCompras_ObterPedidos = X.BaixaPedCompras_ObterPedidos(objBaixaPedCompra)

    Exit Function
    
Erro_BaixaPedCompras_ObterPedidos:
    
    BaixaPedCompras_ObterPedidos = gErr
    
    Select Case gErr

        Case 65494
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145060)
        
    End Select
    
    Exit Function

End Function

Function PedidoComprasBaixar_Batch(colPedCompras As Collection) As Long
'transfere os Pedidos de Compra p/Pedidos de Compra Baixados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoComprasBaixar_Batch
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRASBAIXAR_BATCH", X)
    If lErro <> SUCESSO Then gError 65495
    
    PedidoComprasBaixar_Batch = X.PedidoComprasBaixar_Batch(colPedCompras)

    Exit Function
    
Erro_PedidoComprasBaixar_Batch:
    
    PedidoComprasBaixar_Batch = gErr
    
    Select Case gErr

        Case 65495
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145061)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Baixar1(objPedidoCompra As ClassPedidoCompras, alComando() As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Baixar1
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_BAIXAR1", X)
    If lErro <> SUCESSO Then gError 65496
    
    PedidoCompra_Baixar1 = X.PedidoCompra_Baixar1(objPedidoCompra, alComando())

    Exit Function
    
Erro_PedidoCompra_Baixar1:
    
    PedidoCompra_Baixar1 = gErr
    
    Select Case gErr

        Case 65496
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145062)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompraBaixado_Insere(lComando As Long, objPedidoCompra As ClassPedidoCompras) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompraBaixado_Insere
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRABAIXADO_INSERE", X)
    If lErro <> SUCESSO Then gError 65497
    
    PedidoCompraBaixado_Insere = X.PedidoCompraBaixado_Insere(lComando, objPedidoCompra)

    Exit Function
    
Erro_PedidoCompraBaixado_Insere:
    
    PedidoCompraBaixado_Insere = gErr
    
    Select Case gErr

        Case 65497
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145063)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_LeLock(objPedidoCompra_BD As ClassPedidoCompras) As Long
'Le o Pedido de Compra passado como parametro e locka,se encontrar

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_LeLock
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_LELOCK", X)
    If lErro <> SUCESSO Then gError 65498
    
    PedidoCompra_LeLock = X.PedidoCompra_LeLock(objPedidoCompra_BD)

    Exit Function
    
Erro_PedidoCompra_LeLock:
    
    PedidoCompra_LeLock = gErr
    
    Select Case gErr

        Case 65498
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145064)
        
    End Select
    
    Exit Function

End Function
Function PedidoCompra_LeMod(objPedidoCompra_BD As ClassPedidoCompras, lComando As Long) As Long
'Le o Pedido de Compra passado como parametro e locka,se encontrar

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_LeMod
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_LEMOD", X)
    If lErro <> SUCESSO Then gError 62643
    
    PedidoCompra_LeMod = X.PedidoCompra_LeMod(objPedidoCompra_BD, lComando)

    Exit Function
    
Erro_PedidoCompra_LeMod:
    
    PedidoCompra_LeMod = gErr
    
    Select Case gErr

        Case 62643
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145065)
        
    End Select
    
    Exit Function

End Function

Function ItensPC_Le(objPedidoCompra As ClassPedidoCompras) As Long
'Lê no BD a coleção de ítens do Pedido de Compra e devolve em colItensPedido dentro de objPedidoCompra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPC_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPC_LE", X)
    If lErro <> SUCESSO Then gError 65499
    
    ItensPC_Le = X.ItensPC_Le(objPedidoCompra)

    Exit Function
    
Erro_ItensPC_Le:
    
    ItensPC_Le = gErr
    
    Select Case gErr

        Case 65499
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145066)
        
    End Select
    
    Exit Function

End Function

Function LocalizacaoItemPC_Le(objItemPC As ClassItemPedCompra) As Long
'Le ItemPC na tabela de LocalizacaoItemPC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LocalizacaoItemPC_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_LOCALIZACAOITEMPC_LE", X)
    If lErro <> SUCESSO Then gError 71000
    
    LocalizacaoItemPC_Le = X.LocalizacaoItemPC_Le(objItemPC)

    Exit Function
    
Erro_LocalizacaoItemPC_Le:
    
    LocalizacaoItemPC_Le = gErr
    
    Select Case gErr

        Case 71000
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145067)
        
    End Select
    
    Exit Function

End Function

Function Produto_Le_Residual(objProduto As ClassProduto, dResiduo As Double) As Long
'Verifica o Valor do Residuo para o Produto passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Residual
    
    lErro = Rotina_ObterObjetoFuncao("COM_PRODUTO_LE_RESIDUAL", X)
    If lErro <> SUCESSO Then gError 71001
    
    Produto_Le_Residual = X.Produto_Le_Residual(objProduto, dResiduo)

    Exit Function
    
Erro_Produto_Le_Residual:
    
    Produto_Le_Residual = gErr
    
    Select Case gErr

        Case 71001
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145068)
        
    End Select
    
    Exit Function

End Function

Function Valor_Produtos_Recebido(objPedidoCompra As ClassPedidoCompras, dValorProdutoRecebido As Double)
'Calcula o valor recebido efetivo dos produtos do Pedido de Compra passado como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Valor_Produtos_Recebido
    
    lErro = Rotina_ObterObjetoFuncao("COM_VALOR_PRODUTOS_RECEBIDO", X)
    If lErro <> SUCESSO Then gError 71002
    
    Valor_Produtos_Recebido = X.Valor_Produtos_Recebido(objPedidoCompra, dValorProdutoRecebido)

    Exit Function
    
Erro_Valor_Produtos_Recebido:
    
    Valor_Produtos_Recebido = gErr
    
    Select Case gErr

        Case 71002
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145069)
        
    End Select
    
    Exit Function

End Function

Function PedidoCotacao_Baixar(objPedCotacao As ClassPedidoCotacao, alComando() As Long) As Long
'Baixa o Pedido de Cotação passado em objPedCotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCotacao_Baixar
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOTACAO_BAIXAR", X)
    If lErro <> SUCESSO Then gError 71003
    
    PedidoCotacao_Baixar = X.PedidoCotacao_Baixar(objPedCotacao, alComando())

    Exit Function
    
Erro_PedidoCotacao_Baixar:
    
    PedidoCotacao_Baixar = gErr
    
    Select Case gErr

        Case 71003
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145070)
        
    End Select
    
    Exit Function

End Function

Function Cotacao_Baixa(lCotacao As Long, alComando() As Long) As Long
'Baixa a Cotação e seus Itens

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cotacao_Baixa
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAO_BAIXA", X)
    If lErro <> SUCESSO Then gError 71004
    
    Cotacao_Baixa = X.Cotacao_Baixa(lCotacao, alComando())

    Exit Function
    
Erro_Cotacao_Baixa:
    
    Cotacao_Baixa = gErr
    
    Select Case gErr

        Case 71004
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145071)
        
    End Select
    
    Exit Function

End Function

Function BaixaPedCotacao_ObterPedidos(objBaixaPedCotacao As ClassBaixaPedCotacao) As Long
'Preenche a colecao em objPedCotacao com os pedidos que possuem Data de Emissão preenhida e
'possuem objetos associados (concorrência, Pedido de Compras) baixados,
'e atendem aos criterios passados na seleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BaixaPedCotacao_ObterPedidos
    
    lErro = Rotina_ObterObjetoFuncao("COM_BAIXAPEDCOTACAO_OBTERPEDIDOS", X)
    If lErro <> SUCESSO Then gError 71005
    
    BaixaPedCotacao_ObterPedidos = X.BaixaPedCotacao_ObterPedidos(objBaixaPedCotacao)

    Exit Function
    
Erro_BaixaPedCotacao_ObterPedidos:
    
    BaixaPedCotacao_ObterPedidos = gErr
    
    Select Case gErr

        Case 71005
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145072)
        
    End Select
    
    Exit Function

End Function

Function BaixaPedCotacao_Baixar_Pedidos(colPedCotacao As Collection) As Long
'Faz baixa dos Pedidos de Cotação passados na coleção colPedCotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BaixaPedCotacao_Baixar_Pedidos
    
    lErro = Rotina_ObterObjetoFuncao("COM_BAIXAPEDCOTACAO_BAIXAR_PEDIDOS", X)
    If lErro <> SUCESSO Then gError 71006
    
    BaixaPedCotacao_Baixar_Pedidos = X.BaixaPedCotacao_Baixar_Pedidos(colPedCotacao)

    Exit Function
    
Erro_BaixaPedCotacao_Baixar_Pedidos:
    
    BaixaPedCotacao_Baixar_Pedidos = gErr
    
    Select Case gErr

        Case 71006
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145073)
        
    End Select
    
    Exit Function

End Function

Function ItensReqCompras_Le(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê os Ítens da Requisição Compras passada por objRequisiçãoCompras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqCompras_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSREQCOMPRAS_LE", X)
    If lErro <> SUCESSO Then gError 71007
    
    ItensReqCompras_Le = X.ItensReqCompras_Le(objRequisicaoCompras)

    Exit Function
    
Erro_ItensReqCompras_Le:
    
    ItensReqCompras_Le = gErr
    
    Select Case gErr

        Case 71007
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145074)
        
    End Select
    
    Exit Function

End Function

Function Usuarios_Le_Nome(objUsuarios As ClassUsuarios) As Long
'le o usuário com o Nome passado e coloca no objUsuarios

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Usuarios_Le_Nome
    
    lErro = Rotina_ObterObjetoFuncao("COM_USUARIOS_LE_NOME", X)
    If lErro <> SUCESSO Then gError 71008
    
    Usuarios_Le_Nome = X.Usuarios_Le_Nome(objUsuarios)

    Exit Function
    
Erro_Usuarios_Le_Nome:
    
    Usuarios_Le_Nome = gErr
    
    Select Case gErr

        Case 71008
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145075)
        
    End Select
    
    Exit Function

End Function

Function Usuarios_Le_NomeRed(objUsuarios As ClassUsuarios) As Long
'le o usuário com o Nome passado e coloca no objUsuarios

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Usuarios_Le_NomeRed
    
    lErro = Rotina_ObterObjetoFuncao("COM_USUARIOS_LE_NOMERED", X)
    If lErro <> SUCESSO Then gError 71009
    
    Usuarios_Le_NomeRed = X.Usuarios_Le_NomeRed(objUsuarios)

    Exit Function
    
Erro_Usuarios_Le_NomeRed:
    
    Usuarios_Le_NomeRed = gErr
    
    Select Case gErr

        Case 71009
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145076)
        
    End Select
    
    Exit Function

End Function

Function Comprador_Existe_Relacionamento(objComprador As ClassComprador) As Long
'Verifica se existe relacionamento com as tabelas Concorrencia,ConcorrenciaBaixada, PedidoCompra, PedidoCompraBaixado, Cotacao, CotacaoBaixada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comprador_Existe_Relacionamento
    
    lErro = Rotina_ObterObjetoFuncao("COM_COMPRADOR_EXISTE_RELACIONAMENTO", X)
    If lErro <> SUCESSO Then gError 71010
    
    Comprador_Existe_Relacionamento = X.Comprador_Existe_Relacionamento(objComprador)

    Exit Function
    
Erro_Comprador_Existe_Relacionamento:
    
    Comprador_Existe_Relacionamento = gErr
    
    Select Case gErr

        Case 71010
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145077)
        
    End Select
    
    Exit Function

End Function

Function Comprador_Exclui(objComprador As ClassComprador) As Long
'Exclui o comprador da tabela

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comprador_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_COMPRADOR_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71011
    
    Comprador_Exclui = X.Comprador_Exclui(objComprador)

    Exit Function
    
Erro_Comprador_Exclui:
    
    Comprador_Exclui = gErr
    
    Select Case gErr

        Case 71011
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145078)
        
    End Select
    
    Exit Function

End Function

Function Comprador_Grava(objComprador As ClassComprador) As Long
'Grava o Comprador objComprador

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comprador_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_COMPRADOR_GRAVA", X)
    If lErro <> SUCESSO Then gError 71012
    
    Comprador_Grava = X.Comprador_Grava(objComprador)

    Exit Function
    
Erro_Comprador_Grava:
    
    Comprador_Grava = gErr
    
    Select Case gErr

        Case 71012
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145079)
        
    End Select
    
    Exit Function

End Function

Function Comprador_Le(objComprador As ClassComprador) As Long
'Le um comprador

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comprador_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_COMPRADOR_LE", X)
    If lErro <> SUCESSO Then gError 71013
    
    Comprador_Le = X.Comprador_Le(objComprador)

    Exit Function
    
Erro_Comprador_Le:
    
    Comprador_Le = gErr
    
    Select Case gErr

        Case 71013
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145080)
        
    End Select
    
    Exit Function

End Function

Function Comprador_Le_Todos(colComprador As Collection) As Long
'Lê todos os compradores da gifilialempresa e guarda em colComprador
'OK ** Mal descrito. FilialEmpresa ?

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comprador_Le_Todos
    
    lErro = Rotina_ObterObjetoFuncao("COM_COMPRADOR_LE_TODOS", X)
    If lErro <> SUCESSO Then gError 71014
    
    Comprador_Le_Todos = X.Comprador_Le_Todos(colComprador)

    Exit Function
    
Erro_Comprador_Le_Todos:
    
    Comprador_Le_Todos = gErr
    
    Select Case gErr

        Case 71014
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145081)
        
    End Select
    
    Exit Function

End Function

Function Comprador_Le_Usuario(objComprador As ClassComprador) As Long
'Le o comprador de um usuario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Comprador_Le_Usuario
    
    lErro = Rotina_ObterObjetoFuncao("COM_COMPRADOR_LE_USUARIO", X)
    If lErro <> SUCESSO Then gError 71015
    
    Comprador_Le_Usuario = X.Comprador_Le_Usuario(objComprador)

    Exit Function
    
Erro_Comprador_Le_Usuario:
    
    Comprador_Le_Usuario = gErr
    
    Select Case gErr

        Case 71015
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145082)
        
    End Select
    
    Exit Function

End Function

Function ItensConcorrenciaTodos_Le(objConcorrencia As ClassConcorrencia) As Long
'Lê os itens da concorrencia (baixada ou não) passada como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensConcorrenciaTodos_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSCONCORRENCIATODOS_LE", X)
    If lErro <> SUCESSO Then gError 71016
    
    ItensConcorrenciaTodos_Le = X.ItensConcorrenciaTodos_Le(objConcorrencia)

    Exit Function
    
Erro_ItensConcorrenciaTodos_Le:
    
    ItensConcorrenciaTodos_Le = gErr
    
    Select Case gErr

        Case 71016
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145083)
        
    End Select
    
    Exit Function

End Function

Function PedidosCotacaoTodos_Le_Concorrencia(objConcorrencia As ClassConcorrencia, colItensCotacao As Collection) As Long
'Lê os itens de cotação (baixados ou não) vinculadas a concorrência

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosCotacaoTodos_Le_Concorrencia
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOSCOTACAOTODOS_LE_CONCORRENCIA", X)
    If lErro <> SUCESSO Then gError 71017
    
    PedidosCotacaoTodos_Le_Concorrencia = X.PedidosCotacaoTodos_Le_Concorrencia(objConcorrencia, colItensCotacao)

    Exit Function
    
Erro_PedidosCotacaoTodos_Le_Concorrencia:
    
    PedidosCotacaoTodos_Le_Concorrencia = gErr
    
    Select Case gErr

        Case 71017
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145084)
        
    End Select
    
    Exit Function

End Function

Function TipoTributacaoPadrao_Le(iTipoTrib As Integer) As Long
'Lê os Tipo de Tributação Padrão a partir das Tabelas NaOpPadrao e FiliaisEmpresa
'e devolve em iTipoTrib

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoTributacaoPadrao_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_TIPOTRIBUTACAOPADRAO_LE", X)
    If lErro <> SUCESSO Then gError 71018
    
    TipoTributacaoPadrao_Le = X.TipoTributacaoPadrao_Le(iTipoTrib)

    Exit Function
    
Erro_TipoTributacaoPadrao_Le:
    
    TipoTributacaoPadrao_Le = gErr
    
    Select Case gErr

        Case 71018
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145085)
        
    End Select
    
    Exit Function

End Function

Function TiposTributacao_Le(iTipo As Integer, sDescricao As String) As Long
'Lê Tipos de Tributação a partir do Tipo passado em iTipo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposTributacao_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_TIPOSTRIBUTACAO_LE", X)
    If lErro <> SUCESSO Then gError 71019
    
    TiposTributacao_Le = X.TiposTributacao_Le(iTipo, sDescricao)

    Exit Function
    
Erro_TiposTributacao_Le:
    
    TiposTributacao_Le = gErr
    
    Select Case gErr

        Case 71019
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145086)
        
    End Select
    
    Exit Function

End Function

Function Observacao_Le(objObservacao As ClassObservacao) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Observacao_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_OBSERVACAO_LE", X)
    If lErro <> SUCESSO Then gError 71020
    
    Observacao_Le = X.Observacao_Le(objObservacao)

    Exit Function
    
Erro_Observacao_Le:
    
    Observacao_Le = gErr
    
    Select Case gErr

        Case 71020
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145087)
        
    End Select
    
    Exit Function

End Function

Function ItensReqComprasTodos_Le(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê os Ítens da Requisição Compras (baixada ou não) passada por objRequisiçãoCompras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqComprasTodos_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSREQCOMPRASTODOS_LE", X)
    If lErro <> SUCESSO Then gError 71021
    
    ItensReqComprasTodos_Le = X.ItensReqComprasTodos_Le(objRequisicaoCompras)

    Exit Function
    
Erro_ItensReqComprasTodos_Le:
    
    ItensReqComprasTodos_Le = gErr
    
    Select Case gErr

        Case 71021
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145088)
        
    End Select
    
    Exit Function

End Function

Function CondicaoPagto_Le_DescReduzida(objCondicaoPagto As ClassCondicaoPagto) As Long
'le a Condicao de Pagto a partir da descrição reduzida

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CondicaoPagto_Le_DescReduzida
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONDICAOPAGTO_LE_DESCREDUZIDA", X)
    If lErro <> SUCESSO Then gError 71022
    
    CondicaoPagto_Le_DescReduzida = X.CondicaoPagto_Le_DescReduzida(objCondicaoPagto)

    Exit Function
    
Erro_CondicaoPagto_Le_DescReduzida:
    
    CondicaoPagto_Le_DescReduzida = gErr
    
    Select Case gErr

        Case 71022
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145089)
        
    End Select
    
    Exit Function

End Function

Function CotacaoItemConcorrencia_Le(objItemConcorrencia As ClassItemConcorrencia) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CotacaoItemConcorrencia_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAOITEMCONCORRENCIA_LE", X)
    If lErro <> SUCESSO Then gError 71023
    
    CotacaoItemConcorrencia_Le = X.CotacaoItemConcorrencia_Le(objItemConcorrencia)

    Exit Function
    
Erro_CotacaoItemConcorrencia_Le:
    
    CotacaoItemConcorrencia_Le = gErr
    
    Select Case gErr

        Case 71023
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145090)
        
    End Select
    
    Exit Function

End Function

Function ItemRCItemConcorrencia_Le(objItemConcorrencia As ClassItemConcorrencia) As Long
'Busca em ItemRCItemConcorrencia o link entre os itens de concorrência e os itens de requisição

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemRCItemConcorrencia_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITEMRCITEMCONCORRENCIA_LE", X)
    If lErro <> SUCESSO Then gError 71024
    
    ItemRCItemConcorrencia_Le = X.ItemRCItemConcorrencia_Le(objItemConcorrencia)

    Exit Function
    
Erro_ItemRCItemConcorrencia_Le:
    
    ItemRCItemConcorrencia_Le = gErr
    
    Select Case gErr

        Case 71024
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145091)
        
    End Select
    
    Exit Function

End Function

Function QuantSuplementar_Le(objItemConcorrencia As ClassItemConcorrencia) As Long
'Lê as quantidades Suplementares associadas ao Item de Concorrência passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_QuantSuplementar_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_QUANTSUPLEMENTAR_LE", X)
    If lErro <> SUCESSO Then gError 71025
    
    QuantSuplementar_Le = X.QuantSuplementar_Le(objItemConcorrencia)

    Exit Function
    
Erro_QuantSuplementar_Le:
    
    QuantSuplementar_Le = gErr
    
    Select Case gErr

        Case 71025
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145092)
        
    End Select
    
    Exit Function

End Function

Function Motivo_Le_Descricao(objMotivo As AdmCodigoNome)
'Lê o Motivo a partir da descrição passada na tabela Motivo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Motivo_Le_Descricao
    
    lErro = Rotina_ObterObjetoFuncao("COM_MOTIVO_LE_DESCRICAO", X)
    If lErro <> SUCESSO Then gError 71026
    
    Motivo_Le_Descricao = X.Motivo_Le_Descricao(objMotivo)

    Exit Function
    
Erro_Motivo_Le_Descricao:
    
    Motivo_Le_Descricao = gErr
    
    Select Case gErr

        Case 71026
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145093)
        
    End Select
    
    Exit Function

End Function

Function RequisicoesTodas_Le_Concorrencia(objConcorrencia As ClassConcorrencia, colRequisicoes As Collection) As Long
'Lê as Requisições (baixadas ou não) vinculadas aos ItemConcorrência da Concorrência e devolve elas em colRequisicoes

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicoesTodas_Le_Concorrencia
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICOESTODAS_LE_CONCORRENCIA", X)
    If lErro <> SUCESSO Then gError 71027
    
    RequisicoesTodas_Le_Concorrencia = X.RequisicoesTodas_Le_Concorrencia(objConcorrencia, colRequisicoes)

    Exit Function
    
Erro_RequisicoesTodas_Le_Concorrencia:
    
    RequisicoesTodas_Le_Concorrencia = gErr
    
    Select Case gErr

        Case 71027
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145094)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompraTodos_Le_Concorrencia(objConcorrencia As ClassConcorrencia, colPedidos As Collection) As Long
'Lê os Pedidos vinculados aos ItemConcorrência da Concorrência passada e devolve eles em colPedidos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompraTodos_Le_Concorrencia
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRATODOS_LE_CONCORRENCIA", X)
    If lErro <> SUCESSO Then gError 71028
    
    PedidoCompraTodos_Le_Concorrencia = X.PedidoCompraTodos_Le_Concorrencia(objConcorrencia, colPedidos)

    Exit Function
    
Erro_PedidoCompraTodos_Le_Concorrencia:
    
    PedidoCompraTodos_Le_Concorrencia = gErr
    
    Select Case gErr

        Case 71028
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145095)
        
    End Select
    
    Exit Function

End Function

Function ConfiguraCOM_Gravar(objConfiguraCOM As ClassConfiguraCOM) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ConfiguraCOM_Gravar
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONFIGURACOM_GRAVAR", X)
    If lErro <> SUCESSO Then gError 71029
    
    ConfiguraCOM_Gravar = X.ConfiguraCOM_Gravar(objConfiguraCOM)

    Exit Function
    
Erro_ConfiguraCOM_Gravar:
    
    ConfiguraCOM_Gravar = gErr
    
    Select Case gErr

        Case 71029
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145096)
        
    End Select
    
    Exit Function

End Function
Function ConfiguraCOM_GravarTrans(objConfiguraCOM As ClassConfiguraCOM) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ConfiguraCOM_GravarTrans
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONFIGURACOM_GRAVAR_TRANS", X)
    If lErro <> SUCESSO Then gError 74932
    
    ConfiguraCOM_GravarTrans = X.ConfiguraCOM_GravarTrans(objConfiguraCOM)

    Exit Function
    
Erro_ConfiguraCOM_GravarTrans:
    
    ConfiguraCOM_GravarTrans = gErr
    
    Select Case gErr

        Case 74932
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145097)
        
    End Select
    
    Exit Function

End Function

Function Concorrencia_Automatica(lConcorrencia As Long) As Long
'Devolve o proximo código de concorrencia disponivel para a FilialEmpresa em questao.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Concorrencia_Automatica
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONCORRENCIA_AUTOMATICA", X)
    If lErro <> SUCESSO Then gError 71033
    
    Concorrencia_Automatica = X.Concorrencia_Automatica(lConcorrencia)

    Exit Function
    
Erro_Concorrencia_Automatica:
    
    Concorrencia_Automatica = gErr
    
    Select Case gErr

        Case 71033
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145098)
        
    End Select
    
    Exit Function

End Function

Function TiposProduto_Le_Compravel(colCod_DescReduzida As AdmColCodigoNome) As Long
'Le todos os Codigos e DescReduzida dos tipos de produto que podem ser comprados (Compras=1) da tabela TiposDeProduto e coloca na colecao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposProduto_Le_Compravel
    
    lErro = Rotina_ObterObjetoFuncao("COM_TIPOSPRODUTO_LE_COMPRAVEL", X)
    If lErro <> SUCESSO Then gError 71034
    
    TiposProduto_Le_Compravel = X.TiposProduto_Le_Compravel(colCod_DescReduzida)

    Exit Function
    
Erro_TiposProduto_Le_Compravel:
    
    TiposProduto_Le_Compravel = gErr
    
    Select Case gErr

        Case 71034
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145099)
        
    End Select
    
    Exit Function

End Function

Function Concorrencia_Grava(objConcorrencia As ClassConcorrencia) As Long
'Grava concorrência

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Concorrencia_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONCORRENCIA_GRAVA", X)
    If lErro <> SUCESSO Then gError 71035
    
    Concorrencia_Grava = X.Concorrencia_Grava(objConcorrencia)

    Exit Function
    
Erro_Concorrencia_Grava:
    
    Concorrencia_Grava = gErr
    
    Select Case gErr

        Case 71035
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145100)
        
    End Select
    
    Exit Function

End Function

Function Concorrencia_Grava_EmTrans(objConcorrencia As ClassConcorrencia) As Long
'Grava a Concorrência e seus itens

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Concorrencia_Grava_EmTrans
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONCORRENCIA_GRAVA_EMTRANS", X)
    If lErro <> SUCESSO Then gError 71036
    
    Concorrencia_Grava_EmTrans = X.Concorrencia_Grava_EmTrans(objConcorrencia)

    Exit Function
    
Erro_Concorrencia_Grava_EmTrans:
    
    Concorrencia_Grava_EmTrans = gErr
    
    Select Case gErr

        Case 71036
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145101)
        
    End Select
    
    Exit Function

End Function

Function ItensConcorrencia_Exclui(objConcorrencia As ClassConcorrencia) As Long
'Exclui itens de concorrência e tabelas associadas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensConcorrencia_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSCONCORRENCIA_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71037
    
    ItensConcorrencia_Exclui = X.ItensConcorrencia_Exclui(objConcorrencia)

    Exit Function
    
Erro_ItensConcorrencia_Exclui:
    
    ItensConcorrencia_Exclui = gErr
    
    Select Case gErr

        Case 71037
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145102)
        
    End Select
    
    Exit Function

End Function

Function ItensConcorrencia_Exclui_Relacionamento(lNumIntConc As Long) As Long
'Exclui vínculos do Item de Concorrência com CotaçãoItemConcorrência, ItemRCItemConcorrência
'e QuantidadesSuplementates

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensConcorrencia_Exclui_Relacionamento
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSCONCORRENCIA_EXCLUI_RELACIONAME", X)
    If lErro <> SUCESSO Then gError 71038
    
    ItensConcorrencia_Exclui_Relacionamento = X.ItensConcorrencia_Exclui_Relacionamento(lNumIntConc)

    Exit Function
    
Erro_ItensConcorrencia_Exclui_Relacionamento:
    
    ItensConcorrencia_Exclui_Relacionamento = gErr
    
    Select Case gErr

        Case 71038
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145103)
        
    End Select
    
    Exit Function

End Function

Function ItensConcorrencia_Inclui(objConcorrencia As ClassConcorrencia) As Long
'Grava novos Itens de Concorrência

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensConcorrencia_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSCONCORRENCIA_INCLUI", X)
    If lErro <> SUCESSO Then gError 71039
    
    ItensConcorrencia_Inclui = X.ItensConcorrencia_Inclui(objConcorrencia)

    Exit Function
    
Erro_ItensConcorrencia_Inclui:
    
    ItensConcorrencia_Inclui = gErr
    
    Select Case gErr

        Case 71039
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145104)
        
    End Select
    
    Exit Function

End Function

Function PedCompra_Concorrencia_Grava(objConcorrencia As ClassConcorrencia, colPedidoCompras As Collection) As Long
'Atualiza Concorrência e grava Pedidos de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedCompra_Concorrencia_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDCOMPRA_CONCORRENCIA_GRAVA", X)
    If lErro <> SUCESSO Then gError 71040
    
    PedCompra_Concorrencia_Grava = X.PedCompra_Concorrencia_Grava(objConcorrencia, colPedidoCompras)

    Exit Function
    
Erro_PedCompra_Concorrencia_Grava:
    
    PedCompra_Concorrencia_Grava = gErr
    
    Select Case gErr

        Case 71040
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145105)
        
    End Select
    
    Exit Function

End Function

Function PedidosCompra_Grava_Concorrencia_EmTrans(colPedidoCompras As Collection, objConcorrencia As ClassConcorrencia) As Long
'Grava Pedidos de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosCompra_Grava_Concorrencia_EmTrans
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOSCOMPRA_GRAVA_CONCORRENCIA_EMT", X)
    If lErro <> SUCESSO Then gError 71041
    
    PedidosCompra_Grava_Concorrencia_EmTrans = X.PedidosCompra_Grava_Concorrencia_EmTrans(colPedidoCompras, objConcorrencia)

    Exit Function
    
Erro_PedidosCompra_Grava_Concorrencia_EmTrans:
    
    PedidosCompra_Grava_Concorrencia_EmTrans = gErr
    
    Select Case gErr

        Case 71041
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145106)
        
    End Select
    
    Exit Function

End Function

Function Concorrencia_ObterNumIntAutomatico(lNumInt As Long) As Long
'gera o próximo número interno disponível para Concorrencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Concorrencia_ObterNumIntAutomatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONCORRENCIA_OBTERNUMINTAUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71042
    
    Concorrencia_ObterNumIntAutomatico = X.Concorrencia_ObterNumIntAutomatico(lNumInt)

    Exit Function
    
Erro_Concorrencia_ObterNumIntAutomatico:
    
    Concorrencia_ObterNumIntAutomatico = gErr
    
    Select Case gErr

        Case 71042
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145107)
        
    End Select
    
    Exit Function

End Function

Function ItemConcorrencia_ObterNumIntAutomatico(lNumInt As Long) As Long
'gera o próximo número interno disponível para ItemConcorrencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemConcorrencia_ObterNumIntAutomatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITEMCONCORRENCIA_OBTERNUMINTAUTOMATI", X)
    If lErro <> SUCESSO Then gError 71043
    
    ItemConcorrencia_ObterNumIntAutomatico = X.ItemConcorrencia_ObterNumIntAutomatico(lNumInt)

    Exit Function
    
Erro_ItemConcorrencia_ObterNumIntAutomatico:
    
    ItemConcorrencia_ObterNumIntAutomatico = gErr
    
    Select Case gErr

        Case 71043
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145108)
        
    End Select
    
    Exit Function

End Function

Function CotacaoItemConc_ObterNumIntAutomatico(lNumInt As Long) As Long
'gera o próximo número interno disponível para CotaçãoItenConc

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CotacaoItemConc_ObterNumIntAutomatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAOITEMCONC_OBTERNUMINTAUTOMATIC", X)
    If lErro <> SUCESSO Then gError 71044
    
    CotacaoItemConc_ObterNumIntAutomatico = X.CotacaoItemConc_ObterNumIntAutomatico(lNumInt)

    Exit Function
    
Erro_CotacaoItemConc_ObterNumIntAutomatico:
    
    CotacaoItemConc_ObterNumIntAutomatico = gErr
    
    Select Case gErr

        Case 71044
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145109)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Codigo_Automatico(lCodigo As Long) As Long
'Gera o proximo codigo da Tabela de Pedido Compra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Codigo_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_CODIGO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71045
    
    PedidoCompra_Codigo_Automatico = X.PedidoCompra_Codigo_Automatico(lCodigo)

    Exit Function
    
Erro_PedidoCompra_Codigo_Automatico:
    
    PedidoCompra_Codigo_Automatico = gErr
    
    Select Case gErr

        Case 71045
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145110)
        
    End Select
    
    Exit Function

End Function

Function PedidosCotacao_Le_Requisicao(sProduto As String, colItensCotacao As Collection) As Long
'Lê os Pedidos de Cotacao para os Produto passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosCotacao_Le_Requisicao
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOSCOTACAO_LE_REQUISICAO", X)
    If lErro <> SUCESSO Then gError 71046
    
    PedidosCotacao_Le_Requisicao = X.PedidosCotacao_Le_Requisicao(sProduto, colItensCotacao)

    Exit Function
    
Erro_PedidosCotacao_Le_Requisicao:
    
    PedidosCotacao_Le_Requisicao = gErr
    
    Select Case gErr

        Case 71046
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145111)
        
    End Select
    
    Exit Function

End Function

Function Requisicoes_Le_GeracaoPC(gobjGeracaoPedCompraReq As ClassGeracaoPedCompraReq) As Long
'Obtem as Requisicoes de Compras de acordo com a seleção determinada na tela

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Requisicoes_Le_GeracaoPC
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICOES_LE_GERACAOPC", X)
    If lErro <> SUCESSO Then gError 71047
    
    Requisicoes_Le_GeracaoPC = X.Requisicoes_Le_GeracaoPC(gobjGeracaoPedCompraReq)

    Exit Function
    
Erro_Requisicoes_Le_GeracaoPC:
    
    Requisicoes_Le_GeracaoPC = gErr
    
    Select Case gErr

        Case 71047
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145112)
        
    End Select
    
    Exit Function

End Function

Function PedidoCotacao_Le(objPedidoCotacao As ClassPedidoCotacao) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCotacao_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOTACAO_LE", X)
    If lErro <> SUCESSO Then gError 71048
    
    PedidoCotacao_Le = X.PedidoCotacao_Le(objPedidoCotacao)

    Exit Function
    
Erro_PedidoCotacao_Le:
    
    PedidoCotacao_Le = gErr
    
    Select Case gErr

        Case 71048
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145113)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCotacao_Le(objPedidoCotacao As ClassPedidoCotacao) As Long
'Lê os ítens do Pedido de Cotação

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCotacao_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPEDCOTACAO_LE", X)
    If lErro <> SUCESSO Then gError 71049
    
    ItensPedCotacao_Le = X.ItensPedCotacao_Le(objPedidoCotacao)

    Exit Function
    
Erro_ItensPedCotacao_Le:
    
    ItensPedCotacao_Le = gErr
    
    Select Case gErr

        Case 71049
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145114)
        
    End Select
    
    Exit Function

End Function

Function Cotacao_Le(objCotacao As ClassCotacao) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cotacao_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAO_LE", X)
    If lErro <> SUCESSO Then gError 71050
    
    Cotacao_Le = X.Cotacao_Le(objCotacao)

    Exit Function
    
Erro_Cotacao_Le:
    
    Cotacao_Le = gErr
    
    Select Case gErr

        Case 71050
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145115)
        
    End Select
    
    Exit Function

End Function

Function Cotacao_Inclui(objCotacao As ClassCotacao, colItensPedCotacao As Collection) As Long
'Atualiza as tabelas Cotacao,CotacaoCondPagto e CotacaoProduto a partir
'da Cotacao passada como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cotacao_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAO_INCLUI", X)
    If lErro <> SUCESSO Then gError 71051
    
    Cotacao_Inclui = X.Cotacao_Inclui(objCotacao, colItensPedCotacao)

    Exit Function
    
Erro_Cotacao_Inclui:
    
    Cotacao_Inclui = gErr
    
    Select Case gErr

        Case 71051
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145116)
        
    End Select
    
    Exit Function

End Function

Function CotacoesProduto_Inclui(objCotacao As ClassCotacao, colItensPedCotacao As Collection) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CotacoesProduto_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACOESPRODUTO_INCLUI", X)
    If lErro <> SUCESSO Then gError 71052
    
    CotacoesProduto_Inclui = X.CotacoesProduto_Inclui(objCotacao, colItensPedCotacao)

    Exit Function
    
Erro_CotacoesProduto_Inclui:
    
    CotacoesProduto_Inclui = gErr
    
    Select Case gErr

        Case 71052
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145117)
        
    End Select
    
    Exit Function

End Function

Function CondicoesPagto_Inclui(objCotacao As ClassCotacao) As Long
'Inclui na tabela CotacaoCondPagto as condicoes de pagto de Cotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CondicoesPagto_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONDICOESPAGTO_INCLUI", X)
    If lErro <> SUCESSO Then gError 71053
    
    CondicoesPagto_Inclui = X.CondicoesPagto_Inclui(objCotacao)

    Exit Function
    
Erro_CondicoesPagto_Inclui:
    
    CondicoesPagto_Inclui = gErr
    
    Select Case gErr

        Case 71053
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145118)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCotacao_Inclui(objPedidoCotacao As ClassPedidoCotacao, objCotacaoProduto As ClassCotacaoProduto, colItensPedCotacao As Collection) As Long
'Inclui os itens de Pedido de Cotacao na tabela ItemPedCotacao do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCotacao_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPEDCOTACAO_INCLUI", X)
    If lErro <> SUCESSO Then gError 71054
    
    ItensPedCotacao_Inclui = X.ItensPedCotacao_Inclui(objPedidoCotacao, objCotacaoProduto, colItensPedCotacao)

    Exit Function
    
Erro_ItensPedCotacao_Inclui:
    
    ItensPedCotacao_Inclui = gErr
    
    Select Case gErr

        Case 71054
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145119)
        
    End Select
    
    Exit Function

End Function

Function PedidosCotacao_Inclui(colPedidoCotacao As Collection, objCotacaoProduto As ClassCotacaoProduto, colItensPedCotacao As Collection) As Long
'Inclui os Pedidos de Cotacao na tabela PedidoCotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosCotacao_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOSCOTACAO_INCLUI", X)
    If lErro <> SUCESSO Then gError 71055
    
    PedidosCotacao_Inclui = X.PedidosCotacao_Inclui(colPedidoCotacao, objCotacaoProduto, colItensPedCotacao)

    Exit Function
    
Erro_PedidosCotacao_Inclui:
    
    PedidosCotacao_Inclui = gErr
    
    Select Case gErr

        Case 71055
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145120)
        
    End Select
    
    Exit Function

End Function

Function Cotacao_Automatico(lNumIntCotacao As Long) As Long
'Devolve o proximo número interno de cotacao disponivel.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cotacao_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71056
    
    Cotacao_Automatico = X.Cotacao_Automatico(lNumIntCotacao)

    Exit Function
    
Erro_Cotacao_Automatico:
    
    Cotacao_Automatico = gErr
    
    Select Case gErr

        Case 71056
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145121)
        
    End Select
    
    Exit Function

End Function

Function CotacaoProduto_Automatico(lNumCotacaoProduto As Long) As Long
'Devolve o proximo número interno de cotacao produto disponivel.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CotacaoProduto_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAOPRODUTO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71057
    
    CotacaoProduto_Automatico = X.CotacaoProduto_Automatico(lNumCotacaoProduto)

    Exit Function
    
Erro_CotacaoProduto_Automatico:
    
    CotacaoProduto_Automatico = gErr
    
    Select Case gErr

        Case 71057
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145122)
        
    End Select
    
    Exit Function

End Function

Function Cotacao_Grava(objCotacao As ClassCotacao, colPedidoCotacao As Collection) As Long
'Grava o Pedido de Cotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cotacao_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAO_GRAVA", X)
    If lErro <> SUCESSO Then gError 71058
    
    Cotacao_Grava = X.Cotacao_Grava(objCotacao, colPedidoCotacao)

    Exit Function
    
Erro_Cotacao_Grava:
    
    Cotacao_Grava = gErr
    
    Select Case gErr

        Case 71058
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145123)
        
    End Select
    
    Exit Function

End Function

Function Cotacao_ObjetosAssociados_Lock(objCotacao As ClassCotacao, colPedidoCotacao As Collection) As Long
'Lê e locka os objetos associados a Cotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cotacao_ObjetosAssociados_Lock
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAO_OBJETOSASSOCIADOS_LOCK", X)
    If lErro <> SUCESSO Then gError 71059
    
    Cotacao_ObjetosAssociados_Lock = X.Cotacao_ObjetosAssociados_Lock(objCotacao, colPedidoCotacao)

    Exit Function
    
Erro_Cotacao_ObjetosAssociados_Lock:
    
    Cotacao_ObjetosAssociados_Lock = gErr
    
    Select Case gErr

        Case 71059
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145124)
        
    End Select
    
    Exit Function

End Function

Function Cotacao_Automatica(lCotacao As Long) As Long
'Devolve o proximo código de cotacao disponivel para a FilialEmpresa em questao.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cotacao_Automatica
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAO_AUTOMATICA", X)
    If lErro <> SUCESSO Then gError 71060
    
    Cotacao_Automatica = X.Cotacao_Automatica(lCotacao)

    Exit Function
    
Erro_Cotacao_Automatica:
    
    Cotacao_Automatica = gErr
    
    Select Case gErr

        Case 71060
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145125)
        
    End Select
    
    Exit Function

End Function

Function PedidoCotacaoNumInt_Automatico(lNumIntPedCotacao As Long) As Long
'gera o próximo número interno disponível para PedidoCotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCotacaoNumInt_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOTACAONUMINT_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71061
    
    PedidoCotacaoNumInt_Automatico = X.PedidoCotacaoNumInt_Automatico(lNumIntPedCotacao)

    Exit Function
    
Erro_PedidoCotacaoNumInt_Automatico:
    
    PedidoCotacaoNumInt_Automatico = gErr
    
    Select Case gErr

        Case 71061
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145126)
        
    End Select
    
    Exit Function

End Function

Function PedidoCotacao_Automatico(lPedCotacao As Long) As Long
'Gera o próximo código disponível para PedidoCotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCotacao_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOTACAO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71062
    
    PedidoCotacao_Automatico = X.PedidoCotacao_Automatico(lPedCotacao)

    Exit Function
    
Erro_PedidoCotacao_Automatico:
    
    PedidoCotacao_Automatico = gErr
    
    Select Case gErr

        Case 71062
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145127)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCotacao_Automatico(lNumIntItemPC As Long) As Long
'Gera o próximo NumInt de ItemPedCotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCotacao_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPEDCOTACAO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71063
    
    ItensPedCotacao_Automatico = X.ItensPedCotacao_Automatico(lNumIntItemPC)

    Exit Function
    
Erro_ItensPedCotacao_Automatico:
    
    ItensPedCotacao_Automatico = gErr
    
    Select Case gErr

        Case 71063
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145128)
        
    End Select
    
    Exit Function

End Function

Function ReqCompras_Le_GeracaoCotacao(objGeracaoCotacao As ClassGeracaoCotacao) As Long
'Lê as Requisicoes de Compras a partir da selecao utilizada pelo usuário no tabSelecao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReqCompras_Le_GeracaoCotacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQCOMPRAS_LE_GERACAOCOTACAO", X)
    If lErro <> SUCESSO Then gError 71064
    
    ReqCompras_Le_GeracaoCotacao = X.ReqCompras_Le_GeracaoCotacao(objGeracaoCotacao)

    Exit Function
    
Erro_ReqCompras_Le_GeracaoCotacao:
    
    ReqCompras_Le_GeracaoCotacao = gErr
    
    Select Case gErr

        Case 71064
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145129)
        
    End Select
    
    Exit Function

End Function

Function ItensReqCompras_Le_NaoCotadas(objReqCompra As ClassRequisicaoCompras) As Long
'Lê os Itens da Requisicao de Compras que não foram cotados e adicionar na
'coleção de Itens de objReqCompra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqCompras_Le_NaoCotadas
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSREQCOMPRAS_LE_NAOCOTADAS", X)
    If lErro <> SUCESSO Then gError 71065
    
    ItensReqCompras_Le_NaoCotadas = X.ItensReqCompras_Le_NaoCotadas(objReqCompra)

    Exit Function
    
Erro_ItensReqCompras_Le_NaoCotadas:
    
    ItensReqCompras_Le_NaoCotadas = gErr
    
    Select Case gErr

        Case 71065
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145130)
        
    End Select
    
    Exit Function

End Function

Function CotacaoProdutoItemRC_Inclui(objCotacao As ClassCotacao, objCotacaoProduto As ClassCotacaoProduto) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CotacaoProdutoItemRC_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAOPRODUTOITEMRC_INCLUI", X)
    If lErro <> SUCESSO Then gError 71066
    
    CotacaoProdutoItemRC_Inclui = X.CotacaoProdutoItemRC_Inclui(objCotacao, objCotacaoProduto)

    Exit Function
    
Erro_CotacaoProdutoItemRC_Inclui:
    
    CotacaoProdutoItemRC_Inclui = gErr
    
    Select Case gErr

        Case 71066
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145131)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompras_Le_Numero(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê em RequisiçãoCompras a Requisição passada por objRequisiçãoCompras com Código e FlialEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompras_Le_Numero
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRAS_LE_NUMERO", X)
    If lErro <> SUCESSO Then gError 71067
    
    RequisicaoCompras_Le_Numero = X.RequisicaoCompras_Le_Numero(objRequisicaoCompras)

    Exit Function
    
Erro_RequisicaoCompras_Le_Numero:
    
    RequisicaoCompras_Le_Numero = gErr
    
    Select Case gErr

        Case 71067
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145132)
        
    End Select
    
    Exit Function

End Function

Function Cotacao_Grava_Pedidos(objCotacao As ClassCotacao, colPedidoCotacao As Collection) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cotacao_Grava_Pedidos
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAO_GRAVA_ITENSREQ", X)
    If lErro <> SUCESSO Then gError 71068
    
    Cotacao_Grava_Pedidos = X.Cotacao_Grava_Pedidos(objCotacao, colPedidoCotacao)

    Exit Function
    
Erro_Cotacao_Grava_Pedidos:
    
    Cotacao_Grava_Pedidos = gErr
    
    Select Case gErr

        Case 71068
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145133)
        
    End Select
    
    Exit Function

End Function

Function CotacaoProdutoFornecedor_Le(colCotacaoProduto As Collection, colFornecedorProdutoFF As Collection) As Long
'Verifica se o fornecedor do Item é exclusivo ou preferencial e preenche colCotacaoProdutoFornecedor

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CotacaoProdutoFornecedor_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAOPRODUTOFORNECEDOR_LE", X)
    If lErro <> SUCESSO Then gError 71069
    
    CotacaoProdutoFornecedor_Le = X.CotacaoProdutoFornecedor_Le(colCotacaoProduto, colFornecedorProdutoFF)

    Exit Function
    
Erro_CotacaoProdutoFornecedor_Le:
    
    CotacaoProdutoFornecedor_Le = gErr
    
    Select Case gErr

        Case 71069
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145134)
        
    End Select
    
    Exit Function

End Function

Function GeracaoReqPedVenda_Le_Pedidos(gobjGeracaoReqPedVenda As ClassGeracaoReqPedVenda) As Long
'Busca no BD os Pedidos de Venda que atendam as condicoes informadas no TabSelecao da tela

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoReqPedVenda_Le_Pedidos
    
    lErro = Rotina_ObterObjetoFuncao("COM_GERACAOREQPEDVENDA_LE_PEDIDOS", X)
    If lErro <> SUCESSO Then gError 71070
    
    GeracaoReqPedVenda_Le_Pedidos = X.GeracaoReqPedVenda_Le_Pedidos(gobjGeracaoReqPedVenda)

    Exit Function
    
Erro_GeracaoReqPedVenda_Le_Pedidos:
    
    GeracaoReqPedVenda_Le_Pedidos = gErr
    
    Select Case gErr

        Case 71070
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145135)
        
    End Select
    
    Exit Function

End Function

Function GeracaoReqPedVenda_Grava(objReqCompras As ClassRequisicaoCompras, colPedidoVenda As Collection) As Long
'Busca no BD os Pedidos de Venda de acordo com os parametros selecionados no TabSelecao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoReqPedVenda_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_GERACAOREQPEDVENDA_GRAVA", X)
    If lErro <> SUCESSO Then gError 71071
    
    GeracaoReqPedVenda_Grava = X.GeracaoReqPedVenda_Grava(objReqCompras, colPedidoVenda)

    Exit Function
    
Erro_GeracaoReqPedVenda_Grava:
    
    GeracaoReqPedVenda_Grava = gErr
    
    Select Case gErr

        Case 71071
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145136)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompra_Grava_EmTrans(objRequisicaoCompra As ClassRequisicaoCompras, Optional bGeracao As Boolean = False) As Long
'Grava a Requisição de Compra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompra_Grava_EmTrans
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRA_GRAVA_EMTRANS", X)
    If lErro <> SUCESSO Then gError 71072
    
    RequisicaoCompra_Grava_EmTrans = X.RequisicaoCompra_Grava_EmTrans(objRequisicaoCompra, bGeracao)

    Exit Function
    
Erro_RequisicaoCompra_Grava_EmTrans:
    
    RequisicaoCompra_Grava_EmTrans = gErr
    
    Select Case gErr

        Case 71072
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145137)
        
    End Select
    
    Exit Function

End Function

Function ItensReqCompra_Inclui(objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Insere os Itens da Requisição Compras no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqCompra_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSREQCOMPRA_INCLUI", X)
    If lErro <> SUCESSO Then gError 71073
    
    ItensReqCompra_Inclui = X.ItensReqCompra_Inclui(objRequisicaoCompra)

    Exit Function
    
Erro_ItensReqCompra_Inclui:
    
    ItensReqCompra_Inclui = gErr
    
    Select Case gErr

        Case 71073
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145138)
        
    End Select
    
    Exit Function

End Function

Function ItensReqCompra_Exclui(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Exclui os Itens da Requisição de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqCompra_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSREQCOMPRA_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71074
    
    ItensReqCompra_Exclui = X.ItensReqCompra_Exclui(objRequisicaoCompras)

    Exit Function
    
Erro_ItensReqCompra_Exclui:
    
    ItensReqCompra_Exclui = gErr
    
    Select Case gErr

        Case 71074
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145139)
        
    End Select
    
    Exit Function

End Function

Function FilialEmpresa_Lock(iCodFilial As Integer) As Long
'Lê e Faz "lockshared" na FilialEmpresa passada por iCodFilial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialEmpresa_Lock
    
    lErro = Rotina_ObterObjetoFuncao("COM_FILIALEMPRESA_LOCK", X)
    If lErro <> SUCESSO Then gError 71075
    
    FilialEmpresa_Lock = X.FilialEmpresa_Lock(iCodFilial)

    Exit Function
    
Erro_FilialEmpresa_Lock:
    
    FilialEmpresa_Lock = gErr
    
    Select Case gErr

        Case 71075
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145140)
        
    End Select
    
    Exit Function

End Function

Function Ccl_Lock(sCcl As String) As Long
'Lê e Faz "LockShared" do Ccl

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Ccl_Lock
    
    lErro = Rotina_ObterObjetoFuncao("COM_CCL_LOCK", X)
    If lErro <> SUCESSO Then gError 71076
    
    Ccl_Lock = X.Ccl_Lock(sCcl)

    Exit Function
    
Erro_Ccl_Lock:
    
    Ccl_Lock = gErr
    
    Select Case gErr

        Case 71076
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145141)
        
    End Select
    
    Exit Function

End Function

Function Requisitante_Lock(lCodRequisitante As Long) As Long
'Lê e faz LockShared no Requisitante

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Requisitante_Lock
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISITANTE_LOCK", X)
    If lErro <> SUCESSO Then gError 71077
    
    Requisitante_Lock = X.Requisitante_Lock(lCodRequisitante)

    Exit Function
    
Erro_Requisitante_Lock:
    
    Requisitante_Lock = gErr
    
    Select Case gErr

        Case 71077
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145142)
        
    End Select
    
    Exit Function

End Function

Function Observacao_Lock(sObservacao As String, lObservacao As Long) As Long
'Faz "LockShared" em Observacao caso sObservacao já exista
'e insere nova Observação caso sObservacao não exista.
'Devolve em lObservacao o lNumInt gerado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Observacao_Lock
    
    lErro = Rotina_ObterObjetoFuncao("COM_OBSERVACAO_LOCK", X)
    If lErro <> SUCESSO Then gError 71078
    
    Observacao_Lock = X.Observacao_Lock(sObservacao, lObservacao)

    Exit Function
    
Erro_Observacao_Lock:
    
    Observacao_Lock = gErr
    
    Select Case gErr

        Case 71078
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145143)
        
    End Select
    
    Exit Function

End Function

Function ContaContabil_Lock(sConta As String) As Long
'Lê e Faz Lock na Conta passada por sConta

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ContaContabil_Lock
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONTACONTABIL_LOCK", X)
    If lErro <> SUCESSO Then gError 71079
    
    ContaContabil_Lock = X.ContaContabil_Lock(sConta)

    Exit Function
    
Erro_ContaContabil_Lock:
    
    ContaContabil_Lock = gErr
    
    Select Case gErr

        Case 71079
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145144)
        
    End Select
    
    Exit Function

End Function

Function Observacao_Automatico(lNumInt As Long) As Long
'Gera o proximo numero interno da Tabela de Observacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Observacao_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_OBSERVACAO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71080
    
    Observacao_Automatico = X.Observacao_Automatico(lNumInt)

    Exit Function
    
Erro_Observacao_Automatico:
    
    Observacao_Automatico = gErr
    
    Select Case gErr

        Case 71080
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145145)
        
    End Select
    
    Exit Function

End Function

Function TiposTributacaoCompras_Le(colTributacao As AdmColCodigoNome) As Long
'Lê Tipos de Tributação associadas a Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposTributacaoCompras_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_TIPOSTRIBUTACAOCOMPRAS_LE", X)
    If lErro <> SUCESSO Then gError 71081
    
    TiposTributacaoCompras_Le = X.TiposTributacaoCompras_Le(colTributacao)

    Exit Function
    
Erro_TiposTributacaoCompras_Le:
    
    TiposTributacaoCompras_Le = gErr
    
    Select Case gErr

        Case 71081
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145146)
        
    End Select
    
    Exit Function

End Function

Function ItensReqCompra_Lock_Gravacao(objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Faz lock nas tabelas associados ao item da Requisição Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqCompra_Lock_Gravacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSREQCOMPRA_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then gError 71082
    
    ItensReqCompra_Lock_Gravacao = X.ItensReqCompra_Lock_Gravacao(objRequisicaoCompra)

    Exit Function
    
Erro_ItensReqCompra_Lock_Gravacao:
    
    ItensReqCompra_Lock_Gravacao = gErr
    
    Select Case gErr

        Case 71082
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145147)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompra_Lock_Gravacao(objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Faz Lock nas tabelas Associadas a Requisição de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompra_Lock_Gravacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRA_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then gError 71083
    
    RequisicaoCompra_Lock_Gravacao = X.RequisicaoCompra_Lock_Gravacao(objRequisicaoCompra)

    Exit Function
    
Erro_RequisicaoCompra_Lock_Gravacao:
    
    RequisicaoCompra_Lock_Gravacao = gErr
    
    Select Case gErr

        Case 71083
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145148)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompras_Codigo_Automatico_EmTrans(lCodigo As Long) As Long
'Gera código automático para a próxima Requisição de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompras_Codigo_Automatico_EmTrans
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRAS_CODIGO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71084
    
    RequisicaoCompras_Codigo_Automatico_EmTrans = X.RequisicaoCompras_Codigo_Automatico_EmTrans(lCodigo)

    Exit Function
    
Erro_RequisicaoCompras_Codigo_Automatico_EmTrans:
    
    RequisicaoCompras_Codigo_Automatico_EmTrans = gErr
    
    Select Case gErr

        Case 71084
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145149)
        
    End Select
    
    Exit Function

End Function

Function Cotacoes_Produto_Le(colGeracao As Collection, objProduto As ClassProduto, dQuantidade As Double, iTipoDestino As Integer, lDestino As Long, iFilialDestino As Integer, Optional lFornecExclu As Long = 0, Optional iFilialExclu As Integer = 0) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cotacoes_Produto_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACOES_PRODUTO_LE", X)
    If lErro <> SUCESSO Then gError 71086
    
    Cotacoes_Produto_Le = X.Cotacoes_Produto_Le(colGeracao, objProduto, dQuantidade, iTipoDestino, lDestino, iFilialDestino, lFornecExclu, iFilialExclu)

    Exit Function
    
Erro_Cotacoes_Produto_Le:
    
    Cotacoes_Produto_Le = gErr
    
    Select Case gErr

        Case 71086
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145150)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Codigo_Automatico_EmTrans(lCodigo As Long) As Long
'Gera o proximo codigo da Tabela de Pedido Compra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Codigo_Automatico_EmTrans
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_CODIGO_AUTOMATICO_EMTRA", X)
    If lErro <> SUCESSO Then gError 71087
    
    PedidoCompra_Codigo_Automatico_EmTrans = X.PedidoCompra_Codigo_Automatico_EmTrans(lCodigo)

    Exit Function
    
Erro_PedidoCompra_Codigo_Automatico_EmTrans:
    
    PedidoCompra_Codigo_Automatico_EmTrans = gErr
    
    Select Case gErr

        Case 71087
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145151)
        
    End Select
    
    Exit Function

End Function

Function CotacaoItemConcorrencia_Le_ItemConc(lNumItemCotacao As Long, lNumItemConc As Long, lNumIntDoc As Long) As Long
'Lê o NumIntDoc da Cotação Item Concorrência partir do ItemConcorrência e ItemCotação

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CotacaoItemConcorrencia_Le_ItemConc
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAOITEMCONCORRENCIA_LE_ITEMCONC", X)
    If lErro <> SUCESSO Then gError 71088
    
    CotacaoItemConcorrencia_Le_ItemConc = X.CotacaoItemConcorrencia_Le_ItemConc(lNumItemCotacao, lNumItemConc, lNumIntDoc)

    Exit Function
    
Erro_CotacaoItemConcorrencia_Le_ItemConc:
    
    CotacaoItemConcorrencia_Le_ItemConc = gErr
    
    Select Case gErr

        Case 71088
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145152)
        
    End Select
    
    Exit Function

End Function

Function PedidosCotacao_Le_Cotacao(objCotacao As ClassCotacao, colItensCotacao As Collection) As Long
'Lê os itens de cotação vinculadas a concorrência

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosCotacao_Le_Cotacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOSCOTACAO_LE_COTACAO", X)
    If lErro <> SUCESSO Then gError 71089
    
    PedidosCotacao_Le_Cotacao = X.PedidosCotacao_Le_Cotacao(objCotacao, colItensCotacao)

    Exit Function
    
Erro_PedidosCotacao_Le_Cotacao:
    
    PedidosCotacao_Le_Cotacao = gErr
    
    Select Case gErr

        Case 71089
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145153)
        
    End Select
    
    Exit Function

End Function

Function CotacaoProduto_Le(objCotacao As ClassCotacao) As Long
'Lê Cotações do Produto relacionadas a uma cotação a partir do NumIntDoc passado em objCotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CotacaoProduto_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAOPRODUTO_LE", X)
    If lErro <> SUCESSO Then gError 71090
    
    CotacaoProduto_Le = X.CotacaoProduto_Le(objCotacao)

    Exit Function
    
Erro_CotacaoProduto_Le:
    
    CotacaoProduto_Le = gErr
    
    Select Case gErr

        Case 71090
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145154)
        
    End Select
    
    Exit Function

End Function

Function CotacaoProdutoItemRC_Le_Todas(objCotacao As ClassCotacao) As Long
'ItensRC relacionadas a uma cotação Produto a partir do NumIntDoc das Cotações Produtos passadas
'na coleção de objCotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CotacaoProdutoItemRC_Le_Todas
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAOPRODUTOITEMRC_LE_TODAS", X)
    If lErro <> SUCESSO Then gError 71091
    
    CotacaoProdutoItemRC_Le_Todas = X.CotacaoProdutoItemRC_Le_Todas(objCotacao)

    Exit Function
    
Erro_CotacaoProdutoItemRC_Le_Todas:
    
    CotacaoProdutoItemRC_Le_Todas = gErr
    
    Select Case gErr

        Case 71091
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145155)
        
    End Select
    
    Exit Function

End Function

Function Requisicoes_Le_Cotacao(objGeracaoPedCompraCot As ClassGeracaoPedCompraCot, colRequisicoes As Collection) As Long
'Lê as Requisições associadas a Cotação

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Requisicoes_Le_Cotacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICOES_LE_COTACAO", X)
    If lErro <> SUCESSO Then gError 71092
    
    Requisicoes_Le_Cotacao = X.Requisicoes_Le_Cotacao(objGeracaoPedCompraCot, colRequisicoes)

    Exit Function
    
Erro_Requisicoes_Le_Cotacao:
    
    Requisicoes_Le_Cotacao = gErr
    
    Select Case gErr

        Case 71092
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145156)
        
    End Select
    
    Exit Function

End Function

Function GeracaoPedCompraCot_Le_GeracaoPedCotacao(objGeracaoPedidoCompraCot As ClassGeracaoPedCompraCot) As Long
'Lê as Cotações selecionadas a partir das opções preenchidas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoPedCompraCot_Le_GeracaoPedCotacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_GERACAOPEDCOMPRACOT_LE_GERACAOPEDCOT", X)
    If lErro <> SUCESSO Then gError 71093
    
    GeracaoPedCompraCot_Le_GeracaoPedCotacao = X.GeracaoPedCompraCot_Le_GeracaoPedCotacao(objGeracaoPedidoCompraCot)

    Exit Function
    
Erro_GeracaoPedCompraCot_Le_GeracaoPedCotacao:
    
    GeracaoPedCompraCot_Le_GeracaoPedCotacao = gErr
    
    Select Case gErr

        Case 71093
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145157)
        
    End Select
    
    Exit Function

End Function

Function ItemReqCompra_Le_Produto(objItemReqCompra As ClassItemReqCompras) As Long
'Lê o NumIntDoc do ItemRC a partir do NumIntDoc da Requisição o do Produto de ItemRC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemReqCompra_Le_Produto
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITEMREQCOMPRA_LE_PRODUTO", X)
    If lErro <> SUCESSO Then gError 71094
    
    ItemReqCompra_Le_Produto = X.ItemReqCompra_Le_Produto(objItemReqCompra)

    Exit Function
    
Erro_ItemReqCompra_Le_Produto:
    
    ItemReqCompra_Le_Produto = gErr
    
    Select Case gErr

        Case 71094
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145158)
        
    End Select
    
    Exit Function

End Function

Function Concorrencia_Baixa(objConcorrencia As ClassConcorrencia) As Long
'Baixa a Concorrência passada por objConcorrencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Concorrencia_Baixa
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONCORRENCIA_BAIXA", X)
    If lErro <> SUCESSO Then gError 71095
    
    Concorrencia_Baixa = X.Concorrencia_Baixa(objConcorrencia)

    Exit Function
    
Erro_Concorrencia_Baixa:
    
    Concorrencia_Baixa = gErr
    
    Select Case gErr

        Case 71095
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145159)
        
    End Select
    
    Exit Function

End Function

Function Cotacao_Le_Tudo(objCotacao As ClassCotacao) As Long
'Lê a Cotação passada pelo NumIntDoc e coleções associadas
'(Cotação Produto e Cotação Produto ItemRC)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cotacao_Le_Tudo
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAO_LE_TUDO", X)
    If lErro <> SUCESSO Then gError 71096
    
    Cotacao_Le_Tudo = X.Cotacao_Le_Tudo(objCotacao)

    Exit Function
    
Erro_Cotacao_Le_Tudo:
    
    Cotacao_Le_Tudo = gErr
    
    Select Case gErr

        Case 71096
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145160)
        
    End Select
    
    Exit Function

End Function

Function ProdutoFilial_Le_PontoPedido(objGeracaoPtoPedido As ClassGeracaoPtoPedido) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutoFilial_Le_PontoPedido
    
    lErro = Rotina_ObterObjetoFuncao("COM_PRODUTOFILIAL_LE_PONTOPEDIDO", X)
    If lErro <> SUCESSO Then gError 71097
    
    ProdutoFilial_Le_PontoPedido = X.ProdutoFilial_Le_PontoPedido(objGeracaoPtoPedido)

    Exit Function
    
Erro_ProdutoFilial_Le_PontoPedido:
    
    ProdutoFilial_Le_PontoPedido = gErr
    
    Select Case gErr

        Case 71097
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145161)
        
    End Select
    
    Exit Function

End Function

Function EstoqueProduto_Le_Todos_AlmProdFilial(objEstoqueProduto As ClassEstoqueProduto) As Long
'Lê a soma das Quantidades de todos Almoxarifados que pertence ao Produto passado em objEstoqueProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_Le_Todos_AlmProdFilial
    
    lErro = Rotina_ObterObjetoFuncao("COM_ESTOQUEPRODUTO_LE_TODOS_ALMPRODFILIA", X)
    If lErro <> SUCESSO Then gError 71098
    
    EstoqueProduto_Le_Todos_AlmProdFilial = X.EstoqueProduto_Le_Todos_AlmProdFilial(objEstoqueProduto)

    Exit Function
    
Erro_EstoqueProduto_Le_Todos_AlmProdFilial:
    
    EstoqueProduto_Le_Todos_AlmProdFilial = gErr
    
    Select Case gErr

        Case 71098
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145162)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompras_Codigo_Automatico(lCodigo As Long) As Long
'Gera código automático para a próxima Requisição de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompras_Codigo_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRAS_CODIGO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71099
    
    RequisicaoCompras_Codigo_Automatico = X.RequisicaoCompras_Codigo_Automatico(lCodigo)

    Exit Function
    
Erro_RequisicaoCompras_Codigo_Automatico:
    
    RequisicaoCompras_Codigo_Automatico = gErr
    
    Select Case gErr

        Case 71099
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145163)
        
    End Select
    
    Exit Function

End Function

Function GeracaoRequisicaoPtoPedido_Grava(objRequisicao As ClassRequisicaoCompras) As Long
'Grava a Requisição de Compra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoRequisicaoPtoPedido_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_GERACAOREQUISICAOPTOPEDIDO_GRAVA", X)
    If lErro <> SUCESSO Then gError 71100
    
    GeracaoRequisicaoPtoPedido_Grava = X.GeracaoRequisicaoPtoPedido_Grava(objRequisicao)

    Exit Function
    
Erro_GeracaoRequisicaoPtoPedido_Grava:
    
    GeracaoRequisicaoPtoPedido_Grava = gErr
    
    Select Case gErr

        Case 71100
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145164)
        
    End Select
    
    Exit Function

End Function

Function Concorrencia_Exclui(objConcorrencia As ClassConcorrencia) As Long
'Exclui a Concorrência passada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Concorrencia_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONCORRENCIA_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71101
    
    Concorrencia_Exclui = X.Concorrencia_Exclui(objConcorrencia)

    Exit Function
    
Erro_Concorrencia_Exclui:
    
    Concorrencia_Exclui = gErr
    
    Select Case gErr

        Case 71101
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145165)
        
    End Select
    
    Exit Function

End Function

Function Motivo_Le_Codigo(objMotivo As AdmCodigoNome)
'Lê o Motivo com o código passado cadastrado na tabela Motivo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Motivo_Le_Codigo
    
    lErro = Rotina_ObterObjetoFuncao("COM_MOTIVO_LE_CODIGO", X)
    If lErro <> SUCESSO Then gError 71102
    
    Motivo_Le_Codigo = X.Motivo_Le_Codigo(objMotivo)

    Exit Function
    
Erro_Motivo_Le_Codigo:
    
    Motivo_Le_Codigo = gErr
    
    Select Case gErr

        Case 71102
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145166)
        
    End Select
    
    Exit Function

End Function

Function Concorrencia_Le(objConcorrencia As ClassConcorrencia) As Long
'Lê dados da concorrência passada pelo código em objConcorrencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Concorrencia_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONCORRENCIA_LE", X)
    If lErro <> SUCESSO Then gError 71103
    
    Concorrencia_Le = X.Concorrencia_Le(objConcorrencia)

    Exit Function
    
Erro_Concorrencia_Le:
    
    Concorrencia_Le = gErr
    
    Select Case gErr

        Case 71103
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145167)
        
    End Select
    
    Exit Function

End Function

Function GerPedCompraConc_Le_Requisicoes(objGeracaoPedCompra As ClassGeracaoPedCompraConc, colRequisicoes As Collection) As Long
'Lê as Requisições associadas a concorrência

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GerPedCompraConc_Le_Requisicoes
    
    lErro = Rotina_ObterObjetoFuncao("COM_GERPEDCOMPRACONC_LE_REQUISICOES", X)
    If lErro <> SUCESSO Then gError 71104
    
    GerPedCompraConc_Le_Requisicoes = X.GerPedCompraConc_Le_Requisicoes(objGeracaoPedCompra, colRequisicoes)

    Exit Function
    
Erro_GerPedCompraConc_Le_Requisicoes:
    
    GerPedCompraConc_Le_Requisicoes = gErr
    
    Select Case gErr

        Case 71104
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145168)
        
    End Select
    
    Exit Function

End Function

Function PedidosCotacao_Le_Concorrencia(objConcorrencia As ClassConcorrencia, colItensCotacao As Collection) As Long
'Lê os itens de cotação vinculadas a concorrência

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosCotacao_Le_Concorrencia
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOSCOTACAO_LE_CONCORRENCIA", X)
    If lErro <> SUCESSO Then gError 71105
    
    PedidosCotacao_Le_Concorrencia = X.PedidosCotacao_Le_Concorrencia(objConcorrencia, colItensCotacao)

    Exit Function
    
Erro_PedidosCotacao_Le_Concorrencia:
    
    PedidosCotacao_Le_Concorrencia = gErr
    
    Select Case gErr

        Case 71105
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145169)
        
    End Select
    
    Exit Function

End Function

''Function Alcada_Avalia(sCodUsuario As String, dValor As Double, dtData As Date) As Long
'''Verifica se o usuario possui alcada
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_Alcada_Avalia
''
''    lErro = Rotina_ObterObjetoFuncao("COM_ALCADA_AVALIA", X)
''    If lErro <> SUCESSO Then gError 71106
''
''    Alcada_Avalia = X.Alcada_Avalia(sCodUsuario, dValor, dtData)
''
''    Exit Function
''
''Erro_Alcada_Avalia:
''
''    Alcada_Avalia = gErr
''
''    Select Case gErr
''
''        Case 71106
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145170)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function ValorPCLiberado_Grava(dValor As Double, sCodUsuario As String, dtData) As Long
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_ValorPCLiberado_Grava
''
''    lErro = Rotina_ObterObjetoFuncao("COM_VALORPCLIBERADO_GRAVA", X)
''    If lErro <> SUCESSO Then gError 71107
''
''    ValorPCLiberado_Grava = X.ValorPCLiberado_Grava(dValor, sCodUsuario, dtData)
''
''    Exit Function
''
''Erro_ValorPCLiberado_Grava:
''
''    ValorPCLiberado_Grava = gErr
''
''    Select Case gErr
''
''        Case 71107
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145171)
''
''    End Select
''
''    Exit Function
''
''End Function

Function ComprasConfig_Le_Conteudo(objComprasConfig As ClassComprasConfig) As Long
'Lê o registro em ComprasConfig de acordo com o Código passado.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComprasConfig_Le_Conteudo
    
    lErro = Rotina_ObterObjetoFuncao("COM_COMPRASCONFIG_LE_CONTEUDO", X)
    If lErro <> SUCESSO Then gError 71108
    
    ComprasConfig_Le_Conteudo = X.ComprasConfig_Le_Conteudo(objComprasConfig)

    Exit Function
    
Erro_ComprasConfig_Le_Conteudo:
    
    ComprasConfig_Le_Conteudo = gErr
    
    Select Case gErr

        Case 71108
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145172)
        
    End Select
    
    Exit Function

End Function

Function ItemPedCotacao_Le_PedidoCotacao(objItemPedCotacao As ClassItemPedCotacao, objPedidoCotacao As ClassPedidoCotacao) As Long
'Lê o Código do Pedido de Cotacao a que pertence o Item de Pedido de Cotacao passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemPedCotacao_Le_PedidoCotacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITEMPEDCOTACAO_LE_PEDIDOCOTACAO", X)
    If lErro <> SUCESSO Then gError 71109
    
    ItemPedCotacao_Le_PedidoCotacao = X.ItemPedCotacao_Le_PedidoCotacao(objItemPedCotacao, objPedidoCotacao)

    Exit Function
    
Erro_ItemPedCotacao_Le_PedidoCotacao:
    
    ItemPedCotacao_Le_PedidoCotacao = gErr
    
    Select Case gErr

        Case 71109
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145173)
        
    End Select
    
    Exit Function

End Function

Function ItemNFItemPC_Le(lNumItemPC As Long, colItemNF As Collection) As Long
'Lê todos os itens de Nota Fiscal relacionados ao item de Pedido de compras passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemNFItemPC_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITEMNFITEMPC_LE", X)
    If lErro <> SUCESSO Then gError 71110
    
    ItemNFItemPC_Le = X.ItemNFItemPC_Le(lNumItemPC, colItemNF)

    Exit Function
    
Erro_ItemNFItemPC_Le:
    
    ItemNFItemPC_Le = gErr
    
    Select Case gErr

        Case 71110
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145174)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Atualiza_DataEmissao(objPedidoCompra As ClassPedidoCompras) As Long
'Atualiza data de emissao no BD para a data atual

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Atualiza_DataEmissao
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_ATUALIZA_DATAEMISSAO", X)
    If lErro <> SUCESSO Then gError 71111
    
    PedidoCompra_Atualiza_DataEmissao = X.PedidoCompra_Atualiza_DataEmissao(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_Atualiza_DataEmissao:
    
    PedidoCompra_Atualiza_DataEmissao = gErr
    
    Select Case gErr

        Case 71111
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145175)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Le_Numero(objPedidoCompra As ClassPedidoCompras) As Long
'Busca no BD Pedido de Compra com FilialEmpresa e Codigo passados como parametros

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Le_Numero
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_LE_NUMERO", X)
    If lErro <> SUCESSO Then gError 71112
    
    PedidoCompra_Le_Numero = X.PedidoCompra_Le_Numero(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_Le_Numero:
    
    PedidoCompra_Le_Numero = gErr
    
    Select Case gErr

        Case 71112
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145176)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Le_Todos(objPedidoCompra As ClassPedidoCompras) As Long
'Busca no BD Pedido de Compra com FilialEmpresa e Codigo passados como parametros

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Le_Todos
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_LE_TODOS", X)
    If lErro <> SUCESSO Then gError 71113
    
    PedidoCompra_Le_Todos = X.PedidoCompra_Le_Todos(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_Le_Todos:
    
    PedidoCompra_Le_Todos = gErr
    
    Select Case gErr

        Case 71113
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145177)
        
    End Select
    
    Exit Function

End Function

Function ItensPC_LeTodos(objPedidoCompra As ClassPedidoCompras) As Long
'Lê no BD a coleção de ítens do Pedido de Compra e devolve em colItensPedido dentro de objPedidoCompra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPC_LeTodos
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPC_LETODOS", X)
    If lErro <> SUCESSO Then gError 71114
    
    ItensPC_LeTodos = X.ItensPC_LeTodos(objPedidoCompra)

    Exit Function
    
Erro_ItensPC_LeTodos:
    
    ItensPC_LeTodos = gErr
    
    Select Case gErr

        Case 71114
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145178)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompras_Le(objPedidoCompra As ClassPedidoCompras) As Long
'Lê os dados do Pedido de Compra,cujo numero interno foi passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompras_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRAS_LE", X)
    If lErro <> SUCESSO Then gError 71115
    
    PedidoCompras_Le = X.PedidoCompras_Le(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompras_Le:
    
    PedidoCompras_Le = gErr
    
    Select Case gErr

        Case 71115
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145179)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompraBaixado_Le_Numero(objPedidoCompra As ClassPedidoCompras)
'Busca no BD Pedido de Compra Baixado para a FilialEmpresa e Codigo passados como parametros

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompraBaixado_Le_Numero
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRABAIXADO_LE_NUMERO", X)
    If lErro <> SUCESSO Then gError 71116
    
    PedidoCompraBaixado_Le_Numero = X.PedidoCompraBaixado_Le_Numero(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompraBaixado_Le_Numero:
    
    PedidoCompraBaixado_Le_Numero = gErr
    
    Select Case gErr

        Case 71116
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145180)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCompra_Lock(objPedidoCompra As ClassPedidoCompras) As Long
'OK A função está toda sem comentário
'Le o numero interno dos itens do pedido, guarda na colecao de pedidos e faz lock exclusive dos itens

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCompra_Lock
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPEDCOMPRA_LOCK", X)
    If lErro <> SUCESSO Then gError 71117
    
    ItensPedCompra_Lock = X.ItensPedCompra_Lock(objPedidoCompra)

    Exit Function
    
Erro_ItensPedCompra_Lock:
    
    ItensPedCompra_Lock = gErr
    
    Select Case gErr

        Case 71117
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145181)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompraGerado_Grava(objPedidoCompra As ClassPedidoCompras) As Long
'Grava Pedido de Compra no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompraGerado_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRAGERADO_GRAVA", X)
    If lErro <> SUCESSO Then gError 71118
    
    PedidoCompraGerado_Grava = X.PedidoCompraGerado_Grava(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompraGerado_Grava:
    
    PedidoCompraGerado_Grava = gErr
    
    Select Case gErr

        Case 71118
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145182)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_LeMod_Gravacao(objPedidoCompra As ClassPedidoCompras) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_LeMod_Gravacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_LEMOD_GRAVACAO", X)
    If lErro <> SUCESSO Then gError 71119
    
    PedidoCompra_LeMod_Gravacao = X.PedidoCompra_LeMod_Gravacao(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_LeMod_Gravacao:
    
    PedidoCompra_LeMod_Gravacao = gErr
    
    Select Case gErr

        Case 71119
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145183)
        
    End Select
    
    Exit Function

End Function

Function BloqueioAlcada_Critica(objPedidoCompra As ClassPedidoCompras, iValorAlterado As Integer) As Long
'Inclui no BD um bloqueio de Alcada, se o usuário não tiver alçada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueioAlcada_Critica
    
    lErro = Rotina_ObterObjetoFuncao("COM_BLOQUEIOALCADA_CRITICA", X)
    If lErro <> SUCESSO Then gError 71120
    
    BloqueioAlcada_Critica = X.BloqueioAlcada_Critica(objPedidoCompra, iValorAlterado)

    Exit Function
    
Erro_BloqueioAlcada_Critica:
    
    BloqueioAlcada_Critica = gErr
    
    Select Case gErr

        Case 71120
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145184)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCompra_Altera(objPedidoCompra_BD As ClassPedidoCompras, objPedidoCompra As ClassPedidoCompras) As Long
'Atualiza Item de Pedido de Compra na tabela de ItensPedCompra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCompra_Altera
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPEDCOMPRA_ALTERA", X)
    If lErro <> SUCESSO Then gError 71121
    
    ItensPedCompra_Altera = X.ItensPedCompra_Altera(objPedidoCompra_BD, objPedidoCompra)

    Exit Function
    
Erro_ItensPedCompra_Altera:
    
    ItensPedCompra_Altera = gErr
    
    Select Case gErr

        Case 71121
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145185)
        
    End Select
    
    Exit Function

End Function

Function LocalizacaoItemPC_Altera(objPedidoCompra As ClassPedidoCompras) As Long
'Atualiza a tabela de LocalizacaoItensPC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LocalizacaoItemPC_Altera
    
    lErro = Rotina_ObterObjetoFuncao("COM_LOCALIZACAOITEMPC_ALTERA", X)
    If lErro <> SUCESSO Then gError 71122
    
    LocalizacaoItemPC_Altera = X.LocalizacaoItemPC_Altera(objPedidoCompra)

    Exit Function
    
Erro_LocalizacaoItemPC_Altera:
    
    LocalizacaoItemPC_Altera = gErr
    
    Select Case gErr

        Case 71122
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145186)
        
    End Select
    
    Exit Function

End Function

Function LocalizacaoItemPC_Inclui(objPedidoCompra As ClassPedidoCompras) As Long
'Insere registro na tabela de LocalizacaoItemPC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LocalizacaoItemPC_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_LOCALIZACAOITEMPC_INCLUI", X)
    If lErro <> SUCESSO Then gError 71123
    
    LocalizacaoItemPC_Inclui = X.LocalizacaoItemPC_Inclui(objPedidoCompra)

    Exit Function
    
Erro_LocalizacaoItemPC_Inclui:
    
    LocalizacaoItemPC_Inclui = gErr
    
    Select Case gErr

        Case 71123
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145187)
        
    End Select
    
    Exit Function

End Function

Function LocalizacaoItemPC_Exclui(objPedidoCompra As ClassPedidoCompras) As Long
'Exclui registro da tabela de LocalizacaoItemPC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LocalizacaoItemPC_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_LOCALIZACAOITEMPC_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71124
    
    LocalizacaoItemPC_Exclui = X.LocalizacaoItemPC_Exclui(objPedidoCompra)

    Exit Function
    
Erro_LocalizacaoItemPC_Exclui:
    
    LocalizacaoItemPC_Exclui = gErr
    
    Select Case gErr

        Case 71124
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145188)
        
    End Select
    
    Exit Function

End Function

Function BloqueiosPC_Grava(objPedidoCompra As ClassPedidoCompras) As Long
'Grava o BloqueioPC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPC_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_BLOQUEIOSPC_GRAVA", X)
    If lErro <> SUCESSO Then gError 71125
    
    BloqueiosPC_Grava = X.BloqueiosPC_Grava(objPedidoCompra)

    Exit Function
    
Erro_BloqueiosPC_Grava:
    
    BloqueiosPC_Grava = gErr
    
    Select Case gErr

        Case 71125
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145189)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Baixar(objPedidoCompra As ClassPedidoCompras) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Baixar
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_BAIXAR", X)
    If lErro <> SUCESSO Then gError 71126
    
    PedidoCompra_Baixar = X.PedidoCompra_Baixar(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_Baixar:
    
    PedidoCompra_Baixar = gErr
    
    Select Case gErr

        Case 71126
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145190)
        
    End Select
    
    Exit Function

End Function

Function BloqueiosPC_Le(objPedidoCompra As ClassPedidoCompras) As Long
'Le o BloqueioPC do Pedido de Compra passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPC_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_BLOQUEIOSPC_LE", X)
    If lErro <> SUCESSO Then gError 71127
    
    BloqueiosPC_Le = X.BloqueiosPC_Le(objPedidoCompra)

    Exit Function
    
Erro_BloqueiosPC_Le:
    
    BloqueiosPC_Le = gErr
    
    Select Case gErr

        Case 71127
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145191)
        
    End Select
    
    Exit Function

End Function

Function BloqueioPC_Automatico(lCodigo As Long, iProxSeqBloqueio As Integer) As Long
'retorna o número da próximo Sequencial de Bloqueio disponivel
'OBS: ESTA FUNÇÃO DEVE SER CHAMADA DE DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueioPC_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_BLOQUEIOPC_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71128
    
    BloqueioPC_Automatico = X.BloqueioPC_Automatico(lCodigo, iProxSeqBloqueio)

    Exit Function
    
Erro_BloqueioPC_Automatico:
    
    BloqueioPC_Automatico = gErr
    
    Select Case gErr

        Case 71128
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145192)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Exclui(objPedidoCompra As ClassPedidoCompras)
'Exclui o Pedido de Compra passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71129
    
    PedidoCompra_Exclui = X.PedidoCompra_Exclui(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_Exclui:
    
    PedidoCompra_Exclui = gErr
    
    Select Case gErr

        Case 71129
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145193)
        
    End Select
    
    Exit Function

End Function

Function BloqueiosPC_Exclui(objPedidoCompra As ClassPedidoCompras) As Long
'Exclui todos os bloqueios ligados ao Pedido de Compra passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPC_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_BLOQUEIOSPC_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71130
    
    BloqueiosPC_Exclui = X.BloqueiosPC_Exclui(objPedidoCompra)

    Exit Function
    
Erro_BloqueiosPC_Exclui:
    
    BloqueiosPC_Exclui = gErr
    
    Select Case gErr

        Case 71130
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145194)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Grava(objPedidoCompra As ClassPedidoCompras) As Long
'Grava Pedido de Compra no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_GRAVA", X)
    If lErro <> SUCESSO Then gError 71131
    
    PedidoCompra_Grava = X.PedidoCompra_Grava(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_Grava:
    
    PedidoCompra_Grava = gErr
    
    Select Case gErr

        Case 71131
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145195)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCompra_AlteraBD(objPedidoCompra As ClassPedidoCompras) As Long
'Atualiza Item de Pedido de Compra na tabela de ItensPedCompra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCompra_AlteraBD
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPEDCOMPRA_ALTERABD", X)
    If lErro <> SUCESSO Then gError 71132
    
    ItensPedCompra_AlteraBD = X.ItensPedCompra_AlteraBD(objPedidoCompra)

    Exit Function
    
Erro_ItensPedCompra_AlteraBD:
    
    ItensPedCompra_AlteraBD = gErr
    
    Select Case gErr

        Case 71132
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145196)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCompra_Exclui(objPedidoCompra As ClassPedidoCompras) As Long
'Exclui de ItensPedCompra os itens do PedidoCompra passados como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCompra_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPEDCOMPRA_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71133
    
    ItensPedCompra_Exclui = X.ItensPedCompra_Exclui(objPedidoCompra)

    Exit Function
    
Erro_ItensPedCompra_Exclui:
    
    ItensPedCompra_Exclui = gErr
    
    Select Case gErr

        Case 71133
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145197)
        
    End Select
    
    Exit Function

End Function

Function BloqueiosPC_Insere(objPedidoCompra As ClassPedidoCompras) As Long
'Insere o BloqueioPC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPC_Insere
    
    lErro = Rotina_ObterObjetoFuncao("COM_BLOQUEIOSPC_INSERE", X)
    If lErro <> SUCESSO Then gError 71134
    
    BloqueiosPC_Insere = X.BloqueiosPC_Insere(objPedidoCompra)

    Exit Function
    
Erro_BloqueiosPC_Insere:
    
    BloqueiosPC_Insere = gErr
    
    Select Case gErr

        Case 71134
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145198)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCompra_Insere(objPedidoCompra As ClassPedidoCompras) As Long
'Insere itens do PedidoCompra na tabela ItensPedCompra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCompra_Insere
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPEDCOMPRA_INSERE", X)
    If lErro <> SUCESSO Then gError 71135
    
    ItensPedCompra_Insere = X.ItensPedCompra_Insere(objPedidoCompra)

    Exit Function
    
Erro_ItensPedCompra_Insere:
    
    ItensPedCompra_Insere = gErr
    
    Select Case gErr

        Case 71135
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145199)
        
    End Select
    
    Exit Function

End Function

Function LocalizacaoItensPC_Insere(objPedidoCompra As ClassPedidoCompras) As Long
'Insere dados de distribuicao na tabela de LocalizacaoItensPC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LocalizacaoItensPC_Insere
    
    lErro = Rotina_ObterObjetoFuncao("COM_LOCALIZACAOITENSPC_INSERE", X)
    If lErro <> SUCESSO Then gError 71136
    
    LocalizacaoItensPC_Insere = X.LocalizacaoItensPC_Insere(objPedidoCompra)

    Exit Function
    
Erro_LocalizacaoItensPC_Insere:
    
    LocalizacaoItensPC_Insere = gErr
    
    Select Case gErr

        Case 71136
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145200)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Insere(objPedidoCompra As ClassPedidoCompras) As Long
'Insere PedidoCompra na tabela de PedidoCompra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Insere
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_INSERE", X)
    If lErro <> SUCESSO Then gError 71137
    
    PedidoCompra_Insere = X.PedidoCompra_Insere(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_Insere:
    
    PedidoCompra_Insere = gErr
    
    Select Case gErr

        Case 71137
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145201)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Lock_Gravacao(objPedidoCompra As ClassPedidoCompras) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Lock_Gravacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then gError 71138
    
    PedidoCompra_Lock_Gravacao = X.PedidoCompra_Lock_Gravacao(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_Lock_Gravacao:
    
    PedidoCompra_Lock_Gravacao = gErr
    
    Select Case gErr

        Case 71138
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145202)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCotacaoTodos_Le(objPedidoCotacao As ClassPedidoCotacao) As Long
'Lê os ítens do Pedido de Cotação baixado ou não

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCotacaoTodos_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPEDCOTACAOTODOS_LE", X)
    If lErro <> SUCESSO Then gError 71139
    
    ItensPedCotacaoTodos_Le = X.ItensPedCotacaoTodos_Le(objPedidoCotacao)

    Exit Function
    
Erro_ItensPedCotacaoTodos_Le:
    
    ItensPedCotacaoTodos_Le = gErr
    
    Select Case gErr

        Case 71139
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145203)
        
    End Select
    
    Exit Function

End Function

Function CotacaoTodas_Le(objCotacao As ClassCotacao) As Long
'Lê a tabela Cotação (baixada ou não) passada pelo NumIntDoc

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CotacaoTodas_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAOTODAS_LE", X)
    If lErro <> SUCESSO Then gError 71140
    
    CotacaoTodas_Le = X.CotacaoTodas_Le(objCotacao)

    Exit Function
    
Erro_CotacaoTodas_Le:
    
    CotacaoTodas_Le = gErr
    
    Select Case gErr

        Case 71140
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145204)
        
    End Select
    
    Exit Function

End Function

Function PedidoCotacaoTodos_Le(objPedidoCotacao As ClassPedidoCotacao) As Long
'Lê o Pedido de Cotação (Baixado ou Não) a partir do Código e FilialEmpresa passados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCotacaoTodos_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOTACAOTODOS_LE", X)
    If lErro <> SUCESSO Then gError 71141
    
    PedidoCotacaoTodos_Le = X.PedidoCotacaoTodos_Le(objPedidoCotacao)

    Exit Function
    
Erro_PedidoCotacaoTodos_Le:
    
    PedidoCotacaoTodos_Le = gErr
    
    Select Case gErr

        Case 71141
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145205)
        
    End Select
    
    Exit Function

End Function

Function CotacaoProdutoItemRC_Le(objPedidoCotacao As ClassPedidoCotacao) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CotacaoProdutoItemRC_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAOPRODUTOITEMRC_LE", X)
    If lErro <> SUCESSO Then gError 71142
    
    CotacaoProdutoItemRC_Le = X.CotacaoProdutoItemRC_Le(objPedidoCotacao)

    Exit Function
    
Erro_CotacaoProdutoItemRC_Le:
    
    CotacaoProdutoItemRC_Le = gErr
    
    Select Case gErr

        Case 71142
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145206)
        
    End Select
    
    Exit Function

End Function

Function PedidoCotacao_Atualiza(objPedidoCotacao As ClassPedidoCotacao) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCotacao_Atualiza
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOTACAO_ATUALIZA", X)
    If lErro <> SUCESSO Then gError 71143
    
    PedidoCotacao_Atualiza = X.PedidoCotacao_Atualiza(objPedidoCotacao)

    Exit Function
    
Erro_PedidoCotacao_Atualiza:
    
    PedidoCotacao_Atualiza = gErr
    
    Select Case gErr

        Case 71143
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145207)
        
    End Select
    
    Exit Function

End Function

Function PedidoCotacao_Exclui(objPedidoCotacao As ClassPedidoCotacao) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCotacao_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOTACAO_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71144
    
    PedidoCotacao_Exclui = X.PedidoCotacao_Exclui(objPedidoCotacao)

    Exit Function
    
Erro_PedidoCotacao_Exclui:
    
    PedidoCotacao_Exclui = gErr
    
    Select Case gErr

        Case 71144
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145208)
        
    End Select
    
    Exit Function

End Function

Function ItensCotacao_Atualiza(objPedidoCotacao As ClassPedidoCotacao) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensCotacao_Atualiza
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSCOTACAO_ATUALIZA", X)
    If lErro <> SUCESSO Then gError 71145
    
    ItensCotacao_Atualiza = X.ItensCotacao_Atualiza(objPedidoCotacao)

    Exit Function
    
Erro_ItensCotacao_Atualiza:
    
    ItensCotacao_Atualiza = gErr
    
    Select Case gErr

        Case 71145
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145209)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoComprasTodas_Le(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê as Requisições de compras (baixadas ou não) associadas ao Pedidos de Compra passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoComprasTodas_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRASTODAS_LE", X)
    If lErro <> SUCESSO Then gError 71146
    
    RequisicaoComprasTodas_Le = X.RequisicaoComprasTodas_Le(objRequisicaoCompras)

    Exit Function
    
Erro_RequisicaoComprasTodas_Le:
    
    RequisicaoComprasTodas_Le = gErr
    
    Select Case gErr

        Case 71146
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145210)
        
    End Select
    
    Exit Function

End Function

Function ItemRCItemPC_Le(lNumItemRC As Long, colItemPedido As Collection) As Long
'Lê todos os itens de Pedido de compras (baixados ou não) relacionados ao item de Requisição de compras passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemRCItemPC_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITEMRCITEMPC_LE", X)
    If lErro <> SUCESSO Then gError 71147
    
    ItemRCItemPC_Le = X.ItemRCItemPC_Le(lNumItemRC, colItemPedido)

    Exit Function
    
Erro_ItemRCItemPC_Le:
    
    ItemRCItemPC_Le = gErr
    
    Select Case gErr

        Case 71147
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145211)
        
    End Select
    
    Exit Function

End Function

Function ReqCompraEnv_BaixarEmTrans(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Baixa a Requisição passada por objRequisiçãoCompras
'CHAMADA EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReqCompraEnv_BaixarEmTrans
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQCOMPRAENV_BAIXAREMTRANS", X)
    If lErro <> SUCESSO Then gError 71148
    
    ReqCompraEnv_BaixarEmTrans = X.ReqCompraEnv_BaixarEmTrans(objRequisicaoCompras)

    Exit Function
    
Erro_ReqCompraEnv_BaixarEmTrans:
    
    ReqCompraEnv_BaixarEmTrans = gErr
    
    Select Case gErr

        Case 71148
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145212)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompraEnv_Grava(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Grava a Requisição passada pelo objRequisicaoCompras na tabela RequisiçãoCompras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompraEnv_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRAENV_GRAVA", X)
    If lErro <> SUCESSO Then gError 71149
    
    RequisicaoCompraEnv_Grava = X.RequisicaoCompraEnv_Grava(objRequisicaoCompras)

    Exit Function
    
Erro_RequisicaoCompraEnv_Grava:
    
    RequisicaoCompraEnv_Grava = gErr
    
    Select Case gErr

        Case 71149
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145213)
        
    End Select
    
    Exit Function

End Function

Function ItemReqCompra_Busca_PedidoCompra(iTela As Integer, objPedidoCompras As ClassPedidoCompras, objItemReqCompra As ClassItemReqCompras) As Long
'Busca no BD relacionamento entre as tabelas ItensPedCompra, ItemRCItemPC e ItemReqCompras a partir do objItemReqCompras passado
'retorna a tela de Pedido de Compras que será chamada : Pedidos Associados, Enviado, novo, Baixado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemReqCompra_Busca_PedidoCompra
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITEMREQCOMPRA_BUSCA_PEDIDOCOMPRA", X)
    If lErro <> SUCESSO Then gError 71150
    
    ItemReqCompra_Busca_PedidoCompra = X.ItemReqCompra_Busca_PedidoCompra(iTela, objPedidoCompras, objItemReqCompra)

    Exit Function
    
Erro_ItemReqCompra_Busca_PedidoCompra:
    
    ItemReqCompra_Busca_PedidoCompra = gErr
    
    Select Case gErr

        Case 71150
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145214)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompras_Le(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê em RequisiçãoCompras a Requisição passada por objRequisiçãoCompras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompras_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRAS_LE", X)
    If lErro <> SUCESSO Then gError 71151
    
    RequisicaoCompras_Le = X.RequisicaoCompras_Le(objRequisicaoCompras)

    Exit Function
    
Erro_RequisicaoCompras_Le:
    
    RequisicaoCompras_Le = gErr
    
    Select Case gErr

        Case 71151
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145215)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompraEnv_Baixar(objRequisicaoCompras As ClassRequisicaoCompras) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompraEnv_Baixar
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRAENV_BAIXAR", X)
    If lErro <> SUCESSO Then gError 71152
    
    RequisicaoCompraEnv_Baixar = X.RequisicaoCompraEnv_Baixar(objRequisicaoCompras)

    Exit Function
    
Erro_RequisicaoCompraEnv_Baixar:
    
    RequisicaoCompraEnv_Baixar = gErr
    
    Select Case gErr

        Case 71152
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145216)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompra_Exclui(objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Exclui a Requisição de Compras a partir do NumIntDoc passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompra_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRA_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71153
    
    RequisicaoCompra_Exclui = X.RequisicaoCompra_Exclui(objRequisicaoCompra)

    Exit Function
    
Erro_RequisicaoCompra_Exclui:
    
    RequisicaoCompra_Exclui = gErr
    
    Select Case gErr

        Case 71153
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145217)
        
    End Select
    
    Exit Function

End Function

Function ItensReqCompras_Le_ReqCompras(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê os Ítens da Requisição Compras a partir do código passado em objRequisiçãoCompras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqCompras_Le_ReqCompras
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSREQCOMPRAS_LE_REQCOMPRAS", X)
    If lErro <> SUCESSO Then gError 71154
    
    ItensReqCompras_Le_ReqCompras = X.ItensReqCompras_Le_ReqCompras(objRequisicaoCompras)

    Exit Function
    
Erro_ItensReqCompras_Le_ReqCompras:
    
    ItensReqCompras_Le_ReqCompras = gErr
    
    Select Case gErr

        Case 71154
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145218)
        
    End Select
    
    Exit Function

End Function

Function ItensReqModelo_Le(objReqModelo As ClassRequisicaoModelo) As Long
'Lê os Itens de uma Requisição Modelo a partir do Código dessa Requisição

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqModelo_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSREQMODELO_LE", X)
    If lErro <> SUCESSO Then gError 71155
    
    ItensReqModelo_Le = X.ItensReqModelo_Le(objReqModelo)

    Exit Function
    
Erro_ItensReqModelo_Le:
    
    ItensReqModelo_Le = gErr
    
    Select Case gErr

        Case 71155
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145219)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompra_Grava(objRequisicaoCompra As ClassRequisicaoCompras, Optional bGeracao As Boolean = False) As Long
'Grava a Requisição de Compra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompra_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRA_GRAVA", X)
    If lErro <> SUCESSO Then gError 71156
    
    RequisicaoCompra_Grava = X.RequisicaoCompra_Grava(objRequisicaoCompra, bGeracao)

    Exit Function
    
Erro_RequisicaoCompra_Grava:
    
    RequisicaoCompra_Grava = gErr
    
    Select Case gErr

        Case 71156
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145220)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoModelo_Exclui(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Exclui a Requisição Modelo a partir do NumIntDoc passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoModelo_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOMODELO_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71157
    
    RequisicaoModelo_Exclui = X.RequisicaoModelo_Exclui(objRequisicaoModelo)

    Exit Function
    
Erro_RequisicaoModelo_Exclui:
    
    RequisicaoModelo_Exclui = gErr
    
    Select Case gErr

        Case 71157
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145221)
        
    End Select
    
    Exit Function

End Function

Function ItensReqModelo_Exclui(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Exclui os Itens da Requisição Modelo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqModelo_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSREQMODELO_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71158
    
    ItensReqModelo_Exclui = X.ItensReqModelo_Exclui(objRequisicaoModelo)

    Exit Function
    
Erro_ItensReqModelo_Exclui:
    
    ItensReqModelo_Exclui = gErr
    
    Select Case gErr

        Case 71158
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145222)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoModelo_Grava(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Grava a Requisição Modelo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoModelo_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOMODELO_GRAVA", X)
    If lErro <> SUCESSO Then gError 71159
    
    RequisicaoModelo_Grava = X.RequisicaoModelo_Grava(objRequisicaoModelo)

    Exit Function
    
Erro_RequisicaoModelo_Grava:
    
    RequisicaoModelo_Grava = gErr
    
    Select Case gErr

        Case 71159
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145223)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoModelo_Lock_Gravacao(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Faz Lock nas tabelas Associadas a Requisição Modelo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoModelo_Lock_Gravacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOMODELO_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then gError 71160
    
    RequisicaoModelo_Lock_Gravacao = X.RequisicaoModelo_Lock_Gravacao(objRequisicaoModelo)

    Exit Function
    
Erro_RequisicaoModelo_Lock_Gravacao:
    
    RequisicaoModelo_Lock_Gravacao = gErr
    
    Select Case gErr

        Case 71160
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145224)
        
    End Select
    
    Exit Function

End Function

Function ItensReqModelo_Lock_Gravacao(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Faz lock nas tabelas associados ao item da Requisição Modelo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqModelo_Lock_Gravacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSREQMODELO_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then gError 71161
    
    ItensReqModelo_Lock_Gravacao = X.ItensReqModelo_Lock_Gravacao(objRequisicaoModelo)

    Exit Function
    
Erro_ItensReqModelo_Lock_Gravacao:
    
    ItensReqModelo_Lock_Gravacao = gErr
    
    Select Case gErr

        Case 71161
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145225)
        
    End Select
    
    Exit Function

End Function

Function ItensReqModelo_Inclui(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Insere os Itens da Requisição Modelo no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqModelo_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSREQMODELO_INCLUI", X)
    If lErro <> SUCESSO Then gError 71162
    
    ItensReqModelo_Inclui = X.ItensReqModelo_Inclui(objRequisicaoModelo)

    Exit Function
    
Erro_ItensReqModelo_Inclui:
    
    ItensReqModelo_Inclui = gErr
    
    Select Case gErr

        Case 71162
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145226)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoModelo_Le(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Lê dados da RequisiçãoModelo a partir do Número Interno passado em objRequisiçãoModelo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoModelo_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOMODELO_LE", X)
    If lErro <> SUCESSO Then gError 71163
    
    RequisicaoModelo_Le = X.RequisicaoModelo_Le(objRequisicaoModelo)

    Exit Function
    
Erro_RequisicaoModelo_Le:
    
    RequisicaoModelo_Le = gErr
    
    Select Case gErr

        Case 71163
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145227)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoModelo_Le_Codigo(objRequisicaoModelo As ClassRequisicaoModelo) As Long
'Lê em RequisiçãoModelo a partir do Código passado em objRequisiçãoModelo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoModelo_Le_Codigo
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOMODELO_LE_CODIGO", X)
    If lErro <> SUCESSO Then gError 71164
    
    RequisicaoModelo_Le_Codigo = X.RequisicaoModelo_Le_Codigo(objRequisicaoModelo)

    Exit Function
    
Erro_RequisicaoModelo_Le_Codigo:
    
    RequisicaoModelo_Le_Codigo = gErr
    
    Select Case gErr

        Case 71164
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145228)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoModelo_Codigo_Automatico(lCodigo As Long) As Long
'Gera código automático para a próxima Requisição Modelo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoModelo_Codigo_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOMODELO_CODIGO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71165

    RequisicaoModelo_Codigo_Automatico = X.RequisicaoModelo_Codigo_Automatico(lCodigo)

    Exit Function
    
Erro_RequisicaoModelo_Codigo_Automatico:
    
    RequisicaoModelo_Codigo_Automatico = gErr
    
    Select Case gErr

        Case 71165
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145229)
        
    End Select
    
    Exit Function

End Function

Function Requisitante_Grava(objRequisitante As ClassRequisitante) As Long
'Grava o Requisitante no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Requisitante_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISITANTE_GRAVA", X)
    If lErro <> SUCESSO Then gError 71166

    Requisitante_Grava = X.Requisitante_Grava(objRequisitante)

    Exit Function
    
Erro_Requisitante_Grava:
    
    Requisitante_Grava = gErr
    
    Select Case gErr

        Case 71166
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145230)
        
    End Select
    
    Exit Function

End Function

Function Requisitante_Exclui(objRequisitante As ClassRequisitante) As Long
'Exclui o Requisitante

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Requisitante_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISITANTE_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71167

    Requisitante_Exclui = X.Requisitante_Exclui(objRequisitante)

    Exit Function
    
Erro_Requisitante_Exclui:
    
    Requisitante_Exclui = gErr
    
    Select Case gErr

        Case 71167
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145231)
        
    End Select
    
    Exit Function

End Function

Function Requisitante_Automatico(lCodigo As Long) As Long
'Gera o proximo codigo da Tabela de Requisitante

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Requisitante_Automatico
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISITANTE_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 71168

    Requisitante_Automatico = X.Requisitante_Automatico(lCodigo)

    Exit Function
    
Erro_Requisitante_Automatico:
    
    Requisitante_Automatico = gErr
    
    Select Case gErr

        Case 71168
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145232)
        
    End Select
    
    Exit Function

End Function

Function Requisitante_Valida_Exclusao(objRequisitante As ClassRequisitante, alComando() As Long) As Long
'Verifica se o Requisitante a ser excluído está vinculado a alguma outra tabela

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Requisitante_Valida_Exclusao
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISITANTE_AVALIDA_EXCLUSAO", X)
    If lErro <> SUCESSO Then gError 71169

    Requisitante_Valida_Exclusao = X.Requisitante_Valida_Exclusao(objRequisitante, alComando())

    Exit Function
    
Erro_Requisitante_Valida_Exclusao:
    
    Requisitante_Valida_Exclusao = gErr
    
    Select Case gErr

        Case 71169
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145233)
        
    End Select
    
    Exit Function

End Function

Function TipoBloqueioPC_Grava(objTipoDeBloqueioPC As ClassTipoBloqueioPC) As Long
'Grava no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoBloqueioPC_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_TIPOBLOQUEIOPC_GRAVA", X)
    If lErro <> SUCESSO Then gError 71170

    TipoBloqueioPC_Grava = X.TipoBloqueioPC_Grava(objTipoDeBloqueioPC)

    Exit Function
    
Erro_TipoBloqueioPC_Grava:
    
    TipoBloqueioPC_Grava = gErr
    
    Select Case gErr

        Case 71170
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145234)
        
    End Select
    
    Exit Function

End Function

Function TipoBloqueioPC_Exclui(objTipoDeBloqueioPC As ClassTipoBloqueioPC) As Long
'Exclui do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoBloqueioPC_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("COM_TIPOBLOQUEIOPC_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71171

    TipoBloqueioPC_Exclui = X.TipoBloqueioPC_Exclui(objTipoDeBloqueioPC)

    Exit Function
    
Erro_TipoBloqueioPC_Exclui:
    
    TipoBloqueioPC_Exclui = gErr
    
    Select Case gErr

        Case 71171
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145235)
        
    End Select
    
    Exit Function

End Function

Function TipoDeBloqueioPC_Le(objTipoDeBloqueioPC As ClassTipoBloqueioPC) As Long
'Le todos os dados da tabela TipoDeBloqueioPC, correspondente a chave passada por parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeBloqueioPC_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_TIPOBLOQUEIOPC_LE", X)
    If lErro <> SUCESSO Then gError 71172

    TipoDeBloqueioPC_Le = X.TipoDeBloqueioPC_Le(objTipoDeBloqueioPC)

    Exit Function
    
Erro_TipoDeBloqueioPC_Le:
    
    TipoDeBloqueioPC_Le = gErr
    
    Select Case gErr

        Case 71172
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145236)
        
    End Select
    
    Exit Function

End Function

Function GeracaoPedCompraConc_Le_Concorrencias(objGeracaoPedidoCompra As ClassGeracaoPedCompraConc) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoPedCompraConc_Le_Concorrencias
    
    lErro = Rotina_ObterObjetoFuncao("COM_GERACAOPEDCOMPRACONC_LE_CONCORRENCIA", X)
    If lErro <> SUCESSO Then gError 71178

    GeracaoPedCompraConc_Le_Concorrencias = X.GeracaoPedCompraConc_Le_Concorrencias(objGeracaoPedidoCompra)

    Exit Function
    
Erro_GeracaoPedCompraConc_Le_Concorrencias:
    
    GeracaoPedCompraConc_Le_Concorrencias = gErr
    
    Select Case gErr

        Case 71178
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145237)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCompra_Le_CotacaoItemConcorrencia(objCotacaoItemConc As ClassCotacaoItemConc, objConcorrencia As ClassConcorrencia) As Long
''Lê a Concorrencia a partir do NumIntDoc de CotacaoItemConcorrencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCompra_Le_CotacaoItemConcorrencia
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITENSPEDCOMPRA_LE_COTACAOITEMCONCORR", X)
    If lErro <> SUCESSO Then gError 71256

    ItensPedCompra_Le_CotacaoItemConcorrencia = X.ItensPedCompra_Le_CotacaoItemConcorrencia(objCotacaoItemConc, objConcorrencia)

    Exit Function
    
Erro_ItensPedCompra_Le_CotacaoItemConcorrencia:
    
    ItensPedCompra_Le_CotacaoItemConcorrencia = gErr
    
    Select Case gErr

        Case 71256
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145238)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompra_Exclui_EmTrans(objRequisicaoCompra As ClassRequisicaoCompras) As Long
'Exclui a Requisição de Compras a partir do NumIntDoc passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompra_Exclui_EmTrans
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRA_EXCLUI_EMTRANS", X)
    If lErro <> SUCESSO Then gError 71259

    RequisicaoCompra_Exclui_EmTrans = X.RequisicaoCompra_Exclui_EmTrans(objRequisicaoCompra)

    Exit Function
    
Erro_RequisicaoCompra_Exclui_EmTrans:
    
    RequisicaoCompra_Exclui_EmTrans = gErr
    
    Select Case gErr

        Case 71259
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145239)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoComprasTodas_Le_Numero(objRequisicaoCompras As ClassRequisicaoCompras) As Long
'Lê em RequisiçãoComprasTodas a Requisição(baixada ou não) passada por objRequisiçãoCompras com Código e FlialEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoComprasTodas_Le_Numero
    
    lErro = Rotina_ObterObjetoFuncao("COM_REQUISICAOCOMPRASTODAS_LE_NUMERO", X)
    If lErro <> SUCESSO Then gError 71260

    RequisicaoComprasTodas_Le_Numero = X.RequisicaoComprasTodas_Le_Numero(objRequisicaoCompras)

    Exit Function
    
Erro_RequisicaoComprasTodas_Le_Numero:
    
    RequisicaoComprasTodas_Le_Numero = gErr
    
    Select Case gErr

        Case 71260
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145240)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompraBaixado_Atualiza_DataEmissao(objPedidoCompra As ClassPedidoCompras) As Long
'Atualiza data de emissao no BD, em PedidoCompraBaixado, para a data atual

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompraBaixado_Atualiza_DataEmissao
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRABAIXADO_ATUALIZA_DATAEMI", X)
    If lErro <> SUCESSO Then gError 71261

    PedidoCompraBaixado_Atualiza_DataEmissao = X.PedidoCompraBaixado_Atualiza_DataEmissao(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompraBaixado_Atualiza_DataEmissao:
    
    PedidoCompraBaixado_Atualiza_DataEmissao = gErr
    
    Select Case gErr

        Case 71261
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145241)
        
    End Select
    
    Exit Function

End Function

Function TP_Comprador_Le(objCompradorMaskEdBox As Object, objComprador As ClassComprador, Optional iCria As Integer = 1) As Long
'Lê o Comprador com Código ou NomeRed em objCompradorMaskEdBox.Text
'Devolve em objComprador. Coloca código-NomeReduzido no .Text

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TP_Comprador_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_TP_COMPRADOR_LE", X)
    If lErro <> SUCESSO Then gError 71269

    TP_Comprador_Le = X.TP_Comprador_Le(objCompradorMaskEdBox, objComprador, iCria)

    Exit Function
    
Erro_TP_Comprador_Le:
    
    TP_Comprador_Le = gErr
    
    Select Case gErr

        Case 71269
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145242)
        
    End Select
    
    Exit Function

End Function

Function PedidoCotacao_Atualiza_DataEmissao(objPedidoCotacao As ClassPedidoCotacao) As Long
'Atualiza data de emissao no BD para a data atual

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCotacao_Atualiza_DataEmissao
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOTACAO_ATUALIZA_DATAEMISSAO", X)
    If lErro <> SUCESSO Then gError 71271

    PedidoCotacao_Atualiza_DataEmissao = X.PedidoCotacao_Atualiza_DataEmissao(objPedidoCotacao)

    Exit Function
    
Erro_PedidoCotacao_Atualiza_DataEmissao:
    
    PedidoCotacao_Atualiza_DataEmissao = gErr
    
    Select Case gErr

        Case 71271
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145243)
        
    End Select
    
    Exit Function

End Function

Function PedidoCotacao_Grava_PedidoCompra(colPedidoCompra As Collection) As Long
'Grava os Pedidos de Compra gerados a partir do Pedido de Cotacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCotacao_Grava_PedidoCompra
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOTACAO_GRAVA_PEDIDOCOMPRA", X)
    If lErro <> SUCESSO Then gError 71272

    PedidoCotacao_Grava_PedidoCompra = X.PedidoCotacao_Grava_PedidoCompra(colPedidoCompra)

    Exit Function
    
Erro_PedidoCotacao_Grava_PedidoCompra:
    
    PedidoCotacao_Grava_PedidoCompra = gErr
    
    Select Case gErr

        Case 71272
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145244)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Grava_EmTrans(objPedidoCompra As ClassPedidoCompras) As Long
'Grava Pedido de Compra no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Grava_EmTrans
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRA_GRAVA_EMTRANS", X)
    If lErro <> SUCESSO Then gError 71273

    PedidoCompra_Grava_EmTrans = X.PedidoCompra_Grava_EmTrans(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_Grava_EmTrans:
    
    PedidoCompra_Grava_EmTrans = gErr
    
    Select Case gErr

        Case 71273
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145245)
        
    End Select
    
    Exit Function

End Function

Function ItemPC_Associa_ItemRC(objItemPC As ClassItemPedCompra) As Long
'Faz o link entre ItemRC e ItemPC e atualiza QuantPedida de ItemRC e
'desfaz o link na tabela CotacaoProdutoItemRC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemPC_Associa_ItemRC
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITEMPC_ASSOCIA_ITEMRC", X)
    If lErro <> SUCESSO Then gError 71274

    ItemPC_Associa_ItemRC = X.ItemPC_Associa_ItemRC(objItemPC)

    Exit Function
    
Erro_ItemPC_Associa_ItemRC:
    
    ItemPC_Associa_ItemRC = gErr
    
    Select Case gErr

        Case 71274
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145246)
        
    End Select
    
    Exit Function

End Function

Function ItemPedCotacao_Le_RequisicaoCompra(objItemPedCotacao As ClassItemPedCotacao, colReqCompras As Collection, lCodigo As Long, iFilialEmpresa As Integer) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemPedCotacao_Le_RequisicaoCompra
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITEMPEDCOTACAO_LE_REQUISICAOCOMPRA", X)
    If lErro <> SUCESSO Then gError 71275

    ItemPedCotacao_Le_RequisicaoCompra = X.ItemPedCotacao_Le_RequisicaoCompra(objItemPedCotacao, colReqCompras, lCodigo, iFilialEmpresa)

    Exit Function
    
Erro_ItemPedCotacao_Le_RequisicaoCompra:
    
    ItemPedCotacao_Le_RequisicaoCompra = gErr
    
    Select Case gErr

        Case 71275
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145247)
        
    End Select
    
    Exit Function

End Function

Function Valida_Quantidade_ItemPedCotacao(objPedCotacao As ClassPedidoCotacao) As Long
'Faz a validação do ItemPedCotacao, ou seja, verifica se a quantidade
'do Item de Pedido de Cotacao está toda vinculada à Itens de Requisicao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Valida_Quantidade_ItemPedCotacao
    
    lErro = Rotina_ObterObjetoFuncao("COM_VALIDA_QUANTIDADE_ITEMPEDCOTACAO", X)
    If lErro <> SUCESSO Then gError 71276

    Valida_Quantidade_ItemPedCotacao = X.Valida_Quantidade_ItemPedCotacao(objPedCotacao)

    Exit Function
    
Erro_Valida_Quantidade_ItemPedCotacao:
    
    Valida_Quantidade_ItemPedCotacao = gErr
    
    Select Case gErr

        Case 71276
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145248)
        
    End Select
    
    Exit Function

End Function

Function TiposProduto_Le_Todos(colCod_DescReduzida As AdmColCodigoNome) As Long
'Le todos os Codigos e DescReduzida dos tipos de produto cadastrados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposProduto_Le_Todos
    
    lErro = Rotina_ObterObjetoFuncao("COM_TIPOSPRODUTO_LE_TODOS", X)
    If lErro <> SUCESSO Then gError 71338

    TiposProduto_Le_Todos = X.TiposProduto_Le_Todos(colCod_DescReduzida)

    Exit Function
    
Erro_TiposProduto_Le_Todos:
    
    TiposProduto_Le_Todos = gErr
    
    Select Case gErr

        Case 71338
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145249)
        
    End Select
    
    Exit Function

End Function

Function ItemCotacao_Le(objItemCotacao As ClassItemCotacao) As Long
'Lê o item de cotação com o numintdsoc passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemCotacao_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITEMCOTACAO_LE", X)
    If lErro <> SUCESSO Then gError 71339

    ItemCotacao_Le = X.ItemCotacao_Le(objItemCotacao)

    Exit Function
    
Erro_ItemCotacao_Le:
    
    ItemCotacao_Le = gErr
    
    Select Case gErr

        Case 71339
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145250)
        
    End Select
    
    Exit Function

End Function

Function ItensReqCompra_Le_GeracaoPC(objGeracaoPedCompraReq As ClassGeracaoPedCompraReq) As Long
'Lê os itens de Requisição dos Tipos passados por parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqCompra_Le_GeracaoPC
    
    lErro = Rotina_ObterObjetoFuncao("COM_ITEMCOTACAO_LE", X)
    If lErro <> SUCESSO Then gError 71340

    ItensReqCompra_Le_GeracaoPC = X.ItensReqCompra_Le_GeracaoPC(objGeracaoPedCompraReq)

    Exit Function
    
Erro_ItensReqCompra_Le_GeracaoPC:
    
    ItensReqCompra_Le_GeracaoPC = gErr
    
    Select Case gErr

        Case 71340
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145251)
        
    End Select
    
    Exit Function

End Function

Function PedidosAbertos_Grava_Estatistica(objPedCompras As ClassPedidoCompras) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosAbertos_Grava_Estatistica
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOSABERTOS_GRAVA_ESTATISTICA", X)
    If lErro <> SUCESSO Then gError 71341

    PedidosAbertos_Grava_Estatistica = X.PedidosAbertos_Grava_Estatistica(objPedCompras)

    Exit Function
    
Erro_PedidosAbertos_Grava_Estatistica:
    
    PedidosAbertos_Grava_Estatistica = gErr
    
    Select Case gErr

        Case 71341
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145252)
        
    End Select
    
    Exit Function

End Function

Function FornecedorProduto_Grava_Estatisticas(objNFiscal As ClassNFiscal) As Long
'Calcula a Estastistica e Chama a Função de Gravacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FornecedorProduto_Grava_Estatisticas
    
    lErro = Rotina_ObterObjetoFuncao("COM_FORNECEDORPRODUTO_GRAVA_ESTATISTICAS", X)
    If lErro <> SUCESSO Then gError 71342

    FornecedorProduto_Grava_Estatisticas = X.FornecedorProduto_Grava_Estatisticas(objNFiscal)

    Exit Function
    
Erro_FornecedorProduto_Grava_Estatisticas:
    
    FornecedorProduto_Grava_Estatisticas = gErr
    
    Select Case gErr

        Case 71342
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145253)
        
    End Select
    
    Exit Function

End Function

Function FornecedorProduto_Grava_EstatisticasCom(objNFiscal As ClassNFiscal, colItensPedComprasFechados As Collection) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FornecedorProduto_Grava_EstatisticasCom
    
    lErro = Rotina_ObterObjetoFuncao("COM_FORNECEDORPRODUTO_GRAVA_ESTATCOM", X)
    If lErro <> SUCESSO Then gError 71343

    FornecedorProduto_Grava_EstatisticasCom = X.FornecedorProduto_Grava_EstatisticasCom(objNFiscal, colItensPedComprasFechados)

    Exit Function
    
Erro_FornecedorProduto_Grava_EstatisticasCom:
    
    FornecedorProduto_Grava_EstatisticasCom = gErr
    
    Select Case gErr

        Case 71343
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145254)
        
    End Select
    
    Exit Function

End Function

Function UltimoItemPCFechado_Grava(colPedCompraBaixados As Collection) As Long
'Calcula a Estatistica para o Pedido de Compra Fechado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UltimoItemPCFechado_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_ULTIMOITEMPCFECHADO_GRAVA", X)
    If lErro <> SUCESSO Then gError 71344

    UltimoItemPCFechado_Grava = X.UltimoItemPCFechado_Grava(colPedCompraBaixados)

    Exit Function
    
Erro_UltimoItemPCFechado_Grava:
    
    UltimoItemPCFechado_Grava = gErr
    
    Select Case gErr

        Case 71344
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145255)
        
    End Select
    
    Exit Function

End Function

Function UltimaCotacao_Grava(objPedidoCotacao As ClassPedidoCotacao, objCotacao As ClassCotacao) As Long
'Calcula a Estatistica de Cotacao e Grava no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UltimaCotacao_Grava
    
    lErro = Rotina_ObterObjetoFuncao("COM_ULTIMACOTACAO_GRAVA", X)
    If lErro <> SUCESSO Then gError 71345

    UltimaCotacao_Grava = X.UltimaCotacao_Grava(objPedidoCotacao, objCotacao)

    Exit Function
    
Erro_UltimaCotacao_Grava:
    
    UltimaCotacao_Grava = gErr
    
    Select Case gErr

        Case 71345
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145256)
        
    End Select
    
    Exit Function

End Function

Function Cotacao_Le_Codigo(objCotacao As ClassCotacao) As Long
'Lê a tabela Cotacao utilizando a FilialEmpresa/Codigo como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Cotacao_Le_Codigo
    
    lErro = Rotina_ObterObjetoFuncao("COM_COTACAO_LE_CODIGO", X)
    If lErro <> SUCESSO Then gError 89164
    
    Cotacao_Le_Codigo = X.Cotacao_Le_Codigo(objCotacao)

    Exit Function
    
Erro_Cotacao_Le_Codigo:
    
    Cotacao_Le_Codigo = gErr
    
    Select Case gErr

        Case 89164
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145257)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompraBaixado_Le(objPedidoCompra As ClassPedidoCompras) As Long
'Lê os dados do Pedido de Compra,cujo numero interno foi passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompraBaixado_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDIDOCOMPRABAIXADO_LE", X)
    If lErro <> SUCESSO Then gError 89235
    
    PedidoCompraBaixado_Le = X.PedidoCompraBaixado_Le(objPedidoCompra)

    Exit Function
    
Erro_PedidoCompraBaixado_Le:
    
    PedidoCompraBaixado_Le = gErr
    
    Select Case gErr

        Case 89235
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145258)
        
    End Select
    
    Exit Function

End Function

Function UltimaCotacao_Le_FornecedorProduto(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'retorna os dados da ultima cotacao (maior data de pedido de cotação) do produto/Fornecedor passado como parametro.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UltimaCotacao_Le_FornecedorProduto
    
    lErro = Rotina_ObterObjetoFuncao("COM_ULTIMACOTACAO_LE_FORNECEDORPRODUTO", X)
    If lErro <> SUCESSO Then gError 89329
    
    UltimaCotacao_Le_FornecedorProduto = X.UltimaCotacao_Le_FornecedorProduto(objFornecedorProdutoFF)

    Exit Function
    
Erro_UltimaCotacao_Le_FornecedorProduto:
    
    UltimaCotacao_Le_FornecedorProduto = gErr
    
    Select Case gErr

        Case 89329
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145259)
        
    End Select
    
    Exit Function

End Function

Function UltimaCompra_Le_FornecedorProduto(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'retorna os dados da ultima compra global para o produto/Fornecedor/FilialForn passado como parametro.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UltimaCompra_Le_FornecedorProduto
    
    lErro = Rotina_ObterObjetoFuncao("COM_ULTIMACOMPRA_LE_FORNECEDORPRODUTO", X)
    If lErro <> SUCESSO Then gError 89348
    
    UltimaCompra_Le_FornecedorProduto = X.UltimaCompra_Le_FornecedorProduto(objFornecedorProdutoFF)

    Exit Function
    
Erro_UltimaCompra_Le_FornecedorProduto:
    
    UltimaCompra_Le_FornecedorProduto = gErr
    
    Select Case gErr

        Case 89348
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145260)
        
    End Select
    
    Exit Function

End Function

Function UltimoItemPedCompraFechado_Le(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'retorna os dados do ultimo item de pedido de compra com status fechado para o produto/Fornecedor/FilialForn passado como parametro.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UltimoItemPedCompraFechado_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_ULTIMOITEMPEDCOMPRAFECHADO_LE", X)
    If lErro <> SUCESSO Then gError 89353
    
    UltimoItemPedCompraFechado_Le = X.UltimoItemPedCompraFechado_Le(objFornecedorProdutoFF)

    Exit Function
    
Erro_UltimoItemPedCompraFechado_Le:
    
    UltimoItemPedCompraFechado_Le = gErr
    
    Select Case gErr

        Case 89353
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145261)
        
    End Select
    
    Exit Function

End Function

Function QuantProdutoPedAbertos_Le(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'retorna a Quantidade em Unidade de Compra do Produto Pedido pela Filial em questão e ainda não entregue para o produto/Fornecedor/FilialForn passado como parametro.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_QuantProdutoPedAbertos_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_QUANTPRODUTOPEDABERTOS_LE", X)
    If lErro <> SUCESSO Then gError 89367
    
    QuantProdutoPedAbertos_Le = X.QuantProdutoPedAbertos_Le(objFornecedorProdutoFF)

    Exit Function
    
Erro_QuantProdutoPedAbertos_Le:
    
    QuantProdutoPedAbertos_Le = gErr
    
    Select Case gErr

        Case 89367
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145262)
        
    End Select
    
    Exit Function

End Function

Function TempoRessup_Calcula(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Calcula o Tempo de Ressuprimento e devolve em objFornecedorProdutoFF

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TempoRessup_Calcula
    
    lErro = Rotina_ObterObjetoFuncao("COM_TEMPORESSUP_CALCULA", X)
    If lErro <> SUCESSO Then gError 89369
    
    TempoRessup_Calcula = X.TempoRessup_Calcula(objFornecedorProdutoFF)

    Exit Function
    
Erro_TempoRessup_Calcula:
    
    TempoRessup_Calcula = gErr
    
    Select Case gErr

        Case 89367
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145263)
        
    End Select
    
    Exit Function

End Function

Function Pesquisa_Primeira_Compra_FornProd(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'pesquisa a primeira nota fiscal da FilialEmpresa (data de emissao mais antiga) para o Fornecedor/Produto em questão.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Pesquisa_Primeira_Compra_FornProd
    
    lErro = Rotina_ObterObjetoFuncao("COM_PESQUISA_PRIMEIRA_COMPRA_FORNPROD", X)
    If lErro <> SUCESSO Then gError 89383
    
    Pesquisa_Primeira_Compra_FornProd = X.Pesquisa_Primeira_Compra_FornProd(objFornecedorProdutoFF)

    Exit Function
    
Erro_Pesquisa_Primeira_Compra_FornProd:
    
    Pesquisa_Primeira_Compra_FornProd = gErr
    
    Select Case gErr

        Case 89383
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145264)
        
    End Select
    
    Exit Function

End Function

Function Pesquisa_Ultima_Compra_FornProd(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'pesquisa a ultima nota fiscal da FilialEmpresa (data de emissao mais recente) para o Fornecedor/Produto em questão.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Pesquisa_Ultima_Compra_FornProd
    
    lErro = Rotina_ObterObjetoFuncao("COM_PESQUISA_ULTIMA_COMPRA_FORNPROD", X)
    If lErro <> SUCESSO Then gError 89384
    
    Pesquisa_Ultima_Compra_FornProd = X.Pesquisa_Ultima_Compra_FornProd(objFornecedorProdutoFF)

    Exit Function
    
Erro_Pesquisa_Ultima_Compra_FornProd:
    
    Pesquisa_Ultima_Compra_FornProd = gErr
    
    Select Case gErr

        Case 89384
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145265)
        
    End Select
    
    Exit Function

End Function

Function Pesquisa_Primeira_Compra_FornFilProd(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'pesquisa a primeira nota fiscal da FilialEmpresa (data de emissao mais antiga) para o Fornecedor/FilialForn/Produto em questão.

Dim X As Object, lErro As Long

On Error GoTo Erro_Pesquisa_Primeira_Compra_FornFilProd
    
    lErro = Rotina_ObterObjetoFuncao("COM_PESQUISA_PRIMEIRA_COMPRA_FORNFILPROD", X)
    If lErro <> SUCESSO Then gError 89385
    
    Pesquisa_Primeira_Compra_FornFilProd = X.Pesquisa_Primeira_Compra_FornFilProd(objFornecedorProdutoFF)

    Exit Function
    
Erro_Pesquisa_Primeira_Compra_FornFilProd:
    
    Pesquisa_Primeira_Compra_FornFilProd = gErr
    
    Select Case gErr

        Case 89385
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145266)
        
    End Select
    
    Exit Function

End Function

Function Pesquisa_Ultima_Compra_FornFilProd(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'pesquisa a ultima nota fiscal da FilialEmpresa (data de emissao mais recente) para o Fornecedor/FilialForn/Produto em questão.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Pesquisa_Ultima_Compra_FornFilProd
    
    lErro = Rotina_ObterObjetoFuncao("COM_PESQUISA_ULTIMA_COMPRA_FORNFILPROD", X)
    If lErro <> SUCESSO Then gError 89385
    
    Pesquisa_Ultima_Compra_FornFilProd = X.Pesquisa_Ultima_Compra_FornFilProd(objFornecedorProdutoFF)

    Exit Function
    
Erro_Pesquisa_Ultima_Compra_FornFilProd:
    
    Pesquisa_Ultima_Compra_FornFilProd = gErr
    
    Select Case gErr

        Case 89385
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145267)
        
    End Select
    
    Exit Function

End Function

Function Pesquisa_Primeira_Compra_Forn(objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a primeira nota fiscal da FilialEmpresa (data de emissao mais antiga) para o Fornecedor em questão.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Pesquisa_Primeira_Compra_Forn
    
    lErro = Rotina_ObterObjetoFuncao("COM_PESQUISA_PRIMEIRA_COMPRA_FORN", X)
    If lErro <> SUCESSO Then gError 89398
    
    Pesquisa_Primeira_Compra_Forn = X.Pesquisa_Primeira_Compra_Forn(objFilialFornecedorEst)

    Exit Function
    
Erro_Pesquisa_Primeira_Compra_Forn:
    
    Pesquisa_Primeira_Compra_Forn = gErr
    
    Select Case gErr

        Case 89398
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145268)
        
    End Select
    
    Exit Function

End Function

Function Pesquisa_Ultima_Compra_Forn(objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a ultima nota fiscal da FilialEmpresa (data de emissao mais recente) para o Fornecedor em questão.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Pesquisa_Ultima_Compra_Forn
    
    lErro = Rotina_ObterObjetoFuncao("COM_PESQUISA_ULTIMA_COMPRA_FORN", X)
    If lErro <> SUCESSO Then gError 89399
    
    Pesquisa_Ultima_Compra_Forn = X.Pesquisa_Ultima_Compra_Forn(objFilialFornecedorEst)

    Exit Function
    
Erro_Pesquisa_Ultima_Compra_Forn:
    
    Pesquisa_Ultima_Compra_Forn = gErr
    
    Select Case gErr

        Case 89399
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145269)
        
    End Select
    
    Exit Function

End Function

Function Pesquisa_Primeira_Compra_FornFil(objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a primeira nota fiscal da FilialEmpresa (data de emissao mais antiga) para o Fornecedor/FilialForn em questão.

Dim X As Object, lErro As Long

On Error GoTo Erro_Pesquisa_Primeira_Compra_FornFil
    
    lErro = Rotina_ObterObjetoFuncao("COM_PESQUISA_PRIMEIRA_COMPRA_FORNFIL", X)
    If lErro <> SUCESSO Then gError 89400
    
    Pesquisa_Primeira_Compra_FornFil = X.Pesquisa_Primeira_Compra_FornFil(objFilialFornecedorEst)

    Exit Function
    
Erro_Pesquisa_Primeira_Compra_FornFil:
    
    Pesquisa_Primeira_Compra_FornFil = gErr
    
    Select Case gErr

        Case 89400
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145270)
        
    End Select
    
    Exit Function

End Function

Function Pesquisa_Ultima_Compra_FornFil(objFilialFornecedorEst As ClassFilialFornecedorEst) As Long
'pesquisa a ultima nota fiscal da FilialEmpresa (data de emissao mais recente) para o Fornecedor/FilialForn em questão.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Pesquisa_Ultima_Compra_FornFil
    
    lErro = Rotina_ObterObjetoFuncao("COM_PESQUISA_ULTIMA_COMPRA_FORNFIL", X)
    If lErro <> SUCESSO Then gError 89401
    
    Pesquisa_Ultima_Compra_FornFil = X.Pesquisa_Ultima_Compra_FornFil(objFilialFornecedorEst)

    Exit Function
    
Erro_Pesquisa_Ultima_Compra_FornFil:
    
    Pesquisa_Ultima_Compra_FornFil = gErr
    
    Select Case gErr

        Case 89401
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145271)
        
    End Select
    
    Exit Function

End Function

Function TempoRessupFilial_Calcula(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Calcula o Tempo de Ressuprimento da Filial e devolve em objFornecedorProdutoFF

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TempoRessupFilial_Calcula
    
    lErro = Rotina_ObterObjetoFuncao("COM_TEMPORESSUPFILIAL_CALCULA", X)
    If lErro <> SUCESSO Then gError 89412
    
    TempoRessupFilial_Calcula = X.TempoRessupFilial_Calcula(objFornecedorProdutoFF)

    Exit Function
    
Erro_TempoRessupFilial_Calcula:
    
    TempoRessupFilial_Calcula = gErr
    
    Select Case gErr

        Case 89412
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145272)
        
    End Select
    
    Exit Function

End Function

Function ConcorrenciaN_Le(objConcorrencia As ClassConcorrencia) As Long
'Lê dados da concorrência passada pelo código em objConcorrencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ConcorrenciaN_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_CONCORRENCIAN_LE", X)
    If lErro <> SUCESSO Then gError 89867
    
    ConcorrenciaN_Le = X.ConcorrenciaN_Le(objConcorrencia)

    Exit Function
    
Erro_ConcorrenciaN_Le:
    
    ConcorrenciaN_Le = gErr
    
    Select Case gErr

        Case 89867
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145273)
        
    End Select
    
    Exit Function

End Function

Function PedCompraGerado_Le(objPedidoCompra As ClassPedidoCompras) As Long
'Lê os dados do Pedido de Compra Gerado,cujo numero interno foi passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedCompraGerado_Le
    
    lErro = Rotina_ObterObjetoFuncao("COM_PEDCOMPRAGERADO_LE", X)
    If lErro <> SUCESSO Then gError 89948
    
    PedCompraGerado_Le = X.PedCompraGerado_Le(objPedidoCompra)

    Exit Function
    
Erro_PedCompraGerado_Le:
    
    PedCompraGerado_Le = gErr
    
    Select Case gErr

        Case 89948
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145274)
        
    End Select
    
    Exit Function

End Function

