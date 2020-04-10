VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFornProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Utilizado na função ObterRel_Info_Ult_Compra
Dim objUltFornProdFFCompra As New ClassFornecedorProdutoFF

'Utilizado na função ObterRel_Info_Ult_Cotacao
Dim objUltFornProdFFCotacao As New ClassFornecedorProdutoFF

'Utilizado na função ObterRel_Info_Ult_PC_Fechado
Dim objUltFornProdFFPCFechado As New ClassFornecedorProdutoFF

'Variáveis utilizadas na função ObterRel_Quant_PC_Forn_Produto
Dim dUltQtdPedCompra As Double
Dim iUltFilialEmpresaQtdPC As Integer
Dim sUltProdutoQtdPC As String
Dim lUltFornecedorQtdPC As Long
Dim iUltFilialFornQtdPC As Integer

'Variáveis utilizadas na função ObterRel_Interv_Ressup_Medio
Dim dUltTempoRessupMed As Double
Dim iUltFilialEmpresaIRM As Integer
Dim iUltFilialFornIRM As Integer
Dim lUltFornecedorIRM As Long
Dim sUltProdutoIRM As String

Public Function ObterRel_Quant_PC_Forn_Produto(dQtdPedCompra As Double, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal lFornecedor As Long, ByVal iFilialForn As Integer) As Long
'Chama a função que lê no BD a quantidade total em pedidos abertos para o par Filial Fornecedor X Produto
'Parâmetros (na ordem que devem ser passados: iFilialEmpresa, sProduto, lFornecedor, iFilialForn

Dim lErro
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_ObterRel_Quant_PC_Forn_Produto

    'Se pelo menos um dos parâmetros passados pelo Forprint se alterou
    If iFilialEmpresa <> iUltFilialEmpresaQtdPC Or sProduto <> sUltProdutoQtdPC Or lFornecedor <> lUltFornecedorQtdPC Or iFilialForn <> iUltFilialFornQtdPC Then
    
        'Guarda no obj os parâmetros que serão utilizados pela função
        objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
        objFornecedorProdutoFF.sProduto = sProduto
        objFornecedorProdutoFF.lFornecedor = lFornecedor
        objFornecedorProdutoFF.iFilialForn = iFilialForn
        
        'Lê a quantidade do produto em pedidos de compra para o fornecedor selecionado
        lErro = CF("QuantProdutoPedAbertos_Le", objFornecedorProdutoFF)
        If lErro <> SUCESSO Then gError 79994
        
        'Guarda a quantidade encontrada na variável que será retornada à função chamadora
        dQtdPedCompra = objFornecedorProdutoFF.dQuantPedAbertos
    
    'Se todos os parâmetros são iguais
    Else
    
        'Pega a última quantidade encontrada com esses parâmetros e retonra à função chamadora
        dQtdPedCompra = dUltQtdPedCompra
    
    End If
    
    'Guarda a última quantidade pesquisada
    dUltQtdPedCompra = dQtdPedCompra
    
    'Guarda os últimos parâmetros usados para pesquisa
    iUltFilialEmpresaQtdPC = iFilialEmpresa
    lUltFornecedorQtdPC = lFornecedor
    iUltFilialFornQtdPC = iFilialForn
    sUltProdutoQtdPC = sProduto
    
    ObterRel_Quant_PC_Forn_Produto = SUCESSO
    
    Exit Function

Erro_ObterRel_Quant_PC_Forn_Produto:

    ObterRel_Quant_PC_Forn_Produto = gErr
    
    Select Case gErr
    
        Case 79994
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153156)
        
    End Select

    Exit Function
    
End Function

Public Function ObterRel_Interv_Ressup_Medio(dTempoRessupMed As Double, ByVal iFilialEmpresa As Integer, ByVal lFornecedor As Long, ByVal iFilialForn As Integer, ByVal sProduto As String) As Long
'Chama a função que lê no BD o tempo médio de ressuprimento para o par Fornecedor X Produto
'Parâmetros (na ordem que devem ser passados: iFilialEmpresa, lFornecedor, iFilialForn, sProduto

Dim lErro
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_ObterRel_Interv_Ressup_Medio

    'Se pelo menos um dos parâmetros passados pelo Forprint se alterou
    If iFilialEmpresa <> iUltFilialEmpresaIRM Or sProduto <> sUltProdutoIRM Or lFornecedor <> lUltFornecedorIRM Or iFilialForn <> iUltFilialFornIRM Then

        'Guarda no obj os parâmetros que serão utilizados pela função
        objFornecedorProdutoFF.iFilialEmpresa = iFilialEmpresa
        objFornecedorProdutoFF.lFornecedor = lFornecedor
        objFornecedorProdutoFF.iFilialForn = iFilialForn
        objFornecedorProdutoFF.sProduto = sProduto

        'Lê a quantidade do produto em pedidos de compra para o fornecedor selecionado
        lErro = CF("TempoRessupFilial_Calcula", objFornecedorProdutoFF)
        If lErro <> SUCESSO Then gError 79996

        'Guarda a quantidade encontrada na variável que será retornada à função chamadora
        dTempoRessupMed = objFornecedorProdutoFF.dTempoRessup

    'Se todos os parâmetros são iguais
    Else

        'Pega a última quantidade encontrada com esses parâmetros e retonra à função chamadora
        dTempoRessupMed = dUltTempoRessupMed

    End If

    'Guarda a última quantidade pesquisada
    dUltTempoRessupMed = dTempoRessupMed

    'Guarda os últimos parâmetros usados para pesquisa
    iUltFilialEmpresaIRM = iFilialEmpresa
    lUltFornecedorIRM = lFornecedor
    iUltFilialFornIRM = iFilialForn
    sUltProdutoIRM = sProduto
    
    ObterRel_Interv_Ressup_Medio = SUCESSO

    Exit Function

Erro_ObterRel_Interv_Ressup_Medio:

    ObterRel_Interv_Ressup_Medio = gErr

    Select Case gErr

        Case 79996

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153157)

    End Select

    Exit Function

End Function

Public Function ObterRel_Info_Ult_Compra(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Chama a função que lê no BD os dados da última compra para o Produto / Fornecedor / FilialForn / Filial
'Parâmetros (na ordem que devem ser passados: lFornecedor, iFilialForn, sProduto

Dim lErro

On Error GoTo Erro_ObterRel_Info_Ult_Compra

    'Se pelo menos um dos parâmetros passados pelo Forprint se alterou
    If objFornecedorProdutoFF.iFilialEmpresa <> objUltFornProdFFCompra.iFilialEmpresa Or objFornecedorProdutoFF.lFornecedor <> objUltFornProdFFCompra.lFornecedor Or objFornecedorProdutoFF.iFilialForn <> objUltFornProdFFCompra.iFilialForn Or objFornecedorProdutoFF.sProduto <> objUltFornProdFFCompra.sProduto Then

        'Lê a quantidade do produto em pedidos de compra para o fornecedor selecionado
        lErro = CF("UltimaCompra_Le_FornecedorProduto", objFornecedorProdutoFF)
        If lErro <> SUCESSO Then gError 79998
    
    'Se todos os parâmetros são iguais
    Else

        'Pega os últimos dados encontrados com esses parâmetros e retorna à função chamadora
        Set objFornecedorProdutoFF = objUltFornProdFFCompra

    End If

    'Guarda os dados encontrados para serem utilizados posteriormente em uma pesquisa com os mesmos parâmetros
    Set objUltFornProdFFCompra = objFornecedorProdutoFF

    ObterRel_Info_Ult_Compra = SUCESSO

    Exit Function

Erro_ObterRel_Info_Ult_Compra:

    ObterRel_Info_Ult_Compra = gErr

    Select Case gErr

        Case 79998

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153158)

    End Select

    Exit Function

End Function

Public Function ObterRel_Info_Ult_PC_Fechado(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Chama a função que lê no BD os dados do último pedido de compra recebido para o Produto / Fornecedor / FilialFornecedor
'Parâmetros (na ordem que devem ser passados): iFilialEmpresa, lFornecedor, iFilialForn, sProduto

Dim lErro

On Error GoTo Erro_ObterRel_Info_Ult_PC_Fechado

    'Se pelo menos um dos parâmetros passados pelo Forprint se alterou
    If objFornecedorProdutoFF.iFilialEmpresa <> objUltFornProdFFPCFechado.iFilialEmpresa Or objFornecedorProdutoFF.lFornecedor <> objUltFornProdFFPCFechado.lFornecedor Or objFornecedorProdutoFF.iFilialForn <> objUltFornProdFFPCFechado.iFilialForn Or objFornecedorProdutoFF.sProduto <> objUltFornProdFFPCFechado.sProduto Then
        
        'Lê a quantidade do produto em pedidos de compra para o fornecedor selecionado
        lErro = UltimoItemPedCompraFechado_Le2(objFornecedorProdutoFF)
        If lErro <> SUCESSO Then gError 79999

    'Se todos os parâmetros são iguais
    Else

        'Pega os últimos dados encontrados com esses parâmetros e retorna à função chamadora
        Set objFornecedorProdutoFF = objUltFornProdFFPCFechado

    End If
    
    'Guarda os dados encontrados para serem utilizados posteriormente em uma pesquisa com os mesmos parâmetros
    Set objUltFornProdFFPCFechado = objFornecedorProdutoFF

    ObterRel_Info_Ult_PC_Fechado = SUCESSO

    Exit Function

Erro_ObterRel_Info_Ult_PC_Fechado:

    ObterRel_Info_Ult_PC_Fechado = gErr

    Select Case gErr

        Case 79999

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153159)

    End Select

    Exit Function

End Function

Public Function ObterRel_Info_Ult_Cotacao(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Chama a função que lê no BD os dados do último pedido de compra recebido para o Produto / Fornecedor / FilialFornecedor
'Parâmetros (na ordem que devem ser passados): iFilialEmpresa, lFornecedor, iFilialForn, sProduto

Dim lErro

On Error GoTo Erro_ObterRel_Info_Ult_Cotacao

    'Se pelo menos um dos parâmetros passados pelo Forprint se alterou
    If objFornecedorProdutoFF.iFilialEmpresa <> objUltFornProdFFCotacao.iFilialEmpresa Or objFornecedorProdutoFF.lFornecedor <> objUltFornProdFFCotacao.lFornecedor Or objFornecedorProdutoFF.iFilialForn <> objUltFornProdFFCotacao.iFilialForn Or objFornecedorProdutoFF.sProduto <> objUltFornProdFFCotacao.sProduto Then
        
        'Lê a quantidade do produto em pedidos de compra para o fornecedor selecionado
        lErro = UltimaCotacao_Le_FornecedorProduto2(objFornecedorProdutoFF)
        If lErro <> SUCESSO Then gError 90518

    'Se todos os parâmetros são iguais
    Else

        'Pega os últimos dados encontrados com esses parâmetros e retorna à função chamadora
        Set objFornecedorProdutoFF = objUltFornProdFFCotacao

    End If

    'Guarda os dados encontrados para serem utilizados posteriormente em uma pesquisa com os mesmos parâmetros
    Set objUltFornProdFFCotacao = objFornecedorProdutoFF

    ObterRel_Info_Ult_Cotacao = SUCESSO

    Exit Function

Erro_ObterRel_Info_Ult_Cotacao:

    ObterRel_Info_Ult_Cotacao = gErr

    Select Case gErr

        Case 90518

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153160)

    End Select

    Exit Function

End Function

'*** Fernando, favor subir as funções abaixo para o COMSelect ***

Function UltimoItemPedCompraFechado_Le2(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'retorna os dados do ultimo (mais recente nota fiscal) item de pedido de compra da Filial em questão com status fechado para o produto/Fornecedor/FilialForn passado como parametro.
'essa função é semelhante à função UltimoItemPedCompraFechado_Le, entretanto as quantidades retornadas por esta estão convertidas para a U.M. de compras

Dim lComando As Long
Dim sSelect As String
Dim lErro As Long
Dim dValorDesconto As Double
Dim dPrecoUnitario As Double
Dim dFator As Double
Dim dQuantidade As Double
Dim tFornecedorProdutoFF As typeFornecedorProdutoFF
Dim objProduto As New ClassProduto

On Error GoTo Erro_UltimoItemPedCompraFechado_Le2

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90506

    'Guarda em objProduto o parâmetro que será passado para a função Produto_Le
    objProduto.sCodigo = objFornecedorProdutoFF.sProduto
    
    'Lê os dados do produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 90505
    
    'Se não encontrou o produto => erro
    If lErro = 28030 Then gError 90511
    
    'Monta o select que será utlizado para leitura dos dados no BD
    sSelect = "SELECT NFiscal.DataEntrada, PedidoCompraTodos.Data, ItensNFiscal.Quantidade, ItensPedCompraTodos.Quantidade, ItensPedCompraTodos.UM, ItensNFiscal.UnidadeMed FROM PedidoCompraTodos, ItensPedCompraTodos, ItensNFiscal, NFiscal, ItemNFItemPC  WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItemNFItemPC.ItemNFiscal = ItensNFiscal.NumIntDoc AND ItemNFItemPC.ItemPedCompra = ItensPedCompraTodos.NumIntDoc AND ItensPedCompraTodos.PedCompra = PedidoCompraTodos.NumIntDoc AND ItensPedCompraTodos.Status = 1 AND NFiscal.Fornecedor = ? AND NFiscal.FilialForn = ? AND ItensNFiscal.Produto = ? AND PedidoCompraTodos.FilialEmpresa = ? ORDER BY NFiscal.DataEntrada DESC"
        
    With tFornecedorProdutoFF

        'Inicializa as strings que serão utilizadas
        .sUMQuantPedida = String(STRING_UM_SIGLA, 0)
        .sUMQuantRecebida = String(STRING_UM_SIGLA, 0)

        'Busca em NFiscal, ItensNFiscal os registros ligados ao item passado
        lErro = Comando_Executar(lComando, sSelect, .dtDataReceb, .dtDataPedido, .dQuantRecebida, .dQuantPedida, .sUMQuantPedida, .sUMQuantRecebida, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 90508

    End With

    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90509

    If lErro = AD_SQL_SUCESSO Then

        With tFornecedorProdutoFF
    
            'Se a U.M. da quantidade pedida foi preenchida
            If Len(Trim(.sUMQuantPedida)) > 0 Then
            
                'Converte a unidade de medida da quantidade pedida
                lErro = CF("UM_Conversao", objProduto.iClasseUM, .sUMQuantPedida, objProduto.sSiglaUMCompra, dFator)
                If lErro <> SUCESSO Then gError 90507

                'Quantidade Pedida
                
                'Guarda no obj a quantidade pedida convertida para a U.M. padrão de compras
                objFornecedorProdutoFF.dQuantPedida = .dQuantPedida * dFator
            
            End If
            
            If Len(Trim(.sUMQuantRecebida)) > 0 Then
            
                'Converte a unidade de medida da quantidade recebida
                lErro = CF("UM_Conversao", objProduto.iClasseUM, .sUMQuantRecebida, objProduto.sSiglaUMCompra, dFator)
                If lErro <> SUCESSO Then gError 90507
            
                'Quantidade Recebida
            
                'Guarda no obj a quantida recebida convertida para a U.M. padrão de compras
                objFornecedorProdutoFF.dQuantRecebida = .dQuantRecebida * dFator
            
            End If
            
            'Guarda no obj o restante dos dados lidos
            objFornecedorProdutoFF.dtDataPedido = .dtDataPedido
            objFornecedorProdutoFF.dtDataReceb = .dtDataReceb
            
        End With
    
    Else
    
        objFornecedorProdutoFF.dtDataPedido = DATA_NULA
        objFornecedorProdutoFF.dtDataReceb = DATA_NULA
        

    End If

    'Fecha o comndo
    Call Comando_Fechar(lComando)

    UltimoItemPedCompraFechado_Le2 = SUCESSO

    Exit Function

Erro_UltimoItemPedCompraFechado_Le2:

    UltimoItemPedCompraFechado_Le2 = gErr

    Select Case gErr

        Case 90505, 90507
        
        Case 90506
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90511
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 90508, 90509
            Call Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_ULTIMO_PEDCOMPRA_FECHADO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153161)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function UltimaCotacao_Le_FornecedorProduto2(objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'retorna os dados da ultima cotacao (maior data de pedido de cotação) do produto/Fornecedor/FilialForn feita pela FilialEmpresa passado como parametro.
'essa função é semelhante à função UltimaCotacao_Le_FornecedorProduto, entretanto as quantidades retornadas por esta estão convertidas para a U.M. de compras

Dim dFator As Double
Dim lComando As Long
Dim sSelect As String
Dim lErro As Long
Dim tFornecedorProdutoFF As typeFornecedorProdutoFF
Dim objProduto As New ClassProduto

On Error GoTo Erro_UltimaCotacao_Le_FornecedorProduto2

    'ABre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90512

    objProduto.sCodigo = objFornecedorProdutoFF.sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 90513
    
    'Se não encontrou o produto => erro
    If lErro = 28030 Then gError 90514
    
    'Monta o select que será utilizado para leitura do BD
    sSelect = "SELECT ItensCotacao.CondPagto, ItensCotacao.PrecoUnitario, PedidoCotacao.Data, PedidoCotacao.TipoFrete, CotacaoProduto.Quantidade, CotacaoProduto.UM, CondicoesPagto.DescReduzida FROM PedidoCotacao, ItemPedCotacao, CotacaoProduto, ItensCotacao, CondicoesPagto Where PedidoCotacao.NumIntDoc = ItemPedCotacao.PedidoCotacao AND ItemPedCotacao.CotacaoProduto = CotacaoProduto.NumIntDoc AND ItemPedCotacao.NumIntDoc = ItensCotacao.ItemPedCotacao AND CondicoesPagto.Codigo = ItensCotacao.CondPagto AND PedidoCotacao.Fornecedor = ? And PedidoCotacao.Filial = ? AND CotacaoProduto.Produto = ? AND PedidoCotacao.FilialEmpresa = ? ORDER BY PedidoCotacao.Data DESC"
        
    With tFornecedorProdutoFF

        .sCondPagto = String(STRING_CONDICAO_PAGTO_DESCRICAO_REDUZIDA, 0)
        .sUMQuantUltimaCotacao = String(STRING_UM_SIGLA, 0)

        'Busca em CotacaoItemConcorrencia os registros ligados ao item passado
        lErro = Comando_Executar(lComando, sSelect, .iCondPagto, .dUltimaCotacao, .dtDataUltimaCotacao, .iTipoFreteUltimaCotacao, .dQuantUltimaCotacao, .sUMQuantUltimaCotacao, .sCondPagto, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 90515

    End With

    'Busca o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90516

    If lErro = AD_SQL_SUCESSO Then

        With tFornecedorProdutoFF
    
            'Quantidade da última cotação
            'Se a U.M. da quantidade da ult.cotação foi preenchida => erro
            If Len(Trim(.sUMQuantUltimaCotacao)) > 0 Then
            
                'Converte a unidade de medida da quantidade recebida
                lErro = CF("UM_Conversao", objProduto.iClasseUM, .sUMQuantUltimaCotacao, objProduto.sSiglaUMCompra, dFator)
                If lErro <> SUCESSO Then gError 90517
            
                'Guarda no obj a quantida recebida convertida para a U.M. padrão de compras
                objFornecedorProdutoFF.dQuantUltimaCotacao = .dQuantUltimaCotacao * dFator
            End If
            
            'Guarda no obj os outros dados lidos do BD
            objFornecedorProdutoFF.iCondPagto = .iCondPagto
            objFornecedorProdutoFF.dUltimaCotacao = .dUltimaCotacao
            objFornecedorProdutoFF.dtDataUltimaCotacao = .dtDataUltimaCotacao
            objFornecedorProdutoFF.iTipoFreteUltimaCotacao = .iTipoFreteUltimaCotacao
            objFornecedorProdutoFF.sCondPagto = .sCondPagto
            objFornecedorProdutoFF.sUMQuantUltimaCotacao = .sUMQuantUltimaCotacao
            
        End With
    
    Else
    
        objFornecedorProdutoFF.dtDataUltimaCotacao = DATA_NULA
        
    End If

    'Fecha o comndo
    Call Comando_Fechar(lComando)

    UltimaCotacao_Le_FornecedorProduto2 = SUCESSO

    Exit Function

Erro_UltimaCotacao_Le_FornecedorProduto2:

    UltimaCotacao_Le_FornecedorProduto2 = gErr

    Select Case gErr

        Case 90513, 90517
        
        Case 90512
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90514
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            
        Case 90515, 90516
            Call Rotina_Erro(vbOKOnly, "ERRO_PESQUISA_ULTIMA_COTACAO", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153162)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'*** Fernando, favor transportar as funções acima para o COMSelect
