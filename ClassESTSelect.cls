VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const STRING_UM_PV_SIGLA = 5

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

''Function GeracaoOP_SelecionarItens(objGeracaoOP As ClassGeracaoOP) As Long
'''pesquisa no bd os itens candidatos preenchendo objGeracaoOP.colItens a partir das outras informacoes em
''Dim lErro As Long, lComando As Long, sSelect As String, iStatusNaoAtendido As Integer
'''buffers para receber registros selecionados
''Dim iItensPedidoDeVendaItem As Integer, sItensPedidoDeVendaProduto As String, dtItensPedidoDeVendaDataEntrega As Date
''Dim iPedidosDeVendaFilialEmpresa As Integer, lPedidosDeVendaCodigo As Long, iPedidosDeVendaFilial As Integer, dtPedidosDeVendaDataEmissao As Date
''Dim sClientesNomeReduzido As String
''Dim sProdutosDescricao As String
''On Error GoTo Erro_GeracaoOP_SelecionarItens
''
''    lComando = Comando_Abrir()
''    If lComando = 0 Then Error 27434
''
''    iStatusNaoAtendido = STATUS_NAO_ATENDIDO
''
''    Call GeracaoOP_SelecionarItens1(objGeracaoOP, sSelect)
''
''    sItensPedidoDeVendaProduto = String(STRING_PRODUTO, 0)
''    sClientesNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
''    sProdutosDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
''
''    'executa a preparacao da parte fixa do SELECT
''    lErro = GeracaoOP_SelecionarItens2(lComando, sSelect, iItensPedidoDeVendaItem, _
''    sItensPedidoDeVendaProduto, dtItensPedidoDeVendaDataEntrega, _
''    iPedidosDeVendaFilialEmpresa, lPedidosDeVendaCodigo, iPedidosDeVendaFilial, _
''    dtPedidosDeVendaDataEmissao, sClientesNomeReduzido, sProdutosDescricao, iStatusNaoAtendido)
''    If lErro <> SUCESSO Then Error 27435
''
''    'complementa a passagem dos parametros que variam de acordo com a selecao do usuario
''    'e executa o SELECT
''    lErro = GeracaoOP_SelecionarItens3(lComando, objGeracaoOP)
''    If lErro <> SUCESSO Then Error 27436
''
''    'processa todos os registros retornados pelo SELECT executado acima
''    lErro = GeracaoOP_SelecionarItens4(lComando, objGeracaoOP, iItensPedidoDeVendaItem, _
''        sItensPedidoDeVendaProduto, dtItensPedidoDeVendaDataEntrega, _
''        iPedidosDeVendaFilialEmpresa, lPedidosDeVendaCodigo, iPedidosDeVendaFilial, _
''        dtPedidosDeVendaDataEmissao, sClientesNomeReduzido, sProdutosDescricao)
''    If lErro <> SUCESSO Then Error 27437
''
''    lErro = Comando_Fechar(lComando)
''
''    GeracaoOP_SelecionarItens = SUCESSO
''
''    Exit Function
''
''Erro_GeracaoOP_SelecionarItens:
''
''    GeracaoOP_SelecionarItens = Err
''
''    Select Case Err
''
''        Case 27435, 27436, 27437
''
''        Case 27434
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149619)
''
''    End Select
''
''    lErro = Comando_Fechar(lComando)
''
''    Exit Function
''
''End Function
''
''Sub GeracaoOP_SelecionarItens1(objGeracaoOP As ClassGeracaoOP, sSelect As String)
'''auxiliar a GeracaoOP_SelecionarItens
'''monta o SELECT para obtencao dos itens dinamicamente.
''Dim sFrom As String, sWhere As String, sFields As String, sOrderBy As String
''
''    sFields = "ItensPedidoDeVenda.Item, ItensPedidoDeVenda.Produto, ItensPedidoDeVenda.DataEntrega"
''    sFields = sFields & ", PedidosDeVenda.FilialEmpresa, PedidosDeVenda.Codigo, PedidosDeVenda.Filial, PedidosDeVenda.DataEmissao"
''    sFields = sFields & ", Clientes.NomeReduzido"
''    sFields = sFields & ", Produtos.Descricao"
''
''    sFrom = " FROM PedidosDeVenda, ItensPedidoDeVenda, Produtos, Clientes"
''    sWhere = " WHERE PedidosDeVenda.FilialEmpresa = ItensPedidoDeVenda.FilialEmpresa AND PedidosDeVenda.Codigo = ItensPedidoDeVenda.CodPedido AND PedidosDeVenda.Cliente = Clientes.Codigo AND ItensPedidoDeVenda.Produto = Produtos.Codigo AND ItensPedidoDeVenda.Status = ?"
''
''    'Se entrar como empresa toda pode pegar itens de todas as filiais, senao pegar itens somente da filial corrente
''    If giFilialEmpresa <> EMPRESA_TODA Then
''        sWhere = sWhere & " AND PedidosDeVenda.FilialEmpresa = ?"
''    End If
''
''    If objGeracaoOP.iTodos <> 0 Then
''
''        'Se pedido de venda inicial preenchido
''        If (objGeracaoOP.lPVDe <> 0) Then
''           sWhere = sWhere & " AND PedidosDeVenda.Codigo >= ?"
''        End If
''
''        'Se pedido de venda final preenchido
''        If (objGeracaoOP.lPVAte <> 0) Then
''           sWhere = sWhere & " AND PedidosDeVenda.Codigo <= ?"
''        End If
''
''        'Se cliente do pedido de venda inicial preenchido
''        If (objGeracaoOP.lClientesDe <> 0) Then
''           sWhere = sWhere & " AND PedidosDeVenda.Cliente >= ?"
''        End If
''
''        'Se cliente do pedido de venda final preenchido
''        If (objGeracaoOP.lClientesAte <> 0) Then
''           sWhere = sWhere & " AND PedidosDeVenda.Cliente <= ?"
''        End If
''
''        'Se data de entrega inicial do item preenchida
''        If (objGeracaoOP.dtEntregaDe <> DATA_NULA) Then
''           sWhere = sWhere & " AND ItensPedidoDeVenda.DataEntrega >= ?"
''        End If
''
''        'Se data de entrega final do item preenchida
''        If (objGeracaoOP.dtEntregaAte <> DATA_NULA) Then
''           sWhere = sWhere & " AND ItensPedidoDeVenda.DataEntrega <= ?"
''        End If
''
''        'se produto inicial preenchido
''        If (objGeracaoOP.sProdDe <> "") Then
''           sWhere = sWhere & " AND ItensPedidoDeVenda.Produto <= ?"
''        End If
''
''        'se produto final preenchido
''        If (objGeracaoOP.sProdAte <> "") Then
''           sWhere = sWhere & " AND ItensPedidoDeVenda.Produto <= ?"
''        End If
''
''        'se NAO deve incluir item de pedido de venda com Ordem de Producao associada
''        If (objGeracaoOP.iIncluiPVcomOP = 0) Then
''           sWhere = sWhere & " AND ItensPedidoDeVenda.QuantOP = 0"
''        End If
''
''    End If
''
''    sOrderBy = " ORDER BY " & objGeracaoOP.sOrdenacao
''
''    sSelect = "SELECT " & sFields & sFrom & sWhere & sOrderBy
''
''End Sub
''
''Function GeracaoOP_SelecionarItens2(lComando As Long, sSelect As String, _
''    vItensPedidoDeVendaItem As Variant, vItensPedidoDeVendaProduto As Variant, _
''    vItensPedidoDeVendaDataEntrega As Variant, vPedidosDeVendaFilialEmpresa As Variant, _
''    vPedidosDeVendaCodigo As Variant, vPedidosDeVendaFilial As Variant, _
''    vPedidosDeVendaDataEmissao As Variant, vClientesNomeReduzido As Variant, _
''    vProdutosDescricao As Variant, vStatusNaoAtendido As Variant) As Integer
'''isola a preparacao da parte fixa do SELECT, serve apenas para diminuir o tamanho de GeracaoOP_SelecionarItens
''Dim ret As Integer, lErro As Long
''On Error GoTo Erro_GeracaoOP_SelecionarItens2
''
''    ret = Comando_PrepararInt(lComando, sSelect)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27409
''
''    ret = Comando_BindVarInt(lComando, vItensPedidoDeVendaItem)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27410
''
''    ret = Comando_BindVarInt(lComando, vItensPedidoDeVendaProduto)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27411
''
''    ret = Comando_BindVarInt(lComando, vItensPedidoDeVendaDataEntrega)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27412
''
''    ret = Comando_BindVarInt(lComando, vPedidosDeVendaFilialEmpresa)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27413
''
''    ret = Comando_BindVarInt(lComando, vPedidosDeVendaCodigo)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27414
''
''    ret = Comando_BindVarInt(lComando, vPedidosDeVendaFilial)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27415
''
''    ret = Comando_BindVarInt(lComando, vPedidosDeVendaDataEmissao)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27416
''
''    ret = Comando_BindVarInt(lComando, vClientesNomeReduzido)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27417
''
''    ret = Comando_BindVarInt(lComando, vProdutosDescricao)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27418
''
''    ret = Comando_BindVarInt(lComando, vStatusNaoAtendido)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27419
''
''    GeracaoOP_SelecionarItens2 = SUCESSO
''
''    Exit Function
''
''Erro_GeracaoOP_SelecionarItens2:
''
''    GeracaoOP_SelecionarItens2 = Err
''
''    Select Case Err
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149620)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function GeracaoOP_SelecionarItens3(lComando As Long, objGeracaoOP As ClassGeracaoOP) As Long
'''auxiliar a GeracaoOP_SelecionarItens
'''complementa a passagem dos parametros que variam de acordo com a selecao do usuario e executa o SELECT p/obtencao dos itens
''Dim lErro As Long, ret As Integer, vFilialEmpresa As Variant, vPVDe As Variant, vPVAte As Variant
''Dim vClientesDe As Variant, vClientesAte As Variant, vEntregaDe As Variant, vEntregaAte As Variant
''Dim vProdDe As Variant, vProdAte As Variant, vIncluiPVcomOP As Variant
''On Error GoTo Erro_GeracaoOP_SelecionarItens3
''
''    'Se entrar como empresa toda pode pegar itens de todas as filiais, senao pegar itens somente da filial corrente
''    If giFilialEmpresa <> EMPRESA_TODA Then
''        vFilialEmpresa = giFilialEmpresa
''        ret = Comando_BindVarInt(lComando, vFilialEmpresa)
''        If (ret <> AD_SQL_SUCESSO) Then Error 27420
''    End If
''
''    If objGeracaoOP.iTodos <> 0 Then
''
''        'Se pedido de venda inicial preenchido
''        If (objGeracaoOP.lPVDe <> 0) Then
''            vPVDe = objGeracaoOP.lPVDe
''            ret = Comando_BindVarInt(lComando, vPVDe)
''            If (ret <> AD_SQL_SUCESSO) Then Error 27421
''        End If
''
''        'Se pedido de venda final preenchido
''        If (objGeracaoOP.lPVAte <> 0) Then
''            vPVAte = objGeracaoOP.lPVAte
''            ret = Comando_BindVarInt(lComando, vPVAte)
''            If (ret <> AD_SQL_SUCESSO) Then Error 27422
''        End If
''
''        'Se cliente do pedido de venda inicial preenchido
''        If (objGeracaoOP.lClientesDe <> 0) Then
''            vClientesDe = objGeracaoOP.lClientesDe
''            ret = Comando_BindVarInt(lComando, vClientesDe)
''            If (ret <> AD_SQL_SUCESSO) Then Error 27423
''        End If
''
''        'Se cliente do pedido de venda final preenchido
''        If (objGeracaoOP.lClientesAte <> 0) Then
''            vClientesAte = objGeracaoOP.lClientesAte
''            ret = Comando_BindVarInt(lComando, vClientesAte)
''            If (ret <> AD_SQL_SUCESSO) Then Error 27424
''        End If
''
''        'Se data de entrega inicial do item preenchida
''        If (objGeracaoOP.dtEntregaDe <> DATA_NULA) Then
''            vEntregaDe = objGeracaoOP.dtEntregaDe
''            ret = Comando_BindVarInt(lComando, vEntregaDe)
''            If (ret <> AD_SQL_SUCESSO) Then Error 27425
''        End If
''
''        'Se data de entrega final do item preenchida
''        If (objGeracaoOP.dtEntregaAte <> DATA_NULA) Then
''            vEntregaAte = objGeracaoOP.dtEntregaAte
''            ret = Comando_BindVarInt(lComando, vEntregaAte)
''            If (ret <> AD_SQL_SUCESSO) Then Error 27426
''        End If
''
''        'se produto inicial preenchido
''        If (objGeracaoOP.sProdDe <> "") Then
''            vProdDe = objGeracaoOP.sProdDe
''            ret = Comando_BindVarInt(lComando, vProdDe)
''            If (ret <> AD_SQL_SUCESSO) Then Error 27427
''        End If
''
''        'se produto final preenchido
''        If (objGeracaoOP.sProdAte <> "") Then
''            vProdAte = objGeracaoOP.sProdAte
''            ret = Comando_BindVarInt(lComando, vProdAte)
''            If (ret <> AD_SQL_SUCESSO) Then Error 27428
''        End If
''
''        'se NAO deve incluir item de pedido de venda com Ordem de Producao associada
''        If (objGeracaoOP.iIncluiPVcomOP = 0) Then
''            vIncluiPVcomOP = objGeracaoOP.iIncluiPVcomOP
''            ret = Comando_BindVarInt(lComando, vIncluiPVcomOP)
''            If (ret <> AD_SQL_SUCESSO) Then Error 27429
''        End If
''
''    End If
''
''    ret = Comando_ExecutarInt(lComando)
''    If (ret <> AD_SQL_SUCESSO) Then Error 27430
''
''    GeracaoOP_SelecionarItens3 = SUCESSO
''
''    Exit Function
''
''Erro_GeracaoOP_SelecionarItens3:
''
''    GeracaoOP_SelecionarItens3 = Err
''
''    Select Case Err
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149621)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function GeracaoOP_SelecionarItens4(lComando As Long, objGeracaoOP As ClassGeracaoOP, _
''        iItensPedidoDeVendaItem As Integer, sItensPedidoDeVendaProduto As String, _
''        dtItensPedidoDeVendaDataEntrega As Date, iPedidosDeVendaFilialEmpresa As Integer, _
''        lPedidosDeVendaCodigo As Long, iPedidosDeVendaFilial As Integer, _
''        dtPedidosDeVendaDataEmissao As Date, sClientesNomeReduzido As String, _
''        sProdutosDescricao As String) As Integer
'''auxiliar a GeracaoOP_SelecionarItens
'''inclui os itens lidos na colecao
''Dim lErro As Long, objGeracaoOpItem As ClassGeracaoOpItem, objItemOP As ClassItemOP
''On Error GoTo Erro_GeracaoOP_SelecionarItens4
''
''    lErro = Comando_BuscarPrimeiro(lComando)
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27431
''
''    If lErro = AD_SQL_SEM_DADOS Then Error 27432
''
''    Do While lErro = AD_SQL_SUCESSO
''
''        Set objGeracaoOpItem = New ClassGeracaoOpItem
''
''        objGeracaoOpItem.dtEmissaoPV = dtPedidosDeVendaDataEmissao
''        objGeracaoOpItem.dtEntregaItemPV = dtItensPedidoDeVendaDataEntrega
''        objGeracaoOpItem.iFilialCliente = iPedidosDeVendaFilial
''
''        Set objItemOP = New ClassItemOP
''
''        objItemOP.iFilialPedido = iPedidosDeVendaFilialEmpresa
''        objItemOP.lCodPedido = lPedidosDeVendaCodigo
''        objItemOP.iItemPedido = iItensPedidoDeVendaItem
''        objItemOP.sProduto = sItensPedidoDeVendaProduto
''
''        Set objGeracaoOpItem.objItemOP = objItemOP
''
''        objGeracaoOpItem.sClienteReduzido = sClientesNomeReduzido
''        objGeracaoOpItem.sProdutoDescricao = sProdutosDescricao
''
''        objGeracaoOP.colItens.Add objGeracaoOpItem
''
''        lErro = Comando_BuscarProximo(lComando)
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27433
''
''    Loop
''
''    GeracaoOP_SelecionarItens4 = SUCESSO
''
''    Exit Function
''
''Erro_GeracaoOP_SelecionarItens4:
''
''    GeracaoOP_SelecionarItens4 = Err
''
''    Select Case Err
''
''        Case 27431, 27433
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPV_GERACAO_OP", Err)
''
''        Case 27432
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_ITENSPV_GERACAO_OP", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149622)
''
''    End Select
''
''    Exit Function
''
''End Function
''

Function ESTConfig_Le_Configs(ByVal colConfigs As ColESTConfig) As Long
'Recupera da tabela ESTConfig uma coleção de registros

Dim lErro As Long, lComando As Long
Dim objESTConfig As ClassESTConfig

On Error GoTo Erro_ESTConfig_Le_Configs

    'Abrir comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33663

    'Para cada elemento da coleção
    For Each objESTConfig In colConfigs

        lErro = ESTConfig_Le_Batch(objESTConfig, lComando)
        If lErro <> SUCESSO And lErro <> 33668 Then Error 33664

        'Não foi encontrado registro em ESTConfig
        If lErro = 33668 Then Error 33665

    Next

    'Fechar comando
    Call Comando_Fechar(lComando)

    ESTConfig_Le_Configs = SUCESSO

    Exit Function

Erro_ESTConfig_Le_Configs:

    ESTConfig_Le_Configs = Err

    Select Case Err

        Case 33663
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33664

        Case 33665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTCONFIG_INEXISTENTE", Err, objESTConfig.sCodigo, objESTConfig.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149623)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function ESTConfig_Le_Batch(ByVal objESTConfig As ClassESTConfig, ByVal lComando As Long) As Long
'Lê registro em ESTConfig.

Dim lErro As Long, tESTConfig As typeESTConfig

On Error GoTo Erro_ESTConfig_Le_Batch

    With tESTConfig
        'Inicializar buffers
        .sDescricao = String(STRING_ESTCONFIG_DESCRICAO, 0)
        .sConteudo = String(STRING_ESTCONFIG_CONTEUDO, 0)

        'Ler registo
        lErro = Comando_Executar(lComando, "SELECT Descricao, Tipo, Conteudo FROM ESTConfig WHERE Codigo = ? AND FilialEmpresa = ?", .sDescricao, .iTipo, .sConteudo, objESTConfig.sCodigo, objESTConfig.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 33666

    End With

    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33667

    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then Error 33668

    objESTConfig.sDescricao = tESTConfig.sDescricao
    objESTConfig.iTipo = tESTConfig.iTipo
    objESTConfig.sConteudo = tESTConfig.sConteudo

    ESTConfig_Le_Batch = SUCESSO

    Exit Function

Erro_ESTConfig_Le_Batch:

    ESTConfig_Le_Batch = Err

    Select Case Err

        Case 33666, 33667
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTCONFIG2", Err, objESTConfig.sCodigo, objESTConfig.iFilialEmpresa)

        Case 33668
            'Chave nao encontrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149624)

    End Select

    Exit Function

End Function

Function ItensNFiscalEntDev_Lock_Gravacao(ByVal objNFiscal As ClassNFiscal) As Long
'Faz a validação e os "lock's" dos dados ligados aos ítens da Nota Fiscal

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim objItemNF As ClassItemNF
Dim objItemNFOrig As New ClassItemNF
Dim lNumIntDoc As Long
Dim dQuantidade As Double
Dim sProduto As String
Dim sUnidadeMed As String
Dim dQuantDevolvida As Double
Dim dFator As Double
Dim objProduto As New ClassProduto
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_ItensNFiscalEntDev_Lock_Gravacao

    'Abre os Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 35400

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 35401

    'Se Nota for interna, lock shared na série
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 42694
    If lErro <> SUCESSO Then Error 42695
    
    'Para cada Ítem existente na coleção de Ítens
    For Each objItemNF In objNFiscal.colItensNF
        
        dQuantDevolvida = 0
    
        objProduto.sCodigo = objItemNF.sProduto
        lErro = CF("Produto_Lock", objProduto)
        If lErro <> SUCESSO Then Error 35924
        
        objItemNF.iClasseUM = objProduto.iClasseUM
        objItemNF.iControleEstoque = objProduto.iControleEstoque
    
        'Se o Número Interno do Documento Original estiver preenchido
        If objItemNF.lNumIntDocOrig > 0 Then

            sProduto = String(STRING_PRODUTO, 0)
            sUnidadeMed = String(STRING_UM_SIGLA, 0)
            
            'Lê em ItensNFiscal a quantidade, a Unidade de Medida e o Produto do Ítem com o Número Ínterno = NumIntDocOrig
            lErro = Comando_ExecutarLockado(lComando, "SELECT Quantidade, UnidadeMed, Produto FROM ItensNFiscal WHERE NumIntDoc = ? ", dQuantidade, sUnidadeMed, sProduto, objItemNF.lNumIntDocOrig)
            If lErro <> AD_SQL_SUCESSO Then Error 35402

            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35403
            
            If lErro = AD_SQL_SEM_DADOS Then Error 35406 'Não encontrou
            
            'Faz o "lock" do Ítem
            lErro = Comando_LockShared(lComando)
            If lErro <> AD_SQL_SUCESSO Then Error 35407

            objItemNFOrig.dQuantidade = dQuantidade
            objItemNFOrig.sUnidadeMed = sUnidadeMed

            sUnidadeMed = String(STRING_UM_SIGLA, 0)
            
            If objNFiscal.iFilialEmpresa = 0 Then objNFiscal.iFilialEmpresa = giFilialEmpresa
            
            'Lê a Quantidade, Produto e Unidade de Medida de outros ítens já baixados
            lErro = Comando_Executar(lComando1, "SELECT ItensNFiscal.Quantidade, ItensNFiscal.UnidadeMed FROM ItensNFiscal, NFiscal, TiposDocInfo WHERE ItensNFiscal.NumIntNF = NFiscal.NumIntDoc AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.FilialEmpresa = ? AND TiposDocInfo.NomeTelaNFiscal = ? AND TiposDocInfo.Complementar = 0 AND ItensNFiscal.NumIntDocOrig = ? AND ItensNFiscal.Status = ?", dQuantidade, sUnidadeMed, objNFiscal.iFilialEmpresa, "NFiscalEntDev", objItemNF.lNumIntDocOrig, STATUS_BAIXADO)
            If lErro <> AD_SQL_SUCESSO Then Error 35408

            lErro = Comando_BuscarPrimeiro(lComando1)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35409

            Do While lErro <> AD_SQL_SEM_DADOS

                lErro = CF("UM_Conversao", objItemNF.iClasseUM, sUnidadeMed, objItemNFOrig.sUnidadeMed, dFator)
                If lErro <> SUCESSO Then Error 35410

                'Acumula a quantidade com o valor já convertido
                dQuantidade = dQuantidade * dFator

                dQuantDevolvida = dQuantDevolvida + dQuantidade

                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35411

            Loop

            lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemNF.sUnidadeMed, objItemNFOrig.sUnidadeMed, dFator)
            If lErro <> SUCESSO Then Error 35412

            dQuantidade = objItemNF.dQuantidade * dFator
            dQuantDevolvida = dQuantDevolvida + dQuantidade

            'Verifica se a Quantidade devolvida é maior que a Quantidade do Ítem da NFOriginal
            If objTipoDocInfo.iComplementar = 0 And dQuantDevolvida > objItemNFOrig.dQuantidade Then Error 35413
        End If
    Next

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ItensNFiscalEntDev_Lock_Gravacao = SUCESSO

    Exit Function

Erro_ItensNFiscalEntDev_Lock_Gravacao:

    ItensNFiscalEntDev_Lock_Gravacao = Err

    Select Case Err

        Case 35400, 35401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 35402, 35403, 35408, 35409, 35411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL1", Err)

        Case 35404, 35405
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCALBAIXADAS1", Err)

        Case 35406
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEM_NFORIGINAL_NAO_CADASTRADO2", Err, objItemNF.lNumIntDocOrig)

        Case 35407
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITEM_NFORIGINAL", Err, objItemNF.lNumIntDocOrig)

        Case 35410, 35412, 35924, 42694

        Case 42695
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)

        Case 35413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANT_DEVOLVIDA_A_MAIOR", Err, objItemNF.iItem, dQuantDevolvida, objItemNFOrig.sUnidadeMed, objItemNFOrig.dQuantidade, objItemNFOrig.sUnidadeMed)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149625)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function NFiscal_Le_NumFornCli(ByVal objNFiscal As ClassNFiscal) As Long
'Le a Nota Fiscal usando Serie, NumNotaFiscal, Cliente, FilialCli, Fornecedor, FilialForn, TipoNFiscal

Dim lErro As Long
Dim lComando As Long
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscal_Le_NumFornCli

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 35276

    With tNFiscal
        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)

        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_BUFFER_MAX_TEXTO, 0)
        '.lVolumeMarca = String(STRING_BUFFER_MAX_TEXTO, 0)

        If objNFiscal.iFilialEmpresa = 0 Then objNFiscal.iFilialEmpresa = giFilialEmpresa
        
        'Busca a Nota Fiscal
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, FilialEntrega,DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, NumRecebimento, CodTranspRedesp, DetPagFrete FROM NFiscal WHERE Status <> ? AND FilialEmpresa = ? AND Serie = ? AND NumNotaFiscal = ? AND Cliente = ? AND FilialCli = ? AND Fornecedor = ? AND FilialForn = ? AND TipoNFiscal = ? ORDER BY DataEntrada DESC", _
            .lNumIntDoc, .iStatus, .iFilialEntrega, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .lNumRecebimento, .iCodTranspRedesp, .iDetPagFrete, STATUS_EXCLUIDO, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iTipoNFiscal)
        If lErro <> AD_SQL_SUCESSO Then Error 35277

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35278
        If lErro = AD_SQL_SEM_DADOS Then Error 35279 'Não encontrou

        'Preenche objNFiscal com os dados lidos
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dtDataEmissao = .dtDataEmissao
        objNFiscal.dtDataEntrada = .dtDataEntrada
        objNFiscal.dtDataSaida = .dtDataSaida
        objNFiscal.dtDataVencimento = .dtDataVencimento
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.dValorFrete = .dValorFrete
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorProdutos = .dValorProdutos
        objNFiscal.dValorSeguro = .dValorSeguro
        objNFiscal.dValorTotal = .dValorTotal
        objNFiscal.iCanal = .iCanal
        objNFiscal.iClasseDocCPR = .iClasseDocCPR
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscal.iDetPagFrete = .iDetPagFrete
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        objNFiscal.iFilialEntrega = .iFilialEntrega
        objNFiscal.iStatus = .iStatus
        objNFiscal.iTabelaPreco = .iTabelaPreco
        objNFiscal.lVolumeQuant = .lVolumeQuant
        objNFiscal.lNumIntDoc = .lNumIntDoc
        objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscal.lNumIntTrib = .lNumIntTrib
        objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.sPlaca = .sPlaca
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.lVolumeEspecie = .lVolumeEspecie
        objNFiscal.lVolumeMarca = .lVolumeMarca
        objNFiscal.lNumRecebimento = .lNumRecebimento

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    NFiscal_Le_NumFornCli = SUCESSO

    Exit Function

Erro_NFiscal_Le_NumFornCli:

    NFiscal_Le_NumFornCli = Err

    Select Case Err

        Case 35276
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 35277, 35278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL2", Err)

        Case 35279

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149626)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemNFiscalSaida_Le_NumNFItem(ByVal objItemNF As ClassItemNF) As Long
'Lê o Ítem de Uma Nota Fiscal de Saida através do Número da Nota Fiscal, Série e Numero do Ítem passados

Dim lErro As Long
Dim lComando As Long
Dim tItemNF As typeItemNF

On Error GoTo Erro_ItemNFiscalSaida_Le_NumNFItem

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 35365

    With tItemNF
        .sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sUnidadeMed = String(STRING_UM_SIGLA, 0)

        'Busca o Ítem
        lErro = Comando_Executar(lComando, "SELECT ItensNFiscal.NumIntNF, ItensNFiscal.Status, ItensNFiscal.Produto, ItensNFiscal.UnidadeMed, ItensNFiscal.Quantidade, ItensNFiscal.Almoxarifado, ItensNFiscal.PrecoUnitario, ItensNFiscal.PercDesc, ItensNFiscal.ValorDesconto, ItensNFiscal.DataEntrega, ItensNFiscal.DescricaoItem, ItensNFiscal.ValorAbatComissao, ItensNFiscal.NumIntPedVenda, ItensNFiscal.NumIntItemPedVenda, ItensNFiscal.NumIntDoc, ItensNFiscal.NumIntTrib, ItensNFiscal.NumIntDocOrig FROM NFiscal, ItensNFiscal, TiposDocInfo WHERE NFiscal.Serie = ? AND NFiscal.NumNotaFiscal = ? AND ItensNFiscal.Item = ? AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.FilialEmpresa = ? AND NFiscal.Status <> ? AND ItensNFiscal.Status <> ? AND TiposDocInfo.Tipo = ? AND TiposDocInfo.NomeTelaNFiscal<> ? ORDER BY DataEmissao DESC", _
        .lNumIntNF, .iStatus, .sProduto, .sUnidadeMed, .dQuantidade, .iAlmoxarifado, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, .lNumIntDoc, .lNumIntTrib, .lNumIntDocOrig, objItemNF.sSerieNFOrig, objItemNF.lNumNFOrig, objItemNF.iItem, giFilialEmpresa, STATUS_EXCLUIDO, STATUS_EXCLUIDO, DOCINFO_NF_INT_SAIDA, "NFiscalDev")
        If lErro <> AD_SQL_SUCESSO Then Error 35366

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35367
        If lErro = AD_SQL_SEM_DADOS Then Error 35368 'Não encontrou

        'Preenche o objItemNF com os dados lidos
        objItemNF.dPercDesc = .dPercDesc
        objItemNF.dPrecoUnitario = .dPrecoUnitario
        objItemNF.dQuantidade = .dQuantidade
        objItemNF.dtDataEntrega = .dtDataEntrega
        objItemNF.dValorAbatComissao = .dValorAbatComissao
        objItemNF.dValorDesconto = .dValorDesconto
        objItemNF.iAlmoxarifado = .iAlmoxarifado
        objItemNF.iStatus = .iStatus
        objItemNF.lNumIntDoc = .lNumIntDoc
        objItemNF.lNumIntDocOrig = .lNumIntDocOrig
        objItemNF.lNumIntItemPedVenda = .lNumIntPedVenda
        objItemNF.lNumIntTrib = .lNumIntTrib
        objItemNF.sDescricaoItem = .sDescricaoItem
        objItemNF.sProduto = .sProduto
        objItemNF.sUnidadeMed = .sUnidadeMed
        objItemNF.lNumIntNF = .lNumIntNF

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItemNFiscalSaida_Le_NumNFItem = SUCESSO

    Exit Function

Erro_ItemNFiscalSaida_Le_NumNFItem:

    ItemNFiscalSaida_Le_NumNFItem = Err

    Select Case Err

        Case 35365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 35366, 35367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL2", Err)

        Case 35368

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149627)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TiposDocInfo_Le_SiglaNFOriginal(ByVal iTipoNFiscal As Integer, ByVal objTipoDocInfo As ClassTipoDocInfo) As Long

Dim lErro As Long
Dim lComando As Long
Dim sSiglaNFOriginal As String

On Error GoTo Erro_TiposDocInfo_Le_SiglaNFOriginal

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 35635

    sSiglaNFOriginal = String(STRING_TIPO_DOC_INFO_SIGLA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT SiglaNFOriginal FROM TiposDocInfo WHERE Codigo = ? ", sSiglaNFOriginal, iTipoNFiscal)
    If lErro <> AD_SQL_SUCESSO Then Error 35636

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35637
    If lErro <> AD_SQL_SUCESSO Then Error 35638
    
    objTipoDocInfo.sSigla = sSiglaNFOriginal
    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 27263 Then Error 35639
    
    'Se não encontrou o Tipo de Documento --> erro
    If lErro <> SUCESSO Then Error 35640

    Call Comando_Fechar(lComando)
    
    TiposDocInfo_Le_SiglaNFOriginal = SUCESSO
    
    Exit Function
    
Erro_TiposDocInfo_Le_SiglaNFOriginal:

    TiposDocInfo_Le_SiglaNFOriginal = Err
    
    Select Case Err
    
        Case 35635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 35636, 35637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO1", Err, iTipoNFiscal)
            
        Case 35638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, iTipoNFiscal)
        
        Case 35639
            
        Case 35640
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, sSiglaNFOriginal)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149628)

    End Select
            
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFiscalFatEntrada_Lock_Gravacao(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim iIndice As Long
Dim alComando(0 To 4) As Long
Dim sNome As String
Dim sDescricao As String
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iFilialEmpresa As Integer
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim iFilialEmpresa1 As Integer

On Error GoTo Erro_NFiscalFatEntrada_Lock_Gravacao

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 42685
    Next

    sDescricao = String(STRING_NATUREZAOP_DESCRICAO, 0)

    lErro = Comando_ExecutarLockado(alComando(2), "SELECT Descricao FROM NaturezaOp WHERE Codigo = ? ", sDescricao, objNFiscal.sNaturezaOp)
    If lErro <> AD_SQL_SUCESSO Then Error 42686

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42687
    If lErro = AD_SQL_SEM_DADOS Then Error 42688
    
    lErro = Comando_LockShared(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 42689

    sNome = String(STRING_FILIAL_FORNECEDOR_NOME, 0)

    'Lê a FilialFornecedor
    lErro = Comando_ExecutarLockado(alComando(0), "SELECT Nome FROM FiliaisFornecedores WHERE CodFornecedor = ? AND CodFilial = ? ", sNome, objNFiscal.lFornecedor, objNFiscal.iFilialForn)
    If lErro <> AD_SQL_SUCESSO Then Error 42690

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42691
    If lErro = AD_SQL_SEM_DADOS Then Error 42692 'Não encontrou --> Erro

    'Faz um "lock" na FilialFornecedor
    lErro = Comando_LockShared(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 42693

    'Se Nota for interna, lock shared na série
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 42694
    If lErro <> SUCESSO Then Error 42695

    If objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA Then
        
        If objNFiscal.iFilialEmpresa = 0 Then objNFiscal.iFilialEmpresa = giFilialEmpresa
        
        lErro = Comando_ExecutarLockado(alComando(3), "SELECT FilialEmpresa FROM Serie WHERE Serie =? AND FilialEmpresa = ? ", iFilialEmpresa, objNFiscal.sSerie, objNFiscal.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 42696

        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42697
        If lErro = AD_SQL_SEM_DADOS Then Error 42698
        
        lErro = Comando_LockShared(alComando(3))
        If lErro <> AD_SQL_SUCESSO Then Error 42699
    End If
        
    'Se a Transportadora estiver preenchida
    If objNFiscal.iCodTransportadora > 0 Then
        'Lê a Transportadora
        lErro = Comando_ExecutarLockado(alComando(1), "SELECT Codigo FROM Transportadoras WHERE Codigo = ?", iCodigo, objNFiscal.iCodTransportadora)
        If lErro <> AD_SQL_SUCESSO Then Error 42700

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 42701
        If lErro = AD_SQL_SEM_DADOS Then Error 42702 'Não encontrou --> Erro

        'Faz um "lock" na transportadora
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 42703
    End If

    'Lock shared nos Produtos e Almoxarifados
    For Each objItemNF In objNFiscal.colItensNF
        objProduto.sCodigo = objItemNF.sProduto
        lErro = CF("Produto_Lock", objProduto)
        If lErro <> SUCESSO Then Error 42704
        
        objItemNF.iControleEstoque = objProduto.iControleEstoque
        objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
        
        For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
            
            objProduto.sCodigo = objItemRomaneio.sProduto
            
            lErro = CF("Produto_Lock", objProduto)
            If lErro <> SUCESSO Then Error 42704
        
            objItemRomaneio.iControleEstoque = objProduto.iControleEstoque
            objItemRomaneio.sUMEstoque = objProduto.sSiglaUMEstoque
        Next
        
        If objItemNF.iAlmoxarifado > 0 Then
        
            If objNFiscal.iFilialEmpresa = 0 Then objNFiscal.iFilialEmpresa = giFilialEmpresa
            
            iFilialEmpresa1 = objNFiscal.iFilialEmpresa
            lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa1)
            If lErro <> SUCESSO Then Error 42704
            
            lErro = Comando_ExecutarPos(alComando(3), "SELECT FilialEmpresa FROM Almoxarifado WHERE Codigo = ? AND FilialEmpresa = ? ", 0, iFilialEmpresa, objItemNF.iAlmoxarifado, iFilialEmpresa1)
            If lErro <> AD_SQL_SUCESSO Then Error 42705
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42706
            If lErro = AD_SQL_SEM_DADOS Then Error 42707
        
            lErro = Comando_LockShared(alComando(3))
            If lErro <> AD_SQL_SUCESSO Then Error 42708
        End If
    Next

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalFatEntrada_Lock_Gravacao = SUCESSO
    
    Exit Function
    
Erro_NFiscalFatEntrada_Lock_Gravacao:

    NFiscalFatEntrada_Lock_Gravacao = Err
    
    Select Case Err
    
        Case 42685
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 42690, 42691
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES1", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 42692
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 42693
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISFORNECEDORES1", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 42700, 42701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORA", Err, objNFiscal.iCodTransportadora)

        Case 42702
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objNFiscal.iCodTransportadora)

        Case 42703
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRANSPORTADORA", Err)

        Case 42686, 42687
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP", Err, objNFiscal.sNaturezaOp)

        Case 42688
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, objNFiscal.sNaturezaOp)
        
        Case 42689
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NATUREZAOP", Err, objNFiscal.sNaturezaOp)

        Case 42694, 42704
        
        Case 42695
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)
        
        Case 42696, 42697
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE", Err)
            
        Case 42698
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, objNFiscal.sSerie)

        Case 42699
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERIE", Err)
        
        Case 42705, 42706
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO", Err, objItemNF.iAlmoxarifado)
        
        Case 42707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objItemNF.iAlmoxarifado)
        
        Case 42708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ALMOXARIFADO1", Err, objItemNF.iAlmoxarifado)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149629)

    End Select
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ProdutoKit_Le_Raiz(objProdutoKit As ClassProdutoKit) As Long
'Le a tabela de Kits de Produtos a partir do produto e da versão passados em objProdutoKit e nível = KIT_NIVEL_RAIZ

Dim lErro As Long
Dim lComando As Long
Dim tProdutoKit As typeProdutoKit
Dim sComando_SQL As String
Dim iClasseUM As Integer

On Error GoTo Erro_ProdutoKit_Le_Raiz

    'Abertura comando
    lComando = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 34872

    tProdutoKit.sProdutoRaiz = String(STRING_PRODUTO, 0)
    tProdutoKit.sProduto = String(STRING_PRODUTO, 0)
    tProdutoKit.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)
    tProdutoKit.sVersao = String(STRING_KIT_VERSAO, 0)
                                                                                                                    
    sComando_SQL = "SELECT ProdutoKit.PercentualPerda, ProdutoKit.ProdutoRaiz, ProdutoKit.Versao, ProdutoKit.Nivel, ProdutoKit.Seq, ProdutoKit.Produto, ProdutoKit.SeqPai, ProdutoKit.Quantidade, ProdutoKit.UnidadeMed, ProdutoKit.Composicao, Produtos.ClasseUM FROM ProdutoKit, Produtos WHERE ProdutoKit.ProdutoRaiz = ? AND ProdutoKit.Versao = ? AND ProdutoKit.Nivel = ? AND Produtos.Codigo = ProdutoKit.Produto"
    lErro = Comando_Executar(lComando, sComando_SQL, tProdutoKit.dPercentualPerda, tProdutoKit.sProdutoRaiz, tProdutoKit.sVersao, tProdutoKit.iNivel, tProdutoKit.iSeq, tProdutoKit.sProduto, tProdutoKit.iSeqPai, tProdutoKit.dQuantidade, tProdutoKit.sUnidadeMed, tProdutoKit.iComposicao, iClasseUM, objProdutoKit.sProdutoRaiz, objProdutoKit.sVersao, KIT_NIVEL_RAIZ)
    If lErro <> AD_SQL_SUCESSO Then Error 34873

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 34874

    If lErro = AD_SQL_SEM_DADOS Then Error 34875

    Set objProdutoKit = New ClassProdutoKit

    objProdutoKit.dQuantidade = tProdutoKit.dQuantidade
    objProdutoKit.iComposicao = tProdutoKit.iComposicao
    objProdutoKit.iNivel = tProdutoKit.iNivel
    objProdutoKit.iSeq = tProdutoKit.iSeq
    objProdutoKit.iSeqPai = tProdutoKit.iSeqPai
    objProdutoKit.sProduto = tProdutoKit.sProduto
    objProdutoKit.sUnidadeMed = tProdutoKit.sUnidadeMed
    objProdutoKit.sProdutoRaiz = tProdutoKit.sProdutoRaiz
    objProdutoKit.sVersao = tProdutoKit.sVersao
    objProdutoKit.iClasseUM = iClasseUM
    objProdutoKit.dPercentualPerda = tProdutoKit.dPercentualPerda

    'Fechamento comando
    Call Comando_Fechar(lComando)

    ProdutoKit_Le_Raiz = SUCESSO

    Exit Function

Erro_ProdutoKit_Le_Raiz:

    ProdutoKit_Le_Raiz = Err

    Select Case Err

        Case 34872
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 34873, 34874
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", Err)

        Case 34875

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149630)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutosKit_Le_PrimeiroNivel(ByVal objKit As ClassKit) As Long
'Preenche objKit.colComponentes a partir do produto e da versão passados em objKit e Nivel = KIT_NIVEL_RAIZ + 1

Dim lErro As Long
Dim lComando As Long
Dim tProdutoKit As typeProdutoKit
Dim objProdutoKit As ClassProdutoKit
Dim sComando_SQL As String, sProdutoDescricao As String
Dim iClasseUM As Integer

On Error GoTo Erro_ProdutosKit_Le_PrimeiroNivel

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 34876

    tProdutoKit.sProdutoRaiz = String(STRING_PRODUTO, 0)
    tProdutoKit.sProduto = String(STRING_PRODUTO, 0)
    tProdutoKit.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)
    tProdutoKit.sVersao = String(STRING_KIT_VERSAO, 0)
    sProdutoDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

    sComando_SQL = "SELECT ProdutoKit.ProdutoRaiz, ProdutoKit.Versao, ProdutoKit.Nivel, ProdutoKit.Seq, ProdutoKit.Produto, Produtos.Descricao, Produtos.ClasseUM, ProdutoKit.SeqPai, ProdutoKit.Quantidade, ProdutoKit.UnidadeMed, ProdutoKit.Composicao FROM Produtos, ProdutoKit WHERE Produtos.Codigo = ProdutoKit.Produto AND ProdutoKit.ProdutoRaiz = ? AND ProdutoKit.Versao = ? AND ProdutoKit.Nivel = ?"
    lErro = Comando_Executar(lComando, sComando_SQL, tProdutoKit.sProdutoRaiz, tProdutoKit.sVersao, tProdutoKit.iNivel, tProdutoKit.iSeq, tProdutoKit.sProduto, sProdutoDescricao, iClasseUM, tProdutoKit.iSeqPai, tProdutoKit.dQuantidade, tProdutoKit.sUnidadeMed, tProdutoKit.iComposicao, objKit.sProdutoRaiz, objKit.sVersao, KIT_NIVEL_RAIZ + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 34877

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 34878

    If lErro = AD_SQL_SEM_DADOS Then Error 34879

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objProdutoKit = New ClassProdutoKit

        objProdutoKit.dQuantidade = tProdutoKit.dQuantidade
        objProdutoKit.iComposicao = tProdutoKit.iComposicao
        objProdutoKit.iNivel = tProdutoKit.iNivel
        objProdutoKit.iSeq = tProdutoKit.iSeq
        objProdutoKit.iSeqPai = tProdutoKit.iSeqPai
        objProdutoKit.sProduto = tProdutoKit.sProduto
        objProdutoKit.sProdutoDesc = sProdutoDescricao
        objProdutoKit.iClasseUM = iClasseUM
        objProdutoKit.sUnidadeMed = tProdutoKit.sUnidadeMed
        objProdutoKit.sProdutoRaiz = tProdutoKit.sProdutoRaiz
        objProdutoKit.sVersao = tProdutoKit.sVersao

        objKit.colComponentes.Add objProdutoKit

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 34880

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    ProdutosKit_Le_PrimeiroNivel = SUCESSO

    Exit Function

Erro_ProdutosKit_Le_PrimeiroNivel:

    ProdutosKit_Le_PrimeiroNivel = Err

    Select Case Err

        Case 34876
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 34877, 34878, 34880
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", Err)

        Case 34879

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149631)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFiscalInternaSaidaNaoDev_Le_Numero(ByVal objNFiscalOriginal As ClassNFiscal) As Long
'Pesquisa nas Tabelas [NFiscal e TiposDocInfo] UNION [NotaFiscalBaixadas e TiposDocInfo]
'Busca pela Serie e Número passados
'Filtros : TiposDocInfo.Tipo=DOCINFO_NF_INT_SAIDA
'          NFiscal.Status<>STATUS_EXCLUIDO
'          NomeTelaNFiscal<>"NFiscalDev"
'          FilialEmpresa = giFilialEmpresa
'          Complementar = DOCINFO_NORMAL

'Ordena por DataEmissao, DECENDING
'Lê todos os Dados da Tabela NFiscal

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscalInternaSaidaNaoDev_Le_Numero

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 30762

    With tNFiscal

        .sMensagemNota = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNaturezaOp = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sPlaca = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sPlacaUF = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSerie = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNumPedidoTerc = String(STRING_BUFFER_MAX_TEXTO, 0)
        
        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_BUFFER_MAX_TEXTO, 0)
        '.lVolumeMarca = String(STRING_BUFFER_MAX_TEXTO, 0)
        
        sComandoSQL = "SELECT NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, CodTranspRedesp, DetPagFrete FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.NumNotaFiscal = ? AND NFiscal.Serie = ? AND TiposDocInfo.Tipo = ? AND NFiscal.Status <> ? AND TiposDocInfo.NomeTelaNFiscal <> ? AND NFiscal.FilialEmpresa = ? AND TiposDocInfo.Complementar = ? ORDER BY DataEmissao DESC "

        lErro = Comando_Executar(lComando, sComandoSQL, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .iCodTranspRedesp, .iDetPagFrete, objNFiscalOriginal.lNumNotaFiscal, objNFiscalOriginal.sSerie, DOCINFO_NF_INT_SAIDA, STATUS_EXCLUIDO, "NFiscalDev", giFilialEmpresa, DOCINFO_NORMAL)
        If lErro <> AD_SQL_SUCESSO Then Error 30763

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30764

        'Se não achou --> erro
        If lErro = AD_SQL_SEM_DADOS Then Error 30765

        objNFiscalOriginal.dPesoBruto = .dPesoBruto
        objNFiscalOriginal.dPesoLiq = .dPesoLiq
        objNFiscalOriginal.dtDataEmissao = .dtDataEmissao
        objNFiscalOriginal.dtDataEntrada = .dtDataEntrada
        objNFiscalOriginal.dtDataSaida = .dtDataSaida
        objNFiscalOriginal.dtDataVencimento = .dtDataVencimento
        objNFiscalOriginal.dValorDesconto = .dValorDesconto
        objNFiscalOriginal.dValorFrete = .dValorFrete
        objNFiscalOriginal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscalOriginal.dValorProdutos = .dValorProdutos
        objNFiscalOriginal.dValorSeguro = .dValorSeguro
        objNFiscalOriginal.dValorTotal = .dValorTotal
        objNFiscalOriginal.iCanal = .iCanal
        objNFiscalOriginal.iClasseDocCPR = .iClasseDocCPR
        objNFiscalOriginal.iCodTransportadora = .iCodTransportadora
        objNFiscalOriginal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscalOriginal.iDetPagFrete = .iDetPagFrete
        objNFiscalOriginal.iFilialCli = .iFilialCli
        objNFiscalOriginal.iFilialEntrega = .iFilialEntrega
        objNFiscalOriginal.iFilialForn = .iFilialForn
        objNFiscalOriginal.iStatus = .iStatus
        objNFiscalOriginal.iTabelaPreco = .iTabelaPreco
        objNFiscalOriginal.iTipoNFiscal = .iTipoNFiscal
        objNFiscalOriginal.lVolumeQuant = .lVolumeQuant
        objNFiscalOriginal.lCliente = .lCliente
        objNFiscalOriginal.lFornecedor = .lFornecedor
        objNFiscalOriginal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscalOriginal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscalOriginal.lNumIntTrib = .lNumIntTrib
        objNFiscalOriginal.lNumNotaFiscal = .lNumNotaFiscal
        objNFiscalOriginal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscalOriginal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscalOriginal.sMensagemNota = .sMensagemNota
        objNFiscalOriginal.sNaturezaOp = .sNaturezaOp
        objNFiscalOriginal.sPlaca = .sPlaca
        objNFiscalOriginal.sPlacaUF = .sPlacaUF
        objNFiscalOriginal.sSerie = .sSerie
        objNFiscalOriginal.lVolumeEspecie = .lVolumeEspecie
        objNFiscalOriginal.lVolumeMarca = .lVolumeMarca
        objNFiscalOriginal.lNumIntDoc = .lNumIntDoc
        objNFiscalOriginal.iFilialEmpresa = .iFilialEmpresa

    End With

    'Libera comando
    Call Comando_Fechar(lComando)

    NFiscalInternaSaidaNaoDev_Le_Numero = SUCESSO

    Exit Function

Erro_NFiscalInternaSaidaNaoDev_Le_Numero:

    NFiscalInternaSaidaNaoDev_Le_Numero = Err

    Select Case Err

        Case 30762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30763, 30764
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 30765 'Tratado pela rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149632)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RecebMaterialC_Pesquisa_NotaFiscal(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long) As Long
'Tenta ler a Nota Fiscal correspondente ao Recebimento

Dim lErro As Long
Dim iTipoDoc As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_RecebMaterialC_Pesquisa_NotaFiscal

    If objNFiscal.iTipoNFiscal = DOCINFO_NRCP Then iTipoDoc = DOCINFO_NF_INT_ENTRADA
    If objNFiscal.iTipoNFiscal = DOCINFO_NRCC Then iTipoDoc = DOCINFO_NF_EXTERNA

    'Tenta ler Nota Fiscal correspondente ao Recebimento
    lErro = Comando_Executar(lComando1, "SELECT NFiscal.NumIntDoc FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.Cliente = ? AND NFiscal.FilialCli = ? AND NFiscal.DataEntrada = ? AND NFiscal.Serie = ? AND NFiscal.NumNotaFiscal = ? AND TiposDocInfo.Tipo = ? ", lNumIntDoc, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEntrada, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, iTipoDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 30741

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30742

    'Se encontrou a Nota Fiscal correspondente ao recebimento --> erro
    If lErro = AD_SQL_SUCESSO Then Error 30743
    
    RecebMaterialC_Pesquisa_NotaFiscal = SUCESSO
    
    Exit Function
    
Erro_RecebMaterialC_Pesquisa_NotaFiscal:

    RecebMaterialC_Pesquisa_NotaFiscal = Err
    
    Select Case Err
    
        Case 30721, 30723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCALBAIXADA2", Err)
        
        Case 30741, 30742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 30743
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149633)
            
    End Select
            
    Exit Function
            
End Function

Function RecebMaterialF_Pesquisa_NotaFiscal(ByVal objNFiscal As ClassNFiscal, ByVal lComando1 As Long) As Long

Dim lErro As Long
Dim iTipoDoc As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_RecebMaterialF_Pesquisa_NotaFiscal
    
    'obtem o tipo em que este recebimento se transformará
    If objNFiscal.iTipoNFiscal = DOCINFO_NRFP Then iTipoDoc = DOCINFO_NF_INT_ENTRADA
    If objNFiscal.iTipoNFiscal = DOCINFO_NRFF Then iTipoDoc = DOCINFO_NF_EXTERNA

    'Tenta ler Nota Fiscal correspondente ao Recebimento
    lErro = Comando_Executar(lComando1, "SELECT NFiscal.NumIntDoc FROM NFiscal, TiposDocInfo WHERE NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.Fornecedor = ? AND NFiscal.FilialForn = ? AND NFiscal.DataEntrada = ? AND NFiscal.Serie = ? AND NFiscal.NumNotaFiscal = ? AND TiposDocInfo.Tipo = ? ", lNumIntDoc, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, iTipoDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 30539

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30540

    If lErro = AD_SQL_SUCESSO Then Error 30541

    RecebMaterialF_Pesquisa_NotaFiscal = SUCESSO
    
    Exit Function
    
Erro_RecebMaterialF_Pesquisa_NotaFiscal:

    RecebMaterialF_Pesquisa_NotaFiscal = Err
    
    Select Case Err
    
        Case 30539, 30540
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 30629, 30630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCALBAIXADA2", Err)
        
        Case 30541

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149634)
            
    End Select

    Exit Function

End Function

Function NFiscalEntradaRemessa_Lock_Gravacao(ByVal objNFiscal As ClassNFiscal, ByVal lComando As Long) As Long
'Faz o lock dos campos da NFiszcal Entrada Remessa

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_NFiscalEntradaRemessa_Lock_Gravacao

    If objNFiscal.lFornecedor > 0 Then
        'Verifica se Filial Fornecedor
        lErro = CF("FilialFornecedor_Lock", objNFiscal.lFornecedor, objNFiscal.iFilialForn)
        If lErro <> SUCESSO And lErro <> 42870 Then Error 35611
        
        'Se não encontrou a Filial do Cliente --> erro
        If lErro <> SUCESSO Then Error 35613
    Else
        'Verifica se Filial Cliente existe
        lErro = CF("FilialCliente_Lock", objNFiscal.lCliente, objNFiscal.iFilialCli)
        If lErro <> SUCESSO And lErro <> 43050 Then Error 35614
        
        'Se não encontrou a Filial Cliente --> erro
        If lErro <> SUCESSO Then Error 35616
    End If

    'Se a Transportadora estiver preenchida
    If objNFiscal.iCodTransportadora > 0 Then
        'Verifica se a Transportadora existe
        lErro = CF("Transportadora_Lock1", objNFiscal.iCodTransportadora)
        If lErro <> SUCESSO And lErro <> 43038 Then Error 35618
        
        'Se não encontrou a Transportadora --> erro
        If lErro <> SUCESSO Then Error 35620

    End If
    
    'Para cada Ítem existente na coleção de Ítens
    For Each objItemNF In objNFiscal.colItensNF
        
        objProduto.sCodigo = objItemNF.sProduto
        'Faz o Lock do Produto
        lErro = CF("Produto_Lock", objProduto)
        If lErro <> SUCESSO Then Error 35622
        
        objItemNF.iControleEstoque = objProduto.iControleEstoque
                
        For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
            
            objProduto.sCodigo = objItemRomaneio.sProduto
            
            lErro = CF("Produto_Lock", objProduto)
            If lErro <> SUCESSO Then Error 35622
        
            objItemRomaneio.iControleEstoque = objProduto.iControleEstoque
            objItemRomaneio.sUMEstoque = objProduto.sSiglaUMEstoque
        Next
        
        If objItemNF.iAlmoxarifado > 0 Then
                
            'Faz o Lock do Almoxarifado
            lErro = CF("Almoxarifado_Lock", objItemNF.iAlmoxarifado, lComando)
            If lErro <> SUCESSO Then Error 35802
            
        End If
        
    Next

    NFiscalEntradaRemessa_Lock_Gravacao = SUCESSO
    
    Exit Function
    
Erro_NFiscalEntradaRemessa_Lock_Gravacao:

    NFiscalEntradaRemessa_Lock_Gravacao = Err
    
    Select Case Err
    
        Case 35611, 35614, 35618, 35622, 35802
        
        Case 35613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objNFiscal.lFornecedor, objNFiscal.iFilialCli)
        
        Case 35616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objNFiscal.lCliente, objNFiscal.iFilialCli)

        Case 35620
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objNFiscal.iCodTransportadora)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149635)
            
    End Select
            
    Exit Function
    
End Function

Function Reservas_Le_Codigo(ByVal colCodigo As Collection) As Long
'Lê o código de todas as reservas no Bd

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long

On Error GoTo Erro_Reservas_Le_Codigo

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23917

    'Executar comando SQL
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Reserva WHERE FilialEmpresa = ? ORDER BY Codigo", lCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 23918

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23919

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        colCodigo.Add lCodigo

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23920

    Loop

    Call Comando_Fechar(lComando)

    Reservas_Le_Codigo = SUCESSO

    Exit Function

Erro_Reservas_Le_Codigo:

    Reservas_Le_Codigo = Err

    Select Case Err

        Case 23917
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23918, 23919, 23920
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149636)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Reserva_Le(ByVal objReserva As ClassReserva) As Long
'Lê a reserva a partir do código
'Filtra por FilialEmpresa=giFilialEmpresa

Dim lErro As Long
Dim tReserva As typeReserva
Dim lComando As Long

On Error GoTo Erro_Reserva_Le

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23925

    With tReserva
        .sProduto = String(STRING_PRODUTO, 0)
        .sCodUsuario = String(STRING_SIGLA_USUARIO, 0)
        .sResponsavel = String(STRING_RESPONSAVEL_RESERVA, 0)

        'Lê a Reserva
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Produto, Almoxarifado, TipoDoc, DocOrigem, NumIntOrigem, Quantidade, DataReserva, DataValidade, CodUsuario, Responsavel FROM Reserva WHERE FilialEmpresa = ? AND Codigo=? ", _
            .lNumIntDoc, .iFilialEmpresa, .sProduto, .iAlmoxarifado, .iTipoDoc, .lDocOrigem, .lNumIntOrigem, .dQuantidade, .dtDataReserva, .dtDataValidade, .sCodUsuario, .sResponsavel, giFilialEmpresa, objReserva.lCodigo)

    End With

    If lErro <> AD_SQL_SUCESSO Then Error 23926

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23927

    'Caso não haja nenhum Reserva com este código dessa FilialEmpresa no Bd
    If lErro = AD_SQL_SEM_DADOS Then Error 23928

    'Preenche objeto com os Dados lidos no Banco de Dados
    With tReserva
        objReserva.iFilialEmpresa = .iFilialEmpresa
        objReserva.sProduto = .sProduto
        objReserva.iAlmoxarifado = .iAlmoxarifado
        objReserva.iTipoDoc = .iTipoDoc
        objReserva.lDocOrigem = .lDocOrigem
        objReserva.lNumIntOrigem = .lNumIntOrigem
        objReserva.lNumIntDoc = .lNumIntDoc
        objReserva.dQuantidade = .dQuantidade
        objReserva.dtDataValidade = .dtDataValidade
        objReserva.dtDataReserva = .dtDataReserva
        objReserva.sCodUsuario = .sCodUsuario
        objReserva.sResponsavel = .sResponsavel

    End With

    'Fecha comando
    Call Comando_Fechar(lComando)

    Reserva_Le = SUCESSO

    Exit Function

Erro_Reserva_Le:

    Reserva_Le = Err

    Select Case Err

        Case 23925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23926, 23927
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA", Err)

        Case 23928  'Não existe reserva com Codigo %l na FilialEmpresa%i
                    'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149637)

    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemPedido_Le_Produto(ByVal objItemPedido As ClassItemPedido) As Long
'Lê no Banco de Dados o Item correspondente ao produto e CodPedido

Dim lErro As Long
Dim lComando As Long
Dim tItemPedido As typeItemPedido
Dim colReservaIte As New colReservaItem

On Error GoTo Erro_ItemPedido_Le_Produto

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 30059

    tItemPedido.sDescricao = String(STRING_ITEM_PEDIDO_DESCRICAO, 0)
    tItemPedido.sLote = String(STRING_ITEM_PEDIDO_LOTE, 0)
    tItemPedido.sProdutoDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedido.sProdutoNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    tItemPedido.sUMEstoque = String(STRING_UM_SIGLA_UM_BASE, 0)
    tItemPedido.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)

    lErro = Comando_Executar(lComando, "SELECT Quantidade, PrecoUnitario, PrecoTotal, ItensPedidoDeVenda.ClasseUM, UnidadeMed, ValorDesconto, DataEntrega, ItensPedidoDeVenda.Descricao, Lote, Status, ValorAbatComissao, QuantCancelada, QuantReservada, QuantFaturada, Produtos.Descricao, NomeReduzido, ControleEstoque, SiglaUMEstoque, NumIntDoc FROM ItensPedidoDeVenda, Produtos WHERE ItensPedidoDeVenda.Produto = Produtos.Codigo AND ItensPedidoDeVenda.Produto= ? AND ItensPedidoDeVenda.FilialEmpresa = ? AND CodPedido = ? ", tItemPedido.dQuantidade, tItemPedido.dPrecoUnitario, tItemPedido.dPrecoTotal, tItemPedido.iClasseUM, tItemPedido.sUnidadeMed, tItemPedido.dValorDesconto, tItemPedido.dtDataEntrega, _
        tItemPedido.sDescricao, tItemPedido.sLote, tItemPedido.iStatus, tItemPedido.dValorAbatComissao, tItemPedido.dQuantCancelada, tItemPedido.dQuantReservada, tItemPedido.dQuantFaturada, tItemPedido.sProdutoDescricao, tItemPedido.sProdutoNomeReduzido, tItemPedido.iControleEstoque, tItemPedido.sUMEstoque, tItemPedido.lNumIntDoc, objItemPedido.sProduto, giFilialEmpresa, objItemPedido.lCodPedido)
    If lErro <> AD_SQL_SUCESSO Then Error 30060

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30061

    If lErro = AD_SQL_SEM_DADOS Then Error 30062

    objItemPedido.dQuantidade = tItemPedido.dQuantidade
    objItemPedido.dPrecoUnitario = tItemPedido.dPrecoUnitario
    objItemPedido.dPrecoTotal = tItemPedido.dPrecoTotal
    objItemPedido.iClasseUM = tItemPedido.iClasseUM
    objItemPedido.sUnidadeMed = tItemPedido.sUnidadeMed
    objItemPedido.dValorDesconto = tItemPedido.dValorDesconto
    objItemPedido.dtDataEntrega = tItemPedido.dtDataEntrega
    objItemPedido.sDescricao = tItemPedido.sDescricao
    objItemPedido.sLote = tItemPedido.sLote
    objItemPedido.iStatus = tItemPedido.iStatus
    objItemPedido.dValorAbatComissao = tItemPedido.dValorAbatComissao
    objItemPedido.dQuantCancelada = tItemPedido.dQuantCancelada
    objItemPedido.dQuantFaturada = tItemPedido.dQuantFaturada
    objItemPedido.dQuantReservada = tItemPedido.dQuantReservada
    objItemPedido.sDescricao = tItemPedido.sDescricao
    objItemPedido.sProdutoNomeReduzido = tItemPedido.sProdutoNomeReduzido
    objItemPedido.iControleEstoque = tItemPedido.iControleEstoque
    objItemPedido.sUMEstoque = tItemPedido.sUMEstoque
    objItemPedido.lNumIntDoc = tItemPedido.lNumIntDoc

    'Libera comando
    Call Comando_Fechar(lComando)

    ItemPedido_Le_Produto = SUCESSO

    Exit Function

Erro_ItemPedido_Le_Produto:

    ItemPedido_Le_Produto = Err

    Select Case Err

        Case 30059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30060, 30061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", Err)

        Case 30062 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149638)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Almoxarifados_Le_Cod_NomeRed(ByVal colCodNome As AdmColCodigoNome) As Long
'Lê o código e nome reduzido dos almoxarifados e retorna na coleção

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer
Dim sNomeRed As String

On Error GoTo Erro_Almoxarifados_Le_Cod_NomeRed

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23921

    sNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)

    'Executar comando SQL
    lErro = Comando_Executar(lComando, "SELECT Codigo, NomeReduzido FROM Almoxarifado", iCodigo, sNomeRed)
    If lErro <> AD_SQL_SUCESSO Then Error 23922

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23923

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        colCodNome.Add iCodigo, sNomeRed

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23924

    Loop

    Call Comando_Fechar(lComando)

    Almoxarifados_Le_Cod_NomeRed = SUCESSO

    Exit Function

Erro_Almoxarifados_Le_Cod_NomeRed:

    Almoxarifados_Le_Cod_NomeRed = Err

    Select Case Err

        Case 23921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23922, 23923, 23924
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RESERVA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149639)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function GeracaoOP_SelecionarItens(ByVal objGeracaoOP As ClassGeracaoOP, Optional bExibeSemDados As Boolean = True) As Long
'pesquisa no bd os itens candidatos preenchendo objGeracaoOP.colItens a partir das outras informacoes em

Dim lErro As Long, lComando As Long, sSelect As String, iStatusNaoAtendido As Integer, lNumIntItemPV As Long
'buffers para receber registros selecionados
Dim sItensPedidoDeVendaProduto As String, dtItensPedidoDeVendaDataEntrega As Date
Dim iPedidosDeVendaFilialEmpresa As Integer, lPedidosDeVendaCodigo As Long, iPedidosDeVendaFilial As Integer, dtPedidosDeVendaDataEmissao As Date
Dim sClientesNomeReduzido As String, iProdutosCompras As Integer, iProdutosPCP As Integer
Dim sProdutosDescricao As String, sProdutosSiglaUM As String, iProdutosClasseUM As Integer, sItensPVSiglaUM  As String, dItensPedidoDeVendaQuantFalta As Double
Dim iPrioridade As Integer

On Error GoTo Erro_GeracaoOP_SelecionarItens

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27434

    iStatusNaoAtendido = STATUS_NAO_ATENDIDO
    iProdutosCompras = PRODUTO_COMPRAVEL
    iProdutosPCP = PRODUTO_PCP_PODE

    Call GeracaoOP_SelecionarItens1(objGeracaoOP, sSelect)

    sItensPedidoDeVendaProduto = String(STRING_PRODUTO, 0)
    sClientesNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    sProdutosDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    sProdutosSiglaUM = String(STRING_UM_SIGLA, 0)
    sItensPVSiglaUM = String(STRING_UM_PV_SIGLA, 0)
        
    'executa a preparacao da parte fixa do SELECT
    lErro = GeracaoOP_SelecionarItens2(lComando, sSelect, _
    lNumIntItemPV, sItensPedidoDeVendaProduto, dtItensPedidoDeVendaDataEntrega, dItensPedidoDeVendaQuantFalta, _
    iPedidosDeVendaFilialEmpresa, lPedidosDeVendaCodigo, iPedidosDeVendaFilial, _
    dtPedidosDeVendaDataEmissao, sClientesNomeReduzido, sProdutosDescricao, sProdutosSiglaUM, iProdutosClasseUM, sItensPVSiglaUM, iProdutosCompras, iProdutosPCP, iStatusNaoAtendido, iPrioridade)
    If lErro <> SUCESSO Then Error 27435

    'complementa a passagem dos parametros que variam de acordo com a selecao do usuario
    'e executa o SELECT
    lErro = GeracaoOP_SelecionarItens3(lComando, objGeracaoOP)
    If lErro <> SUCESSO Then Error 27436

    'processa todos os registros retornados pelo SELECT executado acima
    lErro = GeracaoOP_SelecionarItens4(lComando, objGeracaoOP, _
        lNumIntItemPV, sItensPedidoDeVendaProduto, dtItensPedidoDeVendaDataEntrega, dItensPedidoDeVendaQuantFalta, _
        iPedidosDeVendaFilialEmpresa, lPedidosDeVendaCodigo, iPedidosDeVendaFilial, _
        dtPedidosDeVendaDataEmissao, sClientesNomeReduzido, sProdutosDescricao, sProdutosSiglaUM, iProdutosClasseUM, sItensPVSiglaUM, iPrioridade, bExibeSemDados)
    If lErro <> SUCESSO Then Error 27437

    lErro = Comando_Fechar(lComando)

    GeracaoOP_SelecionarItens = SUCESSO

    Exit Function

Erro_GeracaoOP_SelecionarItens:

    GeracaoOP_SelecionarItens = Err

    Select Case Err

        Case 27435, 27436, 27437

        Case 27434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149640)

    End Select

    lErro = Comando_Fechar(lComando)

    Exit Function

End Function

Private Sub GeracaoOP_SelecionarItens1(ByVal objGeracaoOP As ClassGeracaoOP, sSelect As String)
'auxiliar a GeracaoOP_SelecionarItens
'monta o SELECT para obtencao dos itens dinamicamente.

Dim sFrom As String, sWhere As String, sFields As String, sOrderBy As String

    sFields = "ItensPedidoDeVenda.NumIntDoc, ItensPedidoDeVenda.Produto, ItensPedidoDeVenda.DataEntrega"
    sFields = sFields & ", (ItensPedidoDeVenda.Quantidade - (ItensPedidoDeVenda.QuantFaturada + ItensPedidoDeVenda.QuantCancelada +"
    'sFields = sFields & "ItensPedidoDeVenda.QuantOP + ItensPedidoDeVenda.QuantSC + ItensPedidoDeVenda.QuantReservada))"
    sFields = sFields & "ItemPV_QuantEmOP.QuantEmOP + ItensPedidoDeVenda.QuantSC + ItensPedidoDeVenda.QuantReservada))"
    sFields = sFields & ", PedidosDeVenda.FilialEmpresa, PedidosDeVenda.Codigo, PedidosDeVenda.Filial, PedidosDeVenda.DataEmissao"
    sFields = sFields & ", Clientes.NomeReduzido"
    sFields = sFields & ", Produtos.Descricao, Produtos.SiglaUMEstoque, Produtos.ClasseUM, ItensPedidoDeVenda.UnidadeMed, ItensPedidoDeVenda.Prioridade "

    'sFrom = " FROM PedidosDeVenda, ItensPedidoDeVenda, Produtos, Clientes"
    sFrom = " FROM PedidosDeVenda, ItensPedidoDeVenda, Produtos, Clientes,ItemPV_QuantEmOP, ProdutosFilial"
    
    sWhere = " WHERE PedidosDeVenda.FilialEmpresa = ItensPedidoDeVenda.FilialEmpresa AND PedidosDeVenda.Codigo = ItensPedidoDeVenda.CodPedido AND ProdutosFilial.FilialEmpresa = PedidosDeVenda.FilialEmpresa AND ProdutosFilial.Produto = ItensPedidoDeVenda.Produto AND  PedidosDeVenda.Cliente = Clientes.Codigo AND ItensPedidoDeVenda.Produto = Produtos.Codigo "
    sWhere = sWhere & " AND Produtos.Compras <> ? AND Produtos.PCP = ? AND ItensPedidoDeVenda.Status = ? AND  ProdNaFilial = 1 "
    'sWhere = sWhere & " AND (ItensPedidoDeVenda.Quantidade - (ItensPedidoDeVenda.QuantCancelada + ItensPedidoDeVenda.QuantReservada + ItensPedidoDevenda.QuantFaturada + ItensPedidoDeVenda.QuantOP + ItensPedidoDeVenda.QuantSC)) > 0"
    sWhere = sWhere & " AND ItensPedidoDeVenda.NumIntDoc = ItemPV_QuantEmOP.NumIntDoc AND (ItensPedidoDeVenda.Quantidade - (ItensPedidoDeVenda.QuantCancelada + ItensPedidoDeVenda.QuantReservada + ItensPedidoDevenda.QuantFaturada + ItemPV_QuantEmOP.QuantEmOP + ItensPedidoDeVenda.QuantSC)) > 0"
    sWhere = sWhere & " AND NOT EXISTS (Select PedidoDeVenda FROM BloqueiosPV WHERE BloqueiosPV.FilialEmpresa = PedidosDeVenda.FilialEmpresa AND BloqueiosPV.PedidoDeVenda = PedidosDeVenda.Codigo AND BloqueiosPV.TipoDeBloqueio <> 2 AND BloqueiosPV.TipoDeBloqueio <> 3 AND BloqueiosPV.CodUsuarioLib = '')"

    'Se entrar como empresa toda pode pegar itens de todas as filiais, senao pegar itens somente da filial corrente
    If giFilialEmpresa <> EMPRESA_TODA Then
        sWhere = sWhere & " AND PedidosDeVenda.FilialEmpresa = ?"
    End If

    'Se pedido de venda inicial preenchido
    If (objGeracaoOP.lPVDe <> 0) Then
       sWhere = sWhere & " AND PedidosDeVenda.Codigo >= ?"
    End If

    'Se pedido de venda final preenchido
    If (objGeracaoOP.lPVAte <> 0) Then
       sWhere = sWhere & " AND PedidosDeVenda.Codigo <= ?"
    End If

    'Se cliente do pedido de venda inicial preenchido
    If (objGeracaoOP.lClientesDe <> 0) Then
       sWhere = sWhere & " AND PedidosDeVenda.Cliente >= ?"
    End If

    'Se cliente do pedido de venda final preenchido
    If (objGeracaoOP.lClientesAte <> 0) Then
       sWhere = sWhere & " AND PedidosDeVenda.Cliente <= ?"
    End If

    'Se data de entrega inicial do item preenchida
    If (objGeracaoOP.dtEntregaDe <> DATA_NULA) Then
       sWhere = sWhere & " AND ItensPedidoDeVenda.DataEntrega >= ?"
    End If

    'Se data de entrega final do item preenchida
    If (objGeracaoOP.dtEntregaAte <> DATA_NULA) Then
       sWhere = sWhere & " AND ItensPedidoDeVenda.DataEntrega <= ?"
    End If

    'se produto inicial preenchido
    If (objGeracaoOP.sProdDe <> "") Then
       sWhere = sWhere & " AND ItensPedidoDeVenda.Produto >= ?"
    End If

    'se produto final preenchido
    If (objGeracaoOP.sProdAte <> "") Then
       sWhere = sWhere & " AND ItensPedidoDeVenda.Produto <= ?"
    End If

    'se NAO deve incluir item de pedido de venda com Ordem de Producao associada
    If (objGeracaoOP.iIncluiPVcomOP = DESMARCADO) Then
       sWhere = sWhere & " AND ItensPedidoDeVenda.QuantOP = 0"
    End If

    sOrderBy = " ORDER BY " & objGeracaoOP.sOrdenacao

    sSelect = "SELECT " & sFields & sFrom & sWhere & sOrderBy

End Sub

Private Function GeracaoOP_SelecionarItens2(ByVal lComando As Long, ByVal sSelect As String, vNumIntItemPV As Variant, _
    vItensPedidoDeVendaProduto As Variant, vItensPedidoDeVendaDataEntrega As Variant, _
    vItensPedidoDeVendaQuantFalta As Variant, vPedidosDeVendaFilialEmpresa As Variant, _
    vPedidosDeVendaCodigo As Variant, vPedidosDeVendaFilial As Variant, _
    vPedidosDeVendaDataEmissao As Variant, vClientesNomeReduzido As Variant, _
    vProdutosDescricao As Variant, vProdutosSiglaUM As Variant, vProdutosClasseUM As Variant, sItensPVSiglaUM As Variant, _
    vProdutosCompras As Variant, vProdutosPCP As Variant, vStatusNaoAtendido As Variant, viPrioridade As Variant) As Long

'isola a preparacao da parte fixa do SELECT, serve apenas para diminuir o tamanho de GeracaoOP_SelecionarItens
Dim ret As Integer, lErro As Long

On Error GoTo Erro_GeracaoOP_SelecionarItens2

    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then gError 27409

    ret = Comando_BindVarInt(lComando, vNumIntItemPV)
    If (ret <> AD_SQL_SUCESSO) Then gError 27411

    ret = Comando_BindVarInt(lComando, vItensPedidoDeVendaProduto)
    If (ret <> AD_SQL_SUCESSO) Then gError 27411

    ret = Comando_BindVarInt(lComando, vItensPedidoDeVendaDataEntrega)
    If (ret <> AD_SQL_SUCESSO) Then gError 27412

    ret = Comando_BindVarInt(lComando, vItensPedidoDeVendaQuantFalta)
    If (ret <> AD_SQL_SUCESSO) Then gError 22983

    ret = Comando_BindVarInt(lComando, vPedidosDeVendaFilialEmpresa)
    If (ret <> AD_SQL_SUCESSO) Then gError 27413

    ret = Comando_BindVarInt(lComando, vPedidosDeVendaCodigo)
    If (ret <> AD_SQL_SUCESSO) Then gError 27414

    ret = Comando_BindVarInt(lComando, vPedidosDeVendaFilial)
    If (ret <> AD_SQL_SUCESSO) Then gError 27415

    ret = Comando_BindVarInt(lComando, vPedidosDeVendaDataEmissao)
    If (ret <> AD_SQL_SUCESSO) Then gError 27416

    ret = Comando_BindVarInt(lComando, vClientesNomeReduzido)
    If (ret <> AD_SQL_SUCESSO) Then gError 27417

    ret = Comando_BindVarInt(lComando, vProdutosDescricao)
    If (ret <> AD_SQL_SUCESSO) Then gError 27418

    ret = Comando_BindVarInt(lComando, vProdutosSiglaUM)
    If (ret <> AD_SQL_SUCESSO) Then gError 27410
    
    ret = Comando_BindVarInt(lComando, vProdutosClasseUM)
    If (ret <> AD_SQL_SUCESSO) Then gError 90764
    
    ret = Comando_BindVarInt(lComando, sItensPVSiglaUM)
    If (ret <> AD_SQL_SUCESSO) Then gError 90765
    
    ret = Comando_BindVarInt(lComando, viPrioridade)
    If (ret <> AD_SQL_SUCESSO) Then gError 90765

    ret = Comando_BindVarInt(lComando, vProdutosCompras)
    If (ret <> AD_SQL_SUCESSO) Then gError 19998

    ret = Comando_BindVarInt(lComando, vProdutosPCP)
    If (ret <> AD_SQL_SUCESSO) Then gError 19999

    ret = Comando_BindVarInt(lComando, vStatusNaoAtendido)
    If (ret <> AD_SQL_SUCESSO) Then gError 27419

    GeracaoOP_SelecionarItens2 = SUCESSO

    Exit Function

Erro_GeracaoOP_SelecionarItens2:

    GeracaoOP_SelecionarItens2 = gErr

    Select Case gErr

        Case 27409 To 27419, 22983, 90764, 90765, 19998, 19999
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149641)

    End Select

    Exit Function

End Function

Private Function GeracaoOP_SelecionarItens3(ByVal lComando As Long, ByVal objGeracaoOP As ClassGeracaoOP) As Long
'auxiliar a GeracaoOP_SelecionarItens
'complementa a passagem dos parametros que variam de acordo com a selecao do usuario e executa o SELECT p/obtencao dos itens

Dim lErro As Long, ret As Integer, vFilialEmpresa As Variant, vPVDe As Variant, vPVAte As Variant
Dim vClientesDe As Variant, vClientesAte As Variant, vEntregaDe As Variant, vEntregaAte As Variant
Dim vProdDe As Variant, vProdAte As Variant

On Error GoTo Erro_GeracaoOP_SelecionarItens3

    'Se entrar como empresa toda pode pegar itens de todas as filiais, senao pegar itens somente da filial corrente
    If giFilialEmpresa <> EMPRESA_TODA Then
        vFilialEmpresa = giFilialEmpresa
        ret = Comando_BindVarInt(lComando, vFilialEmpresa)
        If (ret <> AD_SQL_SUCESSO) Then Error 27420
    End If

    'Se pedido de venda inicial preenchido
    If (objGeracaoOP.lPVDe <> 0) Then
        vPVDe = objGeracaoOP.lPVDe
        ret = Comando_BindVarInt(lComando, vPVDe)
        If (ret <> AD_SQL_SUCESSO) Then Error 27421
    End If

    'Se pedido de venda final preenchido
    If (objGeracaoOP.lPVAte <> 0) Then
        vPVAte = objGeracaoOP.lPVAte
        ret = Comando_BindVarInt(lComando, vPVAte)
        If (ret <> AD_SQL_SUCESSO) Then Error 27422
    End If

    'Se cliente do pedido de venda inicial preenchido
    If (objGeracaoOP.lClientesDe <> 0) Then
        vClientesDe = objGeracaoOP.lClientesDe
        ret = Comando_BindVarInt(lComando, vClientesDe)
        If (ret <> AD_SQL_SUCESSO) Then Error 27423
    End If

    'Se cliente do pedido de venda final preenchido
    If (objGeracaoOP.lClientesAte <> 0) Then
        vClientesAte = objGeracaoOP.lClientesAte
        ret = Comando_BindVarInt(lComando, vClientesAte)
        If (ret <> AD_SQL_SUCESSO) Then Error 27424
    End If

    'Se data de entrega inicial do item preenchida
    If (objGeracaoOP.dtEntregaDe <> DATA_NULA) Then
        vEntregaDe = objGeracaoOP.dtEntregaDe
        ret = Comando_BindVarInt(lComando, vEntregaDe)
        If (ret <> AD_SQL_SUCESSO) Then Error 27425
    End If

    'Se data de entrega final do item preenchida
    If (objGeracaoOP.dtEntregaAte <> DATA_NULA) Then
        vEntregaAte = objGeracaoOP.dtEntregaAte
        ret = Comando_BindVarInt(lComando, vEntregaAte)
        If (ret <> AD_SQL_SUCESSO) Then Error 27426
    End If

    'se produto inicial preenchido
    If (objGeracaoOP.sProdDe <> "") Then
        vProdDe = objGeracaoOP.sProdDe
        ret = Comando_BindVarInt(lComando, vProdDe)
        If (ret <> AD_SQL_SUCESSO) Then Error 27427
    End If

    'se produto final preenchido
    If (objGeracaoOP.sProdAte <> "") Then
        vProdAte = objGeracaoOP.sProdAte
        ret = Comando_BindVarInt(lComando, vProdAte)
        If (ret <> AD_SQL_SUCESSO) Then Error 27428
    End If

    ret = Comando_ExecutarInt(lComando)
    If (ret <> AD_SQL_SUCESSO) Then Error 27430

    GeracaoOP_SelecionarItens3 = SUCESSO

    Exit Function

Erro_GeracaoOP_SelecionarItens3:

    GeracaoOP_SelecionarItens3 = Err

    Select Case Err
        
        Case 27420 To 27430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149642)

    End Select

    Exit Function

End Function

Private Function GeracaoOP_SelecionarItens4(ByVal lComando As Long, ByVal objGeracaoOP As ClassGeracaoOP, _
        lNumIntItemPV As Long, sItensPedidoDeVendaProduto As String, dtItensPedidoDeVendaDataEntrega As Date, _
        dItensPedidoDeVendaQuantFalta As Double, iPedidosDeVendaFilialEmpresa As Integer, _
        lPedidosDeVendaCodigo As Long, iPedidosDeVendaFilial As Integer, _
        dtPedidosDeVendaDataEmissao As Date, sClientesNomeReduzido As String, _
        sProdutosDescricao As String, sProdutosSiglaUM As String, iProdutosClasseUM As Integer, _
        sItensPVSiglaUM As String, iPrioridade As Integer, Optional bExibeSemDados As Boolean = True) As Integer
'auxiliar a GeracaoOP_SelecionarItens
'inclui os itens lidos na colecao

Dim lErro As Long, objGeracaoOPItem As ClassGeracaoOPItem, objItemOP As ClassItemOP
Dim objItemRomaneiGrade As ClassItemRomaneioGrade
Dim lComando2 As Long
Dim sProdutoGrade As String, dQuantidadeGrade As Double, dQuantCanceladaGrade As Double, dQuantReservadaGrade As Double, dQuantFaturadaGrade As Double, dQuantOPGrade As Double, dQuantSCGrade As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim objKit As New ClassKit
Dim objProduto As New ClassProduto, bComErro As Boolean

On Error GoTo Erro_GeracaoOP_SelecionarItens4

    lComando2 = Comando_Abrir
    If lComando2 = 0 Then Error 27431
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27431

    If bExibeSemDados Then
        If lErro = AD_SQL_SEM_DADOS Then Error 27432
    End If

    Do While lErro = AD_SQL_SUCESSO

        bComErro = False

        Set objGeracaoOPItem = New ClassGeracaoOPItem

        objGeracaoOPItem.dQtdeFalta = dItensPedidoDeVendaQuantFalta
        objGeracaoOPItem.dtEmissaoPV = dtPedidosDeVendaDataEmissao
        objGeracaoOPItem.dtEntregaItemPV = dtItensPedidoDeVendaDataEntrega
        objGeracaoOPItem.iFilialCliente = iPedidosDeVendaFilial
        objGeracaoOPItem.iPrioridade = iPrioridade
        
        Set objItemOP = New ClassItemOP

        objItemOP.iFilialPedido = iPedidosDeVendaFilialEmpresa
        objItemOP.lCodPedido = lPedidosDeVendaCodigo
        objItemOP.sProduto = sItensPedidoDeVendaProduto
        objItemOP.sSiglaUMEstoque = sProdutosSiglaUM
        objItemOP.iClasseUM = iProdutosClasseUM
        objItemOP.sSiglaUM = sItensPVSiglaUM
        
        Set objGeracaoOPItem.objItemOP = objItemOP

        objGeracaoOPItem.sClienteReduzido = sClientesNomeReduzido
        objGeracaoOPItem.sProdutoDescricao = sProdutosDescricao

        sProdutoGrade = String(STRING_PRODUTO, 0)
        lErro = Comando_Executar(lComando2, "SELECT Produto, Quantidade, QuantCancelada, QuantReservada, QuantFaturada, QuantOP, QuantSC FROM ItensPedidoDeVendaGrade WHERE NumIntItemPV = ? ORDER BY Produto", _
            sProdutoGrade, dQuantidadeGrade, dQuantCanceladaGrade, dQuantReservadaGrade, dQuantFaturadaGrade, dQuantOPGrade, dQuantSCGrade, lNumIntItemPV)
        If lErro <> AD_SQL_SUCESSO Then Error 27431
            
        lErro = Comando_BuscarProximo(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27431
        
        Do While lErro = AD_SQL_SUCESSO
        
            Set objItemRomaneiGrade = New ClassItemRomaneioGrade
            
            With objItemRomaneiGrade
                .sProduto = sProdutoGrade
                .dQuantCancelada = dQuantCanceladaGrade
                .dQuantFaturada = dQuantFaturadaGrade
                .dQuantidade = dQuantidadeGrade
                .dQuantOP = dQuantOPGrade
                .dQuantReservada = dQuantReservadaGrade
                .dQuantSC = dQuantSCGrade
            End With
            
            'Preneche o objProduto Filial para poder ler o Almoxarifado Padrao
            objProdutoFilial.sProduto = sProdutoGrade
            objProdutoFilial.iFilialEmpresa = giFilialEmpresa
            
            'Le o AlmoxarifadoPadrao
            lErro = CF("ProdutoFilial_Le", objProdutoFilial)
            If lErro <> SUCESSO And lErro <> 28261 Then Error 27431
            If lErro <> SUCESSO Then Error 27431
            
            objItemRomaneiGrade.iAlmoxarifado = objProdutoFilial.iAlmoxarifado
            
            objProduto.sCodigo = sProdutoGrade
    
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 27431
            If lErro <> SUCESSO Then Error 27431
            
            objItemRomaneiGrade.sUMEstoque = objProduto.sSiglaUMEstoque
        
            'Armazena o Produto Raiz do kit
            objKit.sProdutoRaiz = sProdutoGrade
            
            'Le as Versoes Ativas e a Padrao
            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then Error 27431
            If lErro <> SUCESSO Then bComErro = True
        
            objItemRomaneiGrade.sVersao = objKit.sVersao
            
            objGeracaoOPItem.objItemOP.colItensRomaneioGrade.Add objItemRomaneiGrade
        
            lErro = Comando_BuscarProximo(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27431
        
        Loop
        
        If bComErro = False Then objGeracaoOP.colItens.Add objGeracaoOPItem

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27433

    Loop

    Call Comando_Fechar(lComando2)
    
    GeracaoOP_SelecionarItens4 = SUCESSO

    Exit Function

Erro_GeracaoOP_SelecionarItens4:

    GeracaoOP_SelecionarItens4 = Err

    Select Case Err

        Case 27431, 27433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPV_GERACAO_OP", Err)

        Case 27432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_ITENSPV_GERACAO_OP", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149643)

    End Select

    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Function NFiscalEntrada_Lock_Gravacao(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim iIndice As Long
Dim alComando(0 To 5) As Long
Dim sNome As String
Dim sDescricao As String
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iFilialEmpresa As Integer
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim iFilialEmpresa1 As Integer

On Error GoTo Erro_NFiscalEntrada_Lock_Gravacao

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 35045
    Next

    sDescricao = String(STRING_NATUREZAOP_DESCRICAO, 0)

    lErro = Comando_ExecutarLockado(alComando(2), "SELECT Descricao FROM NaturezaOp WHERE Codigo = ? ", sDescricao, objNFiscal.sNaturezaOp)
    If lErro <> AD_SQL_SUCESSO Then Error 35664

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35665
    If lErro = AD_SQL_SEM_DADOS Then Error 35666
    
    lErro = Comando_LockShared(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then Error 35667

    sNome = String(STRING_FILIAL_FORNECEDOR_NOME, 0)

    'Lê a FilialFornecedor
    lErro = Comando_ExecutarLockado(alComando(0), "SELECT Nome FROM FiliaisFornecedores WHERE CodFornecedor = ? AND CodFilial = ? ", sNome, objNFiscal.lFornecedor, objNFiscal.iFilialForn)
    If lErro <> AD_SQL_SUCESSO Then Error 35046

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35047
    If lErro = AD_SQL_SEM_DADOS Then Error 35048 'Não encontrou --> Erro

    'Faz um "lock" na FilialFornecedor
    lErro = Comando_LockShared(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 35049

    'Se Nota for interna, lock shared na série
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 35676
    If lErro <> SUCESSO Then Error 35677

    If objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA Then
        
        If objNFiscal.iFilialEmpresa = 0 Then objNFiscal.iFilialEmpresa = giFilialEmpresa
        
        lErro = Comando_ExecutarLockado(alComando(3), "SELECT FilialEmpresa FROM Serie WHERE Serie =? AND FilialEmpresa = ? ", iFilialEmpresa, objNFiscal.sSerie, objNFiscal.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 35678

        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35679
        If lErro = AD_SQL_SEM_DADOS Then Error 35680
        
        lErro = Comando_LockShared(alComando(3))
        If lErro <> AD_SQL_SUCESSO Then Error 35681
    End If
        
    'Se a Transportadora estiver preenchida
    If objNFiscal.iCodTransportadora > 0 Then
        'Lê a Transportadora
        lErro = Comando_ExecutarLockado(alComando(1), "SELECT Codigo FROM Transportadoras WHERE Codigo = ?", iCodigo, objNFiscal.iCodTransportadora)
        If lErro <> AD_SQL_SUCESSO Then Error 35050

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 35051
        If lErro = AD_SQL_SEM_DADOS Then Error 35052 'Não encontrou --> Erro

        'Faz um "lock" na transportadora
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then Error 35053
    End If

    'Lock shared nos Produtos e Almoxarifados
    For Each objItemNF In objNFiscal.colItensNF
    
        objProduto.sCodigo = objItemNF.sProduto
        lErro = CF("Produto_Lock_Shared", alComando(5), objProduto)
        If lErro <> SUCESSO Then Error 35682
        
        objItemNF.iControleEstoque = objProduto.iControleEstoque
        objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
        objItemNF.sUMVenda = objProduto.sSiglaUMVenda
        objItemNF.iClasseUM = objProduto.iClasseUM
        objItemNF.iApropriacaoProd = objProduto.iApropriacaoCusto
        
        For Each objItemRomaneioGrade In objItemNF.colItensRomaneioGrade
            
            objProduto.sCodigo = objItemRomaneioGrade.sProduto
    
            'Faz lock do Produto
            lErro = CF("Produto_Lock", objProduto)
            If lErro <> SUCESSO Then Error 30899
        
            objItemRomaneioGrade.iControleEstoque = objProduto.iControleEstoque
        Next
        
        '??? está locando varias vezes o mesmo almoxarifado, fazer igual a NFiscal_Lock_Gravacao em fatgrava. Jones 14/12/99
        If objItemNF.iAlmoxarifado > 0 Then
        
            If objNFiscal.iFilialEmpresa = 0 Then objNFiscal.iFilialEmpresa = giFilialEmpresa
            
            iFilialEmpresa1 = objNFiscal.iFilialEmpresa
            lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa1)
            If lErro <> SUCESSO Then Error 30899
            
            lErro = Comando_ExecutarPos(alComando(3), "SELECT FilialEmpresa FROM Almoxarifado WHERE Codigo = ? AND FilialEmpresa = ? ", 0, iFilialEmpresa, objItemNF.iAlmoxarifado, iFilialEmpresa1)
            If lErro <> AD_SQL_SUCESSO Then Error 35794
            
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 35795
            If lErro = AD_SQL_SEM_DADOS Then Error 35796
        
            lErro = Comando_LockShared(alComando(3))
            If lErro <> AD_SQL_SUCESSO Then Error 35797
        End If
    Next

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalEntrada_Lock_Gravacao = SUCESSO
    
    Exit Function
    
Erro_NFiscalEntrada_Lock_Gravacao:

    NFiscalEntrada_Lock_Gravacao = Err
    
    Select Case Err
    
        Case 35045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 35046, 35047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES1", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 35048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 35049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISFORNECEDORES1", Err, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 35050, 35051
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORA", Err, objNFiscal.iCodTransportadora)

        Case 35052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objNFiscal.iCodTransportadora)

        Case 35053
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRANSPORTADORA", Err)

        Case 35664, 35665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP", Err, objNFiscal.sNaturezaOp)

        Case 35666
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, objNFiscal.sNaturezaOp)
        
        Case 35667
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NATUREZAOP", Err, objNFiscal.sNaturezaOp)

        Case 35676, 35682, 30899
        
        Case 35677
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)
        
        Case 35678, 35679
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE", Err)
            
        Case 35680
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, objNFiscal.sSerie)

        Case 35681
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERIE", Err)
        
        Case 35794, 35795
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO", Err, objItemNF.iAlmoxarifado)
        
        Case 35796
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objItemNF.iAlmoxarifado)
        
        Case 35797
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ALMOXARIFADO1", Err, objItemNF.iAlmoxarifado)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149644)

    End Select
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function MovEstoque_Le_NomeTela_NFiscal(ByVal objMovEstoque As ClassMovEstoque, sTela As String, ByVal objNFiscal As ClassNFiscal) As Long
'Esta função retorna o nome da tela, os dados da Nota Fiscal
'relacionada com o Movimento Estoque passado no parametro

Dim lErro As Long
Dim tNFiscal As typeNFiscal
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iTipo As Integer

On Error GoTo Erro_MovEstoque_Le_NomeTela_NFiscal
    
    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 52882
    Next
    
    sTela = String(STRING_NOME_TELA, 0)
    
    'Lê o Nome da Tela e os dados da Nfiscal relacionados com o Movimento de Estoque
    lErro = Comando_Executar(alComando(1), "SELECT NFiscalTodas.TipoNFiscal, TiposDocInfo.Tipo, TiposDocInfo.NomeTelaNFiscal, NFiscalTodas.NumIntDoc, NFiscalTodas.Status, NFiscalTodas.FilialEmpresa FROM MovimentoEstoque, NFiscalTodas, TiposDocInfo WHERE MovimentoEstoque.Codigo = ? AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.NumIntDocOrigem = NFiscalTodas.NumIntDoc AND NFiscalTodas.TipoNFiscal = TiposDocInfo.Codigo", tNFiscal.iTipoNFiscal, iTipo, sTela, tNFiscal.lNumIntDoc, tNFiscal.iStatus, tNFiscal.iFilialEmpresa, objMovEstoque.lCodigo, objMovEstoque.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 52883

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 52884
    
    objNFiscal.lNumIntDoc = tNFiscal.lNumIntDoc
    objNFiscal.iStatus = tNFiscal.iStatus
    objNFiscal.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objNFiscal.iTipoDocInfo = iTipo
    objNFiscal.iTipoNFiscal = tNFiscal.iTipoNFiscal
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    MovEstoque_Le_NomeTela_NFiscal = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Le_NomeTela_NFiscal:

    MovEstoque_Le_NomeTela_NFiscal = Err
    
    Select Case Err
        
        Case 52882
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 52883, 52884
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149645)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function MovEstoque_Le_NomeTela_ItemNFiscal(ByVal objMovEstoque As ClassMovEstoque, sTela As String, ByVal objNFiscal As ClassNFiscal) As Long
'Esta função retorna o nome da tela, os dados da Nota Fiscal dos Itens relacionado
'relacionada com o Movimento Estoque passado no parametro

Dim lErro As Long
Dim tNFiscal As typeNFiscal
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iTipo As Integer

On Error GoTo Erro_MovEstoque_Le_NomeTela_ItemNFiscal
    
    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 52885
    Next
    
    sTela = String(STRING_NOME_TELA, 0)
    
    'Lê o Nome da Tela e os dados da Nfiscal relacionados com o Movimento de Estoque
    lErro = Comando_Executar(alComando(1), "SELECT NFiscalTodas.TipoNFiscal, TiposDocInfo.Tipo, TiposDocInfo.NomeTelaNFiscal, NFiscalTodas.NumIntDoc, NFiscalTodas.Status, NFiscalTodas.FilialEmpresa, NFiscalTodas.NumRecebimento FROM MovimentoEstoque, NFiscalTodas, ItensNFiscalTodos, TiposDocInfo WHERE MovimentoEstoque.Codigo = ? AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.NumIntDocOrigem = ItensNFiscalTodos.NumIntDoc AND NFiscalTodas.NumIntDoc = ItensNFiscalTodos.NumIntNF AND NFiscalTodas.TipoNFiscal = TiposDocInfo.Codigo", tNFiscal.iTipoNFiscal, iTipo, sTela, tNFiscal.lNumIntDoc, tNFiscal.iStatus, tNFiscal.iFilialEmpresa, tNFiscal.lNumRecebimento, objMovEstoque.lCodigo, objMovEstoque.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 52886

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 52887
    
    objNFiscal.lNumIntDoc = tNFiscal.lNumIntDoc
    objNFiscal.iStatus = tNFiscal.iStatus
    objNFiscal.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objNFiscal.iTipoDocInfo = iTipo
    objNFiscal.iTipoNFiscal = tNFiscal.iTipoNFiscal
    objNFiscal.lNumRecebimento = tNFiscal.lNumRecebimento
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    MovEstoque_Le_NomeTela_ItemNFiscal = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Le_NomeTela_ItemNFiscal:

    MovEstoque_Le_NomeTela_ItemNFiscal = Err
    
    Select Case Err
        
        Case 52885
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 52886, 52887
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149646)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function OrdemDeProducao_Le_SemItens(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'carrega do bd a ordem de producao sem carregar os seus itens

Dim lErro As Long
Dim dtDataEmissao As Date
Dim iNumItens As Integer
Dim iNumItensBaixados As Integer
Dim lComando As Long

On Error GoTo Erro_OrdemDeProducao_Le_SemItens

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 34452

    'Leitura da tabela de Ordens de Produção
    lErro = Comando_Executar(lComando, "SELECT DataEmissao, NumItens, NumItensBaixados FROM OrdensDeProducao WHERE FilialEmpresa = ? AND Codigo = ? ", dtDataEmissao, iNumItens, iNumItensBaixados, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 34453

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 34454

    If lErro <> AD_SQL_SUCESSO Then Error 34455

    objOrdemDeProducao.dtDataEmissao = dtDataEmissao
    objOrdemDeProducao.iNumItens = iNumItens
    objOrdemDeProducao.iNumItensBaixados = iNumItensBaixados

   'Fechamento comando
    Call Comando_Fechar(lComando)

    OrdemDeProducao_Le_SemItens = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Le_SemItens:

    OrdemDeProducao_Le_SemItens = Err

    Select Case Err

        Case 34452
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 34453, 34454
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", Err)

        Case 34455 'Ordem de Producao nao encontrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149647)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OPBaixada_Le_SemItens(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'carrega do bd a ordem de producao baixada sem carregar os seus itens

Dim lErro As Long
Dim dtDataEmissao As Date
Dim iNumItens As Integer
Dim iNumItensBaixados As Integer
Dim lComando As Long

On Error GoTo Erro_OPBaixada_Le_SemItens

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 34456

    'Leitura da tabela de Ordens de Produção
    lErro = Comando_Executar(lComando, "SELECT DataEmissao, NumItens, NumItensBaixados FROM OrdensDeProducaoBaixadas WHERE FilialEmpresa = ? AND Codigo = ? ", dtDataEmissao, iNumItens, iNumItensBaixados, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 34457

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 34458

    If lErro <> AD_SQL_SUCESSO Then Error 34459

    objOrdemDeProducao.dtDataEmissao = dtDataEmissao
    objOrdemDeProducao.iNumItens = iNumItens
    objOrdemDeProducao.iNumItensBaixados = iNumItensBaixados

   'Fechamento comando
    Call Comando_Fechar(lComando)

    OPBaixada_Le_SemItens = SUCESSO

    Exit Function

Erro_OPBaixada_Le_SemItens:

    OPBaixada_Le_SemItens = Err

    Select Case Err

        Case 34456
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 34457, 34458
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", Err)

        Case 34459 'Ordem de Producao nao encontrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149648)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutosKit_Le_PrimeiroNivel_Est(ByVal objKit As ClassKit) As Long
'Preenche objKit.colComponentes a partir dos produtos Inventariados e da versão passados em objKit e Nivel = KIT_NIVEL_RAIZ + 1

Dim lErro As Long
Dim lComando As Long
Dim tProdutoKit As typeProdutoKit
Dim objProdutoKit As ClassProdutoKit
Dim sComando_SQL As String, sProdutoDescricao As String
Dim iClasseUM As Integer

On Error GoTo Erro_ProdutosKit_Le_PrimeiroNivel_Est

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55480

    tProdutoKit.sProdutoRaiz = String(STRING_PRODUTO, 0)
    tProdutoKit.sProduto = String(STRING_PRODUTO, 0)
    tProdutoKit.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)
    tProdutoKit.sVersao = String(STRING_KIT_VERSAO, 0)
    sProdutoDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

    sComando_SQL = "SELECT ProdutoKit.ProdutoRaiz, ProdutoKit.Versao, ProdutoKit.Nivel, ProdutoKit.Seq, ProdutoKit.Produto, Produtos.Descricao, Produtos.ClasseUM, ProdutoKit.SeqPai, ProdutoKit.Quantidade, ProdutoKit.UnidadeMed, ProdutoKit.Composicao FROM Produtos, ProdutoKit WHERE (Produtos.ControleEstoque = ? Or Produtos.ControleEstoque = ?) And  Produtos.Codigo = ProdutoKit.Produto AND ProdutoKit.ProdutoRaiz = ? AND ProdutoKit.Versao = ? AND ProdutoKit.Nivel = ?"
    lErro = Comando_Executar(lComando, sComando_SQL, tProdutoKit.sProdutoRaiz, tProdutoKit.sVersao, tProdutoKit.iNivel, tProdutoKit.iSeq, tProdutoKit.sProduto, sProdutoDescricao, iClasseUM, tProdutoKit.iSeqPai, tProdutoKit.dQuantidade, tProdutoKit.sUnidadeMed, tProdutoKit.iComposicao, PRODUTO_CONTROLE_RESERVA, PRODUTO_CONTROLE_ESTOQUE, objKit.sProdutoRaiz, objKit.sVersao, KIT_NIVEL_RAIZ + 1)
    If lErro <> AD_SQL_SUCESSO Then Error 55481

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55482

    If lErro = AD_SQL_SEM_DADOS Then Error 55484

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objProdutoKit = New ClassProdutoKit

        objProdutoKit.dQuantidade = tProdutoKit.dQuantidade
        objProdutoKit.iComposicao = tProdutoKit.iComposicao
        objProdutoKit.iNivel = tProdutoKit.iNivel
        objProdutoKit.iSeq = tProdutoKit.iSeq
        objProdutoKit.iSeqPai = tProdutoKit.iSeqPai
        objProdutoKit.sProduto = tProdutoKit.sProduto
        objProdutoKit.sProdutoDesc = sProdutoDescricao
        objProdutoKit.iClasseUM = iClasseUM
        objProdutoKit.sUnidadeMed = tProdutoKit.sUnidadeMed
        objProdutoKit.sProdutoRaiz = tProdutoKit.sProdutoRaiz
        objProdutoKit.sVersao = tProdutoKit.sVersao

        objKit.colComponentes.Add objProdutoKit

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55483

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    ProdutosKit_Le_PrimeiroNivel_Est = SUCESSO

    Exit Function

Erro_ProdutosKit_Le_PrimeiroNivel_Est:

    ProdutosKit_Le_PrimeiroNivel_Est = Err

    Select Case Err

        Case 55480
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55481, 55482, 55483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", Err)

        Case 55484

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149649)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OrdemDeProducao_TestaExistencia(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Pesquisa no bd a ordem de producao a partir de FilialEmpresa, Codigo e Data de Emissão

Dim lErro As Long
Dim iNumItens As Integer
Dim lComando As Long

On Error GoTo Erro_OrdemDeProducao_TestaExistencia

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 57718

    'Leitura da tabela de Ordens de Produção
    lErro = Comando_Executar(lComando, "SELECT NumItens FROM OrdensDeProducao WHERE FilialEmpresa = ? AND Codigo = ? And DataEmissao = ?", iNumItens, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo, objOrdemDeProducao.dtDataEmissao)
    If lErro <> AD_SQL_SUCESSO Then Error 57719

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57720
    
    If lErro <> AD_SQL_SUCESSO Then Error 57721

   'Fechamento comando
    Call Comando_Fechar(lComando)

    OrdemDeProducao_TestaExistencia = SUCESSO

    Exit Function

Erro_OrdemDeProducao_TestaExistencia:

    OrdemDeProducao_TestaExistencia = Err

    Select Case Err

        Case 57718
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 57719, 57720
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", Err)

        Case 57721 'Ordem de Producao não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149650)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFiscalRec_Interna_TestaExistencia(ByVal objNFiscal As ClassNFiscal) As Long
'Verifica se foi cadastrada uma Nota Fiscal no BD a partir de FilialEmpresa, Serie, Número, Tipo, Data Entrada

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_NFiscalRec_Interna_TestaExistencia

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 57701

    'Pesquisa a Nota Fiscal com os dados passados
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM NFiscal WHERE FilialEmpresa = ? AND Serie = ? AND NumNotaFiscal = ? AND TipoNFiscal = ? AND Status <> ? And DataEntrada = ?", _
        lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iTipoNFiscal, STATUS_EXCLUIDO, objNFiscal.dtDataEntrada)
    If lErro <> AD_SQL_SUCESSO Then Error 57702

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57703
    
    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 57704

    'Fechar o Comando
    Call Comando_Fechar(lComando)

    NFiscalRec_Interna_TestaExistencia = SUCESSO

    Exit Function

Erro_NFiscalRec_Interna_TestaExistencia:

    NFiscalRec_Interna_TestaExistencia = Err

    Select Case Err

        Case 57701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 57702, 57703
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL2", Err, objNFiscal.iTipoNFiscal, objNFiscal.lCliente, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 57704 'Nota Fiscal não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149651)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function NFiscal_Interna_TestaExistencia(ByVal objNFiscal As ClassNFiscal) As Long
'Verifica se foi cadastrada uma Nota Fiscal no BD a partir de FilialEmpresa, Serie, Número, Tipo, Data Emissão

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_NFiscal_Interna_TestaExistencia

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 57674

    'Pesquisa a Nota Fiscal com os dados passados
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM NFiscal WHERE FilialEmpresa = ? AND Serie = ? AND NumNotaFiscal = ? AND TipoNFiscal = ? AND Status <> ? And DataEmissao = ?", _
        lNumIntDoc, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iTipoNFiscal, STATUS_EXCLUIDO, objNFiscal.dtDataEmissao)
    If lErro <> AD_SQL_SUCESSO Then Error 57675

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 57676
    
    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 57677

    'Fechar o Comando
    Call Comando_Fechar(lComando)

    NFiscal_Interna_TestaExistencia = SUCESSO

    Exit Function

Erro_NFiscal_Interna_TestaExistencia:

    NFiscal_Interna_TestaExistencia = Err

    Select Case Err

        Case 57674
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 57675, 57676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL2", Err, objNFiscal.iTipoNFiscal, objNFiscal.lCliente, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 57677 'Nota Fiscal não cadastrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149652)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Recebimento_Testa_Baixado(ByVal objNFiscal As ClassNFiscal) As Long
'Testa se o Recebimento passado está Baixado se estiver dá mensagem

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long

On Error GoTo Erro_Recebimento_Testa_Baixado

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61046
    
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM NFiscal WHERE Status = ? AND FilialEmpresa = ? AND NumRecebimento =?", _
    lNumIntDoc, STATUS_BAIXADO, giFilialEmpresa, objNFiscal.lNumRecebimento)
    If lErro <> AD_SQL_SUCESSO Then Error 61047
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61048
    
    'Recebimento Baixado
    If lErro = SUCESSO Then Error 61049
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Recebimento_Testa_Baixado = SUCESSO
    
    Exit Function
    
Erro_Recebimento_Testa_Baixado:
    
    Recebimento_Testa_Baixado = Err
    
    Select Case Err
        
        Case 61046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 61047, 61048
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)
    
        Case 61049
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_BAIXADO", Err, objNFiscal.lNumRecebimento)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149653)

    End Select

    Exit Function

End Function

Function RecebimentoF_Verifica_Existencia(ByVal objNFiscal As ClassNFiscal) As Long
'Testa se existe outro recebimento com os memos dados de Número de Nota

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long, lNumRecebimento As Long
Dim dtDataEntrada As Date
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_RecebimentoF_Verifica_Existencia
        
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61085
        
    'Lê o Recebimento Com os Dados Iguais
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumRecebimento FROM NFiscal WHERE FilialEmpresa = ? AND Fornecedor = ? AND FilialForn = ? AND DataEntrada = ? AND TipoNFiscal = ? AND Serie = ? AND NumNotaFiscal = ?", lNumIntDoc, lNumRecebimento, objNFiscal.iFilialEmpresa, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
    If lErro <> AD_SQL_SUCESSO Then Error 61086

    'Le o Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61087
    
    'Se Encontrar --> Erro
    If lErro = AD_SQL_SUCESSO Then Error 61088
        
    'Verifica a se há um Recebimento emitido dentro de um prazo -->
    '--> menor que o período de emissão
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEntrada FROM NFiscal WHERE FilialEmpresa = ? AND Fornecedor = ? AND FilialForn = ? AND TipoNFiscal = ? AND Serie = ? AND NumNotaFiscal = ? AND DataEntrada > ? AND DataEntrada < ?", lNumIntDoc, dtDataEntrada, objNFiscal.iFilialEmpresa, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEntrada - PERIODO_EMISSAO, objNFiscal.dtDataEntrada + PERIODO_EMISSAO)
    If lErro <> AD_SQL_SUCESSO Then Error 61119

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61120

    'Se encontrar
    If lErro <> AD_SQL_SEM_DADOS Then

        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_RECEBMATERIALF_MESMO_NUMERO", objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, dtDataEntrada)
        'Se resposta for, Sai da Rotina
        If vbMsg = vbNo Then Error 61121

    End If
    'Fecha o comando
    Call Comando_Fechar(lComando)

    RecebimentoF_Verifica_Existencia = SUCESSO
    
    Exit Function
    
Erro_RecebimentoF_Verifica_Existencia:
    
    RecebimentoF_Verifica_Existencia = Err
    
    Select Case Err
    
        Case 61085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 61086, 61087, 61119, 61120
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)
    
        Case 61088
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NFEXT_CADASTRADO_FORNECEDOR", Err, lNumRecebimento, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada)
    
        Case 61121 'Já foi Tratado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149654)

    End Select

End Function

Function RecebimentoC_Verifica_Existencia(ByVal objNFiscal As ClassNFiscal) As Long
'Testa se existe outro recebimento com os memos dados de Número de Nota

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long, lNumRecebimento As Long
Dim dtDataEntrada As Date
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_RecebimentoC_Verifica_Existencia
        
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61111
        
    'Lê o Recebimento Com os Dados Iguais
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumRecebimento FROM NFiscal WHERE FilialEmpresa = ? AND Cliente = ? AND FilialCli = ? AND DataEntrada = ? AND TipoNFiscal = ? AND Serie = ? AND NumNotaFiscal = ?", lNumIntDoc, lNumRecebimento, objNFiscal.iFilialEmpresa, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEntrada, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
    If lErro <> AD_SQL_SUCESSO Then Error 61112

    'Le o Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61113
    
    'Se Encontrar --> Erro
    If lErro = AD_SQL_SUCESSO Then Error 61114
    
    'Verifica a se há uma Nota Fiscal emitida dentro de um prazo -->
    '--> menor que o período de emissão
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEntrada FROM NFiscal WHERE FilialEmpresa = ? AND Cliente = ? AND FilialCli = ? AND TipoNFiscal = ? AND Serie = ? AND NumNotaFiscal = ? AND DataEntrada > ? AND DataEntrada < ?", lNumIntDoc, dtDataEntrada, objNFiscal.iFilialEmpresa, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEntrada - PERIODO_EMISSAO, objNFiscal.dtDataEntrada + PERIODO_EMISSAO)
    If lErro <> AD_SQL_SUCESSO Then Error 61116

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61117
    
    'Se encontrar
    If lErro <> AD_SQL_SEM_DADOS Then
        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_RECEBMATERIALC_MESMO_NUMERO", objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, dtDataEntrada)
        If vbMsg = vbNo Then Error 61118 'Se resposta for nao, Sai da Rotina
    End If
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    RecebimentoC_Verifica_Existencia = SUCESSO
    
    Exit Function
    
Erro_RecebimentoC_Verifica_Existencia:
    
    RecebimentoC_Verifica_Existencia = Err
    
    Select Case Err
    
        Case 61111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 61112, 61113, 61116, 61117
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)
    
        Case 61114
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NFEXT_CADASTRADO_CLIENTE", Err, lNumRecebimento, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEntrada)
        
        Case 61118 'já foi Tratado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 149655)

    End Select

End Function

'Copiada de RotinasEST - ClassSELECT
'INCLUINDO CAMPO FilialPedido e Observacao
Function NFiscal_Le_Recebimento(ByVal objNFiscal As ClassNFiscal) As Long
'Le a Nota de Recebimento partir do Código do Recebimento

Dim lErro As Long
Dim tNFiscal As typeNFiscal
Dim lComando As Long

On Error GoTo Erro_NFiscal_Le_Recebimento

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61047
    
    With tNFiscal

        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
        .sSerie = String(STRING_SERIE, 0)
        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
        .sVolumeNumero = String(STRING_NFISCAL_VOLUME_NUMERO, 0)
        .sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0)
        
        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_BUFFER_MAX_TEXTO, 0)
        '.lVolumeMarca = String(STRING_BUFFER_MAX_TEXTO, 0)
        
        'Pesquisa a Nota Fiscal com os dados passados
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, Fornecedor, FilialForn, Cliente, FilialCli, FilialEntrega, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco , NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,VolumeNumero,FreteRespons, TipoNFiscal, Serie, NumNotaFiscal, FilialPedido, Observacao, CodTranspRedesp, DetPagFrete FROM NFiscal WHERE FilialEmpresa = ? AND NumRecebimento = ?  AND  Status <> ? AND Status <> ? AND (TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ? OR TipoNFiscal = ?) ORDER BY DataEntrada DESC", .lNumIntDoc, .iStatus, .lFornecedor, .iFilialForn, .lCliente, .iFilialCli, .iFilialEntrega, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, _
            .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .sVolumeNumero, .iFreteRespons, .iTipoNFiscal, .sSerie, .lNumNotaFiscal, .iFilialPedido, .sObservacao, .iCodTranspRedesp, .iDetPagFrete, giFilialEmpresa, objNFiscal.lNumRecebimento, STATUS_EXCLUIDO, STATUS_CANCELADO, DOCINFO_NRFF, DOCINFO_NRFP, DOCINFO_NRCC, DOCINFO_NRCP, DOCINFO_NRFFCO, DOCINFO_NRFPCO)
        If lErro <> AD_SQL_SUCESSO Then Error 61048
        
    End With
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61049
    
    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 61050

    With tNFiscal
    
        'Preenche objNFiscal com os dados lidos do BD
        objNFiscal.lNumIntDoc = .lNumIntDoc
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dtDataEmissao = .dtDataEmissao
        objNFiscal.dtDataEntrada = .dtDataEntrada
        objNFiscal.dtDataSaida = .dtDataSaida
        objNFiscal.dtDataVencimento = .dtDataVencimento
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.dValorFrete = .dValorFrete
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorProdutos = .dValorProdutos
        objNFiscal.dValorSeguro = .dValorSeguro
        objNFiscal.dValorTotal = .dValorTotal
        objNFiscal.iCanal = .iCanal
        objNFiscal.iClasseDocCPR = .iClasseDocCPR
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscal.iDetPagFrete = .iDetPagFrete
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        objNFiscal.iFilialEntrega = .iFilialEntrega
        objNFiscal.iStatus = .iStatus
        objNFiscal.iTabelaPreco = .iTabelaPreco
        objNFiscal.lVolumeQuant = .lVolumeQuant
        objNFiscal.lFornecedor = .lFornecedor
        objNFiscal.iFilialForn = .iFilialForn
        objNFiscal.lCliente = .lCliente
        objNFiscal.iFilialCli = .iFilialCli
        objNFiscal.iFilialPedido = .iFilialPedido
        objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscal.lNumIntTrib = .lNumIntTrib
        objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.sPlaca = .sPlaca
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.lVolumeEspecie = .lVolumeEspecie
        objNFiscal.lVolumeMarca = .lVolumeMarca
        objNFiscal.iFreteRespons = .iFreteRespons
        objNFiscal.iTipoNFiscal = .iTipoNFiscal
        objNFiscal.sSerie = .sSerie
        objNFiscal.lNumNotaFiscal = .lNumNotaFiscal
        
    End With

    'Fechar o Comando
    Call Comando_Fechar(lComando)
    
    NFiscal_Le_Recebimento = SUCESSO
    
    Exit Function

Erro_NFiscal_Le_Recebimento:

    NFiscal_Le_Recebimento = Err
    
    Select Case Err
        
        Case 61047
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 61048, 61049
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", Err)
        
        Case 61050 'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149656)
     
    End Select
     
    'Fechar o Comando
    Call Comando_Fechar(lComando)
     
    Exit Function

End Function

'??? deveria estar em grava e nao em select. Jones 16/01/2002
Function NFiscalEntradaCancela_Contabilidade(ByVal objNFiscal As ClassNFiscal, ByVal objTipoDocInfo As ClassTipoDocInfo, ByVal dtDataCancelamento As Date) As Long
'Gera um lançamento extorno para a Nota Fiscal ou se ainda não foi contabilizado Exclui

Dim lErro As Long
Dim objContabil As New ClassContabil
Dim objTelaAux As Object

On Error GoTo Erro_NFiscalEntradaCancela_Contabilidade
    
    lErro = objContabil.Contabil_Inicializa_Contabilidade1(objTelaAux, MODULO_ESTOQUE, objTipoDocInfo.sNomeTelaNFiscal, objTipoDocInfo.iSubTipoContabil)
    If lErro <> SUCESSO Then gError 64326
    
    lErro = objContabil.Contabil_Exclui(objNFiscal.lNumIntDoc, , , , dtDataCancelamento)
    If lErro <> SUCESSO Then gError 64327
    
    NFiscalEntradaCancela_Contabilidade = SUCESSO
    
    Exit Function
    
Erro_NFiscalEntradaCancela_Contabilidade:

    NFiscalEntradaCancela_Contabilidade = gErr
    
    Select Case gErr
        
        Case 64326, 64327, 79400
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149657)

    End Select
    
    Exit Function

End Function

Function NFiscalInternaEntrada_Le_Numero(ByVal objNFiscalOriginal As ClassNFiscal) As Long
'Pesquisa nas Tabelas [NFiscal e TiposDocInfo] UNION [NotaFiscalBaixadas e TiposDocInfo]
'Busca pela Serie e Número passados
'Filtros : TiposDocInfo.Tipo=DOCINFO_NF_INT_ENTRADA
'          NFiscal.Status<>STATUS_EXCLUIDO
'Ordena por DataEmissao, DECENDING
'Lê todos os Dados da Tabela NFiscal

Dim lErro As Long
Dim lComando As Long
Dim sComandoSQL As String
Dim tNFiscal As typeNFiscal

On Error GoTo Erro_NFiscalInternaEntrada_Le_Numero

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62141

    With tNFiscal

        .sMensagemNota = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNaturezaOp = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sPlaca = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sPlacaUF = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSerie = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sNumPedidoTerc = String(STRING_BUFFER_MAX_TEXTO, 0)
        
        'Removido por Luiz Nogueira em 21/08/03
        '.lVolumeEspecie = String(STRING_BUFFER_MAX_TEXTO, 0)
        '.lVolumeMarca = String(STRING_BUFFER_MAX_TEXTO, 0)
        
        sComandoSQL = "SELECT NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal, CodTranspRedesp, DetPagFrete FROM NFiscal, TiposDocInfo WHERE FilialEmpresa = ? AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.NumNotaFiscal = ? AND NFiscal.Serie = ? AND TiposDocInfo.Tipo = ? AND NFiscal.Status <> ? ORDER BY DataEmissao DESC "

        lErro = Comando_Executar(lComando, sComandoSQL, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .iCodTranspRedesp, .iDetPagFrete, _
            objNFiscalOriginal.iFilialEmpresa, objNFiscalOriginal.lNumNotaFiscal, objNFiscalOriginal.sSerie, DOCINFO_NF_INT_ENTRADA, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then Error 62142

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62143

        If lErro = AD_SQL_SEM_DADOS Then Error 62144

        objNFiscalOriginal.dPesoBruto = .dPesoBruto
        objNFiscalOriginal.dPesoLiq = .dPesoLiq
        objNFiscalOriginal.dtDataEmissao = .dtDataEmissao
        objNFiscalOriginal.dtDataEntrada = .dtDataEntrada
        objNFiscalOriginal.dtDataSaida = .dtDataSaida
        objNFiscalOriginal.dtDataVencimento = .dtDataVencimento
        objNFiscalOriginal.dValorDesconto = .dValorDesconto
        objNFiscalOriginal.dValorFrete = .dValorFrete
        objNFiscalOriginal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscalOriginal.dValorProdutos = .dValorProdutos
        objNFiscalOriginal.dValorSeguro = .dValorSeguro
        objNFiscalOriginal.dValorTotal = .dValorTotal
        objNFiscalOriginal.iCanal = .iCanal
        objNFiscalOriginal.iClasseDocCPR = .iClasseDocCPR
        objNFiscalOriginal.iCodTransportadora = .iCodTransportadora
        objNFiscalOriginal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscalOriginal.iDetPagFrete = .iDetPagFrete
        objNFiscalOriginal.iFilialCli = .iFilialCli
        objNFiscalOriginal.iFilialEntrega = .iFilialEntrega
        objNFiscalOriginal.iFilialForn = .iFilialForn
        objNFiscalOriginal.iStatus = .iStatus
        objNFiscalOriginal.iTabelaPreco = .iTabelaPreco
        objNFiscalOriginal.iTipoNFiscal = .iTipoNFiscal
        objNFiscalOriginal.lVolumeQuant = .lVolumeQuant
        objNFiscalOriginal.lCliente = .lCliente
        objNFiscalOriginal.lFornecedor = .lFornecedor
        objNFiscalOriginal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscalOriginal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscalOriginal.lNumIntTrib = .lNumIntTrib
        objNFiscalOriginal.lNumNotaFiscal = .lNumNotaFiscal
        objNFiscalOriginal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscalOriginal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscalOriginal.sMensagemNota = .sMensagemNota
        objNFiscalOriginal.sNaturezaOp = .sNaturezaOp
        objNFiscalOriginal.sPlaca = .sPlaca
        objNFiscalOriginal.sPlacaUF = .sPlacaUF
        objNFiscalOriginal.sSerie = .sSerie
        objNFiscalOriginal.lVolumeEspecie = .lVolumeEspecie
        objNFiscalOriginal.lVolumeMarca = .lVolumeMarca
        objNFiscalOriginal.lNumIntDoc = .lNumIntDoc
        objNFiscalOriginal.iFilialEmpresa = .iFilialEmpresa

    End With

    Call Comando_Fechar(lComando)

    NFiscalInternaEntrada_Le_Numero = SUCESSO

    Exit Function

Erro_NFiscalInternaEntrada_Le_Numero:

    NFiscalInternaEntrada_Le_Numero = Err

    Select Case Err

        Case 62141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 62142, 62143
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", Err)

        Case 62144 'Tratado pela rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149658)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function OPGeradora_Le_OPGerada(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim lComando As Long
Dim sCodigo As String

On Error GoTo Erro_OPGeradora_Le_OPGerada
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62634
    
    sCodigo = String(STRING_OPCODIGO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM OrdensDeProducao WHERE OPGeradora = ?", sCodigo, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 62635
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62636
    If lErro <> AD_SQL_SUCESSO Then Error 62637

    objOrdemDeProducao.sCodOPGerar = sCodigo
    
    Call Comando_Fechar(lComando)
    
    OPGeradora_Le_OPGerada = SUCESSO
    
    Exit Function

Erro_OPGeradora_Le_OPGerada:

    OPGeradora_Le_OPGerada = Err
    
    Select Case Err
    
        Case 62634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 62635, 62636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAO", Err)
        
        Case 62637
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149659)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ItensPC_Le_Codigo(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Lê os itens de um Pedido de Compra a partir do código e de FilialEmpresa do Pedido de Compras

Dim lErro As Long
Dim tItemPedido As typeItemPedCompra
Dim lComando As Long
Dim lTransacao As Long
Dim objItemPC As ClassItemPedCompra
Dim dCotacaoMoeda As Double

On Error GoTo Erro_ItensPC_Le_Codigo

    Set objPedidoCompra.colItens = New Collection

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 65837

    tItemPedido.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedido.sProduto = String(STRING_PRODUTO, 0)
    tItemPedido.sObservacao = String(STRING_OBSERVACAO_OBSERVACAO, 0) 'OK o tamanho correto da string é 255
    tItemPedido.sUM = String(STRING_UM_SIGLA, 0)

    'Lê registros da tabela de ItensPedCompra
    lErro = Comando_Executar(lComando, "SELECT PedidoCompra.CotacaoMoeda, ItensPedCompra.NumIntDoc, PedCompra, DataLimite, Produto, DescProduto, Quantidade, QuantRecebida, QuantRecebimento, UM, PrecoUnitario, ItensPedCompra.ValorDesconto, TipoOrigem, NumIntOrigem, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, ItensPedCompra.Status, ItensPedCompra.Observacao, ItensPedCompra.ValorIPI, ItensPedCompra.AliquotaIPI, ItensPedCompra.AliquotaICMS, PedidoCompra.Moeda, PedidoCompra.Taxa FROM ItensPedCompra, PedidoCompra WHERE PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND PedidoCompra.Codigo = ? AND PedidoCompra.FilialEmpresa = ? Order By ItensPedCompra.Produto", dCotacaoMoeda, tItemPedido.lNumIntDoc, tItemPedido.lPedCompra, _
                                                                tItemPedido.dtDataLimite, tItemPedido.sProduto, tItemPedido.sDescProduto, tItemPedido.dQuantidade, tItemPedido.dQuantRecebida, tItemPedido.dQuantRecebimento, tItemPedido.sUM, tItemPedido.dPrecoUnitario, tItemPedido.dValorDesconto, tItemPedido.iTipoOrigem, tItemPedido.lNumIntOrigem, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRebebForaFaixa, tItemPedido.iStatus, tItemPedido.lObservacao, tItemPedido.dValorIPI, tItemPedido.dAliquotaIPI, tItemPedido.dAliquotaICMS, tItemPedido.iMoeda, tItemPedido.dTaxa, objPedidoCompra.lCodigo, objPedidoCompra.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 65838

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65839

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 65840

    Do While lErro = AD_SQL_SUCESSO

        Set objItemPC = New ClassItemPedCompra

        'preenchendo objitem
        objItemPC.dAliquotaICMS = tItemPedido.dAliquotaICMS
        objItemPC.dAliquotaIPI = tItemPedido.dAliquotaIPI
        objItemPC.dPercentMaisReceb = tItemPedido.dPercentMaisReceb
        objItemPC.dPercentMenosReceb = tItemPedido.dPercentMenosReceb
        objItemPC.dPrecoUnitario = tItemPedido.dPrecoUnitario
        objItemPC.dQuantidade = tItemPedido.dQuantidade
        objItemPC.dQuantRecebida = tItemPedido.dQuantRecebida
        objItemPC.dQuantRecebimento = tItemPedido.dQuantRecebimento
        objItemPC.dtDataLimite = tItemPedido.dtDataLimite
        objItemPC.dValorDesconto = tItemPedido.dValorDesconto
        objItemPC.dValorIPI = tItemPedido.dValorIPI
        objItemPC.iRebebForaFaixa = tItemPedido.iRebebForaFaixa
        objItemPC.iStatus = tItemPedido.iStatus
        objItemPC.iTipoOrigem = tItemPedido.iTipoOrigem
        objItemPC.lNumIntDoc = tItemPedido.lNumIntDoc
        objItemPC.lNumIntOrigem = tItemPedido.lNumIntOrigem
        objItemPC.lObservacao = tItemPedido.lObservacao
        objItemPC.sDescProduto = tItemPedido.sDescProduto
        objItemPC.sObservacao = tItemPedido.sObservacao
        objItemPC.sProduto = tItemPedido.sProduto
        objItemPC.sUM = tItemPedido.sUM
        objItemPC.iMoeda = tItemPedido.iMoeda
        objItemPC.dTaxa = tItemPedido.dTaxa
        If objItemPC.dTaxa = 0 Then objItemPC.dTaxa = dCotacaoMoeda
        
        'Insere em colItens os valores de objItemPC
        objPedidoCompra.colItens.Add objItemPC

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65841
    
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensPC_Le_Codigo = SUCESSO

    Exit Function

Erro_ItensPC_Le_Codigo:

    ItensPC_Le_Codigo = gErr

    Select Case gErr

        Case 65837
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65838, 65839, 65841
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)

        Case 65840
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_AUSENCIA_ITENS", gErr, objPedidoCompra.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149660)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensPedCompra_Le_NFiscalReceb(ByVal objNFiscal As ClassNFiscal, ByVal colItemPedCompraInfo As Collection) As Long
'Lê os itens de Pedido de compras não baixados relacionados a um Recebimento e os devolve em colItemPedCompraInfo

Dim lErro As Long
Dim lComando As Long
Dim objItemPedCompraInfo As ClassItemPedCompraInfo
Dim tItemPedCompraInfo As typeItemPedidoCompraInfo
Dim dQuantNFPC As Double
Dim dTaxaItemNFPC As Double
Dim dPrecoUnitario As Double
Dim dCotacaoMoeda As Double

On Error GoTo Erro_ItensPedCompra_Le_NFiscalReceb

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66138
        
    tItemPedCompraInfo.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedCompraInfo.sProduto = String(STRING_PRODUTO, 0)
    tItemPedCompraInfo.sUM = String(STRING_UM_SIGLA, 0)
    
    'Pesquisa os Itens de Pedidos de Compras relacionados a Nota Fiscal a partir do NumIntDoc de NFiscal
    lErro = Comando_Executar(lComando, "SELECT DISTINCT PedidoCompra.CotacaoMoeda, ItensPedCompra.NumIntDoc, PedidoCompra.Codigo, ItensPedCompra.Produto, ItensPedCompra.DescProduto, ItensPedCompra.Quantidade, ItensPedCompra.QuantRecebida, ItensPedCompra.QuantRecebimento, ItensPedCompra.UM, ItensPedCompra.AliquotaICMS, ItensPedCompra.AliquotaIPI, ItemNFItemPC.Quantidade, PedidoCompra.Moeda, PedidoCompra.Taxa, ItemNFItemPC.Taxa, ItensPedCompra.PrecoUnitario FROM PedidoCompra, ItensPedCompra, ItemNFItemPC, ItensNFiscal, NFiscal WHERE PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItensNFiscal.NumIntDoc = ItemNFItemPC.ItemNFiscal AND ItensPedCompra.NumIntDoc = ItemNFItemPC.ItemPedCompra AND NFiscal.NumIntDoc = ? ORDER BY PedidoCompra.Codigo, ItensPedCompra.Produto", _
       dCotacaoMoeda, tItemPedCompraInfo.lNumIntDoc, tItemPedCompraInfo.lPedCompra, tItemPedCompraInfo.sProduto, tItemPedCompraInfo.sDescProduto, tItemPedCompraInfo.dQuantidade, tItemPedCompraInfo.dQuantRecebida, tItemPedCompraInfo.dQuantRecebimento, tItemPedCompraInfo.sUM, tItemPedCompraInfo.dAliquotaICMS, tItemPedCompraInfo.dAliquotaIPI, dQuantNFPC, tItemPedCompraInfo.iMoeda, tItemPedCompraInfo.dTaxa, dTaxaItemNFPC, dPrecoUnitario, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 66139

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66140

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66141
        
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objItemPedCompraInfo = New ClassItemPedCompraInfo

        objItemPedCompraInfo.lNumIntDoc = tItemPedCompraInfo.lNumIntDoc
        objItemPedCompraInfo.lPedCompra = tItemPedCompraInfo.lPedCompra
        objItemPedCompraInfo.sProduto = tItemPedCompraInfo.sProduto
        objItemPedCompraInfo.sDescProduto = tItemPedCompraInfo.sDescProduto
        objItemPedCompraInfo.sUM = tItemPedCompraInfo.sUM
        objItemPedCompraInfo.dQuantReceber = dQuantNFPC + (tItemPedCompraInfo.dQuantidade - tItemPedCompraInfo.dQuantRecebimento - tItemPedCompraInfo.dQuantRecebida)
        objItemPedCompraInfo.dQuantRecebida = dQuantNFPC
        objItemPedCompraInfo.dAliquotaICMS = tItemPedCompraInfo.dAliquotaICMS
        objItemPedCompraInfo.dAliquotaIPI = tItemPedCompraInfo.dAliquotaIPI
        objItemPedCompraInfo.dTaxa = IIf(dTaxaItemNFPC = 0, tItemPedCompraInfo.dTaxa, dTaxaItemNFPC)
        
        If objItemPedCompraInfo.dTaxa = 0 Then
            objItemPedCompraInfo.dTaxa = dCotacaoMoeda
            objItemPedCompraInfo.bTaxaPedido = True
        End If
        
        objItemPedCompraInfo.iMoeda = tItemPedCompraInfo.iMoeda
        objItemPedCompraInfo.dPrecoUnitario = dPrecoUnitario
        
        'Adiciona na coleção
        colItemPedCompraInfo.Add objItemPedCompraInfo

        'Busca no BD o próximo item de Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66142

    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ItensPedCompra_Le_NFiscalReceb = SUCESSO

    Exit Function

Erro_ItensPedCompra_Le_NFiscalReceb:

    ItensPedCompra_Le_NFiscalReceb = gErr

    Select Case gErr

        Case 66138, 66130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66139, 66140, 66142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
        
        Case 66141 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149661)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


Function PedidosCompraEnv_Le_Recebimento(ByVal colPedidoCompra As Collection, ByVal lCodigoFornecedor As Long, ByVal iFilial As Integer, ByVal iFilialCompra As Integer) As Long
'Lê os Pedidos de Compra Enviados que possuem a Filial, o Fornecedor e a FilialCompra passados

Dim lErro As Long
Dim lComando As Long
Dim tPedidoCompras As typePedidoCompras
Dim objPedidoCompras As ClassPedidoCompras

On Error GoTo Erro_PedidosCompraEnv_Le_Recebimento

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 65842

    tPedidoCompras.sContato = String(STRING_CONTATO, 0)
    tPedidoCompras.sTipoFrete = String(STRING_TIPOFRETE, 0)
    tPedidoCompras.sMotivoBaixa = String(STRING_MOTIVO_BAIXA, 0)
    tPedidoCompras.sAlcada = String(STRING_USUARIO_CODIGO, 0)

    'Pesquisa os Pedidos de Compras com Fornecedor, Filial e FilialCompra da tela
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Codigo, Fornecedor, Filial, Comprador, Contato, TipoDestino, FornCliDestino, FilialDestino, Data, DataEmissao, DataEnvio, DataAlteracao, CondicaoPagto, OutrasDespesas, ValorFrete, ValorSeguro, ValorDesconto, ValorTotal, ValorIPI, Observacao, TipoFrete, Transportadora, ProxSeqBloqueio, TipoBaixa, MotivoBaixa, Alcada FROM PedidoCompraDesbloqueado WHERE ((Fornecedor = ? AND Filial = ?) OR (TipoDestino = ? AND FornCliDestino = ? AND FilialDestino = ?)) AND FilialEmpresa = ? AND DataEnvio <> ? ORDER BY Codigo, Data", _
        tPedidoCompras.lNumIntDoc, tPedidoCompras.iFilialEmpresa, tPedidoCompras.lCodigo, tPedidoCompras.lFornecedor, tPedidoCompras.iFilial, tPedidoCompras.iComprador, tPedidoCompras.sContato, tPedidoCompras.iTipoDestino, tPedidoCompras.lFornCliDestino, tPedidoCompras.iFilialDestino, tPedidoCompras.dtData, tPedidoCompras.dtDataEmissao, tPedidoCompras.dtDataEnvio, tPedidoCompras.dtDataAlteracao, tPedidoCompras.iCondicaoPagto, tPedidoCompras.dOutrasDespesas, tPedidoCompras.dValorFrete, tPedidoCompras.dValorSeguro, tPedidoCompras.dValorDesconto, tPedidoCompras.dValorTotal, tPedidoCompras.dValorIPI, tPedidoCompras.lObservacao, tPedidoCompras.sTipoFrete, tPedidoCompras.iTransportadora, tPedidoCompras.iProxSeqBloqueio, tPedidoCompras.iTipoBaixa, tPedidoCompras.sMotivoBaixa, tPedidoCompras.sAlcada, lCodigoFornecedor, iFilial, TIPO_DESTINO_FORNECEDOR, lCodigoFornecedor, iFilial, iFilialCompra, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 65843

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65844

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 65845

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objPedidoCompras = New ClassPedidoCompras

        'Preenche o objPedidoCompras
        With objPedidoCompras

            .lNumIntDoc = tPedidoCompras.lNumIntDoc
            .iFilialEmpresa = tPedidoCompras.iFilialEmpresa
            .lCodigo = tPedidoCompras.lCodigo
            .lFornecedor = tPedidoCompras.lFornecedor
            .iFilial = tPedidoCompras.iFilial
            .iComprador = tPedidoCompras.iComprador
            .sContato = tPedidoCompras.sContato
            .iTipoDestino = tPedidoCompras.iTipoDestino
            .lFornCliDestino = tPedidoCompras.lFornCliDestino
            .iFilialDestino = tPedidoCompras.iFilialDestino
            .dtData = tPedidoCompras.dtData
            .dtDataEmissao = tPedidoCompras.dtDataEmissao
            .dtDataEnvio = tPedidoCompras.dtDataEnvio
            .dtDataAlteracao = tPedidoCompras.dtDataAlteracao
            .iCondicaoPagto = tPedidoCompras.iCondicaoPagto
            .dOutrasDespesas = tPedidoCompras.dOutrasDespesas
            .dValorFrete = tPedidoCompras.dValorFrete
            .dValorSeguro = tPedidoCompras.dValorSeguro
            .dValorDesconto = tPedidoCompras.dValorDesconto
            .dValorTotal = tPedidoCompras.dValorTotal
            .dValorIPI = tPedidoCompras.dValorIPI
            .lObservacao = tPedidoCompras.lObservacao
            .sTipoFrete = tPedidoCompras.sTipoFrete
            .iTransportadora = tPedidoCompras.iTransportadora
            .iProxSeqBloqueio = tPedidoCompras.iProxSeqBloqueio
            .iTipoBaixa = tPedidoCompras.iTipoBaixa
            .sMotivoBaixa = tPedidoCompras.sMotivoBaixa
            .sAlcada = tPedidoCompras.sAlcada

        End With

        colPedidoCompra.Add objPedidoCompras

        'Busca no BD o próximo Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65846
    
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    PedidosCompraEnv_Le_Recebimento = SUCESSO

    Exit Function

Erro_PedidosCompraEnv_Le_Recebimento:

    PedidosCompraEnv_Le_Recebimento = gErr

    Select Case gErr

        Case 65842
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65843, 65844, 65846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 65845 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149662)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCompra_Le_Recebimento(ByVal objNFiscal As ClassNFiscal, ByVal colPedidoCompra As Collection) As Long
'Lê os Pedidos de Compra não baixados Relacionados aos itens da Nota fiscal passada por objNFical

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim lCodigo As Long
Dim iFilialEmpresa As Integer
Dim objPedidoCompra As ClassPedidoCompras

On Error GoTo Erro_PedidoCompra_Le_Recebimento

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66133

    'Pesquisa os Pedidos de Compras associados a Nota Fiscal a partir do NumIntDoc da NF
    lErro = Comando_Executar(lComando, "SELECT DISTINCT PedidoCompra.NumIntDoc, PedidoCompra.Codigo, PedidoCompra.FilialEmpresa FROM PedidoCompra, ItensPedCompra, ItemNFItemPC, ItensNFiscal, NFiscal WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItensNFiscal.NumIntDoc = ItemNFItemPC.ItemNFiscal AND ItensPedCompra.NumIntDoc = ItemNFItemPC.ItemPedCompra AND ItensPedCompra.PedCompra = PedidoCompra.NumIntDoc AND NFiscal.NumIntDoc = ? ORDER BY PedidoCompra.Codigo", _
        lNumIntDoc, lCodigo, iFilialEmpresa, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 66134

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66135
    
    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66136
            
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objPedidoCompra = New ClassPedidoCompras

        objPedidoCompra.lNumIntDoc = lNumIntDoc
        objPedidoCompra.lCodigo = lCodigo
        objPedidoCompra.iFilialEmpresa = iFilialEmpresa

        colPedidoCompra.Add objPedidoCompra

        'Busca no BD o próximo Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66137
    
    Loop
        
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    PedidoCompra_Le_Recebimento = SUCESSO

    Exit Function

Erro_PedidoCompra_Le_Recebimento:

    PedidoCompra_Le_Recebimento = gErr

    Select Case gErr

        Case 66133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66134, 66135, 66137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
            
        Case 66136 'Tratado na rotina chamadora
                                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149663)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PedidoCompras_Le_Valores(ByVal colPedidoCompras As Collection) As Long
'Lê valores dos Pedido de compras passados na coleção colPedidoCompras

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim objPedidoCompras As ClassPedidoCompras
Dim dValorFrete As Double
Dim dValorSeguro As Double
Dim dValorDespesas As Double
Dim dValorDescontos As Double

On Error GoTo Erro_PedidoCompras_Le_Valores
    
    'Abertura do comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 66063
    
    'Abertura do comando
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 66566
    
    For Each objPedidoCompras In colPedidoCompras
    
        'Pesquisa os Pedidos de Compras a partir do código e FilialEmpresa passados
        lErro = Comando_Executar(lComando1, "SELECT ValorFrete, ValorSeguro, OutrasDespesas, ValorDesconto FROM PedidoCompra WHERE Codigo = ? AND FilialEmpresa = ?", dValorFrete, dValorSeguro, dValorDespesas, dValorDescontos, objPedidoCompras.lCodigo, objPedidoCompras.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 66064
    
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66065
    
        'Se não encontrou --> Erro
        If lErro = AD_SQL_SEM_DADOS Then
        
            'Verifica se o Pedido de Compras está baixado
            lErro = Comando_Executar(lComando2, "SELECT ValorFrete FROM PedidoCompraBaixado WHERE Codigo = ? AND FilialEmpresa = ?", dValorFrete, objPedidoCompras.lCodigo, objPedidoCompras.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 65740

            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65745

            'Se encontrou --> Erro
            If lErro = AD_SQL_SUCESSO Then
                gError 66066
            Else
                gError 89228
            End If
    
        End If
        
        objPedidoCompras.dValorFrete = dValorFrete
        objPedidoCompras.dValorSeguro = dValorSeguro
        objPedidoCompras.dOutrasDespesas = dValorDespesas
        objPedidoCompras.dValorDesconto = dValorDescontos
        
    Next

    'Fecha o comando
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    PedidoCompras_Le_Valores = SUCESSO

    Exit Function

Erro_PedidoCompras_Le_Valores:

    PedidoCompras_Le_Valores = gErr

    Select Case gErr

        Case 66063, 66566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66064, 66065
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 65740, 65745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRABAIXADO", gErr)
            
        Case 66066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_BAIXADO", gErr, objPedidoCompras.lCodigo)
            
        Case 89228
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompras.lCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149664)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function RecebimentoFCom_Verifica_Existencia(ByVal objNFiscal As ClassNFiscal) As Long
'Testa se existe outro recebimento com os memos dados de Número de Nota

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long, lNumRecebimento As Long
Dim dtDataEntrada As Date
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_RecebimentoFCom_Verifica_Existencia
        
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66626
        
    'Lê o Recebimento Com os Dados Iguais
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumRecebimento FROM NFiscal WHERE FilialEmpresa = ? AND Fornecedor = ? AND FilialForn = ? AND DataEntrada = ? AND TipoNFiscal = ? AND Serie = ? AND NumNotaFiscal = ?", lNumIntDoc, lNumRecebimento, objNFiscal.iFilialEmpresa, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
    If lErro <> AD_SQL_SUCESSO Then gError 66627

    'Le o Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66628
    
    'Se Encontrar --> Erro
    If lErro = AD_SQL_SUCESSO Then gError 66629
        
    'Verifica a se há um Recebimento emitido dentro de um prazo -->
    '--> menor que o período de emissão
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, DataEntrada FROM NFiscal WHERE FilialEmpresa = ? AND Fornecedor = ? AND FilialForn = ? AND TipoNFiscal = ? AND Serie = ? AND NumNotaFiscal = ? AND DataEntrada > ? AND DataEntrada < ?", lNumIntDoc, dtDataEntrada, objNFiscal.iFilialEmpresa, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEntrada - PERIODO_EMISSAO, objNFiscal.dtDataEntrada + PERIODO_EMISSAO)
    If lErro <> AD_SQL_SUCESSO Then gError 66630

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66631

    'Se encontrar
    If lErro <> AD_SQL_SEM_DADOS Then

        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_RECEBMATERIALFCOM_MESMO_NUMERO", objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, dtDataEntrada)
        'Se resposta for, Sai da Rotina
        If vbMsg = vbNo Then gError 66632

    End If
    'Fecha o comando
    Call Comando_Fechar(lComando)

    RecebimentoFCom_Verifica_Existencia = SUCESSO
    
    Exit Function
    
Erro_RecebimentoFCom_Verifica_Existencia:
    
    RecebimentoFCom_Verifica_Existencia = gErr
    
    Select Case gErr
    
        Case 66626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 66627, 66628, 66630, 66631
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
    
        Case 66629
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NFEXT_CADASTRADO_FORNECEDOR", gErr, lNumRecebimento, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEntrada)
    
        Case 66632 'Já foi Tratado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149665)

    End Select

End Function

Function ItensPedidosCompra_VerificaQuantidade(ByVal colPedidos As Collection) As Long
'Remove da coleção passada os Pedidos de compras que não possuem pelo menos um
'ItemPC com quantidade à Receber

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim iStatus As Integer

On Error GoTo Erro_ItensPedidosCompra_VerificaQuantidade

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66626
        
    'Para cada Pedido de Compras passado
    For iIndice = colPedidos.Count To 1 Step -1
        
        'Verifica se o Pedido tem pelo menos um item com quantidade a receber (Status aberto)
        lErro = Comando_Executar(lComando, "SELECT ItensPedCompra.Status FROM ItensPedCompra, PedidoCompra WHERE ItensPedCompra.PedCompra = PedidoCompra.NumIntDoc And PedidoCompra.NumIntDoc = ? AND ItensPedCompra.Status = ? AND (Quantidade - QuantRecebida -QuantRecebimento)>0", iStatus, colPedidos(iIndice).lNumIntDoc, ITEM_PED_COMPRAS_ABERTO)
        If lErro <> AD_SQL_SUCESSO Then gError 66984
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67247
        
        'Se não encontrou nenhum item
        If lErro = AD_SQL_SEM_DADOS Then
            
            'Remove o Pedido de compras da coleção
            colPedidos.Remove (iIndice)
        
        End If
        
    Next
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensPedidosCompra_VerificaQuantidade = SUCESSO

    Exit Function
    
Erro_ItensPedidosCompra_VerificaQuantidade:
    
    ItensPedidosCompra_VerificaQuantidade = gErr
    
    Select Case gErr
    
        Case 66626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 66984, 67247
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149666)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function NFiscalEntradaCom_Lock_Gravacao(ByVal objNFiscal As ClassNFiscal, ByVal colPedCompras As Collection, ByVal colReqCompras As Collection) As Long

Dim lErro As Long
Dim iIndice As Long
Dim alComando(0 To 7) As Long
Dim sNome As String
Dim sDescricao As String
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iFilialEmpresa As Integer
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim objItemNFItemPC As ClassItemNFItemPC
Dim objItemNFItemRC As ClassItemNFItemRC
Dim lItemNF As Long
Dim lNumIntDoc As Long
Dim objPedidoCompra As ClassPedidoCompras
Dim iFilialEmpresa1 As Integer

On Error GoTo Erro_NFiscalEntradaCom_Lock_Gravacao

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 66510
    Next

    sDescricao = String(STRING_NATUREZAOP_DESCRICAO, 0)

    lErro = Comando_ExecutarLockado(alComando(2), "SELECT Descricao FROM NaturezaOp WHERE Codigo = ? ", sDescricao, objNFiscal.sNaturezaOp)
    If lErro <> AD_SQL_SUCESSO Then gError 66511

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65928
    If lErro = AD_SQL_SEM_DADOS Then gError 66512

    lErro = Comando_LockShared(alComando(2))
    If lErro <> AD_SQL_SUCESSO Then gError 66513

    sNome = String(STRING_FILIAL_FORNECEDOR_NOME, 0)

    'Lê a FilialFornecedor
    lErro = Comando_ExecutarLockado(alComando(0), "SELECT Nome FROM FiliaisFornecedores WHERE CodFornecedor = ? AND CodFilial = ? ", sNome, objNFiscal.lFornecedor, objNFiscal.iFilialForn)
    If lErro <> AD_SQL_SUCESSO Then gError 66514

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66546
    If lErro = AD_SQL_SEM_DADOS Then gError 66515 'Não encontrou --> Erro

    'Faz um "lock" na FilialFornecedor
    lErro = Comando_LockShared(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 66516

    'Se Nota for interna, lock shared na série
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 66517
    If lErro <> SUCESSO Then gError 66518

    If objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA Then

        If objNFiscal.iFilialEmpresa = 0 Then objNFiscal.iFilialEmpresa = giFilialEmpresa
        
        lErro = Comando_ExecutarLockado(alComando(3), "SELECT FilialEmpresa FROM Serie WHERE Serie =? AND FilialEmpresa = ? ", iFilialEmpresa, objNFiscal.sSerie, objNFiscal.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 66519

        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66520
        If lErro = AD_SQL_SEM_DADOS Then gError 66521

        lErro = Comando_LockShared(alComando(3))
        If lErro <> AD_SQL_SUCESSO Then gError 66522
    End If

    'Se a Transportadora estiver preenchida
    If objNFiscal.iCodTransportadora > 0 Then
        'Lê a Transportadora
        lErro = Comando_ExecutarLockado(alComando(1), "SELECT Codigo FROM Transportadoras WHERE Codigo = ?", iCodigo, objNFiscal.iCodTransportadora)
        If lErro <> AD_SQL_SUCESSO Then gError 66523

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 65942
        If lErro = AD_SQL_SEM_DADOS Then gError 66524 'Não encontrou --> Erro

        'Faz um "lock" na transportadora
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 66525
    End If

    'Lock shared nos Produtos e Almoxarifados
    For Each objItemNF In objNFiscal.colItensNF

        objProduto.sCodigo = objItemNF.sProduto
        lErro = CF("Produto_Lock_Shared", alComando(5), objProduto)
        If lErro <> SUCESSO Then gError 66526

        objItemNF.iControleEstoque = objProduto.iControleEstoque
        objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
        objItemNF.sUMVenda = objProduto.sSiglaUMVenda
        objItemNF.iClasseUM = objProduto.iClasseUM
        objItemNF.iApropriacaoProd = objProduto.iApropriacaoCusto

        If objItemNF.iAlmoxarifado > 0 Then

            If objNFiscal.iFilialEmpresa = 0 Then objNFiscal.iFilialEmpresa = giFilialEmpresa
            
            iFilialEmpresa1 = objNFiscal.iFilialEmpresa
            lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa1)
            If lErro <> SUCESSO Then Error 66526
            
            lErro = Comando_ExecutarPos(alComando(3), "SELECT FilialEmpresa FROM Almoxarifado WHERE Codigo = ? AND FilialEmpresa = ? ", 0, iFilialEmpresa, objItemNF.iAlmoxarifado, iFilialEmpresa1)
            If lErro <> AD_SQL_SUCESSO Then gError 66527

            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65947
            If lErro = AD_SQL_SEM_DADOS Then gError 66528

            lErro = Comando_LockShared(alComando(3))
            If lErro <> AD_SQL_SUCESSO Then gError 66529
        End If
    Next

    'Para cada item da Nota Fiscal
    For Each objItemNF In objNFiscal.colItensNF
        
        'Para cada item de Nota Fiscal associado ao item de Pedido de Compras
        For Each objItemNFItemPC In objItemNF.colItemNFItemPC
            
            'Busca no BD o ItemPC
            lErro = Comando_ExecutarPos(alComando(4), "SELECT NumIntDoc FROM ItensPedCompra WHERE NumIntDoc = ?", 0, lItemNF, objItemNFItemPC.lItemPedCompra)
            If lErro <> AD_SQL_SUCESSO Then gError 66530
            
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66531
    
            'Se não encontrar, erro
            If lErro = AD_SQL_SEM_DADOS Then gError 66532
            
            'Faz LockShared em ItensPedCompra
            lErro = Comando_LockShared(alComando(4))
            If lErro <> AD_SQL_SUCESSO Then gError 66533
        
        Next
        
        'Para cada item de nota fiscal associado a um item de Requisição de compras
        For Each objItemNFItemRC In objItemNF.colItemNFItemRC
    
            'Busca no BD o ItensReqCompra
            lErro = Comando_ExecutarPos(alComando(5), "SELECT NumIntDoc FROM ItensReqCompra WHERE NumIntDoc = ?", 0, lItemNF, objItemNFItemRC.lItemReqCompra)
            If lErro <> AD_SQL_SUCESSO Then gError 66534
            
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66535
            
            'Se não encontrar, erro
            If lErro = AD_SQL_SEM_DADOS Then gError 66536
        
            'Faz lockShared em ItensReqCompra
            lErro = Comando_LockShared(alComando(5))
            If lErro <> AD_SQL_SUCESSO Then gError 66537
            
        Next
    
    Next
    
    'Para cada Pedido de Compras em colPedCompras
    For Each objPedidoCompra In colPedCompras
        
        'Pesquisa no BD o Pedido de Compras passado
        lErro = Comando_ExecutarPos(alComando(6), "SELECT NumIntDoc FROM PedidoCompra WHERE Codigo = ? AND FilialEmpresa = ?", 0, lNumIntDoc, objPedidoCompra.lCodigo, objNFiscal.iFilialPedido)
        If lErro <> AD_SQL_SUCESSO Then gError 66538
        
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66539
    
        'Se não encontrou, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 66540
        
        'Faz LockShared no Pedido de Compras
        lErro = Comando_LockShared(alComando(6))
        If lErro <> AD_SQL_SUCESSO Then gError 66541
    
        objPedidoCompra.lNumIntDoc = lNumIntDoc
    
    Next
    
    'Para cada Requisição de Compras em colReqCompras
    For iIndice = 1 To colReqCompras.Count
        
        'Pesquisa no BD a Requisição de Compras passada
        lErro = Comando_ExecutarPos(alComando(7), "SELECT NumIntDoc FROM RequisicaoCompra WHERE Codigo = ? AND FilialCompra = ?", 0, lNumIntDoc, colReqCompras(iIndice), objNFiscal.iFilialPedido)
        If lErro <> AD_SQL_SUCESSO Then gError 66542
        
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66543
        
        'Se não encontrou, erro
        If lErro = AD_SQL_SEM_DADOS Then gError 66544
    
        'Faz lockShared na Requisição de Compras
        lErro = Comando_LockShared(alComando(7))
        If lErro <> AD_SQL_SUCESSO Then gError 66545
    
    Next

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalEntradaCom_Lock_Gravacao = SUCESSO

    Exit Function

Erro_NFiscalEntradaCom_Lock_Gravacao:

    NFiscalEntradaCom_Lock_Gravacao = gErr

    Select Case gErr

        Case 66510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66511, 65928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP", gErr, objNFiscal.sNaturezaOp)

        Case 66512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", gErr, objNFiscal.sNaturezaOp)

        Case 66513
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_NATUREZAOP", gErr, objNFiscal.sNaturezaOp)

        Case 66514, 66546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISFORNECEDORES1", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 66515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 66516
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISFORNECEDORES1", gErr, objNFiscal.lFornecedor, objNFiscal.iFilialForn)

        Case 66517, 66526

        Case 66518
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 66519, 66520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE", gErr)

        Case 66521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objNFiscal.sSerie)

        Case 66522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERIE", gErr)

        Case 66523, 65942
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORA", gErr, objNFiscal.iCodTransportadora)

        Case 66524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objNFiscal.iCodTransportadora)

        Case 66525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRANSPORTADORA", gErr)

        Case 66527, 65947
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO", gErr, objItemNF.iAlmoxarifado)

        Case 66528
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objItemNF.iAlmoxarifado)

        Case 66529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ALMOXARIFADO1", gErr, objItemNF.iAlmoxarifado)

        Case 66530, 66531
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)

        Case 66532
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDCOMPRA_INEXISTENTE", gErr, objItemNFItemPC.lItemPedCompra)

        Case 66533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSPEDCOMPRA", gErr)

        Case 66534, 66535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSREQCOMPRA", gErr, objItemNFItemRC.lItemReqCompra)

        Case 66536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITENSREQCOMPRA_NAO_CADASTRADO", gErr, objItemNFItemRC.lItemReqCompra)

        Case 66537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ITENSREQCOMPRA1", gErr, objItemNFItemRC.lItemReqCompra)

        Case 66538, 66539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 66540
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOCOMPRA_NAO_CADASTRADO", gErr, objPedidoCompra.lCodigo)

        Case 66541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PEDIDOCOMPRA", gErr, objPedidoCompra.lCodigo)

        Case 66542, 66543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", gErr, colReqCompras.Item(iIndice))

        Case 66544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REQUISICAOCOMPRA_NAO_CADASTRADA", gErr, colReqCompras.Item(iIndice))

        Case 66545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_REQUISICAOCOMPRA", gErr, colReqCompras.Item(iIndice))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149667)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ItensPCTodos_Le_Codigo(ByVal objPedidoCompra As ClassPedidoCompras) As Long
'Lê os itens de um Pedido de Compra (baixado ou não) a partir do código e de FilialEmpresa do Pedido de Compras

Dim lErro As Long
Dim tItemPedido As typeItemPedCompra
Dim lComando As Long
Dim lTransacao As Long
Dim objItemPC As ClassItemPedCompra

On Error GoTo Erro_ItensPCTodos_Le_Codigo

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66147

    tItemPedido.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedido.sProduto = String(STRING_PRODUTO, 0)
    tItemPedido.sObservacao = String(STRING_OBSERVACAO, 0) 'OK o tamanho correto da string é 255
    tItemPedido.sUM = String(STRING_UM_SIGLA, 0)

    'Lê registros da tabela de ItensPedCompraTodos
    lErro = Comando_Executar(lComando, "SELECT DISTINCT ItensPedCompraTodos.NumIntDoc, PedCompra, DataLimite, Produto, DescProduto, Quantidade, QuantRecebida, QuantRecebimento, UM, PrecoUnitario, ItensPedCompraTodos.ValorDesconto, TipoOrigem, NumIntOrigem, PercentMaisReceb, PercentMenosReceb, RecebForaFaixa, ItensPedCompraTodos.Status, ItensPedCompraTodos.Observacao, ItensPedCompraTodos.ValorIPI, ItensPedCompraTodos.AliquotaIPI, ItensPedCompraTodos.AliquotaICMS, PedidoCompraTodos.Taxa FROM ItensPedCompraTodos, PedidoCompraTodos WHERE PedidoCompraTodos.NumIntDoc = ItensPedCompraTodos.PedCompra AND PedidoCompraTodos.Codigo = ? AND PedidoCompraTodos.FilialEmpresa = ? Order By ItensPedCompraTodos.Produto", tItemPedido.lNumIntDoc, tItemPedido.lPedCompra, _
        tItemPedido.dtDataLimite, tItemPedido.sProduto, tItemPedido.sDescProduto, tItemPedido.dQuantidade, tItemPedido.dQuantRecebida, tItemPedido.dQuantRecebimento, tItemPedido.sUM, tItemPedido.dPrecoUnitario, tItemPedido.dValorDesconto, tItemPedido.iTipoOrigem, tItemPedido.lNumIntOrigem, tItemPedido.dPercentMaisReceb, tItemPedido.dPercentMenosReceb, tItemPedido.iRebebForaFaixa, tItemPedido.iStatus, tItemPedido.lObservacao, tItemPedido.dValorIPI, tItemPedido.dAliquotaIPI, tItemPedido.dAliquotaICMS, tItemPedido.dTaxa, objPedidoCompra.lCodigo, objPedidoCompra.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 66148

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66149

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66150

    Do While lErro = AD_SQL_SUCESSO

        Set objItemPC = New ClassItemPedCompra

        'preenchendo objitem
        objItemPC.dAliquotaICMS = tItemPedido.dAliquotaICMS
        objItemPC.dAliquotaIPI = tItemPedido.dAliquotaIPI
        objItemPC.dPercentMaisReceb = tItemPedido.dPercentMaisReceb
        objItemPC.dPercentMenosReceb = tItemPedido.dPercentMenosReceb
        objItemPC.dPrecoUnitario = tItemPedido.dPrecoUnitario
        objItemPC.dQuantidade = tItemPedido.dQuantidade
        objItemPC.dQuantRecebida = tItemPedido.dQuantRecebida
        objItemPC.dQuantRecebimento = tItemPedido.dQuantRecebimento
        objItemPC.dtDataLimite = tItemPedido.dtDataLimite
        objItemPC.dValorDesconto = tItemPedido.dValorDesconto
        objItemPC.dValorIPI = tItemPedido.dValorIPI
        objItemPC.iRebebForaFaixa = tItemPedido.iRebebForaFaixa
        objItemPC.iStatus = tItemPedido.iStatus
        objItemPC.iTipoOrigem = tItemPedido.iTipoOrigem
        objItemPC.lNumIntDoc = tItemPedido.lNumIntDoc
        objItemPC.lNumIntOrigem = tItemPedido.lNumIntOrigem
        objItemPC.lObservacao = tItemPedido.lObservacao
        objItemPC.sDescProduto = tItemPedido.sDescProduto
        objItemPC.sObservacao = tItemPedido.sObservacao
        objItemPC.sProduto = tItemPedido.sProduto
        objItemPC.sUM = tItemPedido.sUM
        objItemPC.dTaxa = tItemPedido.dTaxa

        'Insere em colItens os valores de objItemPC
        objPedidoCompra.colItens.Add objItemPC

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66151
    
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensPCTodos_Le_Codigo = SUCESSO

    Exit Function

Erro_ItensPCTodos_Le_Codigo:

    ItensPCTodos_Le_Codigo = gErr

    Select Case gErr

        Case 66147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66148, 66149, 66151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr, objPedidoCompra.lNumIntDoc)

        Case 66150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDCOMPRA_AUSENCIA_ITENS", gErr, objPedidoCompra.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149668)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ReqComprasTodas_Le_PedidoCompra(ByVal objPedidoCompras As ClassPedidoCompras, ByVal colReqCompras As Collection) As Long
'Lê as Requisições de compras (baixadas ou não) associadas ao Pedidos de Compra passado

Dim lErro As Long
Dim lComando As Long
Dim lCodReq As Long
Dim lCodigoFornecedor As Long
Dim iFilial As Integer

On Error GoTo Erro_ReqComprasTodas_Le_PedidoCompra

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66567

    'Lê as Requisicões de compras baixadas ou não relacionadas a Pedidos de Compras
    lErro = Comando_Executar(lComando, "SELECT DISTINCT RequisicaoCompraTodas.Codigo FROM RequisicaoCompraTodas, ItensReqCompraTodos, PedidoCompraTodos, ItensPedCompraTodos, ItemRCItemPC WHERE PedidoCompraTodos.NumIntDoc = ItensPedCompraTodos.PedCompra AND ItensPedCompraTodos.NumIntDoc = ItemRCItemPC.ItemPC AND ItemRCItemPC.ItemRC = ItensReqCompraTodos.NumIntDoc AND ItensReqCompraTodos.ReqCompra = RequisicaoCompraTodas.NumIntDoc AND PedidoCompraTodos.Codigo = ? AND PedidoCompraTodos.FilialEmpresa = ? ORDER BY RequisicaoCompraTodas.Codigo", _
        lCodReq, objPedidoCompras.lCodigo, objPedidoCompras.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 66568

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66569

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66570

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        'Guarda o código da Requisição
        colReqCompras.Add lCodReq

        'Busca no BD o próximo Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66571
    
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ReqComprasTodas_Le_PedidoCompra = SUCESSO

    Exit Function

Erro_ReqComprasTodas_Le_PedidoCompra:

    ReqComprasTodas_Le_PedidoCompra = gErr

    Select Case gErr

        Case 66567
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66568, 66569, 66571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 66570 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149669)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCompra_Le_ReqComprasVinculada(ByVal objPedidoCompras As ClassPedidoCompras, ByVal colReqCompras As Collection) As Long
'Lê as Requisições de compras associadas ao Pedidos de Compra passado

Dim lErro As Long
Dim lComando As Long
Dim lCodReq As Long
Dim lCodigoFornecedor As Long
Dim iFilial As Integer

On Error GoTo Erro_PedidoCompra_Le_ReqComprasVinculada

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 65793

    'Pesquisa os Pedidos de Compras com Fornecedor, Filial e FilialCompra da tela
    lErro = Comando_Executar(lComando, "SELECT DISTINCT RequisicaoCompra.Codigo FROM RequisicaoCompra, ItensReqCompra, PedidoCompra, ItensPedCompra, ItemRCItemPC WHERE PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND ItensPedCompra.NumIntDoc = ItemRCItemPC.ItemPC AND ItemRCItemPC.ItemRC = ItensReqCompra.NumIntDoc AND ItensReqCompra.ReqCompra = RequisicaoCompra.NumIntDoc AND PedidoCompra.Codigo = ? AND PedidoCompra.FilialEmpresa = ? ORDER BY RequisicaoCompra.Codigo", _
        lCodReq, objPedidoCompras.lCodigo, objPedidoCompras.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 65794

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65795

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 65796

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        'Guarda o código da Requisição
        colReqCompras.Add lCodReq

        'Busca no BD o próximo Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65797
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    PedidoCompra_Le_ReqComprasVinculada = SUCESSO

    Exit Function

Erro_PedidoCompra_Le_ReqComprasVinculada:

    PedidoCompra_Le_ReqComprasVinculada = gErr

    Select Case gErr

        Case 65793
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65794, 65795, 65797
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 65796 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149670)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensRequisicaoTodos_Le_PedidoCompra(ByVal objRequisicaoCompras As ClassRequisicaoCompras, ByVal colItemRCInfo As Collection, ByVal lNumIntDocNF As Long) As Long
'Lê todos os Itens de Requisição de Compras (baixados ou não) associadas ao Pedido de Compras

Dim lErro As Long
Dim lComando As Long
Dim objItemReqCompraInfo As ClassItemReqComprasInfo
Dim tItemReqCompraInfo As typeItemReqCompraInfo
Dim iContPC As Integer
Dim iContRC As Integer
Dim lPedCompra As Long, dFator As Double
Dim lReqCompra As Long, iClasseUM As Integer, sSiglaUMCompra As String
Dim dQuantidadePedido As Double
Dim dQuantRecebidaPedido As Double
Dim dQuantRecebidaTotal As Double

On Error GoTo Erro_ItensRequisicaoTodos_Le_PedidoCompra
    
    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66102

    tItemReqCompraInfo.sProduto = String(STRING_PRODUTO, 0)
    tItemReqCompraInfo.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemReqCompraInfo.sUM = String(STRING_UM_SIGLA, 0)
    sSiglaUMCompra = String(STRING_UM_SIGLA, 0)

    'Pesquisa os itens de Requisição de Compras relacionadas a um pedido de compras
    lErro = Comando_Executar(lComando, "SELECT DISTINCT Produtos.SiglaUMCompra, Produtos.ClasseUM, " & _
        "PedidoCompraTodos.Codigo, RequisicaoCompraTodas.Codigo, ItemNFItemPC.Quantidade, ItensReqCompraTodos.NumIntDoc, ItensReqCompraTodos.Produto, ItensReqCompraTodos.DescProduto, ItensReqCompraTodos.UM, ItensReqCompraTodos.QuantPedida, RequisicaoCompraTodas.FilialEmpresa, RequisicaoCompraTodas.Urgente, ItensReqCompraTodos.Quantidade, ItemNFItemRC.Quantidade, ItensReqCompraTodos.QuantCancelada, ItensReqCompraTodos.TipoTributacao, ItensPedCompraTodos.NumIntDoc FROM RequisicaoCompraTodas, ItensReqCompraTodos, PedidoCompraTodos, ItensPedCompraTodos, ItemRCItemPC, ItemNFItemPC, ItemNFItemRC, ItensNFiscal, NFiscal, Produtos WHERE PedidoCompraTodos.NumIntDoc = ItensPedCompraTodos.PedCompra AND ItensPedCompraTodos.NumIntDoc = ItemRCItemPC.ItemPC AND ItemRCItemPC.ItemRC = ItensReqCompraTodos.NumIntDoc AND ItemNFItemPC.ItemNFiscal = ItensNFiscal.NumIntDoc AND ItensReqCompraTodos.ReqCompra = RequisicaoCompraTodas.NumIntDoc AND RequisicaoCompraTodas.Codigo = ? AND " _
        & " ItensReqCompraTodos.Produto = Produtos.Codigo AND ItensPedCompraTodos.NumIntDoc = ItemNFItemPC.ItemPedCompra AND ItemNFItemPC.ItemNFiscal = ItensNFiscal.NumIntDoc AND ItensReqCompraTodos.NumIntDoc = ItemNFItemRC.ItemReqCompra AND ItemNFItemRC.ItemNFiscal = ItensNFiscal.NumIntDoc AND ItensNFiscal.NumIntNF = NFiscal.NumIntDoc AND ItensNFiscal.NumIntNF = ? ORDER BY PedidoCompraTodos.Codigo, RequisicaoCompraTodas.Codigo, RequisicaoCompraTodas.Urgente, ItensReqCompraTodos.Produto", _
        sSiglaUMCompra, iClasseUM, tItemReqCompraInfo.lPedCompra, tItemReqCompraInfo.lReqCompra, dQuantidadePedido, tItemReqCompraInfo.lNumIntDoc, tItemReqCompraInfo.sProduto, tItemReqCompraInfo.sDescProduto, tItemReqCompraInfo.sUM, tItemReqCompraInfo.dQuantPedida, tItemReqCompraInfo.iFilialEmpresa, tItemReqCompraInfo.lUrgente, tItemReqCompraInfo.dQuantidade, tItemReqCompraInfo.dQuantRecebida, tItemReqCompraInfo.dQuantCancelada, tItemReqCompraInfo.iTipoTributacao, tItemReqCompraInfo.lNumIntDocItemPC, objRequisicaoCompras.lCodigo, lNumIntDocNF)
    If lErro <> AD_SQL_SUCESSO Then gError 66103

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66104

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66105

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objItemReqCompraInfo = New ClassItemReqComprasInfo

        lPedCompra = tItemReqCompraInfo.lPedCompra
        lReqCompra = tItemReqCompraInfo.lReqCompra
        
        iContPC = iContPC + 1
        iContRC = iContRC + 1
        
        objItemReqCompraInfo.lPedCompra = tItemReqCompraInfo.lPedCompra
        objItemReqCompraInfo.iItemPC = iContPC
        objItemReqCompraInfo.lNumIntDocItemPC = tItemReqCompraInfo.lNumIntDocItemPC
        objItemReqCompraInfo.sProduto = tItemReqCompraInfo.sProduto
        objItemReqCompraInfo.sDescProduto = tItemReqCompraInfo.sDescProduto
        objItemReqCompraInfo.sUM = tItemReqCompraInfo.sUM
        
        lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUMCompra, tItemReqCompraInfo.sUM, dFator)
        If lErro <> SUCESSO Then gError 66323
        
        objItemReqCompraInfo.dQuantRecebidoPedido = dFator * dQuantidadePedido
        objItemReqCompraInfo.iClasseUM = iClasseUM
        objItemReqCompraInfo.sSiglaUMCompra = sSiglaUMCompra
        
        objItemReqCompraInfo.iFilialEmpresa = tItemReqCompraInfo.iFilialEmpresa
        objItemReqCompraInfo.lReqCompra = tItemReqCompraInfo.lReqCompra
        objItemReqCompraInfo.iUrgente = tItemReqCompraInfo.lUrgente
        objItemReqCompraInfo.iItemRC = iContRC
        
        'A quantidade a Receber desse Item da Nota Fiscal será a Quantidade Total a Receber - a total que foi recebida + a quantidade que foi recebida quando a NF foi gravada
        objItemReqCompraInfo.dQuantReceber = tItemReqCompraInfo.dQuantidade - dQuantRecebidaTotal
        
        'A quantidade recebida será a grava em ItemNFItemRC
        objItemReqCompraInfo.dQuantRecebida = tItemReqCompraInfo.dQuantRecebida
        objItemReqCompraInfo.lNumIntDoc = tItemReqCompraInfo.lNumIntDoc
        objItemReqCompraInfo.iTipoTributacao = tItemReqCompraInfo.iTipoTributacao
        
        colItemRCInfo.Add objItemReqCompraInfo

        'Busca no BD o próximo Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66106
            
        If lPedCompra <> tItemReqCompraInfo.lPedCompra Then iContPC = 0
        If lReqCompra <> tItemReqCompraInfo.lReqCompra Then iContRC = 0
    
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensRequisicaoTodos_Le_PedidoCompra = SUCESSO

    Exit Function

Erro_ItensRequisicaoTodos_Le_PedidoCompra:

    ItensRequisicaoTodos_Le_PedidoCompra = gErr

    Select Case gErr

        Case 66102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66103, 66104, 66106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 66105 'Erro tratado na rotina chamadora
        
        Case 66323

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149671)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensRequisicao_Le_PedidoCompra_NF(ByVal objRequisicaoCompras As ClassRequisicaoCompras, ByVal colItemRCInfo As Collection, ByVal lNumIntNF As Long) As Long
'Lê os Itens de Requisição de Compras não baixados associadas ao Pedido de Compras associados a um item de nota fiscal

Dim lErro As Long
Dim lComando As Long
Dim objItemReqCompraInfo As ClassItemReqComprasInfo
Dim tItemReqCompraInfo As typeItemReqCompraInfo
Dim iContPC As Integer
Dim iContRC As Integer
Dim lPedCompra As Long, dFator As Double
Dim lReqCompra As Long, iClasseUM As Integer, sSiglaUMCompra As String
Dim dQuantidadePedido As Double
Dim dQuantRecebidaPedido As Double

On Error GoTo Erro_ItensRequisicao_Le_PedidoCompra_NF
    
    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 65848

    tItemReqCompraInfo.sProduto = String(STRING_PRODUTO, 0)
    tItemReqCompraInfo.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemReqCompraInfo.sUM = String(STRING_UM_SIGLA, 0)
    sSiglaUMCompra = String(STRING_UM_SIGLA, 0)

    'Pesquisa os itens de Requisição de Compras relacionadas a um pedido de compras
    lErro = Comando_Executar(lComando, "SELECT Produtos.SiglaUMCompra, Produtos.ClasseUM, " & _
        "PedidoCompra.Codigo, RequisicaoCompra.Codigo, ItensPedCompra.Quantidade, ItensPedCompra.QuantRecebida, ItensReqCompra.NumIntDoc, ItensReqCompra.Produto, ItensReqCompra.DescProduto, " & _
        "ItensReqCompra.UM, ItensReqCompra.QuantPedida, RequisicaoCompra.FilialEmpresa, RequisicaoCompra.Urgente, ItensReqCompra.Quantidade, ItensReqCompra.QuantRecebida, ItensReqCompra.QuantCancelada, ItensReqCompra.TipoTributacao, ItensPedCompra.NumIntDoc FROM RequisicaoCompra, ItensReqCompra, PedidoCompra, ItensPedCompra, ItemRCItemPC, ItemNFItemPC, ItemNFItemRC, ItensNFiscal, NFiscal, Produtos WHERE PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND ItensPedCompra.NumIntDoc = ItemRCItemPC.ItemPC AND ItemRCItemPC.ItemRC = ItensReqCompra.NumIntDoc AND ItensReqCompra.ReqCompra = RequisicaoCompra.NumIntDoc AND ItensNFiscal.NumIntNF = NFiscal.NumIntDoc AND NFiscal.NumIntDoc = ? AND RequisicaoCompra.Codigo = ? AND RequisicaoCompra.FilialEmpresa = ? AND ItensReqCompra.Status <> ? AND ItensReqCompra.Produto = Produtos.Codigo" _
        & " AND ItemNFItemPC.ItemPedCompra = ItensPedCompra.NumIntDoc AND ItemNFItemRC.ItemReqCompra = ItensReqCompra.NumIntDoc AND ItemNFItemPC.ItemNFIscal = ItensNFiscal.NumIntDoc ORDER BY PedidoCompra.Codigo, PedidoCompra.FilialEmpresa, RequisicaoCompra.Codigo, RequisicaoCompra.Urgente, ItensReqCompra.Produto", _
        sSiglaUMCompra, iClasseUM, tItemReqCompraInfo.lPedCompra, tItemReqCompraInfo.lReqCompra, dQuantidadePedido, dQuantRecebidaPedido, tItemReqCompraInfo.lNumIntDoc, tItemReqCompraInfo.sProduto, tItemReqCompraInfo.sDescProduto, tItemReqCompraInfo.sUM, tItemReqCompraInfo.dQuantPedida, tItemReqCompraInfo.iFilialEmpresa, tItemReqCompraInfo.lUrgente, tItemReqCompraInfo.dQuantidade, tItemReqCompraInfo.dQuantRecebida, tItemReqCompraInfo.dQuantCancelada, tItemReqCompraInfo.iTipoTributacao, tItemReqCompraInfo.lNumIntDocItemPC, lNumIntNF, objRequisicaoCompras.lCodigo, objRequisicaoCompras.iFilialEmpresa, ITEM_REQ_RECEBIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 65849

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65850

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 65851

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objItemReqCompraInfo = New ClassItemReqComprasInfo

        lPedCompra = tItemReqCompraInfo.lPedCompra
        lReqCompra = tItemReqCompraInfo.lReqCompra
        
        iContPC = iContPC + 1
        iContRC = iContRC + 1
        
        objItemReqCompraInfo.lPedCompra = tItemReqCompraInfo.lPedCompra
        objItemReqCompraInfo.iItemPC = iContPC
        objItemReqCompraInfo.sProduto = tItemReqCompraInfo.sProduto
        objItemReqCompraInfo.sDescProduto = tItemReqCompraInfo.sDescProduto
        objItemReqCompraInfo.sUM = tItemReqCompraInfo.sUM
        
        lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUMCompra, tItemReqCompraInfo.sUM, dFator)
        If lErro <> SUCESSO Then gError 66323
        
        objItemReqCompraInfo.dQuantRecebidoPedido = dQuantidadePedido * dFator
        objItemReqCompraInfo.iClasseUM = iClasseUM
        objItemReqCompraInfo.sSiglaUMCompra = sSiglaUMCompra
        
        objItemReqCompraInfo.iFilialEmpresa = tItemReqCompraInfo.iFilialEmpresa
        objItemReqCompraInfo.lReqCompra = tItemReqCompraInfo.lReqCompra
        objItemReqCompraInfo.iUrgente = tItemReqCompraInfo.lUrgente
        objItemReqCompraInfo.iItemRC = iContRC
        objItemReqCompraInfo.dQuantReceber = tItemReqCompraInfo.dQuantidade - tItemReqCompraInfo.dQuantRecebida
        objItemReqCompraInfo.dQuantRecebida = tItemReqCompraInfo.dQuantRecebida
        objItemReqCompraInfo.lNumIntDoc = tItemReqCompraInfo.lNumIntDoc
        objItemReqCompraInfo.iTipoTributacao = tItemReqCompraInfo.iTipoTributacao
        objItemReqCompraInfo.lNumIntDocItemPC = tItemReqCompraInfo.lNumIntDocItemPC
        
        colItemRCInfo.Add objItemReqCompraInfo

        'Busca no BD o próximo Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65852
            
        If lPedCompra <> tItemReqCompraInfo.lPedCompra Then iContPC = 0
        If lReqCompra <> tItemReqCompraInfo.lReqCompra Then iContRC = 0
    
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensRequisicao_Le_PedidoCompra_NF = SUCESSO

    Exit Function

Erro_ItensRequisicao_Le_PedidoCompra_NF:

    ItensRequisicao_Le_PedidoCompra_NF = gErr

    Select Case gErr

        Case 65848
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65849, 65850, 65852
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 65851 'Erro tratado na rotina chamadora

        Case 66323

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149672)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensRequisicao_Le_PedidoCompra(ByVal objRequisicaoCompras As ClassRequisicaoCompras, ByVal colItemRCInfo As Collection) As Long
'Lê os Itens de Requisição de Compras não baixados associadas ao Pedido de Compras

Dim lErro As Long
Dim lComando As Long
Dim objItemReqCompraInfo As ClassItemReqComprasInfo
Dim tItemReqCompraInfo As typeItemReqCompraInfo
Dim iContPC As Integer
Dim iContRC As Integer, dFator As Double
Dim lPedCompra As Long, iClasseUM As Integer, sSiglaUMCompra As String
Dim lReqCompra As Long
Dim dQuantidadePedido As Double
Dim dQuantRecebidaPedido As Double

On Error GoTo Erro_ItensRequisicao_Le_PedidoCompra
    
    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66102

    tItemReqCompraInfo.sProduto = String(STRING_PRODUTO, 0)
    tItemReqCompraInfo.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemReqCompraInfo.sUM = String(STRING_UM_SIGLA, 0)
    sSiglaUMCompra = String(STRING_UM_SIGLA, 0)

    'Pesquisa os itens de Requisição de Compras relacionadas a um pedido de compras
    lErro = Comando_Executar(lComando, "SELECT Produtos.SiglaUMCompra, Produtos.ClasseUM, PedidoCompra.Codigo, RequisicaoCompra.Codigo, ItensPedCompra.Quantidade, ItensPedCompra.QuantRecebida, ItensReqCompra.NumIntDoc, ItensReqCompra.Produto, ItensReqCompra.DescProduto, ItensReqCompra.UM, ItensReqCompra.QuantPedida, RequisicaoCompra.FilialEmpresa, RequisicaoCompra.Urgente, ItensReqCompra.Quantidade, ItensReqCompra.QuantRecebida, ItensReqCompra.QuantCancelada, ItensReqCompra.TipoTributacao, ItensPedCompra.NumIntDoc FROM RequisicaoCompra, ItensReqCompra, PedidoCompra, ItensPedCompra, ItemRCItemPC, Produtos WHERE PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND ItensPedCompra.NumIntDoc = ItemRCItemPC.ItemPC AND ItemRCItemPC.ItemRC = ItensReqCompra.NumIntDoc AND ItensReqCompra.ReqCompra = RequisicaoCompra.NumIntDoc AND RequisicaoCompra.Codigo = ? AND PedidoCompra.FilialEmpresa = ? AND ItensReqCompra.Status <> ? AND ItensReqCompra.Produto = Produtos.Codigo" _
        & " ORDER BY PedidoCompra.Codigo, PedidoCompra.FilialEmpresa, RequisicaoCompra.Codigo, RequisicaoCompra.Urgente, ItensReqCompra.Produto", _
        sSiglaUMCompra, iClasseUM, tItemReqCompraInfo.lPedCompra, tItemReqCompraInfo.lReqCompra, dQuantidadePedido, dQuantRecebidaPedido, tItemReqCompraInfo.lNumIntDoc, tItemReqCompraInfo.sProduto, tItemReqCompraInfo.sDescProduto, tItemReqCompraInfo.sUM, tItemReqCompraInfo.dQuantPedida, tItemReqCompraInfo.iFilialEmpresa, tItemReqCompraInfo.lUrgente, tItemReqCompraInfo.dQuantidade, tItemReqCompraInfo.dQuantRecebida, tItemReqCompraInfo.dQuantCancelada, tItemReqCompraInfo.iTipoTributacao, tItemReqCompraInfo.lNumIntDocItemPC, objRequisicaoCompras.lCodigo, objRequisicaoCompras.iFilialEmpresa, ITEM_REQ_RECEBIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 66103

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66104

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66105

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objItemReqCompraInfo = New ClassItemReqComprasInfo

        lPedCompra = tItemReqCompraInfo.lPedCompra
        lReqCompra = tItemReqCompraInfo.lReqCompra
        
        iContPC = iContPC + 1
        iContRC = iContRC + 1
        
        objItemReqCompraInfo.lPedCompra = tItemReqCompraInfo.lPedCompra
        objItemReqCompraInfo.iItemPC = iContPC
        objItemReqCompraInfo.sProduto = tItemReqCompraInfo.sProduto
        objItemReqCompraInfo.sDescProduto = tItemReqCompraInfo.sDescProduto
        objItemReqCompraInfo.sUM = tItemReqCompraInfo.sUM
        objItemReqCompraInfo.iFilialEmpresa = tItemReqCompraInfo.iFilialEmpresa
        objItemReqCompraInfo.lReqCompra = tItemReqCompraInfo.lReqCompra
        objItemReqCompraInfo.iUrgente = tItemReqCompraInfo.lUrgente
        objItemReqCompraInfo.iItemRC = iContRC
        objItemReqCompraInfo.dQuantReceber = tItemReqCompraInfo.dQuantidade - tItemReqCompraInfo.dQuantRecebida
        
        lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUMCompra, tItemReqCompraInfo.sUM, dFator)
        If lErro <> SUCESSO Then gError 66323
        
        objItemReqCompraInfo.iClasseUM = iClasseUM
        objItemReqCompraInfo.sSiglaUMCompra = sSiglaUMCompra
        objItemReqCompraInfo.dFator = dFator
        objItemReqCompraInfo.dQuantRecebida = tItemReqCompraInfo.dQuantPedida
        objItemReqCompraInfo.dQuantRecebidoPedido = dFator * (dQuantidadePedido - dQuantRecebidaPedido)
        
        objItemReqCompraInfo.lNumIntDoc = tItemReqCompraInfo.lNumIntDoc
        objItemReqCompraInfo.iTipoTributacao = tItemReqCompraInfo.iTipoTributacao
        objItemReqCompraInfo.lNumIntDocItemPC = tItemReqCompraInfo.lNumIntDocItemPC
        
        colItemRCInfo.Add objItemReqCompraInfo

        'Busca no BD o próximo Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66106
            
        If lPedCompra <> tItemReqCompraInfo.lPedCompra Then iContPC = 0
        If lReqCompra <> tItemReqCompraInfo.lReqCompra Then iContRC = 0
    
    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItensRequisicao_Le_PedidoCompra = SUCESSO

    Exit Function

Erro_ItensRequisicao_Le_PedidoCompra:

    ItensRequisicao_Le_PedidoCompra = gErr

    Select Case gErr

        Case 66102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66103, 66104, 66106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 66105 'Erro tratado na rotina chamadora

        Case 66323

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149673)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCompraTodos_Le_NFiscal(ByVal objNFiscal As ClassNFiscal, ByVal colPedidoCompra As Collection) As Long
'Lê os Pedidos de Compra Relacionados aos itens da Nota fiscal passada por objNFical

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim lCodigo As Long
Dim iFilialEmpresa As Integer
Dim objPedidoCompra As ClassPedidoCompras

On Error GoTo Erro_PedidoCompraTodos_Le_NFiscal

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 65853

    'Pesquisa os Pedidos de Compras associados a Nota Fiscal a partir do NumIntDoc da NF
    lErro = Comando_Executar(lComando, "SELECT DISTINCT PedidoCompraTodos.NumIntDoc, PedidoCompraTodos.Codigo, PedidoCompraTodos.FilialEmpresa FROM PedidoCompraTodos, ItensPedCompraTodos, ItemNFItemPC, ItensNFiscal, NFiscal WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItensNFiscal.NumIntDoc = ItemNFItemPC.ItemNFiscal AND ItensPedCompraTodos.NumIntDoc = ItemNFItemPC.ItemPedCompra AND ItensPedCompraTodos.PedCompra = PedidoCompraTodos.NumIntDoc AND NFiscal.NumIntDoc = ? ORDER BY PedidoCompraTodos.Codigo", _
        lNumIntDoc, lCodigo, iFilialEmpresa, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 65854

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65855
    
    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66117
            
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objPedidoCompra = New ClassPedidoCompras

        objPedidoCompra.lNumIntDoc = lNumIntDoc
        objPedidoCompra.lCodigo = lCodigo
        objPedidoCompra.iFilialEmpresa = iFilialEmpresa

        colPedidoCompra.Add objPedidoCompra

        'Busca no BD o próximo Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65857
    
    Loop
        
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    PedidoCompraTodos_Le_NFiscal = SUCESSO

    Exit Function

Erro_PedidoCompraTodos_Le_NFiscal:

    PedidoCompraTodos_Le_NFiscal = gErr

    Select Case gErr

        Case 65853
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65854, 65855, 65857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
            
        Case 66117 'Tratado na rotina chamadora
                                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149674)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PedidoCompra_Le_NFiscal(ByVal objNFiscal As ClassNFiscal, ByVal colPedidoCompra As Collection) As Long
'Lê os Pedidos de Compra não baixados Relacionados aos itens da Nota fiscal passada por objNFical

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim lCodigo As Long
Dim iFilialEmpresa As Integer
Dim objPedidoCompra As ClassPedidoCompras

On Error GoTo Erro_PedidoCompra_Le_NFiscal

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66133

    'Pesquisa os Pedidos de Compras associados a Nota Fiscal a partir do NumIntDoc da NF
    lErro = Comando_Executar(lComando, "SELECT DISTINCT PedidoCompra.NumIntDoc, PedidoCompra.Codigo, PedidoCompra.FilialEmpresa FROM PedidoCompra WHERE DataEnvio <> ? AND Fornecedor = ? AND Filial = ? AND FilialEmpresa = ? ORDER BY PedidoCompra.Codigo", _
        lNumIntDoc, lCodigo, iFilialEmpresa, DATA_NULA, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iFilialPedido)
    If lErro <> AD_SQL_SUCESSO Then gError 66134

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66135
    
    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66136
            
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objPedidoCompra = New ClassPedidoCompras

        objPedidoCompra.lNumIntDoc = lNumIntDoc
        objPedidoCompra.lCodigo = lCodigo
        objPedidoCompra.iFilialEmpresa = iFilialEmpresa

        colPedidoCompra.Add objPedidoCompra

        'Busca no BD o próximo Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66137
    
    Loop
        
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    PedidoCompra_Le_NFiscal = SUCESSO

    Exit Function

Erro_PedidoCompra_Le_NFiscal:

    PedidoCompra_Le_NFiscal = gErr

    Select Case gErr

        Case 66133
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66134, 66135, 66137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
            
        Case 66136 'Tratado na rotina chamadora
                                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149675)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ItensPedCompraTodos_Le_NFiscal(ByVal objNFiscal As ClassNFiscal, ByVal colItemPedCompraInfo As Collection) As Long
'Lê todos os itens de Pedido de compras (baixados ou não) relacionados a nota fiscal e os devolve em colItemPedCompraInfo

Dim lErro As Long
Dim lComando As Long
Dim objItemPedCompraInfo As ClassItemPedCompraInfo
Dim tItemPedCompraInfo As typeItemPedidoCompraInfo
Dim dQuantidadeNFPC As Double
Dim dPrecoUnitario As Double

On Error GoTo Erro_ItensPedCompraTodos_Le_NFiscal

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 65858
        
    tItemPedCompraInfo.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedCompraInfo.sProduto = String(STRING_PRODUTO, 0)
    tItemPedCompraInfo.sUM = String(STRING_UM_SIGLA, 0)
    
    'Pesquisa os Itens de Pedidos de Compras relacionados a Nota Fiscal a partir do NumIntDoc de NFiscal
    lErro = Comando_Executar(lComando, "SELECT DISTINCT PedidoCompraTodos.Moeda, ItemNFItemPC.Taxa, ItensPedCompraTodos.PrecoUnitario, ItensPedCompraTodos.NumIntDoc, PedidoCompraTodos.Codigo, ItensPedCompraTodos.Produto, ItensPedCompraTodos.DescProduto, ItensPedCompraTodos.Quantidade, ItensPedCompraTodos.QuantRecebida, ItensPedCompraTodos.UM, ItensPedCompraTodos.AliquotaICMS, ItensPedCompraTodos.AliquotaIPI, ItemNFItemPC.Quantidade FROM PedidoCompraTodos, ItensPedCompraTodos, ItemNFItemPC, ItensNFiscal, NFiscal WHERE PedidoCompraTodos.NumIntDoc = ItensPedCompraTodos.PedCompra AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItensNFiscal.NumIntDoc = ItemNFItemPC.ItemNFiscal AND ItensPedCompraTodos.NumIntDoc = ItemNFItemPC.ItemPedCompra AND NFiscal.NumIntDoc = ? ORDER BY PedidoCompraTodos.Codigo, ItensPedCompraTodos.Produto", _
        tItemPedCompraInfo.iMoeda, tItemPedCompraInfo.dTaxa, dPrecoUnitario, tItemPedCompraInfo.lNumIntDoc, tItemPedCompraInfo.lPedCompra, tItemPedCompraInfo.sProduto, tItemPedCompraInfo.sDescProduto, tItemPedCompraInfo.dQuantidade, tItemPedCompraInfo.dQuantRecebida, tItemPedCompraInfo.sUM, tItemPedCompraInfo.dAliquotaICMS, tItemPedCompraInfo.dAliquotaIPI, dQuantidadeNFPC, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 65859

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65860

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 65861
        
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objItemPedCompraInfo = New ClassItemPedCompraInfo

        objItemPedCompraInfo.lNumIntDoc = tItemPedCompraInfo.lNumIntDoc
        objItemPedCompraInfo.lPedCompra = tItemPedCompraInfo.lPedCompra
        objItemPedCompraInfo.sProduto = tItemPedCompraInfo.sProduto
        objItemPedCompraInfo.sDescProduto = tItemPedCompraInfo.sDescProduto
        objItemPedCompraInfo.sUM = tItemPedCompraInfo.sUM
        objItemPedCompraInfo.dQuantReceber = tItemPedCompraInfo.dQuantidade
        objItemPedCompraInfo.dQuantRecebida = dQuantidadeNFPC
        objItemPedCompraInfo.dAliquotaICMS = tItemPedCompraInfo.dAliquotaICMS
        objItemPedCompraInfo.dAliquotaIPI = tItemPedCompraInfo.dAliquotaIPI
        objItemPedCompraInfo.dPrecoUnitario = dPrecoUnitario
        objItemPedCompraInfo.iMoeda = tItemPedCompraInfo.iMoeda
        objItemPedCompraInfo.dTaxa = tItemPedCompraInfo.dTaxa
        
        'Adiciona na coleção
        colItemPedCompraInfo.Add objItemPedCompraInfo

        'Busca no BD o próximo item de Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 65862

    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ItensPedCompraTodos_Le_NFiscal = SUCESSO

    Exit Function

Erro_ItensPedCompraTodos_Le_NFiscal:

    ItensPedCompraTodos_Le_NFiscal = gErr

    Select Case gErr

        Case 65858, 66130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 65859, 65860, 65862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
        
        Case 65861 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149676)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


Function ItensPedCompra_Le_NFiscal(ByVal objNFiscal As ClassNFiscal, ByVal colItemPedCompraInfo As Collection) As Long
'Lê os itens de Pedido de compras não baixados relacionados a nota fiscal e os devolve em colItemPedCompraInfo

Dim lErro As Long
Dim lComando As Long
Dim objItemPedCompraInfo As ClassItemPedCompraInfo
Dim tItemPedCompraInfo As typeItemPedidoCompraInfo
Dim dQuantidadeNFPC As Double

On Error GoTo Erro_ItensPedCompra_Le_NFiscal

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66138
        
    tItemPedCompraInfo.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedCompraInfo.sProduto = String(STRING_PRODUTO, 0)
    tItemPedCompraInfo.sUM = String(STRING_UM_SIGLA, 0)
    
    'Pesquisa os Itens de Pedidos de Compras relacionados a Nota Fiscal a partir do NumIntDoc de NFiscal
    lErro = Comando_Executar(lComando, "SELECT DISTINCT ItensPedCompra.NumIntDoc, ItensPedCompra.PedCompra, ItensPedCompra.Produto, ItensPedCompra.DescProduto, ItensPedCompra.Quantidade, ItensPedCompra.QuantRecebida, ItensPedCompra.UM, ItensPedCompra.AliquotaICMS, ItensPedCompra.AliquotaIPI, ItemNFItemPC.Quantidade FROM PedidoCompra, ItensPedCompra, NFiscal, ItensNFiscal, ItemNFItemPC WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND ItensNFiscal.NumIntDoc = ItemNFItemPC.ItemNFiscal AND ItensPedCompra.PedCompra = PedidoCompra.NumIntDoc AND ItensPedCompra.NumIntDoc = ItemNFItemPC.ItemPedCompra AND NFiscal.NumIntDoc = ? ORDER BY ItensPedCompra.PedCompra, ItensPedCompra.Produto", _
        tItemPedCompraInfo.lNumIntDoc, tItemPedCompraInfo.lPedCompra, tItemPedCompraInfo.sProduto, tItemPedCompraInfo.sDescProduto, tItemPedCompraInfo.dQuantidade, tItemPedCompraInfo.dQuantRecebida, tItemPedCompraInfo.sUM, tItemPedCompraInfo.dAliquotaICMS, tItemPedCompraInfo.dAliquotaIPI, dQuantidadeNFPC, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 66139

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66140

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66141
        
    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        Set objItemPedCompraInfo = New ClassItemPedCompraInfo

        objItemPedCompraInfo.lNumIntDoc = tItemPedCompraInfo.lNumIntDoc
        objItemPedCompraInfo.lPedCompra = tItemPedCompraInfo.lPedCompra
        objItemPedCompraInfo.sProduto = tItemPedCompraInfo.sProduto
        objItemPedCompraInfo.sDescProduto = tItemPedCompraInfo.sDescProduto
        objItemPedCompraInfo.sUM = tItemPedCompraInfo.sUM
        objItemPedCompraInfo.dQuantReceber = tItemPedCompraInfo.dQuantidade
        objItemPedCompraInfo.dQuantRecebida = dQuantidadeNFPC
        objItemPedCompraInfo.dAliquotaICMS = tItemPedCompraInfo.dAliquotaICMS
        objItemPedCompraInfo.dAliquotaIPI = tItemPedCompraInfo.dAliquotaIPI
        
        'Adiciona na coleção
        colItemPedCompraInfo.Add objItemPedCompraInfo

        'Busca no BD o próximo item de Pedido de Compras
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66142

    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ItensPedCompra_Le_NFiscal = SUCESSO

    Exit Function

Erro_ItensPedCompra_Le_NFiscal:

    ItensPedCompra_Le_NFiscal = gErr

    Select Case gErr

        Case 66138, 66130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66139, 66140, 66142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
        
        Case 66141 'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149677)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ItemPedidoCompra_Le_QuantSobra(ByVal objItemPedCompraInfo As ClassItemPedCompraInfo, ByVal iFilialEmpresa As Integer, dQuantidadeSobra As Double) As Long
'Lê a Soma das Quantidades de Itens para o Pedido de Compra desvinculado a Requisição de Compra
'iFilialEmpresa que originou o Pedido de Compra

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim dQuantidade As Double
Dim dQuantRecebida As Double
Dim dQuantidadeTotal As Double
Dim dQuantRecebidaTotal As Double

On Error GoTo Erro_ItemPedidoCompra_Le_QuantSobra

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 66108
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 66109
    
    'Lê a soma das quantidades de Itens de Pedido de Compra vinculados com Requisicao de Compras
    lErro = Comando_Executar(lComando, "SELECT ItemRCItemPC.Quantidade, ItemRCItemPC.QuantRecebida FROM ItensPedCompra, ItemRCItemPC WHERE ItemRCItemPC.ItemPC = ItensPedCompra.NumIntDoc AND ItensPedCompra.PedCompra = ? AND ItensPedCompra.Produto = ?", dQuantidade, dQuantRecebida, objItemPedCompraInfo.lPedCompra, objItemPedCompraInfo.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 66110
    
    'Buscar o primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66111
    
    'Lê a soma das quantidades de Itens de Pedido de Compra desvinculados com Requisicao de Compras
    lErro = Comando_Executar(lComando2, "SELECT ItensPedCompraTodos.Quantidade, ItensPedCompraTodos.QuantRecebida FROM ItensPedCompraTodos, PedidoCompraTodos WHERE ItensPedCompraTodos.Produto = ? AND ItensPedCompraTodos.PedCompra = PedidoCompraTodos.NumIntDoc AND PedidoCompraTodos.Codigo = ? AND PedidoCompraTodos.FilialEmpresa = ? AND ItensPedCompraTodos.NumIntDoc NOT IN (SELECT ItemPC FROM ItemRCItemPC)", dQuantidadeTotal, dQuantRecebidaTotal, objItemPedCompraInfo.sProduto, objItemPedCompraInfo.lPedCompra, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 66112

    'Buscar o primeiro
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66113

    'Retorna a quantidade de Sobra que é a Quantidade - Quantidade Recebida
    dQuantidadeSobra = (dQuantidadeTotal - dQuantRecebidaTotal) + (dQuantidade - dQuantRecebida)
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    ItemPedidoCompra_Le_QuantSobra = SUCESSO

    Exit Function

Erro_ItemPedidoCompra_Le_QuantSobra:

    ItemPedidoCompra_Le_QuantSobra = gErr

    Select Case gErr

        Case 66108, 66109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66110, 66111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 66112, 66113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149678)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function ReqCompras_Le_PedidoCompras(ByVal objPedidoCompra As ClassPedidoCompras, ByVal lReqCompra As Long, ByVal colPedCompras As Collection) As Long
'Lê os Pedidos de Compras que possuem associação com uma determinada Requisição passada em lReqCompra

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long

On Error GoTo Erro_ReqCompras_Le_PedidoCompras
    
    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 65856
    
    'Pesquisa o vínculo entre Item Pedido de Compras e Item de Requisição de Compras
    lErro = Comando_Executar(lComando, "SELECT DISTINCT PedidoCompra.Codigo FROM RequisicaoCompra, ItensReqCompra, PedidoCompra, ItensPedCompra, ItemRCItemPC WHERE PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND ItensPedCompra.NumIntDoc = ItemRCItemPC.ItemPC AND ItemRCItemPC.ItemRC = ItensReqCompra.NumIntDoc AND ItensReqCompra.ReqCompra = RequisicaoCompra.NumIntDoc AND RequisicaoCompra.Codigo = ? AND PedidoCompra.Codigo = ? AND PedidoCompra.FilialEmpresa = ? ", lCodigo, lReqCompra, objPedidoCompra.lCodigo, objPedidoCompra.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 66130

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66128

    'Se não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 66119
                   
    Do While lErro = AD_SQL_SUCESSO
        
        'Guarda o código do Pedido associado a Requisição
        colPedCompras.Add lCodigo
    
        'Busca próximo vínculo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 66127
        
    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    ReqCompras_Le_PedidoCompras = SUCESSO

    Exit Function

Erro_ReqCompras_Le_PedidoCompras:

    ReqCompras_Le_PedidoCompras = gErr

    Select Case gErr

        Case 65856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 66119
        
        Case 66127, 66128, 66130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149679)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovEstoque_Le_Produto(ByVal objItemMovEstoque As ClassItemMovEstoque) As Long
'Lê o MovEstoque a partir do Código, Filial Empresa e Produto.

Dim lErro As Long
Dim lComando As Long
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovEstoque_Le_Produto
    
    'Abertura de comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 78186
    
    tItemMovEstoque.sProduto = String(STRING_PRODUTO, 0)
    tItemMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemMovEstoque.sCcl = String(STRING_CCL, 0)
    tItemMovEstoque.sOPCodigo = String(STRING_OPCODIGO, 0)
    tItemMovEstoque.sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)
    tItemMovEstoque.sContaContabilEst = String(STRING_CONTA, 0)
    tItemMovEstoque.sContaContabilAplic = String(STRING_CONTA, 0)
    
    'Le o Item do Movimento com o Filtro: Produto, FilialEmpresa e Código
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo, NumIntDoc, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic FROM MovimentoEstoque WHERE FilialEmpresa = ? AND Codigo = ? AND Produto = ? AND NumIntDocEst = 0", tItemMovEstoque.iFilialEmpresa, tItemMovEstoque.lCodigo, tItemMovEstoque.lNumIntDoc, tItemMovEstoque.dCusto, tItemMovEstoque.iApropriacao, tItemMovEstoque.sProduto, tItemMovEstoque.sSiglaUM, tItemMovEstoque.dQuantidade, tItemMovEstoque.iAlmoxarifado, tItemMovEstoque.iTipoMov, tItemMovEstoque.lNumIntDocOrigem, tItemMovEstoque.iTipoNumIntDocOrigem, tItemMovEstoque.dtData, tItemMovEstoque.sCcl, tItemMovEstoque.lNumIntDocEst, tItemMovEstoque.lCliente, tItemMovEstoque.lFornecedor, tItemMovEstoque.sOPCodigo, tItemMovEstoque.sDocOrigem, tItemMovEstoque.sContaContabilEst, tItemMovEstoque.sContaContabilAplic, _
    objItemMovEstoque.iFilialEmpresa, objItemMovEstoque.lCodigo, objItemMovEstoque.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 78187

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78188

    If lErro = AD_SQL_SEM_DADOS Then gError 78189

    objItemMovEstoque.iFilialEmpresa = tItemMovEstoque.iFilialEmpresa
    objItemMovEstoque.lCodigo = tItemMovEstoque.lCodigo
    objItemMovEstoque.lNumIntDoc = tItemMovEstoque.lNumIntDoc
    objItemMovEstoque.dCusto = tItemMovEstoque.dCusto
    objItemMovEstoque.iApropriacao = tItemMovEstoque.iApropriacao
    objItemMovEstoque.sProduto = tItemMovEstoque.sProduto
    objItemMovEstoque.sSiglaUM = tItemMovEstoque.sSiglaUM
    objItemMovEstoque.dQuantidade = tItemMovEstoque.dQuantidade
    objItemMovEstoque.iAlmoxarifado = tItemMovEstoque.iAlmoxarifado
    objItemMovEstoque.iTipoMov = tItemMovEstoque.iTipoMov
    objItemMovEstoque.lNumIntDocOrigem = tItemMovEstoque.lNumIntDocOrigem
    objItemMovEstoque.iTipoNumIntDocOrigem = tItemMovEstoque.iTipoNumIntDocOrigem
    objItemMovEstoque.dtData = tItemMovEstoque.dtData
    objItemMovEstoque.sCcl = tItemMovEstoque.sCcl
    objItemMovEstoque.lNumIntDocEst = tItemMovEstoque.lNumIntDocEst
    objItemMovEstoque.lCliente = tItemMovEstoque.lCliente
    objItemMovEstoque.lFornecedor = tItemMovEstoque.lFornecedor
    objItemMovEstoque.sOPCodigo = tItemMovEstoque.sOPCodigo
    objItemMovEstoque.sDocOrigem = tItemMovEstoque.sDocOrigem
    objItemMovEstoque.sContaContabilEst = tItemMovEstoque.sContaContabilEst
    objItemMovEstoque.sContaContabilAplic = tItemMovEstoque.sContaContabilAplic

    Call Comando_Fechar(lComando)

    MovEstoque_Le_Produto = SUCESSO

    Exit Function

Erro_MovEstoque_Le_Produto:

    MovEstoque_Le_Produto = gErr

    Select Case gErr

        Case 78186
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 78187, 78188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 78189 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149680)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovEstoque_Le_NumIntDoc(ByVal objItemMovEstoque As ClassItemMovEstoque) As Long
'Lê o MovEstoque a partir do NumIntDoc.

Dim lErro As Long
Dim lComando As Long
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovEstoque_Le_NumIntDoc
    
    'Abertura de comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 78190
    
    tItemMovEstoque.sProduto = String(STRING_PRODUTO, 0)
    tItemMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemMovEstoque.sCcl = String(STRING_CCL, 0)
    tItemMovEstoque.sOPCodigo = String(STRING_OPCODIGO, 0)
    tItemMovEstoque.sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)
    tItemMovEstoque.sContaContabilEst = String(STRING_CONTA, 0)
    tItemMovEstoque.sContaContabilAplic = String(STRING_CONTA, 0)
    
    'Le o Item do Movimento com o Filtro: Produto, FilialEmpresa e Código
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo, NumIntDoc, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic FROM MovimentoEstoque WHERE NumIntDoc = ? AND NumIntDocEst = 0", tItemMovEstoque.iFilialEmpresa, tItemMovEstoque.lCodigo, tItemMovEstoque.lNumIntDoc, tItemMovEstoque.dCusto, tItemMovEstoque.iApropriacao, tItemMovEstoque.sProduto, tItemMovEstoque.sSiglaUM, tItemMovEstoque.dQuantidade, tItemMovEstoque.iAlmoxarifado, tItemMovEstoque.iTipoMov, tItemMovEstoque.lNumIntDocOrigem, tItemMovEstoque.iTipoNumIntDocOrigem, tItemMovEstoque.dtData, tItemMovEstoque.sCcl, tItemMovEstoque.lNumIntDocEst, tItemMovEstoque.lCliente, tItemMovEstoque.lFornecedor, tItemMovEstoque.sOPCodigo, tItemMovEstoque.sDocOrigem, tItemMovEstoque.sContaContabilEst, tItemMovEstoque.sContaContabilAplic, _
    objItemMovEstoque.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 78191

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78192

    If lErro = AD_SQL_SEM_DADOS Then gError 78193

    objItemMovEstoque.iFilialEmpresa = tItemMovEstoque.iFilialEmpresa
    objItemMovEstoque.lCodigo = tItemMovEstoque.lCodigo
    objItemMovEstoque.lNumIntDoc = tItemMovEstoque.lNumIntDoc
    objItemMovEstoque.dCusto = tItemMovEstoque.dCusto
    objItemMovEstoque.iApropriacao = tItemMovEstoque.iApropriacao
    objItemMovEstoque.sProduto = tItemMovEstoque.sProduto
    objItemMovEstoque.sSiglaUM = tItemMovEstoque.sSiglaUM
    objItemMovEstoque.dQuantidade = tItemMovEstoque.dQuantidade
    objItemMovEstoque.iAlmoxarifado = tItemMovEstoque.iAlmoxarifado
    objItemMovEstoque.iTipoMov = tItemMovEstoque.iTipoMov
    objItemMovEstoque.lNumIntDocOrigem = tItemMovEstoque.lNumIntDocOrigem
    objItemMovEstoque.iTipoNumIntDocOrigem = tItemMovEstoque.iTipoNumIntDocOrigem
    objItemMovEstoque.dtData = tItemMovEstoque.dtData
    objItemMovEstoque.sCcl = tItemMovEstoque.sCcl
    objItemMovEstoque.lNumIntDocEst = tItemMovEstoque.lNumIntDocEst
    objItemMovEstoque.lCliente = tItemMovEstoque.lCliente
    objItemMovEstoque.lFornecedor = tItemMovEstoque.lFornecedor
    objItemMovEstoque.sOPCodigo = tItemMovEstoque.sOPCodigo
    objItemMovEstoque.sDocOrigem = tItemMovEstoque.sDocOrigem
    objItemMovEstoque.sContaContabilEst = tItemMovEstoque.sContaContabilEst
    objItemMovEstoque.sContaContabilAplic = tItemMovEstoque.sContaContabilAplic

    Call Comando_Fechar(lComando)

    MovEstoque_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_MovEstoque_Le_NumIntDoc:

    MovEstoque_Le_NumIntDoc = gErr

    Select Case gErr

        Case 78190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 78191, 78192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 78193 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149681)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ApropriacaoInsumo_Le_NumIntDocOrigem(ByVal lNumIntDocOrigem As Long, ByVal colApropriacaoInsumo As Collection) As Long
'Lê a tabela de RastreamentoMovto através do Movimento de Estoque

Dim lErro As Long
Dim tApropricacaoInsumosProd As typeApropricacaoInsumosProd
Dim objApropriacaoInsumo As New ClassApropriacaoInsumosProd
Dim lComando As Long

On Error GoTo Erro_ApropriacaoInsumo_Le_NumIntDocOrigem
    
    'Abertura de comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 78430
    
    tApropricacaoInsumosProd.sProduto = String(STRING_PRODUTO, 0)
    
    'Lê o Rastreamento Movto
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, NumIntDocOrigem, NumIntReqProd, Produto, Quantidade, Automatico FROM ApropriacaoInsumosProd WHERE NumIntDocOrigem = ?", tApropricacaoInsumosProd.lNumIntDoc, tApropricacaoInsumosProd.lNumIntDocOrigem, tApropricacaoInsumosProd.lNumIntReqProd, tApropricacaoInsumosProd.sProduto, tApropricacaoInsumosProd.dQuantidade, tApropricacaoInsumosProd.iAutomatico, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 78431
           
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78432

    Do While lErro = AD_SQL_SUCESSO
        
        Set objApropriacaoInsumo = New ClassApropriacaoInsumosProd
        
        'passa para o objeto
        objApropriacaoInsumo.lNumIntDoc = tApropricacaoInsumosProd.lNumIntDoc
        objApropriacaoInsumo.lNumIntDocOrigem = tApropricacaoInsumosProd.lNumIntDocOrigem
        objApropriacaoInsumo.lNumIntReqProd = tApropricacaoInsumosProd.lNumIntReqProd
        objApropriacaoInsumo.sProduto = tApropricacaoInsumosProd.sProduto
        objApropriacaoInsumo.dQuantidade = tApropricacaoInsumosProd.dQuantidade
        objApropriacaoInsumo.iAutomatico = tApropricacaoInsumosProd.iAutomatico

        colApropriacaoInsumo.Add objApropriacaoInsumo
                
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78433
    
    Loop
            
    Call Comando_Fechar(lComando)
    
    ApropriacaoInsumo_Le_NumIntDocOrigem = SUCESSO
        
    Exit Function
    
Erro_ApropriacaoInsumo_Le_NumIntDocOrigem:

    ApropriacaoInsumo_Le_NumIntDocOrigem = gErr
    
    Select Case gErr
            
        Case 78430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 78431, 78432, 78433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_RASTREAMENTOMOVTO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149682)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ReqProducao_TestaExistencia(ByVal objItemMovEst As ClassItemMovEstoque) As Long
'Pesquisa no bd a ordem de producao a partir de FilialEmpresa, Codigo e Data de Emissão
'Função alterada por Wagner 09/09/05 (Validação a nível de item de OP)

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ReqProducao_TestaExistencia

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 71414

    'Leitura da tabela de Ordens de Produção
    lErro = Comando_Executar(lComando, "SELECT ME.NumIntDoc FROM MovimentoEstoque AS ME, ItensOrdemProducao AS IOP WHERE IOP.NumIntDoc = ME.NumIntDocOrigem AND IOP.FilialEmpresa = ? AND (ME.TipoMov = ? OR ME.TipoMov = ?) And ME.TipoNumIntDocOrigem = ? AND IOP.Codigo = ? AND IOP.Produto = ?", lNumIntDoc, giFilialEmpresa, MOV_EST_REQ_PRODUCAO, MOV_EST_REQ_PRODUCAO_BENEF3, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, objItemMovEst.sOPCodigo, objItemMovEst.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71415

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71416
    
    If lErro <> AD_SQL_SUCESSO Then gError 71417

   'Fechamento comando
    Call Comando_Fechar(lComando)

    ReqProducao_TestaExistencia = SUCESSO

    Exit Function

Erro_ReqProducao_TestaExistencia:

    ReqProducao_TestaExistencia = gErr

    Select Case gErr

        Case 71414
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 71415, 71416
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 71417

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149683)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PedidoCompras_Le_Valores1(ByVal colPedidoCompras As Collection) As Long
'Lê valores dos Pedido de compras passados na coleção colPedidoCompras

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim objPedidoCompras As ClassPedidoCompras
Dim dValorFrete As Double
Dim dValorSeguro As Double
Dim dValorDespesas As Double
Dim dValorDescontos As Double

On Error GoTo Erro_PedidoCompras_Le_Valores1
    
    'Abertura do comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 89221
    
    'Abertura do comando
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 89222
    
    For Each objPedidoCompras In colPedidoCompras
    
        'Pesquisa os Pedidos de Compras a partir do código e FilialEmpresa passados
        lErro = Comando_Executar(lComando1, "SELECT ValorFrete, ValorSeguro, OutrasDespesas, ValorDesconto FROM PedidoCompra WHERE Codigo = ? AND FilialEmpresa = ?", dValorFrete, dValorSeguro, dValorDespesas, dValorDescontos, objPedidoCompras.lCodigo, objPedidoCompras.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 89223
    
        lErro = Comando_BuscarPrimeiro(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89224
    
        'Se não encontrou --> Erro
        If lErro = AD_SQL_SUCESSO Then
        
            'Verifica se o Pedido de Compras está baixado
            lErro = Comando_Executar(lComando2, "SELECT ValorFrete FROM PedidoCompraBaixado WHERE Codigo = ? AND FilialEmpresa = ?", dValorFrete, objPedidoCompras.lCodigo, objPedidoCompras.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 89225

            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89226

            'Se encontrou --> pedido baixado
            If lErro = AD_SQL_SUCESSO Then
                gError 89227
            Else
                'pedido não encontrado
                gError 89229
            End If
    
        End If
        
        objPedidoCompras.dValorFrete = dValorFrete
        objPedidoCompras.dValorSeguro = dValorSeguro
        objPedidoCompras.dOutrasDespesas = dValorDespesas
        objPedidoCompras.dValorDesconto = dValorDescontos
        
    Next

    'Fecha o comando
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    PedidoCompras_Le_Valores1 = SUCESSO

    Exit Function

Erro_PedidoCompras_Le_Valores1:

    PedidoCompras_Le_Valores1 = gErr

    Select Case gErr

        Case 89221, 89222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89223, 89224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", gErr)

        Case 89225, 89226
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRABAIXADO", gErr)
            
        Case 89227
            
        Case 89229
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149684)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function PedidoCompra_Le_EnvComQuantReceber(ByVal colPedidos As Collection, ByVal lCodigoFornecedor As Long, ByVal iFilial As Integer, ByVal iFilialCompra As Integer) As Long
'Le os pedidos de compra enviados para o fornecedor em questão com quantidade a receber

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim lCodigo As Long
Dim lNumIntDoc As Long
Dim iFilialEmpresa As Integer
Dim objPedidoCompra As ClassPedidoCompras

On Error GoTo Erro_PedidoCompra_Le_EnvComQuantReceber

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89267
        
    'Verifica se o Pedido tem pelo menos um item com quantidade a receber (Status aberto)
    lErro = Comando_Executar(lComando, "SELECT DISTINCT PedidoCompra.NumIntDoc, PedidoCompra.Codigo, PedidoCompra.FilialEmpresa FROM ItensPedCompra, PedidoCompra WHERE ItensPedCompra.PedCompra = PedidoCompra.NumIntDoc AND Fornecedor = ? AND Filial = ? AND FilialEmpresa = ? AND DataEnvio <> ? AND ItensPedCompra.Status = ? AND (Quantidade - QuantRecebida -QuantRecebimento)>0 ORDER BY PedidoCompra.Codigo", lNumIntDoc, lCodigo, iFilialEmpresa, lCodigoFornecedor, iFilial, iFilialCompra, DATA_NULA, ITEM_PED_COMPRAS_ABERTO)
    If lErro <> AD_SQL_SUCESSO Then gError 89268

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89269
    
    Do While lErro = AD_SQL_SUCESSO
        
        Set objPedidoCompra = New ClassPedidoCompras

        objPedidoCompra.lNumIntDoc = lNumIntDoc
        objPedidoCompra.lCodigo = lCodigo
        objPedidoCompra.iFilialEmpresa = iFilialEmpresa
        
        'Remove o Pedido de compras da coleção
        colPedidos.Add objPedidoCompra
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89270
    
    Loop
        
    'Fecha o comando
    Call Comando_Fechar(lComando)

    PedidoCompra_Le_EnvComQuantReceber = SUCESSO

    Exit Function
    
Erro_PedidoCompra_Le_EnvComQuantReceber:
    
    PedidoCompra_Le_EnvComQuantReceber = gErr
    
    Select Case gErr
    
        Case 89267
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 89268, 89269, 89270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDCOMPRA", gErr)
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149685)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Embalagem_Le(ByVal objEmbalagem As ClassEmbalagem) As Long
'Lê a embalagem a partir do código passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim tEmbalagem As typeEmbalagem

On Error GoTo Erro_Embalagem_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 82760
    
    tEmbalagem.sDescricao = String(STRING_EMBALAGEM_DESCRICAO, 0)
    tEmbalagem.sSigla = String(STRING_EMBALAGEM_SIGLA, 0)
    tEmbalagem.sProduto = String(STRING_PRODUTO, 0)
    
    'Busca no BD uma embalagem com mesmo código
    lErro = Comando_Executar(lComando, "SELECT Codigo,Descricao, Sigla, Capacidade, Peso, Produto FROM Embalagens WHERE Codigo=?", tEmbalagem.iCodigo, tEmbalagem.sDescricao, tEmbalagem.sSigla, tEmbalagem.dCapacidade, tEmbalagem.dPeso, tEmbalagem.sProduto, objEmbalagem.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 82761
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82762
    
    'Se nao encontrou ==> erro
    If lErro <> AD_SQL_SUCESSO Then gError 82763
    
    objEmbalagem.dCapacidade = tEmbalagem.dCapacidade
    objEmbalagem.dPeso = tEmbalagem.dPeso
    objEmbalagem.iCodigo = tEmbalagem.iCodigo
    objEmbalagem.sDescricao = tEmbalagem.sDescricao
    objEmbalagem.sSigla = tEmbalagem.sSigla
    objEmbalagem.sProduto = tEmbalagem.sProduto
    
    'Por Leo
    lErro = EmbalagensExpedicao_Le(objEmbalagem)
    If lErro <> AD_SQL_SUCESSO Then gError 103444
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Embalagem_Le = SUCESSO
    
    Exit Function
    
Erro_Embalagem_Le:

    Embalagem_Le = gErr
    
    Select Case gErr
    
        Case 82760
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 82761, 82762
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)
            
        Case 82763
            'Erro tratado na rotina chamadora
            
        Case 103444
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149686)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Embalagem_Le_Todas(ByVal colEmbalagem As Collection) As Long
'Lê todas as embalagens e guarda em colEmbalagem

Dim lErro As Long
Dim lComando As Long
Dim objEmbalagem As New ClassEmbalagem
Dim tEmbalagem As typeEmbalagem

On Error GoTo Erro_Embalagem_Le_Todas

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 82727

    tEmbalagem.sDescricao = String(STRING_EMBALAGEM_DESCRICAO, 0)
    tEmbalagem.sSigla = String(STRING_EMBALAGEM_SIGLA, 0)

    'Seleciona todas as embalagens
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao,Sigla,Capacidade, Peso FROM Embalagens ORDER BY Codigo", tEmbalagem.iCodigo, tEmbalagem.sDescricao, tEmbalagem.sSigla, tEmbalagem.dCapacidade, tEmbalagem.dPeso)
    If lErro <> AD_SQL_SUCESSO Then gError 82728

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82729
    
    'Se não encontrou a embalagem ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 82731
    
    Do While lErro = AD_SQL_SUCESSO

        'Guarda na coleção as embalagens
        Set objEmbalagem = New ClassEmbalagem

        objEmbalagem.dCapacidade = tEmbalagem.dCapacidade
        objEmbalagem.dPeso = tEmbalagem.dPeso
        objEmbalagem.iCodigo = tEmbalagem.iCodigo
        objEmbalagem.sDescricao = tEmbalagem.sDescricao
        objEmbalagem.sSigla = tEmbalagem.sSigla
        
        'Adiciona na colecao
        colEmbalagem.Add objEmbalagem

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82730

    Loop

    'fecha o comando
    Call Comando_Fechar(lComando)

    Embalagem_Le_Todas = SUCESSO

    Exit Function

Erro_Embalagem_Le_Todas:

    Embalagem_Le_Todas = gErr

    Select Case gErr

        Case 82727
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 82728, 82729, 82730
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr)
        
        Case 82731
            'Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149687)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Calcula_MnemonicoFPreco(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal sProduto As String, Optional ByVal objExeExp As Object) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dCusto As Double
Dim sTipo As String
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim dFator As Double
Dim dtDataUltEnt As Date
Dim dAliquotaICMSUltEnt As Double
Dim dFrete As Double
Dim dAliquota As Double

On Error GoTo Erro_Calcula_MnemonicoFPreco

    Select Case objMnemonicoValor.sMnemonico
                       
        Case MNEMONICOFPRECO_CUSTO_PRODUCAO
            
            If Len(Trim(sProduto)) > 0 Then
            
                objProduto.sCodigo = sProduto
                
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
                If lErro = 28030 Then gError 92417
    
                sTipo = STRING_QUANT_DISPONIVEL_NOSSA
    
                'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
                lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, sTipo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'Incluído por Luiz Nogueira em 15/04/04
                'O custo retornado está na U.M. de estoque
                'Para calcular preço de venda é preciso calcular o custo na U.M. venda
                lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMVenda, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                objMnemonicoValor.colValor.Add dCusto * dFator
                    
            Else
            
                objMnemonicoValor.colValor.Add 0
                    
            End If
                    
        Case MNEMONICOFPRECO_CUSTO_PRODUCAO_PROD
            
            lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProdutoFormatado, iPreenchido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            objProduto.sCodigo = sProdutoFormatado
            
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

            If lErro = 28030 Then gError 92420

            sTipo = STRING_QUANT_DISPONIVEL_NOSSA

            'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
            lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, sTipo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
            objMnemonicoValor.colValor.Add dCusto
            
        Case MNEMONICOFPRECO_CUSTOROTEIROINSUMOSMAQ
            lErro = CF("CustoRoteiro_InsumosMaquina", objMnemonicoValor, objExeExp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Case MNEMONICOFPRECO_CUSTOROTEIROMAOOBRA
            lErro = CF("CustoRoteiro_MaoDeObra", objMnemonicoValor, objExeExp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Case MNEMONICOFPRECO_CUSTOMP
            lErro = CF("CustoMP", objMnemonicoValor, objExeExp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Case MNEMONICOFPRECO_CUSTO_KIT_MP_PARAM
            lErro = CF("Custo_Kit_MP", objMnemonicoValor, sProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Case MNEMONICOFPRECO_CUSTO_KIT_MP
            lErro = CF("Custo_Kit_MP", objMnemonicoValor, sProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Case MNEMONICOFPRECO_CUSTO_KIT_MP2_PARAM
            lErro = CF("Custo_Kit_MP2", objMnemonicoValor, sProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Case MNEMONICOFPRECO_CUSTOREP_KIT_MP_PARAM
            lErro = CF("CustoRep_Kit_MP", objMnemonicoValor, sProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Case MNEMONICOFPRECO_CUSTOREP_KIT_MP
            lErro = CF("CustoRep_Kit_MP", objMnemonicoValor, sProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Case MNEMONICOFPRECO_CUSTOEMB
            lErro = CF("CustoEmb", objMnemonicoValor, objExeExp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Case MNEMONICOFPRECO_ICMSALIQTABPRECO
            lErro = CF("ICMSAliqTabPreco", objMnemonicoValor, objExeExp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Case MNEMONICOFPRECO_CUSTODIRETO
            lErro = CF("CustoDireto", objMnemonicoValor, objExeExp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Case MNEMONICOFPRECO_CUSTOFIXO
            lErro = CF("CustoFixo", objMnemonicoValor, objExeExp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Case MNEMONICOFPRECO_VALORCATEGORIA
            lErro = CF("FPreco_ObtemValorCateg", objMnemonicoValor, objExeExp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
               
        Case MNEMONICOFPRECO_FRETEFILCLI
            lErro = CF("FPreco_ObtemFreteFilCli", objMnemonicoValor, objExeExp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Case MNEMONICOFPRECO_CUSTO_ULT_ENTRADA
            
            If Len(Trim(sProduto)) > 0 Then
            
                objProduto.sCodigo = sProduto
                
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
                If lErro = 28030 Then gError 178581
    
                lErro = CF("Produto_ObterDadosUltEnt", objProduto, giFilialEmpresa, dCusto, dtDataUltEnt, dAliquotaICMSUltEnt, dFrete)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                'O custo retornado está na U.M. de estoque
                'Para calcular preço de venda é preciso calcular o custo na U.M. venda
                lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMVenda, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                objMnemonicoValor.colValor.Add dCusto * dFator
                    
            Else
            
                objMnemonicoValor.colValor.Add 0
                    
            End If
            
        Case MNEMONICOFPRECO_ALIQUOTA_PRODUTO

            lErro = CF("NCM_Le_Aliquota_ComProduto", objMnemonicoValor.vParam(1), objMnemonicoValor.vParam(2), dAliquota)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
            objMnemonicoValor.colValor.Add dAliquota
            
        Case Else
            gError 92413

    End Select

    Calcula_MnemonicoFPreco = SUCESSO

    Exit Function

Erro_Calcula_MnemonicoFPreco:

    Calcula_MnemonicoFPreco = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 92413
'            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_NAO_ENCONTRADO", gErr, objMnemonicoValor.sMnemonico)
                
        Case 92417, 92420, 178581
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149688)

    End Select

    Exit Function

End Function

'Por Leo
Function EmbalagensExpedicao_Le(ByVal objEmbalagem As ClassEmbalagem) As Long
'Lê todas as embalagensexpedicao e guarda em objembalagensexpedicao.colEmbalagensexpedicao através do campo embalagem

Dim lErro As Long
Dim lComando As Long
Dim objEmbalagensExpedicao As New ClassEmbalagensExpedicao
Dim iSequencial As Integer
Dim sProduto As String
Dim dQuantidade As Double

On Error GoTo Erro_EmbalagensExpedicao_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 103440
    
    Set objEmbalagem.colEmbExpedicao = New Collection
    
    sProduto = String(STRING_PRODUTO, 0)
    
    'Seleciona todas as embalagens
    lErro = Comando_Executar(lComando, "SELECT Sequencial, Produto, Quantidade FROM EmbalagensExpedicao WHERE Embalagem = ? ORDER BY Sequencial", iSequencial, sProduto, dQuantidade, objEmbalagem.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 103441

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103442
        
    Do While lErro = AD_SQL_SUCESSO

        'Guarda na coleção as embalagens
        Set objEmbalagensExpedicao = New ClassEmbalagensExpedicao

        objEmbalagensExpedicao.dQuantidade = dQuantidade
        objEmbalagensExpedicao.iEmbalagem = objEmbalagem.iCodigo
        objEmbalagensExpedicao.sProduto = sProduto
        objEmbalagensExpedicao.iSequencial = iSequencial
        
        'Adiciona na colecao
        objEmbalagem.colEmbExpedicao.Add objEmbalagensExpedicao

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 103443

    Loop

    'fecha o comando
    Call Comando_Fechar(lComando)

    EmbalagensExpedicao_Le = SUCESSO

    Exit Function

Erro_EmbalagensExpedicao_Le:

    EmbalagensExpedicao_Le = gErr

    Select Case gErr

        Case 103440
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 103441, 103442, 103443
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGENSEXPEDICAO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149689)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Calcula_MnemonicoFPreco2(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal sProduto As String, ByVal objExeExp As ClassExeExp) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dCusto As Double
Dim sTipo As String
Dim sProdutoFormatado As String
Dim iPreenchido As Integer

On Error GoTo Erro_Calcula_MnemonicoFPreco2

'    MsgBox ("Calcula_MnemonicoFPreco2:" & objMnemonicoValor.sMnemonico)
    
    lErro = CF("Calcula_MnemonicoFPreco", objMnemonicoValor, sProduto, objExeExp)
    If lErro <> SUCESSO And lErro <> 92413 Then gError 106695
    
'    MsgBox ("lerro: " & CStr(lErro) & " Calcula_MnemonicoFPreco2:" & objMnemonicoValor.sMnemonico)
    
    If lErro <> SUCESSO Then

        Select Case Trim(objMnemonicoValor.sMnemonico)

            Case MNEMONICOFPRECO_QUANTIDADE
                objMnemonicoValor.colValor.Add CDbl(objExeExp.objContexto.dQuantidade)
            
            Case MNEMONICOFPRECO_CLIENTE
                objMnemonicoValor.colValor.Add CDbl(objExeExp.objContexto.lCliente)
            
            Case MNEMONICOFPRECO_FILIALCLI
                objMnemonicoValor.colValor.Add CDbl(objExeExp.objContexto.iFilialCli)
                        
            Case MNEMONICOFPRECO_PRODUTO
                objMnemonicoValor.colValor.Add sProduto
            
            Case MNEMONICOFPRECO_PRECOPRATICADO
                lErro = CF("PrecoPraticado", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106679

            Case MNEMONICOFPRECO_VALPRES
                lErro = CF("ValPres", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106680

            Case MNEMONICOFPRECO_VALFUT
                lErro = CF("ValFut", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106681

            Case MNEMONICOFPRECO_DVVTOTAL
                objMnemonicoValor.colValor.Add CDbl(0)

            Case MNEMONICOFPRECO_PRECOTABELA
                lErro = CF("PrecoTabela", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106686
            
            Case MNEMONICOFPRECO_PRECOTABELA_PARAM
                lErro = CF("PrecoTabela_Param", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106686
            
            Case MNEMONICOFPRECO_TABELAPRECO
                objMnemonicoValor.colValor.Add CDbl(objExeExp.objContexto.iTabelaPreco)
            
            Case MNEMONICOFPRECO_TABELAPRECOPADRAO
                lErro = CF("TabelaPrecoPadrao", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106686

            Case MNEMONICOFPRECO_DEVDUV
                lErro = CF("DevDuv", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106687

            Case MNEMONICOFPRECO_DIASCONDPAGTO
                lErro = CF("DiasCondPagto", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106690

            Case MNEMONICOFPRECO_CONDPAGTOCLIENTE
                lErro = CF("CondPagtoCliente", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106691

            Case MNEMONICOFPRECO_PRECOULTPV
                lErro = CF("PrecoUltPV", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106692

            Case MNEMONICOFPRECO_PRECOULTNF
                lErro = CF("PrecoUltNF", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106692

            Case MNEMONICOFPRECO_COMISSOESPROPCLIPROD
                lErro = CF("ComissoesPropCliProd", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106693

            Case MNEMONICOFPRECO_COMISSOESTERCCLIPROD
                lErro = CF("ComissoesTercCliProd", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106694

            Case MNEMONICOFPRECO_ENCCOMISSOESPROPCLIPROD
                lErro = CF("EncComissoesPropCliProd", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106693

            Case MNEMONICOFPRECO_ENCCOMISSOESTERCCLIPROD
                lErro = CF("EncComissoesTercCliProd", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106694

            Case MNEMONICOFPRECO_ICMSALIQFILCLI
                lErro = CF("ICMSAliqFilCli", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106694
            
            Case MNEMONICOFPRECO_ICMSALIQFILCLIPROD
                lErro = CF("ICMSAliqFilCliProd", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106694
            
            Case MNEMONICOFPRECO_IPIALIQUOTA
                lErro = CF("FormPreco_IPIAliquota", objMnemonicoValor, objExeExp)
                If lErro <> SUCESSO Then gError 106694
            
            Case Else
                gError 92413

        End Select

    End If

    Calcula_MnemonicoFPreco2 = SUCESSO

    Exit Function

Erro_Calcula_MnemonicoFPreco2:

    Calcula_MnemonicoFPreco2 = gErr

    Select Case gErr

        Case 106679 To 106695

        Case 92413

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149690)

    End Select

    Exit Function

End Function

Public Function Escaninhos_Le_Terceiros(ByVal colEscaninhos As Collection) As Long
'Le todos os escaninhos que podem ser de de 3ºs

Dim lErro As Long
Dim lComando As Long
Dim tEscaninho As typeEscaninho
Dim objEscaninho As ClassEscaninho

On Error GoTo Erro_Escaninhos_Le_Terceiros

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 119705

    'inicializa a string p/ receber o nome do escaninho
    tEscaninho.sNome = String(STRING_NOME_ESCANINHO, 0)

    'Lê os dados dos Escaninhos que podem se de terceiros / nosso
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, RastroEstoqueInicial FROM Escaninhos WHERE Terceiros = ?", tEscaninho.iCodigo, tEscaninho.sNome, tEscaninho.iRastroEstoqueInicial, ESCANINHO_TERCEIROS)
    If lErro <> AD_SQL_SUCESSO Then gError 119706

    'busca o 1º registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119707

    'se não encontrar os escaninhos, retorna erro
    If lErro = AD_SQL_SEM_DADOS Then gError 119709

    'enquanto houver escaninho
    Do While lErro = AD_SQL_SUCESSO
    
        'instancia o obj
        Set objEscaninho = New ClassEscaninho
    
        'guarda os dados lidos no obj
        objEscaninho.iCodigo = tEscaninho.iCodigo
        objEscaninho.sNome = tEscaninho.sNome
        objEscaninho.iRastroEstoqueInicial = tEscaninho.iRastroEstoqueInicial
        
        'adiciona o objescaninho na coleção
        colEscaninhos.Add objEscaninho
    
        'busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 119708
    
    Loop
    
    'Fechamento do comando
    Call Comando_Fechar(lComando)

   Escaninhos_Le_Terceiros = SUCESSO

    Exit Function

Erro_Escaninhos_Le_Terceiros:

    Escaninhos_Le_Terceiros = gErr

    Select Case gErr

        Case 119705
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 119706, 119707, 119708
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESCANINHOS", gErr)

        Case 119709 'sem dados // tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149691)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function GradeCategoria_Le(ByVal objGrade As ClassGrade) As Long
'Carrega em objGrade as Categorias referentes ao produto cujo código está contido
'em objGrade

Dim lErro As Long
Dim sCategoria As String
Dim lComando As Long
Dim iPosicao As Integer
Dim objGradeCategoria As ClassGradeCategoria
Dim iSeq As Integer

On Error GoTo Erro_GradeCategoria_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 122723
        
    sCategoria = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    
    'Seleciona em GradeCategoria as categorias referentes ao código armazenado em objGrade
    lErro = Comando_Executar(lComando, "SELECT Categoria, Posicao, Seq FROM GradeCategoria WHERE Grade=? ORDER BY Posicao, Seq", sCategoria, iPosicao, iSeq, objGrade.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 122724

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122725

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objGradeCategoria = New ClassGradeCategoria
        
        'Preenche o objGradeCategoria
        objGradeCategoria.sCategoria = sCategoria
        objGradeCategoria.iPosicao = iPosicao
        objGradeCategoria.iSeq = iSeq
        
        'Adiciona na Coleção o ObjGradeCategoria
        objGrade.colCategoria.Add objGradeCategoria
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122726

    Loop

    Call Comando_Fechar(lComando)

    GradeCategoria_Le = SUCESSO

    Exit Function

Erro_GradeCategoria_Le:

    GradeCategoria_Le = gErr

    Select Case gErr

        Case 122723
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 122724, 122725, 122726
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRADE_CATEGORIA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149692)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Grade_Grava_Lock(ByVal objGrade As ClassGrade) As Long
'Faz o Lock Shared dos registros referentes às Categorias presentes em objGrade,
'na tabela CategoriaProduto

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim sCategoria As String
Dim objGradeCategoria As ClassGradeCategoria

On Error GoTo Erro_Grade_Grava_Lock
     
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 122713
 
    'Para todas as categorias armazenadas em objGrade
    For Each objGradeCategoria In objGrade.colCategoria
        
        sCategoria = String(STRING_GRADE_CODIGO, 0)

        lErro = Comando_ExecutarPos(lComando, "SELECT Categoria FROM CategoriaProduto WHERE Categoria=?", 0, sCategoria, objGradeCategoria.sCategoria)
        If lErro <> AD_SQL_SUCESSO Then gError 122714

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122715

        If lErro = AD_SQL_SEM_DADOS Then gError 122716

        'Efetua o lock da categoria em CategoriaProduto
        lErro = Comando_LockShared(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 122717

    Next
    
    Call Comando_Fechar(lComando)
    
    Grade_Grava_Lock = SUCESSO
       
    Exit Function
    
Erro_Grade_Grava_Lock:

    Grade_Grava_Lock = gErr

    Select Case gErr
    
        Case 122713
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 122714, 122715
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTO", gErr)
            
        Case 122716
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_INEXISTENTE", gErr)
            
        Case 122717
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIAPRODUTO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149693)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Grade_Verifica_Utilizacao(ByVal objGrade As ClassGrade) As Long
'Verifica se a Grade passada em objGrade está sendo utilizada por algum produto
'Se estiver -> ERRO

Dim lErro As Long
Dim lComando As Long
Dim sGrade As String

On Error GoTo Erro_Grade_Verifica_Utilizacao
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 122741

    sGrade = String(STRING_GRADE_CODIGO, 0)
        
    'Verifica se grade está sendo usada por algum produto da tabela Produtos
    lErro = Comando_Executar(lComando, "SELECT Grade FROM Produtos WHERE Grade=?", sGrade, objGrade.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 122742
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 122743
    
    'Se estiver sendo usada -> ERRO
    If lErro = AD_SQL_SUCESSO Then gError 122744
    
    Call Comando_Fechar(lComando)
    
    Grade_Verifica_Utilizacao = SUCESSO

    Exit Function

Erro_Grade_Verifica_Utilizacao:

    Grade_Verifica_Utilizacao = gErr

    Select Case gErr

        Case 122741
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 122742, 122743
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
            
        Case 122744
            Call Rotina_Erro(vbOKOnly, "ERRO_GRADE_USO_NAO_ALTERAVEL", gErr, objGrade.sCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149694)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function


Function ItemPedidoGrade_Le_Produto(objItemPedido As ClassItemPedido) As Long

Dim lErro As Long
Dim tItemPedidoGrade As typeItemRomaneioGrade
Dim lComando As Long
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_ItemPedidoGrade_Le_Produto

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86379

    tItemPedidoGrade.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemPedidoGrade.sProduto = String(STRING_PRODUTO, 0)
    tItemPedidoGrade.sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)

    lErro = Comando_Executar(lComando, "SELECT ItensPedidoDeVendaGrade.NumIntDoc, NumIntItemPV, ItensPedidoDeVendaGrade.Produto, ItensPedidoDeVendaGrade.Quantidade, ItensPedidoDeVendaGrade.QuantCancelada, ItensPedidoDeVendaGrade.QuantReservada, ItensPedidoDeVendaGrade.QuantFaturada, ItensPedidoDeVendaGrade.QuantOP, ItensPedidoDeVendaGrade.QuantSC FROM ItensPedidoDeVendaGrade, ItensPedidoDeVenda, Produtos WHERE ItensPedidoDeVenda.NumIntDoc = ItensPedidoDeVendaGrade.NumIntItemPV AND ItensPedidoDeVendaGrade.Produto = Produtos.Codigo AND ItensPedidoDeVendaGrade.Produto= ? AND ItensPedidoDeVenda.FilialEmpresa = ? AND CodPedido = ? ", tItemPedidoGrade.lNumIntDoc, tItemPedidoGrade.lNumIntItemPV, tItemPedidoGrade.sProduto, tItemPedidoGrade.dQuantidade, tItemPedidoGrade.dQuantCancelada, tItemPedidoGrade.dQuantReservada, tItemPedidoGrade.dQuantFaturada, tItemPedidoGrade.dQuantOP, tItemPedidoGrade.dQuantSC, objItemPedido.sProduto, giFilialEmpresa, objItemPedido.lCodPedido)
    If lErro <> AD_SQL_SUCESSO Then gError 86380
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86381
    
    If lErro = AD_SQL_SEM_DADOS Then gError 86382
    
    objItemPedido.iPossuiGrade = MARCADO
    
    Set objItemRomaneio = New ClassItemRomaneioGrade
    
    objItemRomaneio.dQuantAFaturar = tItemPedidoGrade.dQuantAFaturar
    objItemRomaneio.dQuantCancelada = tItemPedidoGrade.dQuantCancelada
    objItemRomaneio.dQuantFaturada = tItemPedidoGrade.dQuantFaturada
    objItemRomaneio.dQuantidade = tItemPedidoGrade.dQuantidade
    objItemRomaneio.dQuantOP = tItemPedidoGrade.dQuantOP
    objItemRomaneio.dQuantReservada = tItemPedidoGrade.dQuantReservada
    objItemRomaneio.dQuantSC = tItemPedidoGrade.dQuantSC
    objItemRomaneio.iAlmoxarifado = tItemPedidoGrade.iAlmoxarifado
    objItemRomaneio.iControleEstoque = tItemPedidoGrade.iControleEstoque
    objItemRomaneio.lNumIntDoc = tItemPedidoGrade.lNumIntDoc
    objItemRomaneio.lNumIntItemPV = tItemPedidoGrade.lNumIntItemPV
    objItemRomaneio.sDescricao = tItemPedidoGrade.sDescricao
    objItemRomaneio.sProduto = tItemPedidoGrade.sProduto
    objItemRomaneio.sUMEstoque = tItemPedidoGrade.sSiglaUMEstoque

    objItemPedido.colItensRomaneioGrade.Add objItemRomaneio

    Call Comando_Fechar(lComando)
    
    ItemPedidoGrade_Le_Produto = SUCESSO
    
    Exit Function
    
Erro_ItemPedidoGrade_Le_Produto:

    ItemPedidoGrade_Le_Produto = gErr
    
    Select Case gErr
    
        Case 86379
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 86380, 86381
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPVGRADE", gErr)
        
        Case 86382
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149695)
            
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function PrestServ_Le(ByVal objPrestServ As ClassPrestServ) As Long
'Lê os dados do requisitante, cujo código foi passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim tPrestServ As typePrestServ
Dim sNome As String
Dim sNomeReduzido As String

On Error GoTo Erro_PrestServ_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 49081

    tPrestServ.sNome = String(STRING_PRESTSERV_NOME, 0)
    tPrestServ.sNomeReduzido = String(STRING_PRESTSERV_NOMERED, 0)

    'Lê os dados da tabela PrestServ
    lErro = Comando_Executar(lComando, "SELECT  Nome, NomeReduzido, Fornecedor FROM PrestServ WHERE Codigo = ?", tPrestServ.sNome, tPrestServ.sNomeReduzido, tPrestServ.lFornecedor, objPrestServ.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49082

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49083

    'PrestServ nao esta cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error ERRO_OBJETO_NAO_CADASTRADO

    'carrega os dados em objPrestServ
    objPrestServ.sNome = tPrestServ.sNome
    objPrestServ.sNomeReduzido = tPrestServ.sNomeReduzido
    objPrestServ.lFornecedor = tPrestServ.lFornecedor

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    PrestServ_Le = SUCESSO

    Exit Function

Erro_PrestServ_Le:

    PrestServ_Le = Err

    Select Case Err

        Case 49081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49082, 49083
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRESTSERV", Err, objPrestServ.lCodigo)

        Case ERRO_OBJETO_NAO_CADASTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149696)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PrestServ_Le_NomeReduzido(ByVal objPrestServ As ClassPrestServ) As Long
'Lê os dados do requisitante, cujo nome reduzido foi passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tPrestServ As typePrestServ

On Error GoTo Erro_PrestServ_Le_NomeReduzido

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51149

    tPrestServ.sNome = String(STRING_PRESTSERV_NOME, 0)

    'Lê os dados da tabela PrestServ baseado no Nome Reduzido
    lErro = Comando_Executar(lComando, "SELECT  Codigo,Nome,Fornecedor FROM PrestServ WHERE NomeReduzido = ?", tPrestServ.lCodigo, tPrestServ.sNome, tPrestServ.lFornecedor, objPrestServ.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then Error 51150

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 51151
    If lErro = AD_SQL_SEM_DADOS Then Error ERRO_OBJETO_NAO_CADASTRADO 'Não encontrou

    'carrega os dados em objPrestServ
    objPrestServ.lCodigo = tPrestServ.lCodigo
    objPrestServ.sNome = tPrestServ.sNome
    objPrestServ.lFornecedor = tPrestServ.lFornecedor

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    PrestServ_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_PrestServ_Le_NomeReduzido:

    PrestServ_Le_NomeReduzido = Err

    Select Case Err

        Case 51149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 51150, 51151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRESTSERV1", Err, objPrestServ.sNomeReduzido)

        Case ERRO_OBJETO_NAO_CADASTRADO 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149697)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PrestServ_Lock(ByVal lCodPrestServ As Long) As Long
'Lê e faz LockShared no PrestServ

Dim lErro As Long
Dim lComando As Long
Dim lCodigo As Long

On Error GoTo Erro_PrestServ_Lock

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61734

    'Tenta ler o PrestServ passado pelo Código
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo FROM PrestServ WHERE Codigo = ?", 0, lCodigo, lCodPrestServ)
    If lErro <> AD_SQL_SUCESSO Then Error 61735

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61736

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then Error 61737

    'Faz "LockShared" no PrestServ
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 61739

    'Fecha Comando
    Call Comando_Fechar(lComando)

    PrestServ_Lock = SUCESSO

    Exit Function

Erro_PrestServ_Lock:

    PrestServ_Lock = Err

    Select Case Err

        Case 61734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61735, 61736
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRESTSERV", Err, lCodPrestServ)

        Case 61737 'PrestServ não cadastrado

        Case 61739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRESTSERV", Err, lCodPrestServ)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149698)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PrestServ_Valida_Exclusao(ByVal objPrestServ As ClassPrestServ, alComando() As Long) As Long
'Verifica se o PrestServ a ser excluído está vinculado a alguma outra tabela

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_PrestServ_Valida_Exclusao

    'verifica se PrestServ esta relacionado com alguma Ordem de Producao
    lErro = Comando_Executar(alComando(1), "SELECT PrestServ FROM OrdensDeProducao WHERE PrestServ =?", lCodigo, objPrestServ.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49079

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49098
    
    'O requisitante está relacionado com Ordem de Producao
    If lErro = AD_SQL_SUCESSO Then Error 49099
    
    'verifica se PrestServ esta relacionado com alguma Ordem de Producao Baixada
    lErro = Comando_Executar(alComando(1), "SELECT PrestServ FROM OrdensDeProducaoBaixadas WHERE PrestServ =?", lCodigo, objPrestServ.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 49291

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49292
    
    'O requisitante está relacionado com Ordem de Producao Baixada
    If lErro = AD_SQL_SUCESSO Then Error 49293
    
    PrestServ_Valida_Exclusao = SUCESSO
    
    Exit Function
    
Erro_PrestServ_Valida_Exclusao:

    PrestServ_Valida_Exclusao = Err
    
    Select Case Err
    
        Case 49079, 49098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA1", Err, objPrestServ.lCodigo)
            
        Case 49099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRESTSERV_VINCULADO_REQUISICAOCOMPRA", Err, objPrestServ.lCodigo)
        
        Case 49291, 49292
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRABAIXADA1", Err, objPrestServ.lCodigo)
        
        Case 49293
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRESTSERV_VINCULADO_REQCOMPRABAIXADA", Err, objPrestServ.lCodigo)
            
        Case 49294, 49295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOMODELO1", Err, objPrestServ.lCodigo)
        
        Case 49296
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRESTSERV_VINCULADO_REQCOMPRAMODELO", Err, objPrestServ.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149699)

    End Select
    
    Exit Function

End Function

Function Remove_Pedido_Compra(ByVal objPedidoCompras As ClassPedidoCompras, ByVal colPedidoCompra As Collection, ByVal iIndice As Integer) As Long

On Error GoTo Erro_Remove_Pedido_Compra

    'Remove os PCs com destino <> empresa ou outra filialempresa
    If (objPedidoCompras.iTipoDestino <> TIPO_DESTINO_EMPRESA) Or (objPedidoCompras.iTipoDestino = TIPO_DESTINO_EMPRESA And objPedidoCompras.iFilialDestino <> giFilialEmpresa) Then
        colPedidoCompra.Remove (iIndice)
    End If

    Remove_Pedido_Compra = SUCESSO

    Exit Function

Erro_Remove_Pedido_Compra:

    Remove_Pedido_Compra = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149700)

    End Select

    Exit Function

End Function

Function ImportEstoqueInicial_Le_Todos(ByVal colEstoqueProduto As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim tEstoqueProduto As typeEstoqueProduto
Dim sProdutoAntigo As String
Dim sProdutoConvertido As String, iPreenchido As Integer
Dim iFilialEmpresa As Integer

On Error GoTo Erro_ImportEstoqueInicial_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 129856

    With tEstoqueProduto

        .sProduto = String(STRING_PRODUTO, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        .sLocalizacaoFisica = String(STRING_LOCALIZACAO_FISICA, 0)
        sProdutoAntigo = String(30, 0)

        lErro = Comando_Executar(lComando, "SELECT Produto, ProdutoAntigo, Almoxarifado, FilialEmpresa, LocalizacaoFisica, ContaContabil, QuantidadeInicial, SaldoInicial, DataInicial, " & _
                                                "QuantInicialConsig, ValorInicialConsig, QuantInicialConsig3, ValorInicialConsig3, " & _
                                                "QuantInicialOutras , ValorInicialOutras, QuantInicialOutras3, ValorInicialOutras3, " & _
                                                "QuantInicialConserto , ValorInicialConserto, QuantInicialConserto3, ValorInicialConserto3, " & _
                                                "QuantInicialDemo , ValorInicialDemo, QuantInicialDemo3, ValorInicialDemo3, " & _
                                                "QuantInicialBenef , ValorInicialBenef, QuantInicialBenef3, ValorInicialBenef3 " & _
                                            "FROM ImportEstoqueInicial", .sProduto, sProdutoAntigo, .iAlmoxarifado, iFilialEmpresa, .sLocalizacaoFisica, .sContaContabil, .dQuantidadeInicial, .dSaldoInicial, .dtDataInicial, _
                                                .dQuantInicialConsig, .dValorInicialConsig, .dQuantInicialConsig3, .dValorInicialConsig3, _
                                                .dQuantInicialOutras, .dValorInicialOutras, .dQuantInicialOutras3, .dValorInicialOutras3, _
                                                .dQuantInicialConserto, .dValorInicialConserto, .dQuantInicialConserto3, .dValorInicialConserto3, _
                                                .dQuantInicialDemo, .dValorInicialDemo, .dQuantInicialDemo3, .dValorInicialDemo3, _
                                                .dQuantInicialBenef, .dValorInicialBenef, .dQuantInicialBenef3, .dValorInicialBenef3)
                                        
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 129857

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129858
    
    Do While lErro = AD_SQL_SUCESSO

        Set objEstoqueProduto = New ClassEstoqueProduto

        With tEstoqueProduto

'            If Len(Trim(.sProduto)) = 0 Then
'                lErro = ProdutoConversao_Le(sProdutoAntigo, sProdutoConvertido)
'                If lErro <> SUCESSO Then gError 129871
'
'                objEstoqueProduto.sProduto = sProdutoConvertido
'            End If

            lErro = CF("Produto_Formata", .sProduto, sProdutoConvertido, iPreenchido)
            If lErro <> SUCESSO Then gError 27682
            
            If iFilialEmpresa = 0 Then iFilialEmpresa = giFilialEmpresa
            
            objEstoqueProduto.sProduto = sProdutoConvertido
            objEstoqueProduto.iFilialEmpresa = iFilialEmpresa
            objEstoqueProduto.iAlmoxarifado = .iAlmoxarifado
            objEstoqueProduto.sLocalizacaoFisica = .sLocalizacaoFisica
            objEstoqueProduto.sContaContabil = .sContaContabil
            objEstoqueProduto.dQuantidadeInicial = .dQuantidadeInicial
            objEstoqueProduto.dSaldoInicial = .dSaldoInicial
            objEstoqueProduto.dtDataInicial = .dtDataInicial
            objEstoqueProduto.dQuantInicialConsig = .dQuantInicialConsig
            objEstoqueProduto.dValorInicialConsig = .dValorInicialConsig
            objEstoqueProduto.dQuantInicialConsig3 = .dQuantInicialConsig3
            objEstoqueProduto.dValorInicialConsig3 = .dValorInicialConsig3
            objEstoqueProduto.dQuantInicialOutras = .dQuantInicialOutras
            objEstoqueProduto.dValorInicialOutras = .dValorInicialOutras
            objEstoqueProduto.dQuantInicialOutras3 = .dQuantInicialOutras3
            objEstoqueProduto.dValorInicialOutras3 = .dValorInicialOutras3
            objEstoqueProduto.dQuantInicialConserto = .dQuantInicialConserto
            objEstoqueProduto.dValorInicialConserto = .dValorInicialConserto
            objEstoqueProduto.dQuantInicialConserto3 = .dQuantInicialConserto3
            objEstoqueProduto.dValorInicialConserto3 = .dValorInicialConserto3
            objEstoqueProduto.dQuantInicialDemo = .dQuantInicialDemo
            objEstoqueProduto.dValorInicialDemo = .dValorInicialDemo
            objEstoqueProduto.dQuantInicialDemo3 = .dQuantInicialDemo3
            objEstoqueProduto.dValorInicialDemo3 = .dValorInicialDemo3
            objEstoqueProduto.dQuantInicialBenef = .dQuantInicialBenef
            objEstoqueProduto.dValorInicialBenef = .dValorInicialBenef
            objEstoqueProduto.dQuantInicialBenef3 = .dQuantInicialBenef3
            objEstoqueProduto.dValorInicialBenef3 = .dValorInicialBenef3
            
        End With
        
        'Adiciona na colecao
        colEstoqueProduto.Add objEstoqueProduto

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129859

    Loop

    'fecha o comando
    Call Comando_Fechar(lComando)

    ImportEstoqueInicial_Le_Todos = SUCESSO

    Exit Function

Erro_ImportEstoqueInicial_Le_Todos:

    ImportEstoqueInicial_Le_Todos = gErr

    Select Case gErr

        Case 129856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 129857 To 129859
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTESTOQUEINICIAL", gErr)

        Case 129871

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149701)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ImportRastroInicial_Le(ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal colRastreamento As Collection) As Long

Dim lErro As Long
Dim lComando As Long
Dim objRastroIni As New ClassRastroEstIni
Dim tRastroIni As typeRastroEstIni
Dim sProdutoAntigo As String
Dim sProdutoConvertido As String

On Error GoTo Erro_ImportRastroInicial_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 129861

    With tRastroIni

        .sProduto = String(STRING_PRODUTO, 0)
        .sLote = String(STRING_LOTE_RASTREAMENTO, 0)
        sProdutoAntigo = String(30, 0)

        lErro = Comando_Executar(lComando, "SELECT DataEntrada, Almoxarifado, Escaninho, FilialOP, Lote, Produto, ProdutoAntigo FROM ImportRastroInicial WHERE Almoxarifado = ? AND Produto = ? ", _
                                                    .dtDataEntrada, .iAlmoxarifado, .iEscaninho, .iFilialOP, .sLote, .sProduto, sProdutoAntigo, objEstoqueProduto.iAlmoxarifado, objEstoqueProduto.sProduto)
                                        
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 129862

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129863
    
    Do While lErro = AD_SQL_SUCESSO

        Set objRastroIni = New ClassRastroEstIni
        
        With tRastroIni
        
            objRastroIni.dtDataEntrada = .dtDataEntrada
            objRastroIni.iAlmoxarifado = .iAlmoxarifado
            objRastroIni.iEscaninho = .iEscaninho
            objRastroIni.iFilialOP = .iFilialOP
            objRastroIni.sLote = .sLote
            
            If Len(Trim(.sProduto)) = 0 Then
                lErro = ProdutoConversao_Le(sProdutoAntigo, sProdutoConvertido)
                If lErro <> SUCESSO Then gError 129870
                
                objRastroIni.sProduto = sProdutoConvertido
            End If
            
        End With
        
        'Adiciona na colecao
        colRastreamento.Add objRastroIni

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129864

    Loop

    'fecha o comando
    Call Comando_Fechar(lComando)

    ImportRastroInicial_Le = SUCESSO

    Exit Function

Erro_ImportRastroInicial_Le:

    ImportRastroInicial_Le = gErr

    Select Case gErr

        Case 129861
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 129862 To 129864
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTROTINASINICIAL", gErr)

        Case 129870

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149702)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoConversao_Le(ByVal sProdutoAntigo As String, sProdutoCorporator As String) As Long

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_ProdutoConversao_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 129865

    sProdutoCorporator = String(STRING_PRODUTO, 0)

    lErro = Comando_Executar(lComando, "SELECT ProdutoCorporator FROM ProdutoConversao WHERE ProdutoAntigo = ? ORDER BY ProdutoCorporator ", sProdutoCorporator, sProdutoAntigo)
    If lErro <> AD_SQL_SUCESSO Then gError 129866

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129867
    
    'fecha o comando
    Call Comando_Fechar(lComando)

    ProdutoConversao_Le = SUCESSO

    Exit Function

Erro_ProdutoConversao_Le:

    ProdutoConversao_Le = gErr

    Select Case gErr

        Case 129865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 129866
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOCONVERSAO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149703)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Custo_Kit_MP(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal sProduto As String) As Long
'parametros:Produto
'retorna o custo das materias primas necessarias para fabricar o produto,
'incluindo as materias primas para os subprodutos, recursivamente

Dim lErro As Long
Dim colMP As New Collection, dCustoProd As Double
Dim dCustoMP As Double
Dim objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP
Dim objProduto As New ClassProduto
Dim sTipo As String
Dim iPreenchido As Integer

On Error GoTo Erro_Custo_Kit_MP

    If sProduto = "" Then
    
        lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 133017
        
    End If
    
    lErro = CF("CustoDireto_ObtemMP", giFilialEmpresa, Year(gdtDataAtual), sProduto, 1, colMP, FORMACAO_PRECO_ROTINA_CALCPRECO)
    If lErro <> SUCESSO Then gError 133014
    
    'para cada materia prima
    For Each objCustoDirFabrPlanMP In colMP
            
        objProduto.sCodigo = objCustoDirFabrPlanMP.sProdutoMP
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError 133013
            
        sTipo = STRING_QUANT_DISPONIVEL_NOSSA
            
        'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
        lErro = CF("Calcula_Custo_Atual", objProduto, dCustoMP, sTipo)
        If lErro <> SUCESSO Then gError 133020
            
        'com a qtde e custo unitario, acumulo o custo do insumo no custo do produto.
        dCustoProd = dCustoProd + (dCustoMP * objCustoDirFabrPlanMP.dQtde)
    
    Next
    
    objMnemonicoValor.colValor.Add dCustoProd
    
    Custo_Kit_MP = SUCESSO

    Exit Function

Erro_Custo_Kit_MP:

    Custo_Kit_MP = gErr

    Select Case gErr

        Case 133013, 133014, 133015, 133020

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149704)

    End Select

    Exit Function

End Function

Function CustoRep_Kit_MP(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal sProduto As String) As Long
'parametros:Produto
'retorna o custo das materias primas necessarias para fabricar o produto,
'incluindo as materias primas para os subprodutos, recursivamente

Dim lErro As Long
Dim colMP As New Collection, dCustoProd As Double
Dim dCustoMP As Double
Dim objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP
Dim objProduto As New ClassProduto
Dim sTipo As String
Dim iPreenchido As Integer

On Error GoTo Erro_CustoRep_Kit_MP

    If sProduto = "" Then
    
        lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 133021
        
    End If
    
    lErro = CF("CustoDireto_ObtemMP", giFilialEmpresa, Year(gdtDataAtual), sProduto, 1, colMP, FORMACAO_PRECO_ROTINA_CALCPRECO)
    If lErro <> SUCESSO Then gError 133022
    
    'para cada materia prima
    For Each objCustoDirFabrPlanMP In colMP
            
        objProduto.sCodigo = objCustoDirFabrPlanMP.sProdutoMP
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError 133023
            
        'com a qtde e custo unitario, acumulo o custo do insumo no custo do produto.
        dCustoProd = dCustoProd + (objProduto.dCustoReposicao * objCustoDirFabrPlanMP.dQtde)
            
    Next
    
    objMnemonicoValor.colValor.Add dCustoProd
    
    CustoRep_Kit_MP = SUCESSO

    Exit Function

Erro_CustoRep_Kit_MP:

    CustoRep_Kit_MP = gErr

    Select Case gErr

        Case 133021 To 133023

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149705)

    End Select

    Exit Function

End Function

Function PrecoPrazo_Customizado(dPrecoPrazo As Double) As Long
    PrecoPrazo_Customizado = SUCESSO
End Function

Public Function Taxa_Producao_UM_Padrao_Obtem(sUM As String) As Long

    sUM = TAXA_CONSUMO_TEMPO_PADRAO

    Taxa_Producao_UM_Padrao_Obtem = SUCESSO
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function DIInfo_Le(ByVal objDIInfo As ClassDIInfo) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 5) As Long
Dim tDIInfo As typeDIInfo

On Error GoTo Erro_DIInfo_Le

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184501
    Next
    
    'Alocação de espaço no buffer
    tDIInfo.sNumero = String(STRING_DI_NUMERO, 0)
    tDIInfo.sDescricao = String(STRING_DI_DESCRICAO, 0)
    tDIInfo.sProcessoTrading = String(STRING_DI_PROCESSO_TRADING, 0)
    tDIInfo.sUFDesembaraco = String(STRING_ESTADO_SIGLA, 0)
    tDIInfo.sLocalDesembaraco = String(STRING_DI_LOCAL_DESEMBARACO, 0)
    tDIInfo.sCodExportador = String(STRING_DI_COD_EXPORTADOR, 0)
    tDIInfo.sCNPJAdquir = String(STRING_CGC, 0)
    tDIInfo.sUFAdquir = String(STRING_ESTADO_SIGLA, 0)

    'Le a tabelaDIInfo
    With tDIInfo
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Numero, Data, FilialEmpresa, Status, Descricao, FornTrading, " & _
                "FilialFornTrading, ProcessoTrading, Moeda, TaxaMoeda, PesoBrutoKG, PesoLiqKG, ValorMercadoriaMoeda, " & _
                "ValorFreteInternacMoeda, ValorSeguroInternacMoeda, ValorMercadoriaEmReal, ValorFreteInternacEmReal, ValorSeguroInternacEmReal, " & _
                "IIValor, IPIValor, PISValor, COFINSValor, ICMSValor, ValorDespesas, Moeda2, TaxaMoeda2, DataDesembaraco, UFDesembaraco, LocalDesembaraco, MoedaMercadoria, MoedaFrete, MoedaSeguro, MoedaItens, CodExportador, ViaTransp, Intermedio, CNPJAdquir, UFAdquir FROM DIInfo WHERE Numero= ?", _
                .lNumIntDoc, .sNumero, .dtData, .iFilialEmpresa, .iStatus, .sDescricao, .lFornTrading, .iFilialFornTrading, .sProcessoTrading, .iMoeda1, _
                .dTaxaMoeda1, .dPesoBrutoKG, .dPesoLiqKG, .dValorMercadoriaMoeda, .dValorFreteInternacMoeda, _
                .dValorSeguroInternacMoeda, .dValorMercadoriaEmReal, .dValorFreteInternacEmReal, .dValorSeguroInternacEmReal, _
                .dIIValor, .dIPIValor, .dPISValor, .dCOFINSValor, .dICMSValor, .dValorDespesas, .iMoeda2, .dTaxaMoeda2, .dtDataDesembaraco, .sUFDesembaraco, .sLocalDesembaraco, .iMoedaMercadoria, .iMoedaFrete, .iMoedaSeguro, .iMoedaItens, .sCodExportador, .iViaTransp, .iIntermedio, .sCNPJAdquir, .sUFAdquir, objDIInfo.sNumero)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 184502

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184503

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    With objDIInfo
        .lNumIntDoc = tDIInfo.lNumIntDoc
        .sNumero = tDIInfo.sNumero
        .dtData = tDIInfo.dtData
        .iFilialEmpresa = tDIInfo.iFilialEmpresa
        .iStatus = tDIInfo.iStatus
        .sDescricao = tDIInfo.sDescricao
        .lFornTrading = tDIInfo.lFornTrading
        .iFilialFornTrading = tDIInfo.iFilialFornTrading
        .sProcessoTrading = tDIInfo.sProcessoTrading
        .iMoeda1 = tDIInfo.iMoeda1
        .dTaxaMoeda1 = tDIInfo.dTaxaMoeda1
        .iMoeda2 = tDIInfo.iMoeda2
        .dTaxaMoeda2 = tDIInfo.dTaxaMoeda2
        .dPesoBrutoKG = tDIInfo.dPesoBrutoKG
        .dPesoLiqKG = tDIInfo.dPesoLiqKG
        .dValorMercadoriaMoeda = tDIInfo.dValorMercadoriaMoeda
        .dValorFreteInternacMoeda = tDIInfo.dValorFreteInternacMoeda
        .dValorSeguroInternacMoeda = tDIInfo.dValorSeguroInternacMoeda
        .dValorMercadoriaEmReal = tDIInfo.dValorMercadoriaEmReal
        .dValorFreteInternacEmReal = tDIInfo.dValorFreteInternacEmReal
        .dValorSeguroInternacEmReal = tDIInfo.dValorSeguroInternacEmReal
        .dIIValor = tDIInfo.dIIValor
        .dIPIValor = tDIInfo.dIPIValor
        .dPISValor = tDIInfo.dPISValor
        .dCOFINSValor = tDIInfo.dCOFINSValor
        .dICMSValor = tDIInfo.dICMSValor
        .dValorDespesas = tDIInfo.dValorDespesas
        .iMoedaFrete = tDIInfo.iMoedaFrete
        .iMoedaItens = tDIInfo.iMoedaItens
        .iMoedaMercadoria = tDIInfo.iMoedaMercadoria
        .iMoedaSeguro = tDIInfo.iMoedaSeguro
        .dtDataDesembaraco = tDIInfo.dtDataDesembaraco
        .sUFDesembaraco = tDIInfo.sUFDesembaraco
        .sLocalDesembaraco = tDIInfo.sLocalDesembaraco
        .sCodExportador = tDIInfo.sCodExportador
        
        'nfe 3.10
        .iViaTransp = tDIInfo.iViaTransp
        .iIntermedio = tDIInfo.iIntermedio
        .sCNPJAdquir = tDIInfo.sCNPJAdquir
        .sUFAdquir = tDIInfo.sUFAdquir

    End With
    
    lErro = DIInfo_Le_Adicoes(objDIInfo, alComando)
    If lErro <> SUCESSO Then gError 184662
    
    lErro = DIInfo_Le_Despesas(objDIInfo, alComando)
    If lErro <> SUCESSO Then gError 184663
    
    lErro = DIInfo_Le_ItensPCDI(objDIInfo)
    If lErro <> SUCESSO Then gError 210608
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    DIInfo_Le = SUCESSO

    Exit Function

Erro_DIInfo_Le:

    DIInfo_Le = gErr

    Select Case gErr

        Case 184501
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 184502, 184503
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIINFO", gErr)

        Case 184504 'Sem dados -> Tratado na rotina chamadora
        
        Case 184662, 184663, 210608
        
        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184505)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function DIInfo_Le_Adicoes(ByVal objDIInfo As ClassDIInfo, alComando() As Long) As Long

Dim lErro As Long, tAdicao As typeAdicaoDI, objAdicao As ClassAdicaoDI

On Error GoTo Erro_DIInfo_Le_Adicoes

    With tAdicao
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sCodFabricante = String(STRING_DI_AD_COD_FABR, 0)
        .sNumDrawback = String(STRING_DRAWBACK_NUMERO, 0)
        lErro = Comando_Executar(alComando(2), "SELECT NumIntDoc, Seq, IPICodigo, IIAliquota, IPIAliquota, PISAliquota, COFINSAliquota, " & _
            "ICMSAliquota, IIValor, IPIValor, PISValor, COFINSValor, ICMSValor, IPIBase, PISBase, COFINSBase, ICMSBase, ValorAduaneiro, DespesaAduaneira, TaxaSiscomex, CodFabricante, ICMSPercRedBase, NumDrawback FROM AdicaoDI WHERE NumIntDI = ? ORDER BY Seq", _
            .lNumIntDoc, .iSeq, .sIPICodigo, .dIIAliquota, .dIPIAliquota, .dPISAliquota, .dCOFINSAliquota, _
            .dICMSAliquota, .dIIValor, .dIPIValor, .dPISValor, .dCOFINSValor, .dICMSValor, .dIPIBase, .dPISBase, .dCOFINSBase, .dICMSBase, .dValorAduaneiro, .dDespesaAduaneira, .dTaxaSiscomex, .sCodFabricante, .dICMSPercRedBase, .sNumDrawback, objDIInfo.lNumIntDoc)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 184649
    
    'Busca Primeiro
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184650
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objAdicao = New ClassAdicaoDI
    
        With objAdicao
            .lNumIntDoc = tAdicao.lNumIntDoc
            
            .lNumIntDI = objDIInfo.lNumIntDoc
            
            .iSeq = tAdicao.iSeq
            .sIPICodigo = tAdicao.sIPICodigo
            .dIIAliquota = tAdicao.dIIAliquota
            .dIPIAliquota = tAdicao.dIPIAliquota
            .dPISAliquota = tAdicao.dPISAliquota
            .dCOFINSAliquota = tAdicao.dCOFINSAliquota
            .dICMSAliquota = tAdicao.dICMSAliquota
            .dIIValor = tAdicao.dIIValor
            .dIPIValor = tAdicao.dIPIValor
            .dPISValor = tAdicao.dPISValor
            .dCOFINSValor = tAdicao.dCOFINSValor
            .dICMSValor = tAdicao.dICMSValor
            .dIPIBase = tAdicao.dIPIBase
            .dPISBase = tAdicao.dPISBase
            .dCOFINSBase = tAdicao.dCOFINSBase
            .dICMSBase = tAdicao.dICMSBase
            .dValorAduaneiro = tAdicao.dValorAduaneiro
            .dDespesaAduaneira = tAdicao.dDespesaAduaneira
            .dTaxaSiscomex = tAdicao.dTaxaSiscomex
            .sCodFabricante = tAdicao.sCodFabricante
            .dICMSPercRedBase = tAdicao.dICMSPercRedBase
            .sNumDrawback = tAdicao.sNumDrawback
        End With
        
        objDIInfo.colAdicoesDI.Add objAdicao
        
        lErro = AdicaoDI_Le_Itens(objAdicao, alComando)
        If lErro <> SUCESSO Then gError 184651
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184652
    
    Loop
    
    DIInfo_Le_Adicoes = SUCESSO
    
    Exit Function
    
Erro_DIInfo_Le_Adicoes:

    DIInfo_Le_Adicoes = gErr

    Select Case gErr

        Case 184651
        
        Case 184649, 184650, 184652
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ADICAODI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184653)

    End Select
    
    Exit Function

End Function

Private Function AdicaoDI_Le_Itens(ByVal objAdicaoDI As ClassAdicaoDI, alComando() As Long) As Long

Dim lErro As Long, tItem As typeItemAdicaoDI, objItem As ClassItemAdicaoDI

On Error GoTo Erro_AdicaoDI_Le_Itens


    With tItem
        .sProduto = String(STRING_PRODUTO, 0)
        .sUM = String(STRING_UM_SIGLA, 0)
        .sDescricao = String(STRING_DESCRICAO_ITEM, 0)
        
        lErro = Comando_Executar(alComando(3), "SELECT NumIntDoc, Seq, Produto, UM, Quantidade, ValorUnitFOBNaMoeda, ValorUnitFOBEmReal, ValorUnitCIFNaMoeda, ValorUnitCIFEmReal, " & _
            "ValorTotalFOBNaMoeda , ValorTotalFOBEmReal, ValorTotalCIFNaMoeda, ValorTotalCIFEmReal, TotalCIFEmRealManual, PesoBruto, PesoLiq, Descricao, ValorUnitTrib, IPIUnidadePadraoValor FROM ItensAdicaoDI WHERE NumIntAdicaoDI = ? ORDER BY Seq", _
            .lNumIntDoc, .iSeq, .sProduto, .sUM, .dQuantidade, .dValorUnitFOBNaMoeda, .dValorUnitFOBEmReal, .dValorUnitCIFNaMoeda, .dValorUnitCIFEmReal, _
            .dValorTotalFOBNaMoeda, .dValorTotalFOBEmReal, .dValorTotalCIFNaMoeda, .dValorTotalCIFEmReal, .iTotalCIFEmRealManual, .dPesoBruto, .dPesoLiq, .sDescricao, .dValorUnitTrib, .dIPIUnidadePadraoValor, objAdicaoDI.lNumIntDoc)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 184654
    
    'Busca Primeiro
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184655
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objItem = New ClassItemAdicaoDI
    
        With objItem
            .lNumIntDoc = tItem.lNumIntDoc
            
            .lNumIntAdicaoDI = objAdicaoDI.lNumIntDoc
            .iAdicao = objAdicaoDI.iSeq
            
            .iSeq = tItem.iSeq
            .sProduto = tItem.sProduto
            .sUM = tItem.sUM
            .dQuantidade = tItem.dQuantidade
            .dValorUnitFOBNaMoeda = tItem.dValorUnitFOBNaMoeda
            .dValorUnitFOBEmReal = tItem.dValorUnitFOBEmReal
            .dValorUnitCIFNaMoeda = tItem.dValorUnitCIFNaMoeda
            .dValorUnitCIFEmReal = tItem.dValorUnitCIFEmReal
            .dValorTotalFOBNaMoeda = tItem.dValorTotalFOBNaMoeda
            .dValorTotalFOBEmReal = tItem.dValorTotalFOBEmReal
            .dValorTotalCIFNaMoeda = tItem.dValorTotalCIFNaMoeda
            .dValorTotalCIFEmReal = tItem.dValorTotalCIFEmReal
            .iTotalCIFEmRealManual = tItem.iTotalCIFEmRealManual
            .dPesoBruto = tItem.dPesoBruto
            .dPesoLiq = tItem.dPesoLiq
            .sDescricao = tItem.sDescricao
            .dValorUnitTrib = tItem.dValorUnitTrib
            .dIPIUnidadePadraoValor = tItem.dIPIUnidadePadraoValor
        End With
        
        objAdicaoDI.colItensAdicaoDI.Add objItem
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184656
    
    Loop
    
    AdicaoDI_Le_Itens = SUCESSO
    
    Exit Function
    
Erro_AdicaoDI_Le_Itens:

    AdicaoDI_Le_Itens = gErr

    Select Case gErr

        Case 184654, 184655, 184656
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENS_ADICAODI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184657)

    End Select
    
    Exit Function

End Function

Private Function DIInfo_Le_Despesas(ByVal objDIInfo As ClassDIInfo, alComando() As Long) As Long

Dim lErro As Long, tDespesa As typeImportCompl, objDespesa As ClassImportCompl

On Error GoTo Erro_DIInfo_Le_Despesas

    With tDespesa
        .sDescricao = String(STRING_IMPORTCOMPL_DESCRICAO, 0)
        lErro = Comando_Executar(alComando(4), "SELECT NumIntDoc, Seq, Tipo, Descricao, Valor, Perc, Dias FROM ImportCompl WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ? ORDER BY Seq", _
            .lNumIntDoc, .iSeq, .iTipo, .sDescricao, .dValor, .dPerc, .iDias, IMPORTCOMPL_ORIGEM_DI, objDIInfo.lNumIntDoc)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 184658
    
    lErro = Comando_BuscarProximo(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184659
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objDespesa = New ClassImportCompl
    
        With objDespesa
            .lNumIntDoc = tDespesa.lNumIntDoc
            
            .iTipoDocOrigem = IMPORTCOMPL_ORIGEM_DI
            .lNumIntDocOrigem = objDIInfo.lNumIntDoc
            
            .iSeq = tDespesa.iSeq
            .iTipo = tDespesa.iTipo
            .sDescricao = tDespesa.sDescricao
            .dValor = tDespesa.dValor
            .dPerc = tDespesa.dPerc
            .iDias = tDespesa.iDias
        End With
    
        objDIInfo.colDespesasDI.Add objDespesa
        
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184660
    
    Loop
    
    DIInfo_Le_Despesas = SUCESSO
    
    Exit Function
    
Erro_DIInfo_Le_Despesas:

    DIInfo_Le_Despesas = gErr

    Select Case gErr

        Case 184658, 184659, 184660
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DESPESAS_DI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184661)

    End Select
    
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TiposImportCompl_Le(ByVal objTipoImportCompl As ClassTipoImportCompl) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTipoImportCompl As typeTipoImportCompl

On Error GoTo Erro_TiposImportCompl_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 196550

    'Alocação de espaço no buffer
    tTipoImportCompl.sDescReduzida = String(50, 0)
    tTipoImportCompl.sDescricao = String(250, 0)

    'Le a tabelaTiposImportCompl
    lErro = Comando_Executar(lComando, "SELECT Codigo, DescReduzida, Descricao, IncluiValorProdutos, IncluiBaseICMS, ImpressaoNaNF, " & _
                "SeqImpressaoNF, LinhaPadraoNaTela, ComAliquota, PodeSerDespAduaneira, TipoRateio,Gerencial,AceitaDias,AceitaPerc,AceitaValor,IncluiNoValorAduaneiro FROM TiposImportCompl WHERE Codigo= ? ", _
                tTipoImportCompl.iCodigo, tTipoImportCompl.sDescReduzida, tTipoImportCompl.sDescricao, tTipoImportCompl.iIncluiValorProdutos, _
                tTipoImportCompl.iIncluiBaseICMS, tTipoImportCompl.iImpressaoNaNF, tTipoImportCompl.iSeqImpressaoNF, tTipoImportCompl.iLinhaPadraoNaTela, tTipoImportCompl.iComAliquota, _
                tTipoImportCompl.iPodeSerDespAduaneira, tTipoImportCompl.iTipoRateio, tTipoImportCompl.iGerencial, tTipoImportCompl.iAceitaDias, tTipoImportCompl.iAceitaPerc, tTipoImportCompl.iAceitaValor, tTipoImportCompl.iIncluiNoValorAduaneiro, _
                objTipoImportCompl.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 196551

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 196552

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objTipoImportCompl.iCodigo = tTipoImportCompl.iCodigo
    objTipoImportCompl.sDescReduzida = tTipoImportCompl.sDescReduzida
    objTipoImportCompl.sDescricao = tTipoImportCompl.sDescricao
    objTipoImportCompl.iIncluiValorProdutos = tTipoImportCompl.iIncluiValorProdutos
    objTipoImportCompl.iIncluiBaseICMS = tTipoImportCompl.iIncluiBaseICMS
    objTipoImportCompl.iImpressaoNaNF = tTipoImportCompl.iImpressaoNaNF
    objTipoImportCompl.iSeqImpressaoNF = tTipoImportCompl.iSeqImpressaoNF
    objTipoImportCompl.iLinhaPadraoNaTela = tTipoImportCompl.iLinhaPadraoNaTela
    objTipoImportCompl.iComAliquota = tTipoImportCompl.iComAliquota
    objTipoImportCompl.iPodeSerDespAduaneira = tTipoImportCompl.iPodeSerDespAduaneira
    objTipoImportCompl.iTipoRateio = tTipoImportCompl.iTipoRateio
    objTipoImportCompl.iGerencial = tTipoImportCompl.iGerencial
    objTipoImportCompl.iAceitaDias = tTipoImportCompl.iAceitaDias
    objTipoImportCompl.iAceitaPerc = tTipoImportCompl.iAceitaPerc
    objTipoImportCompl.iAceitaValor = tTipoImportCompl.iAceitaValor
    objTipoImportCompl.iIncluiNoValorAduaneiro = tTipoImportCompl.iIncluiNoValorAduaneiro

    'Fecha Comando
    Call Comando_Fechar(lComando)

    TiposImportCompl_Le = SUCESSO

    Exit Function

Erro_TiposImportCompl_Le:

    TiposImportCompl_Le = gErr

    Select Case gErr

        Case 196550
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 196551, 196552
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSIMPORTCOMPL", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196553)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function NFiscal_Le_NFImportacao(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long, iIndice As Integer, objNFImportacao As New ClassNFImportacao
Dim alComando(1 To 5) As Long, objDIInfo As New ClassDIInfo, lNumIntDI As Long, objTipoImportCompl As ClassTipoImportCompl
Dim sDINumero As String, tItemAdicao As typeItemAdicaoDIItemNF, objItemNF As ClassItemNF
Dim iItem As Integer, iSeqItemAdicaoDI As Integer, iSeqAdicaoDI As Integer, objItemAdicaoDIItemNF As ClassItemAdicaoDIItemNF
Dim dIPIAliquota As Double, dICMSAliquota As Double, objImportCompl As ClassImportCompl, dICMSPercRedBase As Double
Dim lICNumIntDoc As Long, iICSeq As Integer, iICTipo As Integer, sICDescricao As String, dICValorBase As Double, dICAliquota As Double, dICValor As Double, iICManual As Integer

On Error GoTo Erro_NFiscal_Le_NFImportacao

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184704
    Next
    
    sDINumero = String(STRING_DI_NUMERO, 0)
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDI, DIInfo.Numero FROM NFImportacao, DIInfo WHERE NFImportacao.NumIntDI = DIInfo.NumIntDoc AND NFImportacao.NumIntNF = ?", lNumIntDI, sDINumero, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 184730
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184731
    If lErro <> AD_SQL_SUCESSO Then gError 184739

    'Lê o DIInfo que está sendo Passado
    objDIInfo.sNumero = sDINumero
    lErro = CF("DIInfo_Le", objDIInfo)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 184732

    Set objNFImportacao.objDIInfo = objDIInfo
    
    With tItemAdicao
        lErro = Comando_Executar(alComando(2), "SELECT ItemAdicaoDIItemNF.NumIntItemNF, ItemAdicaoDIItemNF.NumIntItemAdicaoDI, ItemAdicaoDIItemNF.ValorAduaneiro, ItemAdicaoDIItemNF.ValorII, ItemAdicaoDIItemNF.DespImpValorRateado, ItensNFiscal.Item, ItensAdicaoDI.Seq, AdicaoDI.Seq, AdicaoDI.IPIAliquota, AdicaoDI.ICMSAliquota, AdicaoDI.ICMSPercRedBase FROM ItemAdicaoDIItemNF, ItensNFiscal, ItensAdicaoDI, AdicaoDI WHERE ItensAdicaoDI.NumIntAdicaoDI = AdicaoDI.NumIntDoc AND ItensAdicaoDI.NumIntDoc = ItemAdicaoDIItemNF.NumIntItemAdicaoDI AND ItemAdicaoDIItemNF.NumIntItemNF = ItensNFiscal.NumIntDoc AND ItensNFiscal.NumIntNF = ? ORDER BY ItensNFiscal.Item", _
            .lNumIntItemNF, .lNumIntItemAdicaoDI, .dValorAduaneiro, .dValorII, .dDespImpValorRateado, iItem, iSeqItemAdicaoDI, iSeqAdicaoDI, dIPIAliquota, dICMSAliquota, dICMSPercRedBase, objNFiscal.lNumIntDoc)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 184733
    
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184734
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objItemNF = objNFiscal.colItensNF(iItem)
    
        objItemNF.dDespImpValorRateado = tItemAdicao.dDespImpValorRateado
        objItemNF.objTributacaoItemNF.dICMSAliquotaAdicaoDI = dICMSAliquota
        objItemNF.objTributacaoItemNF.dICMSPercRedBaseAdicaoDI = dICMSPercRedBase
        
        Set objItemAdicaoDIItemNF = New ClassItemAdicaoDIItemNF
        
        With objItemAdicaoDIItemNF
        
            .dValorAduaneiro = tItemAdicao.dValorAduaneiro
            .dValorII = tItemAdicao.dValorII
            .dDespImpValorRateado = tItemAdicao.dDespImpValorRateado
            
            .iAdicao = iSeqAdicaoDI
            .iItemAdicao = iSeqItemAdicaoDI
            
            .sProduto = objItemNF.sProduto
            .sDescricao = objItemNF.sDescricaoItem
            .sUM = objItemNF.sUnidadeMed
            .dQuantidade = objItemNF.dQuantidade
            
            .dPrecoUnitario = objItemNF.dPrecoUnitario
            
            .dIPIAliquotaAdicaoDI = dIPIAliquota
            .dICMSAliquotaAdicaoDI = dICMSAliquota
            .dICMSPercRedBaseAdicaoDI = dICMSPercRedBase
            
        End With
        
        objNFImportacao.colItensNF.Add objItemAdicaoDIItemNF
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184735
    
    Loop
    
    sICDescricao = String(250, 0)
    lErro = Comando_Executar(alComando(3), "SELECT NumIntDoc, Seq, Tipo, Descricao, ValorBase, Aliquota, Valor, Manual FROM ImportCompl WHERE TipoDocOrigem = ? AND NumIntDocOrigem = ? ORDER BY Seq", _
        lICNumIntDoc, iICSeq, iICTipo, sICDescricao, dICValorBase, dICAliquota, dICValor, iICManual, IMPORTCOMPL_ORIGEM_NF, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 184736
    
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184737
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objImportCompl = New ClassImportCompl
        
        With objImportCompl
                            
            .iSeq = iICSeq
            .iTipo = iICTipo
            .sDescricao = sICDescricao
            .dValorBase = dICValorBase
            .dValor = dICValor
            .iManual = iICManual
        
        End With
        
        Set objTipoImportCompl = New ClassTipoImportCompl
    
        objTipoImportCompl.iCodigo = objImportCompl.iTipo
        
        lErro = CF("TiposImportCompl_Le", objTipoImportCompl)
        If lErro <> SUCESSO And lErro <> ERRO_ITEM_NAO_CADASTRADO Then gError 184703
        If lErro <> SUCESSO Then gError 184704
        
        Set objImportCompl.objTipoImportCompl = objTipoImportCompl
        
        objNFImportacao.colComplNF.Add objImportCompl
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184738
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Set objNFiscal.objNFImportacao = objNFImportacao
    
    NFiscal_Le_NFImportacao = SUCESSO
    
    Exit Function
    
Erro_NFiscal_Le_NFImportacao:

    NFiscal_Le_NFImportacao = gErr

    Select Case gErr

        Case 184739
        
        Case 184704
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184703)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function PV_FaltaEstoque_Preenche_Trat_Cust(ByVal objCombo As Object) As Long
    PV_FaltaEstoque_Preenche_Trat_Cust = SUCESSO
End Function

Function ItemNFiscalSaida_Le_NumNFItem1(ByVal objItemNF As ClassItemNF) As Long
'Lê o Ítem de Uma Nota Fiscal de Saida através do Número da Nota Fiscal, Série e Numero do Ítem passados

Dim lErro As Long
Dim lComando As Long
Dim tItemNF As typeItemNF

On Error GoTo Erro_ItemNFiscalSaida_Le_NumNFItem1

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 199743

    With tItemNF
        .sDescricaoItem = String(STRING_ITEMNF_DESCRICAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sUnidadeMed = String(STRING_UM_SIGLA, 0)

        'Busca o Ítem
        lErro = Comando_Executar(lComando, "SELECT ItensNFiscal.NumIntNF, ItensNFiscal.Status, ItensNFiscal.Produto, ItensNFiscal.UnidadeMed, ItensNFiscal.Quantidade, ItensNFiscal.Almoxarifado, ItensNFiscal.PrecoUnitario, ItensNFiscal.PercDesc, ItensNFiscal.ValorDesconto, ItensNFiscal.DataEntrega, ItensNFiscal.DescricaoItem, ItensNFiscal.ValorAbatComissao, ItensNFiscal.NumIntPedVenda, ItensNFiscal.NumIntItemPedVenda, ItensNFiscal.NumIntDoc, ItensNFiscal.NumIntTrib, ItensNFiscal.NumIntDocOrig FROM NFiscal, ItensNFiscal, TiposDocInfo WHERE NFiscal.Serie = ? AND NFiscal.NumNotaFiscal = ? AND ItensNFiscal.Item = ? AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.FilialEmpresa = ? AND NFiscal.Status <> ? AND ItensNFiscal.Status <> ? AND TiposDocInfo.Tipo = ? ORDER BY DataEmissao DESC", _
        .lNumIntNF, .iStatus, .sProduto, .sUnidadeMed, .dQuantidade, .iAlmoxarifado, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, .dValorAbatComissao, .lNumIntPedVenda, .lNumIntItemPedVenda, .lNumIntDoc, .lNumIntTrib, .lNumIntDocOrig, objItemNF.sSerieNFOrig, objItemNF.lNumNFOrig, objItemNF.iItem, giFilialEmpresa, STATUS_EXCLUIDO, STATUS_EXCLUIDO, DOCINFO_NF_INT_SAIDA)
        If lErro <> AD_SQL_SUCESSO Then gError 199744

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199745
        
        If lErro = AD_SQL_SEM_DADOS Then gError 199746 'Não encontrou

        'Preenche o objItemNF com os dados lidos
        objItemNF.dPercDesc = .dPercDesc
        objItemNF.dPrecoUnitario = .dPrecoUnitario
        objItemNF.dQuantidade = .dQuantidade
        objItemNF.dtDataEntrega = .dtDataEntrega
        objItemNF.dValorAbatComissao = .dValorAbatComissao
        objItemNF.dValorDesconto = .dValorDesconto
        objItemNF.iAlmoxarifado = .iAlmoxarifado
        objItemNF.iStatus = .iStatus
        objItemNF.lNumIntDoc = .lNumIntDoc
        objItemNF.lNumIntDocOrig = .lNumIntDocOrig
        objItemNF.lNumIntItemPedVenda = .lNumIntPedVenda
        objItemNF.lNumIntTrib = .lNumIntTrib
        objItemNF.sDescricaoItem = .sDescricaoItem
        objItemNF.sProduto = .sProduto
        objItemNF.sUnidadeMed = .sUnidadeMed

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ItemNFiscalSaida_Le_NumNFItem1 = SUCESSO

    Exit Function

Erro_ItemNFiscalSaida_Le_NumNFItem1:

    ItemNFiscalSaida_Le_NumNFItem1 = gErr

    Select Case gErr

        Case 199743
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 199744, 199745
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL2", gErr)

        Case 199746

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 199747)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Custo_Kit_MP2(ByVal objMnemonicoValor As ClassMnemonicoValor, ByVal sProduto As String) As Long
'parametros:Produto
'retorna o custo das materias primas necessarias para fabricar o produto,
'incluindo as materias primas para os subprodutos, recursivamente

Dim lErro As Long, dCusto As Double
Dim colMP As New Collection, dCustoProd As Double
Dim dCustoMP As Double
Dim objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP
Dim objProduto As New ClassProduto, iAux As Integer
Dim sTipo As String, iIndice As Integer, sCriterio As String
Dim iPreenchido As Integer, sPrioridade As String
Dim dFator As Double, lComando As Long
Dim dtDataUltEnt As Date
Dim dAliquotaICMSUltEnt As Double
Dim dFrete As Double
Dim dAliquota As Double, sProdMP As String
Dim bGeraLog As Boolean, bGeraLogR As Boolean, sAux As String, sDirLog As String, dPercIPI As Double
Dim bSomaFrete As Boolean, dPercFrete As Double, bSomaIPI As Boolean, iNumConfig As Integer, sConfig As String
Dim colProdAux As New Collection, objProdAux As ClassProduto, dPercFreteProd As Double, sTipoCusto As String
Dim lRetorno As Long, iPos As Integer, iPosAnt As Integer

On Error GoTo Erro_Custo_Kit_MP2

    sAux = objMnemonicoValor.vParam(2)

    bGeraLog = False
    bGeraLogR = False
    bSomaFrete = False
    bSomaIPI = False
    dPercFrete = 0

    If Len(Trim(sAux)) > 3 Then
    
        sPrioridade = Replace(Trim(left(sAux, 3)), "0", "")
        
        iNumConfig = (Len(Trim(sAux)) / 3) - 1
        For iAux = 1 To iNumConfig
            sConfig = Mid(sAux, iAux * 3 + 1, 3)
            Select Case left(sConfig, 1)
                Case "L"
                    If Mid(sConfig, 2, 1) = "1" Then bGeraLogR = True
                    If Mid(sConfig, 3, 1) = "1" Then bGeraLog = True
                Case "F"
                    bSomaFrete = True
                    dPercFrete = StrParaDbl(right(sConfig, 2)) / 100
                Case "I"
                    If sConfig = "IPI" Then bSomaIPI = True
            End Select
        Next
        
    Else
        sPrioridade = sAux
    End If
    
    If bGeraLog Then
        sDirLog = String(512, 0)
        lRetorno = GetPrivateProfileString("Geral", "ArqLog", "C:\SGE\Programa\", sDirLog, 512, "ADM100.INI")
        sDirLog = left(sDirLog, lRetorno)
        
        iPos = InStr(1, sDirLog, "/")
        If iPos = 0 Then iPos = InStr(1, sDirLog, "\")
        iPosAnt = iPos
        Do While iPos <> 0
            iPos = InStr(iPosAnt + 1, sDirLog, "/")
            If iPos = 0 Then iPos = InStr(iPosAnt + 1, sDirLog, "\")
            If iPos <> 0 Then iPosAnt = iPos
        Loop
    
        If iPosAnt <> 0 Then sDirLog = left(sDirLog, iPosAnt)
    End If

    If bGeraLog Then Open sDirLog & "FP_CUSTO_MP_" & Trim(sProduto) & ".txt" For Output As #1
    
    If bGeraLog Then Print #1, "INSUMO;PROD_PAI;MP;UM;QTDE;CUSTO_UNIT;%FRETE;%IPI;CUSTO_TOTAL;TIPO_CUSTO"
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 199743

    If sProduto = "" Then

        lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 133017

    End If

    lErro = CF("CustoDireto_ObtemMP", giFilialEmpresa, Year(gdtDataAtual), sProduto, 1, colMP, FORMACAO_PRECO_ROTINA_CALCPRECO)
    If lErro <> SUCESSO Then gError 133014

    'sPrioridade = objMnemonicoValor.vParam(2)
    
    'para cada materia prima
    For Each objCustoDirFabrPlanMP In colMP

        dCustoMP = 0
        dCusto = 0
        dPercIPI = 0
        dPercFreteProd = 0
        
        sProdMP = objCustoDirFabrPlanMP.sProdutoMP
         
        objProduto.sCodigo = sProdMP

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError 133013
        
        If bSomaIPI Then dPercIPI = objProduto.dIPIAliquota
        If objProduto.iNatureza = NATUREZA_PROD_MATERIA_PRIMA Then dPercFreteProd = dPercFrete
        
        For iAux = 1 To Len(sPrioridade)
        
            sCriterio = Mid(sPrioridade, iAux, 1)
            
            Select Case sCriterio
            
                Case "1" 'custo como na tela produto x almoxarifado
                    sTipo = STRING_QUANT_DISPONIVEL_NOSSA
            
                    'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
                    lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, sTipo)
                    If lErro <> SUCESSO Then gError 133020
                    
                    sTipoCusto = "Custo Estoque"

                Case "2" 'custo ultima entrada
                
                    lErro = CF("Produto_ObterDadosUltEnt", objProduto, giFilialEmpresa, dCusto, dtDataUltEnt, dAliquotaICMSUltEnt, dFrete)
                    If lErro <> SUCESSO Then gError 178579
                    
                    sTipoCusto = "NF Entrada"
            
                Case "3" 'custo cadastrado para formacao de precos
                    lErro = CF("CustoDiretoProd_ObtemValor", giFilialEmpresa, sProdMP, 0, 0, 0, dCusto, lComando)
                    If lErro <> SUCESSO Then gError 178579
                    
                    sTipoCusto = "Manual - FP"
                                
            End Select
            
            If dCusto <> 0 Then
            
                'O custo retornado está na U.M. de estoque
                'Para calcular preço de venda é preciso calcular o custo na U.M. venda
                lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMVenda, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 178582
                        
                dCustoMP = dCusto * dFator * (1 + dPercFreteProd) * (1 + dPercIPI)
                
                Exit For
                
            End If
            
        Next
        
        'com a qtde e custo unitario, acumulo o custo do insumo no custo do produto.
        dCustoProd = dCustoProd + (dCustoMP * objCustoDirFabrPlanMP.dQtde)

        If bGeraLog Then Print #1, """" & Trim(objCustoDirFabrPlanMP.sProdutoPai1) & """" & ";" & """" & Trim(objCustoDirFabrPlanMP.sProdutoPai) & """" & ";" & """" & Trim(objCustoDirFabrPlanMP.sProdutoMP) & """" & ";" & objProduto.sSiglaUMVenda & ";" & CStr(objCustoDirFabrPlanMP.dQtde) & ";" & CStr(dCusto) & ";" & CStr(dPercFreteProd) & ";" & CStr(dPercIPI) & ";" & CStr((dCustoMP * objCustoDirFabrPlanMP.dQtde)) & ";" & sTipoCusto

    Next
    
    If bGeraLog Then Close #1

    objMnemonicoValor.colValor.Add dCustoProd

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Custo_Kit_MP2 = SUCESSO

    Exit Function

Erro_Custo_Kit_MP2:

    Custo_Kit_MP2 = gErr

    Select Case gErr

        Case 133013, 133014, 133015, 133020

        Case 199743
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149704)

    End Select

    Call Comando_Fechar(lComando)
    
    If bGeraLog Then Close #1
    
    Exit Function

End Function

Public Function NFiscalEntrada_Verifica_Existencia2_Cust(ByVal lErro As Long, ByVal objNFiscal As ClassNFiscal, vbMsgRes As VbMsgBoxResult, Optional ByVal bSemAviso As Boolean = False) As Long
    'Nota Fiscal Externa Igual
    If lErro = 35384 Then
        If Not bSemAviso Then vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_EXTERNA_CONTAB", objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
    'Nota Fiscal Interna Igual
    ElseIf lErro = 42083 Then
        If Not bSemAviso Then vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALTERACAO_NFISCAL_INTERNA_CONTAB", objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao)
    End If
    NFiscalEntrada_Verifica_Existencia2_Cust = SUCESSO
End Function

Public Function PreencheProduto_Customizado(ByVal objCombo As Object, ByVal sProdPreenchido As String) As Long
    PreencheProduto_Customizado = SUCESSO
End Function


'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function ProdutoCNAE_Le(ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long
Dim tProdutoCNAE As typeProdutoCNAE
Dim objProdutoCNAE As New ClassProdutoCNAE
Dim tCodTribMun As typeCodTribMun
Dim objCodTribMun As ClassCodTribMun

On Error GoTo Erro_ProdutoCNAE_Le

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 208099
    Next
    
    Set objProduto.objProdutoCNAE = objProdutoCNAE
    Set objProdutoCNAE.colCidades = New Collection

    'Alocação de espaço no buffer
    tProdutoCNAE.sISSQN = String(STRING_ISSQN_CODIGO, 0)
    tProdutoCNAE.sCNAE = String(STRING_CNAE_CODIGO, 0)

    'Le a tabela ProdutoCNAE
    lErro = Comando_Executar(alComando(1), "SELECT C.CNAE, P.ISSQN, C.LocServCliente, C.LocIncidImpCliente FROM ProdutoCNAE AS C, Produtos AS P WHERE P.Codigo = C.Produto AND C.Produto= ? ", _
                tProdutoCNAE.sCNAE, tProdutoCNAE.sISSQN, tProdutoCNAE.iLocServCliente, tProdutoCNAE.iLocIncidImpCliente, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 208100

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208101

    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objProduto.sISSQN = tProdutoCNAE.sISSQN
    objProdutoCNAE.sProduto = objProduto.sCodigo
    objProdutoCNAE.sCNAE = tProdutoCNAE.sCNAE
    objProdutoCNAE.iLocServCliente = tProdutoCNAE.iLocServCliente
    objProdutoCNAE.iLocIncidImpCliente = tProdutoCNAE.iLocIncidImpCliente
    
    tCodTribMun.sCodTribMun = String(STRING_CODTRIBMUN, 0)
    
    'Le a tabela ProdutoCNAE
    lErro = Comando_Executar(alComando(2), "SELECT Cidade, CodTribMun, Aliquota FROM CodTribMun WHERE Produto= ? ORDER BY Seq", _
                tCodTribMun.lCidade, tCodTribMun.sCodTribMun, tCodTribMun.dAliquota, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 208102

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208103
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objCodTribMun = New ClassCodTribMun
        
        objCodTribMun.sProduto = objProduto.sCodigo
        objCodTribMun.dAliquota = tCodTribMun.dAliquota
        objCodTribMun.lCidade = tCodTribMun.lCidade
        objCodTribMun.sCodTribMun = tCodTribMun.sCodTribMun
    
        objProdutoCNAE.colCidades.Add objCodTribMun
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208104
    
    Loop

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ProdutoCNAE_Le = SUCESSO

    Exit Function

Erro_ProdutoCNAE_Le:

    ProdutoCNAE_Le = gErr

    Select Case gErr

        Case 208099
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208100, 208101
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOCNAE", gErr)

        Case 208102 To 208104
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CODTRIBMUN", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208105)

    End Select

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function CodServMun_Le(ByVal objCodServMun As ClassCodServMun) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCodServMun As typeCodServMun

On Error GoTo Erro_CodServMun_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208134

    'Alocação de espaço no buffer
    tCodServMun.sISSQNBase = String(STRING_ISSQN_CODIGO, 0)
    tCodServMun.sDescricao1 = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tCodServMun.sDescricao2 = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaCodServMun
    lErro = Comando_Executar(lComando, "SELECT ISSQNBase, PadraoISSQN, Descricao1, Descricao2, " & _
                "Aliquota FROM CodServMun WHERE CodIBGE= ?  AND CodServ= ? ", _
                tCodServMun.sISSQNBase, tCodServMun.iPadraoISSQN, _
                tCodServMun.sDescricao1, tCodServMun.sDescricao2, tCodServMun.dAliquota, _
                objCodServMun.sCodIBGE, objCodServMun.sCodServ)
    If lErro <> AD_SQL_SUCESSO Then gError 208135

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208136

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objCodServMun.sISSQNBase = tCodServMun.sISSQNBase
    objCodServMun.iPadraoISSQN = tCodServMun.iPadraoISSQN
    objCodServMun.sDescricao1 = tCodServMun.sDescricao1
    objCodServMun.sDescricao2 = tCodServMun.sDescricao2
    objCodServMun.dAliquota = tCodServMun.dAliquota

    'Fecha Comando
    Call Comando_Fechar(lComando)

    CodServMun_Le = SUCESSO

    Exit Function

Erro_CodServMun_Le:

    CodServMun_Le = gErr

    Select Case gErr

        Case 208134
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208135, 208136
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CODSERVMUN", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208137)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function CodServMun_Le_Padrao(ByVal objCodServMun As ClassCodServMun) As Long

Dim lErro As Long
Dim lComando As Long
Dim tCodServMun As typeCodServMun

On Error GoTo Erro_CodServMun_Le_Padrao

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208138

    'Alocação de espaço no buffer
    tCodServMun.sCodServ = String(STRING_CODTRIBMUN, 0)
    tCodServMun.sDescricao1 = String(STRING_OBSERVACAO_OBSERVACAO, 0)
    tCodServMun.sDescricao2 = String(STRING_OBSERVACAO_OBSERVACAO, 0)

    'Le a tabelaCodServMun
    lErro = Comando_Executar(lComando, "SELECT CodServ, Descricao1, Descricao2, " & _
                "Aliquota FROM CodServMun WHERE CodIBGE = ? AND ISSQNBase= ?  AND PadraoISSQN = 1 ", _
                tCodServMun.sCodServ, tCodServMun.sDescricao1, tCodServMun.sDescricao2, tCodServMun.dAliquota, _
                objCodServMun.sCodIBGE, objCodServMun.sISSQNBase)
    If lErro <> AD_SQL_SUCESSO Then gError 208139

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208140

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objCodServMun.sCodServ = tCodServMun.sCodServ
    objCodServMun.sDescricao1 = tCodServMun.sDescricao1
    objCodServMun.sDescricao2 = tCodServMun.sDescricao2
    objCodServMun.dAliquota = tCodServMun.dAliquota

    'Fecha Comando
    Call Comando_Fechar(lComando)

    CodServMun_Le_Padrao = SUCESSO

    Exit Function

Erro_CodServMun_Le_Padrao:

    CodServMun_Le_Padrao = gErr

    Select Case gErr

        Case 208138
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208139, 208140
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CODSERVMUN", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208141)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ProdutoGrade_Le_Seq(ByVal sSegProdAux As String, ByVal sFiltroProdAux As String, lSeqProd As Long) As Long

Dim lErro As Long
Dim lComando As Long, sAux As String

On Error GoTo Erro_ProdutoGrade_Le_Seq

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208176

    'Alocação de espaço no buffer
    sAux = String(STRING_PRODUTO, 0)

    'Le a tabelaCodServMun
    lErro = Comando_Executar(lComando, "SELECT MAX(" & sSegProdAux & ") FROM Produtos WHERE " & sFiltroProdAux, sAux)
    If lErro <> AD_SQL_SUCESSO Then gError 208177

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208178

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then
        lSeqProd = 0
    Else
        lSeqProd = StrParaLong(sAux)
    End If

    'Fecha Comando
    Call Comando_Fechar(lComando)

    ProdutoGrade_Le_Seq = SUCESSO

    Exit Function

Erro_ProdutoGrade_Le_Seq:

    ProdutoGrade_Le_Seq = gErr

    Select Case gErr

        Case 208176
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208177, 208178
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208179)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ProdutoGrade_Le_ProdCat(ByVal objProduto As ClassProduto, bExisteProd As Boolean) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long, iIndice As Integer
Dim bAchou As Boolean, sProduto As String, sProdAux As String
Dim sNomeReduzido As String, sDescricao As String, sFigura As String
Dim objProdCat As ClassProdutoCategoria

On Error GoTo Erro_ProdutoGrade_Le_ProdCat

    bAchou = False

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 208180
    Next

    'Alocação de espaço no buffer
    sProduto = String(STRING_PRODUTO, 0)
    sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    sFigura = String(STRING_PRODUTO_FIGURA, 0)

    'Le a tabela Produtos
    'PS-> Nivel está vindo com a posição final do segmento anterior
    lErro = Comando_Executar(alComando(1), "SELECT Codigo, NomeReduzido, Descricao, Figura FROM Produtos WHERE Codigo LIKE ? ", sProduto, sNomeReduzido, sDescricao, sFigura, left(objProduto.sCodigo, objProduto.iNivel) & "%")
    If lErro <> AD_SQL_SUCESSO Then gError 208181

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208182

    'Sem Dados
    Do While lErro = AD_SQL_SUCESSO

        bAchou = True
        For Each objProdCat In objProduto.colCategoriaItem
        
            sProdAux = String(STRING_PRODUTO, 0)
        
            'Le a tabela Produtos
            lErro = Comando_Executar(alComando(2), "SELECT Produto FROM ProdutoCategoria WHERE Produto = ? AND Categoria = ? AND Item = ?", sProdAux, sProduto, objProdCat.sCategoria, objProdCat.sItem)
            If lErro <> AD_SQL_SUCESSO Then gError 208183
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208184
            
            If lErro <> AD_SQL_SUCESSO Then
                bAchou = False
                Exit For
            End If
        
        Next
        
        If bAchou Then
            objProduto.sCodigo = sProduto
            objProduto.sNomeReduzido = sNomeReduzido
            objProduto.sDescricao = sDescricao
            objProduto.sFigura = sFigura
            Exit Do 'Já existe
        End If

        'Busca Próximo
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208185

    Loop
    
    bExisteProd = bAchou

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ProdutoGrade_Le_ProdCat = SUCESSO

    Exit Function

Erro_ProdutoGrade_Le_ProdCat:

    ProdutoGrade_Le_ProdCat = gErr

    Select Case gErr

        Case 208180
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208181, 208182, 208185
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)
        
        Case 208183, 208184
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOCATEGORIA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208186)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function RomaneioGrade_Le_Tela(sNomeTela As String, objRomaneioGrade As ClassRomaneioGrade) As Long

Dim lErro As Long
Dim objGrade As New ClassGrade
Dim objProduto As New ClassProduto

On Error GoTo Erro_RomaneioGrade_Le_Tela

    sNomeTela = "RomaneioGrade"

    'Passa o Produto para o obj
    objProduto.sCodigo = objRomaneioGrade.objObjetoTela.sProduto
            
    lErro = CF("Customiza_RomaneioGrade_TrataParam", objProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    'Lê o Produto passado
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
    objGrade.sCodigo = objProduto.sGrade
    
    lErro = CF("Grade_Le", objGrade)
    If lErro <> SUCESSO And lErro <> 86275 Then gError ERRO_SEM_MENSAGEM
    
    If objGrade.iLayout = GRADE_LAYOUT_COMBOS Then
        sNomeTela = "RomaneioGrade2"
    End If

    RomaneioGrade_Le_Tela = SUCESSO

    Exit Function

Erro_RomaneioGrade_Le_Tela:

    RomaneioGrade_Le_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208214)

    End Select


    Exit Function

End Function

Function FilialForn_Le_UltCompra_Prod(ByVal lFornecedor As Long, ByVal iFilial As Integer, ByVal sProduto As String, sUM As String, dPreco As Double) As Long

Dim lErro As Long
Dim lComando As Long
Dim dPrecoAux As Double
Dim sUMAux As String

On Error GoTo Erro_FilialForn_Le_UltCompra_Prod

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 208415
    
    sUMAux = String(STRING_UM_SIGLA, 0)

    lErro = Comando_Executar(lComando, "SELECT I.UnidadeMed, I.PrecoUnitario FROM NFiscal AS N, TiposDocInfo AS T, ItensNFiscal AS I WHERE N.NumIntDoc = I.NumIntNF AND N.TipoNFiscal = T.Codigo AND N.Status <> 7 AND N.Status <> 5 AND T.Compras = 1 AND N.Fornecedor = ? AND N.FilialForn = ? AND I.Produto = ? ORDER BY DataEmissao DESC ", _
        sUMAux, dPrecoAux, lFornecedor, iFilial, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 208416

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208417

    If lErro <> AD_SQL_SEM_DADOS Then
        dPreco = dPrecoAux
        sUM = sUMAux
    End If

    'Fecha comando
    Call Comando_Fechar(lComando)

    FilialForn_Le_UltCompra_Prod = SUCESSO

    Exit Function

Erro_FilialForn_Le_UltCompra_Prod:

    FilialForn_Le_UltCompra_Prod = gErr

    Select Case gErr

        Case 208415
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 208416, 208417
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208418)

    End Select

    'Fecha comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function InvCliForn_Le(ByVal objInvCliForn As ClassInvCliForn) As Long

Dim lErro As Long
Dim lComando As Long
Dim tInvCliForn As typeInvCliForn

On Error GoTo Erro_InvCliForn_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209518

    'Alocação de espaço no buffer
    tInvCliForn.sUsuario = String(STRING_USUARIO, 0)
    tInvCliForn.sObs = String(STRING_MAXIMO, 0)

    'Le a tabelaInvCliForn
    If objInvCliForn.lNumIntDoc <> 0 Then
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Data, Escaninho, TipoCliForn,CliForn, Filial, Usuario, DataGravacao, HoraGravacao, Obs FROM InvCliForn WHERE NumIntDoc = ? ", _
                    tInvCliForn.lNumIntDoc, tInvCliForn.iFilialEmpresa, tInvCliForn.dtData, tInvCliForn.iEscaninho, tInvCliForn.iTipoCliForn, tInvCliForn.lCliForn, tInvCliForn.iFilial, tInvCliForn.sUsuario, tInvCliForn.dtDataGravacao, tInvCliForn.dHoraGravacao, _
                    tInvCliForn.sObs, objInvCliForn.lNumIntDoc)
    Else
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialEmpresa, Data, Escaninho, TipoCliForn,CliForn, Filial, Usuario, DataGravacao, HoraGravacao, Obs FROM InvCliForn WHERE FilialEmpresa = ? AND Data= ?  AND Escaninho= ?  AND TipoCliForn= ?  AND CliForn= ?  AND Filial= ? ", _
                    tInvCliForn.lNumIntDoc, tInvCliForn.iFilialEmpresa, tInvCliForn.dtData, tInvCliForn.iEscaninho, tInvCliForn.iTipoCliForn, tInvCliForn.lCliForn, tInvCliForn.iFilial, tInvCliForn.sUsuario, tInvCliForn.dtDataGravacao, tInvCliForn.dHoraGravacao, _
                    tInvCliForn.sObs, objInvCliForn.iFilialEmpresa, objInvCliForn.dtData, objInvCliForn.iEscaninho, objInvCliForn.iTipoCliForn, objInvCliForn.lCliForn, objInvCliForn.iFilial)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 209519

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209520

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objInvCliForn.lNumIntDoc = tInvCliForn.lNumIntDoc
    objInvCliForn.iFilialEmpresa = tInvCliForn.iFilialEmpresa
    objInvCliForn.dtData = tInvCliForn.dtData
    objInvCliForn.iEscaninho = tInvCliForn.iEscaninho
    objInvCliForn.iTipoCliForn = tInvCliForn.iTipoCliForn
    objInvCliForn.lCliForn = tInvCliForn.lCliForn
    objInvCliForn.iFilial = tInvCliForn.iFilial
    objInvCliForn.sUsuario = tInvCliForn.sUsuario
    objInvCliForn.dtDataGravacao = tInvCliForn.dtDataGravacao
    objInvCliForn.dHoraGravacao = tInvCliForn.dHoraGravacao
    objInvCliForn.sObs = tInvCliForn.sObs

    'Fecha Comando
    Call Comando_Fechar(lComando)

    InvCliForn_Le = SUCESSO

    Exit Function

Erro_InvCliForn_Le:

    InvCliForn_Le = gErr

    Select Case gErr

        Case 209518
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209519, 209520
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVCLIFORN", gErr)

        Case ERRO_LEITURA_SEM_DADOS

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209521)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function InvCliFornItens_Le(ByVal objInv As ClassInvCliForn) As Long

Dim lErro As Long
Dim lComando As Long
Dim tInvCliFornItens As typeInvCliFornItens
Dim objInvCliFornItens As ClassInvCliFornItens

On Error GoTo Erro_InvCliFornItens_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209522

    'Alocação de espaço no buffer
    tInvCliFornItens.sProduto = String(STRING_PRODUTO, 0)
    tInvCliFornItens.sObs = String(STRING_MAXIMO, 0)

    'Le a tabelaInvCliFornItens
    lErro = Comando_Executar(lComando, "SELECT Seq, Produto, QtdCliData, QtdEncontCliData, QtdData, " & _
                "QtdAcerto, Obs FROM InvCliFornItens WHERE NumIntInvCliForn = ? ORDER BY Seq ", _
                tInvCliFornItens.iSeq, tInvCliFornItens.sProduto, tInvCliFornItens.dQtdCliData, _
                tInvCliFornItens.dQtdEncontCliData, tInvCliFornItens.dQtdData, tInvCliFornItens.dQtdAcerto, tInvCliFornItens.sObs, _
                objInv.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209523

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209524

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objInvCliFornItens = New ClassInvCliFornItens

        objInvCliFornItens.lNumIntInvCliForn = objInv.lNumIntDoc
        objInvCliFornItens.iSeq = tInvCliFornItens.iSeq
        objInvCliFornItens.sProduto = tInvCliFornItens.sProduto
        objInvCliFornItens.dQtdCliData = tInvCliFornItens.dQtdCliData
        objInvCliFornItens.dQtdEncontCliData = tInvCliFornItens.dQtdEncontCliData
        objInvCliFornItens.dQtdData = tInvCliFornItens.dQtdData
        objInvCliFornItens.dQtdAcerto = tInvCliFornItens.dQtdAcerto
        objInvCliFornItens.sObs = tInvCliFornItens.sObs
        
        objInv.colItens.Add objInvCliFornItens
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209525
        
    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    InvCliFornItens_Le = SUCESSO

    Exit Function

Erro_InvCliFornItens_Le:

    InvCliFornItens_Le = gErr

    Select Case gErr

        Case 209522
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209523 To 209525
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVCLIFORNITENS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209526)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function InvCliForn_Le_Prod_Filtro(ByVal objInv As ClassInvCliForn) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 4) As Long
Dim iSeq As Integer, objEstTerc As ClassEstoqueTerc
Dim sTabela As String, sCampoQtd As String, sMacroQtd As String, sWhere As String
Dim iMes As Integer, iAno As Integer, dtDataIniMes As Date
Dim tInvItem As typeInvCliFornItens, objInvItem As ClassInvCliFornItens
Dim colEntrada As New Collection, colSaida As New Collection, colCampos As New Collection
Dim colDistribuicao As Collection

On Error GoTo Erro_InvCliForn_Le_Prod_Filtro

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209527
    Next
    
    iAno = Year(objInv.dtData)
    iMes = Month(objInv.dtData)
    dtDataIniMes = DateAdd("d", 1 - Day(objInv.dtData), objInv.dtData)
    
    lErro = EstoqueTerc_Obtem_Dados_Escaminho(objInv.iEscaninho, sTabela, sCampoQtd)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sWhere = "(QuantInicial" & sCampoQtd & " <> 0 "
    sMacroQtd = "(QuantInicial" & sCampoQtd
    For iIndice = 1 To iMes
        sWhere = sWhere & "OR SaldoQuant" & sCampoQtd & CStr(iIndice) & " <> 0 "
        If iIndice < iMes Then
            'Soma o saldo dos meses mais a qtd inicial até o mês anterior a data
            sMacroQtd = sMacroQtd & " + SaldoQuant" & sCampoQtd & CStr(iIndice)
        End If
    Next
    sWhere = sWhere & ")"
    sMacroQtd = sMacroQtd & ")"
    
    'Alocação de espaço no buffer
    tInvItem.sProduto = String(STRING_PRODUTO, 0)

    'Le o saldo até o início do mês dos produtos que tiverem movimentação até o mês atual
    lErro = Comando_Executar(alComando(1), "SELECT Produto, " & sMacroQtd & " FROM SldMesEst" & sTabela & " WHERE FilialEmpresa = ? AND Ano = ? AND " & sWhere, tInvItem.sProduto, tInvItem.dQtdData, objInv.iFilialEmpresa, iAno)
    If lErro <> AD_SQL_SUCESSO Then gError 209528

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209529

    iSeq = 0
    Do While lErro <> AD_SQL_SEM_DADOS
    
        iSeq = iSeq + 1
    
        Set objInvItem = New ClassInvCliFornItens

        objInvItem.iSeq = 1
        objInvItem.sProduto = tInvItem.sProduto
        objInvItem.dQtdData = tInvItem.dQtdData
        
        lErro = Comando_Executar(alComando(2), "SELECT SUM(QuantEnt" & sCampoQtd & " - QuantSai" & sCampoQtd & ") FROM SldDiaEst WHERE FilialEmpresa = ? AND Produto = ? AND Data BETWEEN ? AND ? ", tInvItem.dQtdData, objInv.iFilialEmpresa, tInvItem.sProduto, dtDataIniMes, objInv.dtData)
        If lErro <> AD_SQL_SUCESSO Then gError 209530
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209531
        
        objInvItem.dQtdData = objInvItem.dQtdData + tInvItem.dQtdData
        
        Set colDistribuicao = New Collection
        
        lErro = CF("EstoqueTerc_Le_Distrib", alComando(3), alComando(4), sTabela, sCampoQtd, objInv.iFilialEmpresa, objInvItem.sProduto, objInv.dtData, colDistribuicao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Set objInvItem.colDistribuicao = colDistribuicao
        
        For Each objEstTerc In objInvItem.colDistribuicao
            If objEstTerc.iTipoCliForn = objInv.iTipoCliForn And objEstTerc.lCliForn = objInv.lCliForn And objEstTerc.iFilial = objInv.iFilial Then
                objInvItem.dQtdCliData = objEstTerc.dQuantidade
                If objInvItem.dQtdCliData > QTDE_ESTOQUE_DELTA Then objInvItem.iSeq = 0
            End If
            objInvItem.dQtdDistrib = objInvItem.dQtdDistrib + objEstTerc.dQuantidade
        Next
        
        colEntrada.Add objInvItem
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209532
        
    Loop
    
    colCampos.Add "iSeq"
    colCampos.Add "sProduto"
    
    Call Ordena_Colecao(colEntrada, colSaida, colCampos)
    
    For iIndice = objInv.colItens.Count To 1 Step -1
        objInv.colItens.Remove iIndice
    Next
    For Each objInvItem In colSaida
        objInv.colItens.Add objInvItem
    Next

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InvCliForn_Le_Prod_Filtro = SUCESSO

    Exit Function

Erro_InvCliForn_Le_Prod_Filtro:

    InvCliForn_Le_Prod_Filtro = gErr

    Select Case gErr

        Case 209527
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209528, 209529, 209532
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr)

        Case 209530, 209531
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209533)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function EstoqueTerc_Le_Distrib(ByVal lComando1 As Long, lComando2 As Long, ByVal sTabela As String, ByVal sCampoQtd As String, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal dtData As Date, colDistribuicao As Collection) As Long

Dim lErro As Long, iIndice As Integer, sMacroQtd As String
Dim iSeq As Integer, objEstTerc As ClassEstoqueTerc
Dim iMes As Integer, iAno As Integer, dtDataIniMes As Date
Dim lForn As Long, lCli As Long, iFilForn As Integer, iFilCli As Integer, dSaldo As Double

On Error GoTo Erro_EstoqueTerc_Le_Distrib
   
    iAno = Year(dtData)
    iMes = Month(dtData)
    dtDataIniMes = DateAdd("d", 1 - Day(dtData), dtData)
    
    sMacroQtd = "(QuantInicial" & sCampoQtd
    For iIndice = 1 To iMes - 1
        'Soma o saldo dos meses mais a qtd inicial até o mês anterior a data
        sMacroQtd = sMacroQtd & " + SaldoQuant" & sCampoQtd & CStr(iIndice)
    Next
    sMacroQtd = sMacroQtd & ")"

    lErro = Comando_Executar(lComando1, "SELECT Fornecedor, FilialForn, Cliente, FilialCli, " & sMacroQtd & " FROM SldMesEst" & sTabela & "Terc WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ?", lForn, iFilForn, lCli, iFilCli, dSaldo, iFilialEmpresa, iAno, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 209534

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209535

    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Se não foi identificado nem cliente nem fornecedor nem exibe,
        'só tem em Sld para análise
        If lForn <> 0 Or lCli <> 0 Then
        
            Set objEstTerc = New ClassEstoqueTerc
            
            If lForn > 0 Then
                objEstTerc.iTipoCliForn = TIPO_TERC_FORNECEDOR
                objEstTerc.lCliForn = lForn
                objEstTerc.iFilial = iFilForn
            Else
                objEstTerc.iTipoCliForn = TIPO_TERC_CLIENTE
                objEstTerc.lCliForn = lCli
                objEstTerc.iFilial = iFilCli
            End If
            objEstTerc.dQuantidade = dSaldo
            
            lErro = Comando_Executar(lComando2, "SELECT SUM(QuantEnt" & sCampoQtd & " - QuantSai" & sCampoQtd & ") FROM SldDiaEstTerc WHERE Produto = ? AND FilialEmpresa = ? AND Data BETWEEN ? AND ? AND Fornecedor = ? AND FilialForn = ? AND Cliente = ? AND FilialCli = ?", dSaldo, sProduto, iFilialEmpresa, dtDataIniMes, dtData, lForn, iFilForn, lCli, iFilCli)
            If lErro <> AD_SQL_SUCESSO Then gError 209536
        
            'Busca Primeiro
            lErro = Comando_BuscarPrimeiro(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209537
            
            objEstTerc.dQuantidade = objEstTerc.dQuantidade + dSaldo
            
            If objEstTerc.dQuantidade <> 0 Then colDistribuicao.Add objEstTerc
            
        End If
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209538
        
    Loop

    EstoqueTerc_Le_Distrib = SUCESSO

    Exit Function

Erro_EstoqueTerc_Le_Distrib:

    EstoqueTerc_Le_Distrib = gErr

    Select Case gErr

        Case 209534, 209535, 209538
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr)

        Case 209536, 209537
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209539)

    End Select
    
    Exit Function

End Function

Public Function EstoqueTerc_Le_Saldo(ByVal objEstTerc As ClassEstoqueTerc, Optional ByVal lCmd_P1 As Long = 0, Optional ByVal lCmd_P2 As Long = 0) As Long

Dim lErro As Long, iIndice As Integer, sMacroQtd As String
Dim alComando(0 To 1) As Long
Dim iSeq As Integer, iMes As Integer, iAno As Integer, dtDataIniMes As Date
Dim sCampoQtd As String, sTabela As String
Dim lForn As Long, lCli As Long, iFilForn As Integer, iFilCli As Integer
Dim dSaldo1 As Double, dSaldo2 As Double

On Error GoTo Erro_EstoqueTerc_Le_Saldo

    'Executa a abertura do Comando
    If lCmd_P1 = 0 Then
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 209625
        Next
    Else
        alComando(0) = lCmd_P1
        alComando(1) = lCmd_P2
    End If
    
    iAno = Year(objEstTerc.dtData)
    iMes = Month(objEstTerc.dtData)
    dtDataIniMes = DateAdd("d", 1 - Day(objEstTerc.dtData), objEstTerc.dtData)
    
    lErro = EstoqueTerc_Obtem_Dados_Escaminho(objEstTerc.iEscaninho, sTabela, sCampoQtd)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = EstoqueTerc_Obtem_Dados_Cliente(objEstTerc, lCli, iFilCli, lForn, iFilForn)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    sMacroQtd = "(QuantInicial" & sCampoQtd
    For iIndice = 1 To iMes - 1
        'Soma o saldo dos meses mais a qtd inicial até o mês anterior a data
        sMacroQtd = sMacroQtd & " + SaldoQuant" & sCampoQtd & CStr(iIndice)
    Next
    sMacroQtd = sMacroQtd & ")"

    lErro = Comando_Executar(alComando(0), "SELECT " & sMacroQtd & " FROM SldMesEst" & sTabela & "Terc WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ? AND Fornecedor = ? AND FilialForn = ? AND Cliente = ? AND FilialCli = ?", dSaldo1, objEstTerc.iFilialEmpresa, iAno, objEstTerc.sProduto, lForn, iFilForn, lCli, iFilCli)
    If lErro <> AD_SQL_SUCESSO Then gError 209626

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209627
    
    If lErro = AD_SQL_SUCESSO Then
        
        lErro = Comando_Executar(alComando(1), "SELECT SUM(QuantEnt" & sCampoQtd & " - QuantSai" & sCampoQtd & ") FROM SldDiaEstTerc WHERE Produto = ? AND FilialEmpresa = ? AND Data BETWEEN ? AND ? AND Fornecedor = ? AND FilialForn = ? AND Cliente = ? AND FilialCli = ?", dSaldo2, objEstTerc.sProduto, objEstTerc.iFilialEmpresa, dtDataIniMes, objEstTerc.dtData, lForn, iFilForn, lCli, iFilCli)
        If lErro <> AD_SQL_SUCESSO Then gError 209628
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209629
        
    End If
    
    objEstTerc.dQuantidade = dSaldo1 + dSaldo2
    
    'Fecha Comando
    If lCmd_P1 = 0 Then
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    End If
    
    EstoqueTerc_Le_Saldo = SUCESSO

    Exit Function

Erro_EstoqueTerc_Le_Saldo:

    EstoqueTerc_Le_Saldo = gErr

    Select Case gErr

        Case 209625
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209626, 209627
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr)

        Case 209628, 209629
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209630)

    End Select
    
    'Fecha Comando
    If lCmd_P1 = 0 Then
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    End If
    
    Exit Function

End Function

Private Function EstoqueTerc_Obtem_Dados_Escaminho(ByVal iEscaninho As Integer, sTabela As String, sCampoQtd As String) As Long

On Error GoTo Erro_EstoqueTerc_Obtem_Dados_Escaminho

    Select Case iEscaninho
    
        Case 2 '2   Conserto - Nosso em Poder de Terceiros
            sTabela = "2"
            sCampoQtd = "Conserto"
    
        Case 3 '3   Consignação - Nosso em Poder de Terceiros
            sTabela = "2"
            sCampoQtd = "Consig"
    
        Case 4 '4   Demonstração - Nosso em Poder de Terceiros
            sTabela = "2"
            sCampoQtd = "Demo"
    
        Case 5 '5   Outros - Nosso em Poder de Terceiros
            sTabela = "2"
            sCampoQtd = "Outros"
    
        Case 6 '6   Beneficiamento - Nosso em Poder de Terceiros
            sTabela = "2"
            sCampoQtd = "Benef"
        
        Case 7 '7   Conserto - De Terceiros em Nosso Poder
            sTabela = "1"
            sCampoQtd = "Conserto3"
    
        Case 8 '8   Consignação - De Terceiros em Nosso Poder
            sTabela = "1"
            sCampoQtd = "Consig3"
    
        Case 9 '9   Demonstração - De Terceiros em Nosso Poder
            sTabela = "1"
            sCampoQtd = "Demo3"
    
        Case 10 '10  Outros - De Terceiros em Nosso Poder
            sTabela = "1"
            sCampoQtd = "Outros3"
    
        Case 11 '11  Beneficiamento - De Terceiros em Nosso Poder
            sTabela = "1"
            sCampoQtd = "Benef3"
    
    End Select
    
    EstoqueTerc_Obtem_Dados_Escaminho = SUCESSO

    Exit Function

Erro_EstoqueTerc_Obtem_Dados_Escaminho:

    EstoqueTerc_Obtem_Dados_Escaminho = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209623)

    End Select
    
    Exit Function

End Function

Private Function EstoqueTerc_Obtem_Dados_Cliente(ByVal objEstTerc As ClassEstoqueTerc, lCli As Long, iFilCli As Integer, lForn As Long, iFilForn As Integer) As Long

On Error GoTo Erro_EstoqueTerc_Obtem_Dados_Cliente

    If objEstTerc.iTipoCliForn = TIPO_TERC_CLIENTE Then
        lCli = objEstTerc.lCliForn
        iFilCli = objEstTerc.iFilial
        lForn = 0
        iFilForn = 0
    Else
        lCli = 0
        iFilCli = 0
        lForn = objEstTerc.lCliForn
        iFilForn = objEstTerc.iFilial
    End If

    EstoqueTerc_Obtem_Dados_Cliente = SUCESSO

    Exit Function

Erro_EstoqueTerc_Obtem_Dados_Cliente:

    EstoqueTerc_Obtem_Dados_Cliente = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209624)

    End Select
    
    Exit Function
    
End Function

Public Function Importa_Xml_NFe_Loc_Prod(ByVal iFilialEmpresa As Integer, ByVal lForn As Long, ByVal iFilForn As Integer, ByVal sProdXml As String, ByVal sEANXML As String, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 6) As Long
Dim sProduto As String, sProdutoFormatado As String, iProdutoPreenchido As Integer

On Error GoTo Erro_Importa_Xml_NFe_Loc_Prod

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209857
    Next
    
    lErro = CF("Produto_Formata", sProdXml, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Tentativa 1 de localizar o produto
    sProduto = String(STRING_MAXIMO, 0)
    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos WHERE Codigo = ?", sProduto, sProdutoFormatado)
    If lErro <> AD_SQL_SUCESSO Then gError 209858
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209859

    If lErro <> AD_SQL_SUCESSO Then
        'Tentativa 2 de localizar o produto
        sProduto = String(STRING_MAXIMO, 0)
        lErro = Comando_Executar(alComando(2), "SELECT Produto FROM FornecedorProdutoFF WHERE FilialEmpresa = ? AND Fornecedor = ? AND FilialForn = ? AND ProdutoFornecedor = ?", sProduto, iFilialEmpresa, lForn, iFilForn, sProdXml)
        If lErro <> AD_SQL_SUCESSO Then gError 209860
        
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209861

        If lErro <> AD_SQL_SUCESSO Then
        
            If Len(Trim(sEANXML)) > 0 Then
        
                 'Tentativa 3 de localizar o produto
                sProduto = String(STRING_MAXIMO, 0)
                lErro = Comando_Executar(alComando(6), "SELECT Codigo FROM Produtos WHERE CodigoBarras = ?", sProduto, sEANXML)
                If lErro <> AD_SQL_SUCESSO Then gError 209862
                
                lErro = Comando_BuscarPrimeiro(alComando(6))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209863
                
            End If
            
            If lErro <> AD_SQL_SUCESSO Then
                
                'Tentativa 4 de localizar o produto
                sProduto = String(STRING_MAXIMO, 0)
                sProdutoFormatado = Replace(Replace(Replace(Replace(Replace(Replace(sProdXml, " ", ""), ".", ""), "_", ""), "-", ""), "/", ""), "\", "")
                lErro = Comando_Executar(alComando(3), "SELECT Codigo FROM Produtos WHERE REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(Codigo, ' ', ''), '.', ''), '_', ''), '-', ''), '/', ''), '\', '') = ?", sProduto, sProdutoFormatado)
                If lErro <> AD_SQL_SUCESSO Then gError 209862
                
                lErro = Comando_BuscarPrimeiro(alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209863
                
                If lErro <> AD_SQL_SUCESSO Then
                    'Tentativa 5 de localizar o produto
                    sProduto = String(STRING_MAXIMO, 0)
                    lErro = Comando_Executar(alComando(4), "SELECT Codigo FROM Produtos WHERE NomeReduzido = ?", sProduto, sProdXml)
                    If lErro <> AD_SQL_SUCESSO Then gError 209890
                    
                    lErro = Comando_BuscarPrimeiro(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209891
                    
                    If lErro <> AD_SQL_SUCESSO Then
                        'Tentativa 6 de localizar o produto
                        sProduto = String(STRING_MAXIMO, 0)
                        lErro = Comando_Executar(alComando(5), "SELECT Codigo FROM Produtos WHERE Referencia = ?", sProduto, sProdXml)
                        If lErro <> AD_SQL_SUCESSO Then gError 209892
                        
                        lErro = Comando_BuscarPrimeiro(alComando(5))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209893
                        
                        If lErro <> AD_SQL_SUCESSO Then gError 209864
                    End If
                End If
            End If
        End If
    End If
    
    objProduto.sCodigo = sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                   
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Importa_Xml_NFe_Loc_Prod = SUCESSO

    Exit Function

Erro_Importa_Xml_NFe_Loc_Prod:

    Importa_Xml_NFe_Loc_Prod = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 209857
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209858, 209859, 209862, 209863, 209890 To 209893
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case 209860, 209861
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)

        Case 209864
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPOTNFE_PRODUTO_NAO_CADASTRADO", gErr, sProdXml, sEANXML)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209865)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Public Function NFiscal_ImpParc_Le_Mae(ByVal objNF As ClassNFiscal, objNFOrig As ClassNFiscal) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 1) As Long
Dim lNumIntDoc As Long, lNumNotaFiscal As Long, sSerie As String

On Error GoTo Erro_NFiscal_ImpParc_Le_Mae

    'Se for uma importação remetida em parcelas pega a NF mãe (primeira) como original
    If Not (objNF.objNFImportacao Is Nothing) Then
    
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 211200
        Next
          
        sSerie = String(STRING_SERIE, 0)
        lErro = Comando_Executar(alComando(1), "SELECT N.NumIntDoc, N.NumNotaFiscal, N.Serie FROM NFImportacao AS I, NFiscal AS N WHERE I.NumIntNf = N.NumIntDoc AND N.TipoNfiscal = ? AND Status <> ? AND I.NumIntDI = ?", lNumIntDoc, lNumNotaFiscal, sSerie, DOCINFO_NFIEIMPSE, STATUS_CANCELADO, objNF.objNFImportacao.objDIInfo.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 211201
          
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211202
        
        If lErro <> AD_SQL_SUCESSO Then gError 211203
        
        objNFOrig.lNumIntDoc = lNumIntDoc
        objNFOrig.lNumNotaFiscal = lNumNotaFiscal
        objNFOrig.sSerie = sSerie
        
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If

    NFiscal_ImpParc_Le_Mae = SUCESSO

    Exit Function

Erro_NFiscal_ImpParc_Le_Mae:

    NFiscal_ImpParc_Le_Mae = gErr

    Select Case gErr

        Case 211200
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211201, 211202
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
            
        Case 211203
            Call Rotina_Erro(vbOKOnly, "ERRO_NFIMPPARC_SEM_MAE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211204)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Public Function NFe_Chv_Valida(ByVal sChvNFe As String, bExisteNF As Boolean, bExisteXml As Boolean, bTrazer As Boolean, lNumIntNF As Long) As Long

Dim lErro As Long
Dim objNF As New ClassNFiscal
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_NFe_Chv_Valida

    bExisteNF = False
    bExisteXml = False
    bTrazer = False
    lNumIntNF = 0

    If Len(Trim(sChvNFe)) > 0 Then
    
        If Len(Trim(sChvNFe)) <> STRING_NFE_CHNFE Then gError 211266
        
        objNF.sChvNFe = sChvNFe

        lErro = CF("NFiscal_Le_Pela_ChvNFe", objNF)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then
            lNumIntNF = objNF.lNumIntDoc
            bExisteNF = True
            
            vbResult = Rotina_Aviso(vbYesNo, "AVISO_JA_EXISTE_NFE_COM_ESSA_CHVNFE")
            If vbResult = vbYes Then bTrazer = True
            
        Else
        
            lErro = CF("ImportNFeXml_Le_Pela_ChvNFe", objNF)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
            If lErro = SUCESSO Then
                bExisteXml = True
                
                vbResult = Rotina_Aviso(vbYesNo, "AVISO_JA_EXISTE_XML_COM_ESSA_CHVNFE")
                If vbResult = vbYes Then bTrazer = True
            
            End If
        
        End If
        
    End If
    
    NFe_Chv_Valida = SUCESSO

    Exit Function

Erro_NFe_Chv_Valida:

    NFe_Chv_Valida = gErr

    Select Case gErr
    
        Case 211266
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CHV_TAM_INVALIDO", gErr)

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211267)

    End Select

    Exit Function
    
End Function

Public Function NFe_Chv_Monta_NF(ByVal objTela As Object, ByVal objNF As ClassNFiscal, ByVal sChvNFe As String, Optional gcolPedidoCompra As Collection) As Long
'Gera o objNF com os dados de ImportXmlNFe a partir da chave passada

Dim lErro As Long, objItemNF As ClassItemNF

On Error GoTo Erro_NFe_Chv_Monta_NF

    If gobjCRFAT.iNFImportXml <> 0 Then
        
        objNF.sChvNFe = sChvNFe
    
        '====> TEM QUE MONTAR O objNF COM OS DADOS IMPORTADOS PARA CHAVE
        lErro = NFe_Chv_Monta_NF1(objTela, objNF, sChvNFe, giFilialEmpresa)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objNF.colItens.Count <> 0 Then
    
            Call Chama_Tela_Modal("NFImportConf", objNF, gcolPedidoCompra)
            If giRetornoTela <> vbOK Then gError ERRO_SEM_MENSAGEM
        
            Call objTela.ValorTotal_Calcula
            
        End If
    
    End If
    
    NFe_Chv_Monta_NF = SUCESSO

    Exit Function

Erro_NFe_Chv_Monta_NF:

    NFe_Chv_Monta_NF = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211668)

    End Select

    Exit Function
    
End Function

Public Function NFiscal_ImpCC_Le_Orig(ByVal objNF As ClassNFiscal, objNFOrig As ClassNFiscal) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 1) As Long
Dim lNumIntDoc As Long, lNumNotaFiscal As Long, sSerie As String
Dim lFornecedor As Long, iFilialForn As Integer

On Error GoTo Erro_NFiscal_ImpCC_Le_Orig

    'Se for uma importação remetida em parcelas pega a NF mãe (primeira) como original
    If Not (objNF.objNFImportacao Is Nothing) Then
    
        'Abertura de Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 211326
        Next
          
        sSerie = String(STRING_SERIE, 0)
        lErro = Comando_Executar(alComando(1), "SELECT N.NumIntDoc, N.NumNotaFiscal, N.Serie, N.Fornecedor, N.FilialForn FROM NFImportacao AS I, NFiscal AS N WHERE I.NumIntNf = N.NumIntDoc AND N.TipoNfiscal IN (?,?) AND Status <> ? AND I.NumIntDI = ?", lNumIntDoc, lNumNotaFiscal, sSerie, lFornecedor, iFilialForn, DOCINFO_NFIEIMP, DOCINFO_NFIEIMPSE, STATUS_CANCELADO, objNF.objNFImportacao.objDIInfo.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 211327
          
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211328
        
        If lErro <> AD_SQL_SUCESSO Then gError 211329
        
        objNFOrig.lNumIntDoc = lNumIntDoc
        objNFOrig.lNumNotaFiscal = lNumNotaFiscal
        objNFOrig.sSerie = sSerie
        objNFOrig.lFornecedor = lFornecedor
        objNFOrig.iFilialForn = iFilialForn
        
        'Fecha Comando
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If

    NFiscal_ImpCC_Le_Orig = SUCESSO

    Exit Function

Erro_NFiscal_ImpCC_Le_Orig:

    NFiscal_ImpCC_Le_Orig = gErr

    Select Case gErr

        Case 211326
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211327, 211328
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
            
        Case 211329
            Call Rotina_Erro(vbOKOnly, "ERRO_NFIMPCC_SEM_NFORIG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211330)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Public Function DIInfo_Le_ItensPCDI(ByVal objDIInfo As ClassDIInfo) As Long

Dim lErro As Long, tItem As typeItemPCDI, objItemPCDI As ClassItemPCDI
Dim iIndice As Integer
Dim alComando(1 To 1) As Long

On Error GoTo Erro_DIInfo_Le_ItensPCDI


    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 210609
    Next

    With tItem
        .sProdutoPC = String(STRING_PRODUTO, 0)
        .sUMPC = String(STRING_UM_SIGLA, 0)
        .sDescProdPC = String(STRING_DESCRICAO_ITEM, 0)
        
        lErro = Comando_Executar(alComando(1), "SELECT NumIntDI, Seq, FilialEmpresa, CodigoPC, DataPC, ProdutoPC, DescProdutoPC, UMPC, QuantPC FROM ItensPCDI WHERE NumIntDI = ? ORDER BY Seq", _
            .lNumIntDI, .iSeq, .iFilialEmpresa, .lCodigoPC, .dtDataPC, .sProdutoPC, .sDescProdPC, .sUMPC, .dQuantPC, objDIInfo.lNumIntDoc)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 210610
    
    'Busca Primeiro
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210611
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objItemPCDI = New ClassItemPCDI
    
        With objItemPCDI
            .lNumIntDI = tItem.lNumIntDI
            .iSeq = tItem.iSeq
            .iFilialEmpresa = tItem.iFilialEmpresa
            .lCodigoPC = tItem.lCodigoPC
            .dtDataPC = tItem.dtDataPC
            .sProdutoPC = tItem.sProdutoPC
            .sDescProdPC = tItem.sDescProdPC
            .sUMPC = tItem.sUMPC
            .dQuantPC = tItem.dQuantPC
        End With
        
        objDIInfo.colItensPC.Add objItemPCDI
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210612
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    DIInfo_Le_ItensPCDI = SUCESSO
    
    Exit Function
    
Erro_DIInfo_Le_ItensPCDI:

    DIInfo_Le_ItensPCDI = gErr

    Select Case gErr

        Case 210609
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 210610 To 210612
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPCDI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210613)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function ItensPCDI_Le_QuantPC(ByVal sNumeroDI As String, ByVal objItemPCDI As ClassItemPCDI, dQuantPC As Double) As Long
'le a quantidade do produto de ItensPC que esta em DI mas que ainda nao esta associada a nota fiscal

Dim lErro As Long, tItem As typeItemPCDI
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim lNumIntDI As Long

On Error GoTo Erro_ItensPCDI_Le_QuantPC

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 210617
    Next
    
    'Le a tabelaDIInfo
    lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc FROM DIInfo WHERE Numero= ? ", _
                lNumIntDI, sNumeroDI)
    If lErro <> AD_SQL_SUCESSO Then gError 210625

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210626
    
    If lErro = AD_SQL_SUCESSO Then

        'Le a tabela ItensPCDI para pegar todas as refrencias ao codigo pc e ao produto que ja estejam em DI que ainda nao estejam baixadas.
        lErro = Comando_Executar(alComando(1), "SELECT SUM(QuantPC) FROM ItensPCDI WHERE CodigoPC = ? AND ProdutoPC = ? AND NumIntDI <> ? AND NOT EXISTS (SELECT NumINtDI FROM NFImportacao WHERE NFImportacao.NumIntDI= ItensPCDI.NumIntDI)", _
                    dQuantPC, objItemPCDI.lCodigoPC, objItemPCDI.sProdutoPC, lNumIntDI)
        If lErro <> AD_SQL_SUCESSO Then gError 210615

    Else
    
        'Le a tabela ItensPCDI para pegar todas as refrencias ao codigo pc e ao produto que ja estejam em DI que ainda nao estejam baixadas.
        lErro = Comando_Executar(alComando(1), "SELECT SUM(QuantPC) FROM ItensPCDI WHERE CodigoPC = ? AND ProdutoPC = ? AND NOT EXISTS (SELECT NumINtDI FROM NFImportacao WHERE NFImportacao.NumIntDI= ItensPCDI.NumIntDI)", _
                    dQuantPC, objItemPCDI.lCodigoPC, objItemPCDI.sProdutoPC)
        If lErro <> AD_SQL_SUCESSO Then gError 210627
    
    End If

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210616


    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    ItensPCDI_Le_QuantPC = SUCESSO
    
    Exit Function
    
Erro_ItensPCDI_Le_QuantPC:

    ItensPCDI_Le_QuantPC = gErr

    Select Case gErr

        Case 210617
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 210615 To 210616, 210627
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPCDI", gErr)
        
        Case 210625, 210626
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIINFO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210618)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function NFe_Chv_Monta_NF1(ByVal objTela As Object, ByVal objNF As ClassNFiscal, ByVal sChvNFe As String, Optional ByVal iFilialEmpresa As Integer = 0) As Long
'Gera o objNF com os dados de ImportXmlNFe a partir da chave passada

Dim lErro As Long, sSQL As String, sSQL2 As String
Dim alComando(0 To 35) As Long, iIndice As Integer

Dim lCodCliente As Long, iCodFilial As Integer, lCodFornecedor As Long
Dim objItemNF As ClassItemNF, objTribTab As ClassTribTab
Dim objTipoDocInfo As ClassTipoDocInfo
Dim objFilialEmpresa As New AdmFiliais, sNFs As String
Dim objProduto As ClassProduto
Dim lNumNF As Long, sSerie As String, dtDataEmissao As Date, sCGCDestinatario As String, sNomeDestinatario As String, dNFICMSBase As Double, dNFICMSValor As Double, dNFICMSSubstBase As Double, dNFICMSSubstValor As Double, dValorProdutos As Double, dValorFrete As Double, dValorSeguro As Double, dValorDescontos As Double, dNFIPIValor As Double, dNFPISValor As Double, dNFCofinsValor As Double, dValorOutrasDespesas As Double
Dim dValorTotalNF As Double, sMsg(1 To 4) As String, sTipoES As String, sTipoCN As String, lNumIntNF As Long, sCGCTransportadora As String, dPesoLiq As Double, dPesoBruto As Double, sEspecie As String, lVolumeQuant As Long, sNomeEmitente As String, sNomeTransportadora As String
Dim lItem As Long, sProduto As String, sDescProd As String, sCFOP As String, sUMXml As String, dQuantidade As Double, dPrecoUnitario As Double, dPrecoTotal As Double, sEAN As String, sInfAdProd As String, dDesconto As Double, dFrete As Double, dSeguro As Double, dDespesas As Double, sISSCST As String, dISSBase As Double, dISSAliquota As Double, dISSValor As Double, iISSListServ As Integer, sISSCidadeIBGE As String, iOrigemMercadoria As Integer
Dim sICMSCST As String, sICMSCSOSN As String, iICMSBaseModalidade As Integer, dICMSBase As Double, dICMSPercRedBase As Double, dICMSAliquota As Double, dICMSValor As Double, iICMSSubstBaseModalidade As Integer, dICMSSubstBase As Double, dICMSSubstPercMVA As Double, dICMSSubstPercRedBase As Double, dICMSSubstAliquota As Double, dICMSSubstValor As Double
Dim dICMSSTCobrAntBase As Double, dICMSSTCobrAntValor As Double, dICMSPercBaseOperacaoPropria As Double, sICMSUFDevidoST As String, dICMSvBCSTRet As Double, dICMSvICMSSTRet As Double, dICMSvBCSTDest As Double, dICMSvICMSSTDest As Double, dICMSpCredSN As Double, dICMSvCredSN As Double, sIPICST As String, sIPIEnquadramentoClasse As String
Dim sIPIEnquadramentoCodigo As String, sIPISeloCodigo As String, sIPICNPJProdutor As String, lIPISeloQtde As Long, dIPIBase As Double, dIPIAliquota As Double, dIPIValor As Double, dIPIUnidadePadraoQtde As Double, dIPIUnidadePadraoValor As Double, dIIValor As Double, dIIBase As Double, dIIDespAduaneira As Double, dIIIOF As Double, sPISCST As String, iPISTipoCalculo As Integer
Dim dPISBase As Double, dPISAliquota As Double, dPISValor As Double, dPISQtde As Double, dPISAliquotaValor As Double, iPISSTTipoCalculo As Integer, dPISSTBase As Double, dPISSTAliquota As Double, dPISSTValor As Double, dPISSTQtde As Double, dPISSTAliquotaValor As Double, sCOFINSCST As String, iCOFINSTipoCalculo As Integer, dCOFINSBase As Double, dCOFINSAliquota As Double
Dim dCOFINSValor As Double, dCOFINSQtde As Double, dCOFINSAliquotaValor As Double, iCOFINSSTTipoCalculo As Integer, dCOFINSSTBase As Double, dCOFINSSTAliquota As Double, dCOFINSSTValor As Double, dCOFINSSTQtde As Double, dCOFINSSTAliquotaValor As Double
Dim dtDataParc As Date, dValorParc As Double
Dim objParcela As ClassParcelaPagar, iTipoNFiscal As Integer
Dim sProdAux As String, sDescProdAux As String, sUMAux As String, sOrigemMercadoriaAux As String, sNomeRedAux As String, lContAux As Long, sNCMAux As String
Dim colTabelaPrecoItem As New Collection
Dim iICMSTipoSimples As Integer, iQtdNFsImportadas As Integer
Dim sProdEstProd As String, colRastreamento As New Collection, dFator As Double
Dim objEstoqueProduto As ClassEstoqueProduto, objItemAloc As ClassItemNFAlocacao
Dim objCliente As ClassCliente, sCFOPPrincipal As String
Dim colEndereco As Collection, lProxCodCliente As Long
Dim objEndereco As New ClassEndereco, sMaxCFOP As String, sCGCAnt As String
Dim dtDataES As Date, sHoraES As String, dISSVlr As Double, dISSBC As Double, dPISVlr_ISS As Double, dCofinsVlr_ISS As Double, dVlrServico As Double, dPISRet As Double, dCofinsRet As Double, dCSLLRet As Double, dIRRFBC As Double, dIRRFVlr As Double, dINSSBCRet As Double, dINSSRet As Double, iExisteEndRetirada As Integer, iExisteEndEntrega As Integer, iExisteTransp As Integer, iExisteISSQNtot As Integer, iExisteRetTrib As Integer, iExisteCobr As Integer
Dim sNome As String, sNomeRed As String, sIE As String, sISUF As String, sIM As String, iCRT As Integer, sCodPais As String, sPais As String, sUF As String, sCodMunicipio As String, sMunicipio As String, sLogradouro As String, sNumero As String, sBairro As String, sComplemento As String, sCEP As String, stelefone As String, sEmail As String
Dim sNCM As String, iIndTot As Integer, sEXTIPI As String, sEANTrib As String, dQtdTrib As Double, sUMTrib As String, dValorUnitTrib As Double
Dim sMunicipioAux As String, sCodMunicipioAux As String, lProxCodCidade As Long, sCGCEmitente As String
Dim bMesmaEmpresa As Boolean, iAux As Integer, sCFOPInterno As String, sProdXml As String, iTipoTributacao As Integer
Dim dNFIPIBase As Double, iProdutoPreenchido As Integer, bDestinatario As Boolean, sModelo As String
Dim dTotTrib As Double, dTotTribItem As Double, dBaseICMSAntesReducao As Double, sCGCFilialEmp As String, lNumIntNFOrig As Long, sUFEmitente As String
Dim dICMSInterestVlrFCPUFDestNF As Double, dICMSInterestVlrUFDestNF As Double, dICMSInterestVlrUFRemetNF As Double, dvIPIDevolNF As Double, dvFCPNF As Double, dvFCPSTNF As Double, dvFCPSTRetNF As Double
Dim dICMSInterestBCUFDest As Double, dICMSInterestPercFCPUFDest As Double, dICMSInterestAliqUFDest As Double, dICMSInterestPercPartilha As Double, dICMSInterestVlrFCPUFDest As Double, dICMSInterestVlrUFDest As Double, dICMSInterestVlrUFRemet As Double, dICMSInterestAliq As Double, iExisteICMSUFDest As Integer
Dim sindEscala As String, sCNPJFab As String, scBenef As String, dvBCFCP As Double, dpFCP As Double, dvFCP As Double, dvBCFCPST As Double, dpFCPST As Double, dvFCPST As Double
Dim dvBCFCPSTRet As Double, dpFCPSTRet As Double, dvFCPSTRet As Double, dICMSSTCobrAntAliquota As Double, dvBCFCPUFDest As Double

On Error GoTo Erro_NFe_Chv_Monta_NF1

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213475
    Next
    
    If iFilialEmpresa = 0 Then iFilialEmpresa = giFilialEmpresa
    
    objFilialEmpresa.iCodFilial = iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sSQL = "SELECT Modelo, DataES, HoraES, ISSVlr, ISSBC, PISVlr_ISS, CofinsVlr_ISS, VlrServico, PISRet, CofinsRet, CSLLRet, IRRFBC, IRRFVlr, INSSBCRet, INSSRet, ExisteEndRetirada, ExisteEndEntrega, ExisteTransp, ExisteISSQNtot, ExisteRetTrib, ExisteCobr, NumNF, Serie, DataEmissao, CGCEmitente, CGCDestinatario, NomeDestinatario, NF.ICMSBase, NF.ICMSValor, NF.ICMSSubstBase, NF.ICMSSubstValor, ValorProdutos, ValorFrete, ValorSeguro, ValorDescontos, NF.IPIValor, NF.PISValor, NF.CofinsValor, ValorOutrasDespesas, "
    sSQL = sSQL & "ValorTotalNF, SUBSTRING(Mensagem,1,250), SUBSTRING(Mensagem,251,250), SUBSTRING(Mensagem,501,250), SUBSTRING(Mensagem,751,250), TipoES, TipoCN, CGCTransportadora, PesoLiq, PesoBruto, Especie, VolumeQuant, NomeEmitente, NomeTransportadora, TotTrib, ICMSInterestVlrFCPUFDest,ICMSInterestVlrUFDest,ICMSInterestVlrUFRemet,vIPIDevol,vFCP,vFCPST,vFCPSTRet "
    sSQL = sSQL & " FROM ImportNFEXML NF WHERE FilialEmpresa = ? AND ChvNFe = ? AND CGCEmitente <> ? AND NumIntNF =0 AND ChvNFe NOT IN (SELECT ChvNFe FROM NFiscal WHERE FilialEmpresa = ?)"
    
    sSQL2 = "SELECT NCM, IndTot, EXTIPI, EANTrib, QtdTrib, UMTrib, ValorUnitTrib, Item, Produto, DescProd, CFOP, UM, Quantidade, PrecoUnitario, PrecoTotal, EAN, InfAdProd, Desconto, Frete, Seguro, Despesas, ISSCST, ISSBase, ISSAliquota, ISSValor, ISSListServ, ISSCidadeIBGE, OrigemMercadoria, "
    sSQL2 = sSQL2 & "ICMSCST, ICMSCSOSN, ICMSBaseModalidade, ICMSBase, ICMSPercRedBase, ICMSAliquota, ICMSValor, ICMSSubstBaseModalidade, ICMSSubstBase, ICMSSubstPercMVA, ICMSSubstPercRedBase, ICMSSubstAliquota, ICMSSubstValor, "
    sSQL2 = sSQL2 & "ICMSSTCobrAntBase, ICMSSTCobrAntValor, ICMSPercBaseOperacaoPropria, ICMSUFDevidoST, ICMSvBCSTRet, ICMSvICMSSTRet, ICMSvBCSTDest, ICMSvICMSSTDest, ICMSpCredSN, ICMSvCredSN, IPICST, IPIEnquadramentoClasse, "
    sSQL2 = sSQL2 & "IPIEnquadramentoCodigo, IPISeloCodigo, IPICNPJProdutor, IPISeloQtde, IPIBase, IPIAliquota, IPIValor, IPIUnidadePadraoQtde, IPIUnidadePadraoValor, IIValor, IIBase, IIDespAduaneira, IIIOF, PISCST, PISTipoCalculo, "
    sSQL2 = sSQL2 & "PISBase, PISAliquota, PISValor, PISQtde, PISAliquotaValor, PISSTTipoCalculo, PISSTBase, PISSTAliquota, PISSTValor, PISSTQtde, PISSTAliquotaValor, COFINSCST, COFINSTipoCalculo, COFINSBase, COFINSAliquota, "
    sSQL2 = sSQL2 & "CofinsValor , COFINSQtde, COFINSAliquotaValor, COFINSSTTipoCalculo, COFINSSTBase, COFINSSTAliquota, COFINSSTValor, COFINSSTQtde, COFINSSTAliquotaValor, TotTrib, "
    sSQL2 = sSQL2 & "ICMSInterestBCUFDest , ICMSInterestPercFCPUFDest, ICMSInterestAliqUFDest, ICMSInterestPercPartilha, ICMSInterestVlrFCPUFDest, ICMSInterestVlrUFDest, ICMSInterestVlrUFRemet, ICMSInterestAliq, ExisteICMSUFDest, indEscala, CNPJFab, cBenef, vBCFCP, pFCP, vFCP, vBCFCPST, pFCPST, vFCPST, vBCFCPSTRet, pFCPSTRet, vFCPSTRet, ICMSSTCobrAntAliquota, vBCFCPUFDest "
    sSQL2 = sSQL2 & "FROM ImportNFEItensXML WHERE ChvNFe = ? ORDER BY Item"
    
    sHoraES = String(STRING_MAXIMO, 0)
    sSerie = String(STRING_MAXIMO, 0)
    sCGCDestinatario = String(STRING_MAXIMO, 0)
    sCGCEmitente = String(STRING_MAXIMO, 0)
    sNomeDestinatario = String(1000, 0)
    For iIndice = 1 To 4
        sMsg(iIndice) = String(STRING_MAXIMO, 0)
    Next
    sTipoES = String(STRING_MAXIMO, 0)
    sTipoCN = String(STRING_MAXIMO, 0)
    sCGCTransportadora = String(STRING_MAXIMO, 0)
    sEspecie = String(STRING_MAXIMO, 0)
    sNomeEmitente = String(1000, 0)
    sNomeTransportadora = String(1000, 0)
    sModelo = String(STRING_MAXIMO, 0)
    
    lErro = Comando_Executar(alComando(0), sSQL, _
        sModelo, dtDataES, sHoraES, dISSVlr, dISSBC, dPISVlr_ISS, dCofinsVlr_ISS, dVlrServico, dPISRet, dCofinsRet, dCSLLRet, dIRRFBC, dIRRFVlr, dINSSBCRet, dINSSRet, iExisteEndRetirada, iExisteEndEntrega, iExisteTransp, iExisteISSQNtot, iExisteRetTrib, iExisteCobr, _
        lNumNF, sSerie, dtDataEmissao, sCGCEmitente, sCGCDestinatario, sNomeDestinatario, dNFICMSBase, dNFICMSValor, dNFICMSSubstBase, dNFICMSSubstValor, dValorProdutos, dValorFrete, dValorSeguro, dValorDescontos, dNFIPIValor, dNFPISValor, dNFCofinsValor, dValorOutrasDespesas, _
        dValorTotalNF, sMsg(1), sMsg(2), sMsg(3), sMsg(4), sTipoES, sTipoCN, sCGCTransportadora, dPesoLiq, dPesoBruto, sEspecie, lVolumeQuant, sNomeEmitente, sNomeTransportadora, dTotTrib, dICMSInterestVlrFCPUFDestNF, dICMSInterestVlrUFDestNF, dICMSInterestVlrUFRemetNF, dvIPIDevolNF, dvFCPNF, dvFCPSTNF, dvFCPSTRetNF, _
        iFilialEmpresa, sChvNFe, Trim(objFilialEmpresa.sCgc), iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 213476
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213477
    
    If lErro <> AD_SQL_SUCESSO Then gError 213478
        
    'Inicializa a parte tributária
    Call objNF.Inicializa_Tributacao
    Call objNF.objTributacaoNF.Coloca_Auto

    sCFOPPrincipal = ""
    
    bMesmaEmpresa = False
    bDestinatario = False
    
    lErro = Comando_Executar(alComando(9), "SELECT FilialEmpresa FROM FiliaisEmpresa WHERE CGC = ?", iAux, sCGCEmitente)
    If lErro <> AD_SQL_SUCESSO Then gError 213479
    
    lErro = Comando_BuscarProximo(alComando(9))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213480

    If lErro = AD_SQL_SUCESSO Then bMesmaEmpresa = True
    
    sCGCFilialEmp = String(STRING_CGC, 0)
    
    lErro = Comando_Executar(alComando(31), "SELECT CGC FROM FiliaisEmpresa WHERE FilialEmpresa = ?", sCGCFilialEmp, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 213481
    
    lErro = Comando_BuscarProximo(alComando(31))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213482

    If sCGCFilialEmp = sCGCDestinatario Then bDestinatario = True

    objNF.sCGCEmitente = sCGCEmitente
    objNF.lNumNotaFiscal = lNumNF
    objNF.sSerie = sSerie & "-e"
    objNF.dtDataEmissao = dtDataEmissao
    objNF.iNFe = MARCADO
    For iIndice = 1 To 4
        objNF.sMensagemNota = objNF.sMensagemNota & sMsg(iIndice)
    Next
    objNF.iFilialEmpresa = iFilialEmpresa
    objNF.iNaoVerificaCredito = NAO_VERIFICA_CREDITO_CLIENTE
    objNF.iStatus = STATUS_BAIXADO
    objNF.dtDataReferencia = DATA_NULA
    objNF.iFreteRespons = FRETE_EMITENTE
    objNF.dtDataRegistro = Date
    objNF.dtDataVencimento = DATA_NULA
    
    objNF.dValorItens = dValorProdutos
    objNF.dValorProdutos = Arredonda_Moeda(dValorProdutos - dValorDescontos)
    objNF.dValorFrete = dValorFrete
    objNF.dValorSeguro = dValorSeguro
    objNF.dValorOutrasDespesas = dValorOutrasDespesas
    objNF.dValorDescontoItens = dValorDescontos
    objNF.dValorTotal = dValorTotalNF
            
    objNF.dtDataEntrada = IIf(dtDataES <> DATA_NULA, dtDataES, dtDataEmissao)
    objNF.dtHoraEntrada = StrParaDate(sHoraES)
    objNF.dtDataSaida = DATA_NULA
    objNF.dtHoraSaida = 0
    
    objNF.dPesoLiq = dPesoLiq
    objNF.dPesoBruto = dPesoBruto
    'objNF. = sEspecie ,
    objNF.lVolumeQuant = lVolumeQuant
    
    objNF.objTributacaoNF.iTipoTributacao = 0
    objNF.objTributacaoNF.dICMSBase = dNFICMSBase
    objNF.objTributacaoNF.iICMSBaseManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dICMSValor = dNFICMSValor + dvFCPNF
    objNF.objTributacaoNF.iICMSValorManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dICMSSubstBase = dNFICMSSubstBase
    objNF.objTributacaoNF.iICMSSubstBaseManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dICMSSubstValor = dNFICMSSubstValor + dvFCPSTNF
    objNF.objTributacaoNF.iICMSSubstValorManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dIPIValor = dNFIPIValor
    objNF.objTributacaoNF.iIPIValorManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dPISValor = dNFPISValor
    objNF.objTributacaoNF.iPISCreditoManual = VAR_PREENCH_AUTOMATICO
    objNF.objTributacaoNF.dCOFINSValor = dNFCofinsValor
    objNF.objTributacaoNF.iCOFINSCreditoManual = VAR_PREENCH_AUTOMATICO
    
    objNF.objTributacaoNF.dISSValor = dISSVlr
    objNF.objTributacaoNF.iISSValorManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dISSBase = dISSBC
    objNF.objTributacaoNF.iISSBaseManual = VAR_PREENCH_MANUAL
    'dPISVlr_ISS
    'dCofinsVlr_ISS
    'dVlrServico
    objNF.objTributacaoNF.dPISRetido = dPISRet
    objNF.objTributacaoNF.iPISRetidoManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dCOFINSRetido = dCofinsRet
    objNF.objTributacaoNF.iCOFINSRetidoManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dCSLLRetido = dCSLLRet
    objNF.objTributacaoNF.iCSLLRetidoManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dIRRFBase = dIRRFBC
    objNF.objTributacaoNF.iIRRFBaseManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dIRRFValor = dIRRFVlr
    objNF.objTributacaoNF.iIRRFValorManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dINSSBase = dINSSBCRet
    objNF.objTributacaoNF.iINSSBaseManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dINSSValor = dINSSRet
    objNF.objTributacaoNF.iINSSValorManual = VAR_PREENCH_MANUAL
    If dINSSRet <> 0 Then objNF.objTributacaoNF.iINSSRetido = 1
    objNF.objTributacaoNF.iINSSRetidoManual = VAR_PREENCH_MANUAL
    
    objNF.objTributacaoNF.iTotTribTipo = 0
    objNF.objTributacaoNF.iTotTribTipoManual = VAR_PREENCH_MANUAL
    objNF.objTributacaoNF.dTotTrib = dTotTrib
    objNF.objTributacaoNF.iTotTribManual = VAR_PREENCH_MANUAL
    
    lErro = Comando_Executar(alComando(33), "SELECT P.NumIntNF FROM ImportNFeDocRefXml AS O, NFeFedProtNFe AS P, NFiscal AS N WHERE P.ChNFe = O.ChvNFeOrig AND O.ChvNFe = ? AND P.ChNFe <> ? AND P.NumIntNF = N.NumIntDoc AND N.FilialEmpresa=?", lNumIntNFOrig, sChvNFe, "", iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 213483
    
    lErro = Comando_BuscarProximo(alComando(33))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213484
    
    If lErro = AD_SQL_SUCESSO Then
        objNF.lNumIntNotaOriginal = lNumIntNFOrig
    Else
    
        lErro = Comando_Executar(alComando(34), "SELECT N.NumIntDoc FROM ImportNFeDocRefXml AS O, NFiscal AS N WHERE N.ChvNFe = O.ChvNFeOrig AND O.ChvNFe = ? AND N.ChvNFe <> ? AND O.Tipo = ? AND N.FilialEmpresa = ?", lNumIntNFOrig, sChvNFe, "", "NFe", iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 213485
        
        lErro = Comando_BuscarProximo(alComando(34))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213486
        
        If lErro = AD_SQL_SUCESSO Then objNF.lNumIntNotaOriginal = lNumIntNFOrig
    
    End If
    
    sUFEmitente = String(STRING_ESTADO_SIGLA, 0)
    
    lErro = Comando_Executar(alComando(35), "SELECT UF FROM ImportNFeEndXml WHERE Tipo = 1 AND ChvNFe = ?", sUFEmitente, sChvNFe)
    If lErro <> AD_SQL_SUCESSO Then gError 213485
    
    lErro = Comando_BuscarProximo(alComando(35))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213486
    
    sProdXml = String(STRING_MAXIMO, 0)
    sDescProd = String(1000, 0)
    sCFOP = String(STRING_MAXIMO, 0)
    sUMXml = String(STRING_MAXIMO, 0)
    sEAN = String(STRING_MAXIMO, 0)
    sInfAdProd = String(STRING_INFOADIC_ITEM_MSG, 0)
    sISSCST = String(STRING_MAXIMO, 0)
    sISSCidadeIBGE = String(STRING_MAXIMO, 0)
    sICMSCST = String(STRING_MAXIMO, 0)
    sICMSCSOSN = String(STRING_MAXIMO, 0)
    sICMSUFDevidoST = String(STRING_MAXIMO, 0)
'    sICMSpCredSN = String(STRING_MAXIMO, 0)
    sIPICST = String(STRING_MAXIMO, 0)
    sIPIEnquadramentoClasse = String(STRING_MAXIMO, 0)
    sIPIEnquadramentoCodigo = String(STRING_MAXIMO, 0)
    sIPISeloCodigo = String(STRING_MAXIMO, 0)
    sIPICNPJProdutor = String(STRING_MAXIMO, 0)
    sPISCST = String(STRING_MAXIMO, 0)
    sCOFINSCST = String(STRING_MAXIMO, 0)
    sNCM = String(STRING_MAXIMO, 0)
    sEXTIPI = String(STRING_MAXIMO, 0)
    sEANTrib = String(STRING_MAXIMO, 0)
    sUMTrib = String(STRING_MAXIMO, 0)
    sindEscala = String(STRING_MAXIMO, 0)
    sCNPJFab = String(STRING_MAXIMO, 0)
    scBenef = String(STRING_MAXIMO, 0)
                    
    lErro = Comando_Executar(alComando(1), sSQL2, _
        sNCM, iIndTot, sEXTIPI, sEANTrib, dQtdTrib, sUMTrib, dValorUnitTrib, _
        lItem, sProdXml, sDescProd, sCFOP, sUMXml, dQuantidade, dPrecoUnitario, dPrecoTotal, sEAN, sInfAdProd, dDesconto, dFrete, dSeguro, dDespesas, sISSCST, dISSBase, dISSAliquota, dISSValor, iISSListServ, sISSCidadeIBGE, iOrigemMercadoria, _
        sICMSCST, sICMSCSOSN, iICMSBaseModalidade, dICMSBase, dICMSPercRedBase, dICMSAliquota, dICMSValor, iICMSSubstBaseModalidade, dICMSSubstBase, dICMSSubstPercMVA, dICMSSubstPercRedBase, dICMSSubstAliquota, dICMSSubstValor, _
        dICMSSTCobrAntBase, dICMSSTCobrAntValor, dICMSPercBaseOperacaoPropria, sICMSUFDevidoST, dICMSvBCSTRet, dICMSvICMSSTRet, dICMSvBCSTDest, dICMSvICMSSTDest, dICMSpCredSN, dICMSvCredSN, sIPICST, sIPIEnquadramentoClasse, _
        sIPIEnquadramentoCodigo, sIPISeloCodigo, sIPICNPJProdutor, lIPISeloQtde, dIPIBase, dIPIAliquota, dIPIValor, dIPIUnidadePadraoQtde, dIPIUnidadePadraoValor, dIIValor, dIIBase, dIIDespAduaneira, dIIIOF, sPISCST, iPISTipoCalculo, _
        dPISBase, dPISAliquota, dPISValor, dPISQtde, dPISAliquotaValor, iPISSTTipoCalculo, dPISSTBase, dPISSTAliquota, dPISSTValor, dPISSTQtde, dPISSTAliquotaValor, sCOFINSCST, iCOFINSTipoCalculo, dCOFINSBase, dCOFINSAliquota, _
        dCOFINSValor, dCOFINSQtde, dCOFINSAliquotaValor, iCOFINSSTTipoCalculo, dCOFINSSTBase, dCOFINSSTAliquota, dCOFINSSTValor, dCOFINSSTQtde, dCOFINSSTAliquotaValor, dTotTribItem, _
        dICMSInterestBCUFDest, dICMSInterestPercFCPUFDest, dICMSInterestAliqUFDest, dICMSInterestPercPartilha, dICMSInterestVlrFCPUFDest, dICMSInterestVlrUFDest, dICMSInterestVlrUFRemet, dICMSInterestAliq, iExisteICMSUFDest, sindEscala, sCNPJFab, scBenef, dvBCFCP, dpFCP, dvFCP, dvBCFCPST, dpFCPST, dvFCPST, dvBCFCPSTRet, dpFCPSTRet, dvFCPSTRet, dICMSSTCobrAntAliquota, dvBCFCPUFDest, _
        sChvNFe)
    If lErro <> AD_SQL_SUCESSO Then gError 213487
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213488
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItemNF = New ClassItemNF
    
        'Inicializa parte tributária
        Call objItemNF.Inicializa_Tributacao
        Call objItemNF.objTributacaoItemNF.Coloca_Auto
        
        'Obtém dados fixos
        objItemNF.dtDataEntrega = DATA_NULA
    
        'Obtém dados do Registro
        objItemNF.iItem = lItem
        
        objItemNF.sProdutoXml = sProdXml
        objItemNF.sDescricaoItem = sDescProd
        objItemNF.dQuantidade = dQuantidade
        objItemNF.dPrecoUnitario = dPrecoUnitario
        objItemNF.dValorDesconto = dDesconto
        objItemNF.dValorTotalXml = dPrecoTotal
        objItemNF.sEANXML = sEAN
        objItemNF.sUnidadeMedXml = sUMXml
        
        objItemNF.objInfoAdicDocItem.sMsg = sInfAdProd
        
        objNF.colItensNF.Add1 objItemNF
        
        'objItemNF.objInfoAdicDocItem.iIncluiValorTotal = IIf(iIndTot = 0, 0, IIf(sCFOP = "5902" Or sCFOP = "6902", 0, iIndTot))
        objItemNF.objInfoAdicDocItem.iIncluiValorTotal = iIndTot
        
        sProduto = ""
        
        If bMesmaEmpresa Then
        
            lErro = CF("Produto_Formata", sProdXml, sProduto, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            objItemNF.sUnidadeMed = sUMXml
            
        Else
        
            sUMAux = ""
            
            lErro = Importa_Xml_NFe_Loc_Prod2(iFilialEmpresa, sCGCEmitente, sProdXml, sEAN, sProduto, sUMXml, sUMAux, IIf(sModelo = "57", True, False), bDestinatario)
            If lErro <> SUCESSO And lErro <> ERRO_SEM_MENSAGEM Then gError ERRO_SEM_MENSAGEM
        
            If lErro = SUCESSO And sUMAux <> "" Then objItemNF.sUnidadeMed = sUMAux
            
        End If
        
        If sProduto <> "" Then
        
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            If lErro = SUCESSO Then
            
                objItemNF.objTributacaoItemNF.sProduto = objProduto.sCodigo
                objItemNF.objTributacaoItemNF.iExTIPI = objProduto.iExTIPI
                objItemNF.objTributacaoItemNF.sGenero = objProduto.sGenero
                objItemNF.objTributacaoItemNF.iProdutoEspecifico = objProduto.iProdutoEspecifico
                objItemNF.objTributacaoItemNF.sIPICodProduto = objProduto.sIPICodigo
                objItemNF.objTributacaoItemNF.sProdutoDescricao = objProduto.sDescricao
        
                objItemNF.sProduto = sProduto
        
            End If
        
        End If
        
        If objNF.sNaturezaOp = "" Then objNF.sNaturezaOp = sCFOP
        
        If sModelo <> "57" Then
        
            sCFOPInterno = String(STRING_NATUREZAOP_CODIGO, 0)
            lErro = Comando_Executar(alComando(12), "SELECT NatOpInterna, TipoTributacao FROM ImportNFeXml NXML, ImportNFeItensXml IXML WHERE CGCEmitente = ? AND NXML.ChvNFe = IXML.ChvNFe AND IXML.CFOP = ? AND NXML.FilialEmpresa = ? AND IXML.Produto = ? AND NatOpInterna <> '' ORDER BY NXML.DataEmissao DESC, NXML.ChvNFe DESC", sCFOPInterno, iTipoTributacao, sCGCEmitente, sCFOP, iFilialEmpresa, sProdXml)
            If lErro <> AD_SQL_SUCESSO Then gError 213489
            
            lErro = Comando_BuscarProximo(alComando(12))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213490
            If lErro <> AD_SQL_SUCESSO Then sCFOPInterno = ""
            
        End If
            
        If sCFOPInterno = "" And sModelo = "57" Then

            sCFOPInterno = String(STRING_NATUREZAOP_CODIGO, 0)
            If bDestinatario Then
                lErro = Comando_Executar(alComando(32), "SELECT NatOpInterna, TipoTributacao FROM ImportNFeXml NXML, ImportNFeItensXml IXML WHERE CGCEmitente = ? AND CGCDestinatario = ? AND Modelo = 57 AND NXML.ChvNFe = IXML.ChvNFe AND IXML.CFOP = ? AND NXML.FilialEmpresa = ? AND IXML.Produto = ? AND NatOpInterna <> '' ORDER BY NXML.DataEmissao DESC, NXML.ChvNFe DESC", sCFOPInterno, iTipoTributacao, sCGCEmitente, sCGCFilialEmp, sCFOP, iFilialEmpresa, sProdXml)
            Else
                lErro = Comando_Executar(alComando(32), "SELECT NatOpInterna, TipoTributacao FROM ImportNFeXml NXML, ImportNFeItensXml IXML WHERE CGCEmitente = ? AND CGCRemetente = ? AND Modelo = 57 AND NXML.ChvNFe = IXML.ChvNFe AND IXML.CFOP = ? AND NXML.FilialEmpresa = ? AND IXML.Produto = ? AND NatOpInterna <> '' ORDER BY NXML.DataEmissao DESC, NXML.ChvNFe DESC", sCFOPInterno, iTipoTributacao, sCGCEmitente, sCGCFilialEmp, sCFOP, iFilialEmpresa, sProdXml)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 213491
            
            lErro = Comando_BuscarProximo(alComando(32))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213492
            If lErro <> AD_SQL_SUCESSO Then sCFOPInterno = ""
        
        End If
        
        If sCFOPInterno = "" Then
        
            'tentar obter cfop e tipotrib de nf registrada anteriormente com mesmo cfop externo p/mesmo fornecedor
            sCFOPInterno = String(STRING_NATUREZAOP_CODIGO, 0)
            lErro = Comando_Executar(alComando(11), "SELECT TD.NaturezaOpInterna, TD.TipoTributacao FROM TributacaoDoc TD, NFiscal NF, ImportNFeXml NXML, ImportNFeItensXml IXML, FiliaisFornecedores FF WHERE NF.Fornecedor = FF.CodFornecedor AND NF.FilialForn = FF.CodFilial AND CGCEmitente = FF.CGC AND CGCEmitente = ? AND TD.TipoDoc = 0 AND TD.NumIntDoc = NF.NumIntDoc AND NF.NumIntDoc = NXML.NumIntNF AND NXML.ChvNFe = IXML.ChvNFe AND IXML.CFOP = ? AND NF.FilialEmpresa = ? AND IXML.Item = 1 ORDER BY NF.DataEmissao DESC, NF.NumIntDoc DESC", sCFOPInterno, iTipoTributacao, sCGCEmitente, sCFOP, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 213493
            
            lErro = Comando_BuscarProximo(alComando(11))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213494
            If lErro <> AD_SQL_SUCESSO Then sCFOPInterno = ""
        
        End If
        
        If sCFOPInterno = "" Then
        
            'tentar obter cfop e tipotrib de nf registrada anteriormente com mesmo cfop externo p/mesmo cliente
            sCFOPInterno = String(STRING_NATUREZAOP_CODIGO, 0)
            lErro = Comando_Executar(alComando(11), "SELECT TD.NaturezaOpInterna, TD.TipoTributacao FROM TributacaoDoc TD, NFiscal NF, ImportNFeXml NXML, ImportNFeItensXml IXML, FiliaisClientes FC WHERE NF.Cliente = FC.CodCliente AND NF.FilialCli = FC.CodFilial AND CGCEmitente = FC.CGC AND CGCEmitente = ? AND TD.TipoDoc = 0 AND TD.NumIntDoc = NF.NumIntDoc AND NF.NumIntDoc = NXML.NumIntNF AND NXML.ChvNFe = IXML.ChvNFe AND IXML.CFOP = ? AND NF.FilialEmpresa = ? AND IXML.Item = 1 ORDER BY NF.DataEmissao DESC, NF.NumIntDoc DESC", sCFOPInterno, iTipoTributacao, sCGCEmitente, sCFOP, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 213495
            
            lErro = Comando_BuscarProximo(alComando(11))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213496
            If lErro <> AD_SQL_SUCESSO Then sCFOPInterno = ""
        
        End If
        
        If sCFOPInterno = "" Then
        
            'tentar obter cfop e tipotrib de nf registrada anteriormente com mesmo cfop externo independentemente de cliente ou fornecedor
            sCFOPInterno = String(STRING_NATUREZAOP_CODIGO, 0)
            lErro = Comando_Executar(alComando(11), "SELECT TD.NaturezaOpInterna, TD.TipoTributacao FROM TributacaoDoc TD, NFiscal NF, ImportNFeXml NXML, ImportNFeItensXml IXML WHERE TD.TipoDoc = 0 AND TD.NumIntDoc = NF.NumIntDoc AND NF.NumIntDoc = NXML.NumIntNF AND NXML.ChvNFe = IXML.ChvNFe AND IXML.CFOP = ? AND NF.FilialEmpresa = ? AND IXML.Item = 1 ORDER BY NF.DataEmissao DESC, NF.NumIntDoc DESC", sCFOPInterno, iTipoTributacao, sCFOP, iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 213497
            
            lErro = Comando_BuscarProximo(alComando(11))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213498
            If lErro <> AD_SQL_SUCESSO Then sCFOPInterno = ""
        
        End If
        
        If sCFOPInterno = "" And sModelo <> "57" Then

            sCFOPInterno = String(STRING_NATUREZAOP_CODIGO, 0)
            lErro = Comando_Executar(alComando(10), "SELECT CFOPInterno FROM ImportNFeCFOP WHERE TipoAtividadeEmp = ? AND CFOPExterno = ?", sCFOPInterno, objFilialEmpresa.iTipoTribAtividade, sCFOP)
            If lErro <> AD_SQL_SUCESSO Then gError 213499

            lErro = Comando_BuscarProximo(alComando(10))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213500
            If lErro <> AD_SQL_SUCESSO Then sCFOPInterno = ""

        End If
        
        If sCFOPInterno = "" Then
        
            If sModelo = "57" Then
                If objFilialEmpresa.objEndereco.sSiglaEstado = sUFEmitente Then
                    sCFOPInterno = "1"
                Else
                    sCFOPInterno = "2"
                End If
            Else
                Select Case left(sCFOP, 1)
                
                    Case "5"
                        sCFOPInterno = "1"
                        
                    Case "6"
                        sCFOPInterno = "2"
                    
                    Case "7"
                        sCFOPInterno = "3"
                    
                End Select
            End If
            
            If sModelo = "57" Then
                sCFOPInterno = sCFOPInterno & "352"
            Else
            
                Select Case objFilialEmpresa.iTipoTribAtividade
            
                    Case 1
                        sCFOPInterno = sCFOPInterno & "102"
                        
                    Case 2
                        sCFOPInterno = sCFOPInterno & "101"
                    
                    Case 3
                        sCFOPInterno = sCFOPInterno & "103"
            
                End Select
            
            End If
            
        End If
        
        If iTipoTributacao = 0 Then
        
            If sModelo = "57" Then
                iTipoTributacao = 27
            Else
                Select Case objFilialEmpresa.iTipoTribAtividade
            
                    Case 1
                        iTipoTributacao = 4
                        
                    Case 2
                        iTipoTributacao = 5
                    
                    Case 3
                        iTipoTributacao = 6
            
                End Select
            End If
        
        End If
        
        objItemNF.sCFOPXml = sCFOP
        objItemNF.objTributacaoItemNF.sNaturezaOp = sCFOPInterno
        objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.iTipoTributacao = iTipoTributacao
        objItemNF.objTributacaoItemNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
        
        If objNF.objTributacaoNF.iTipoTributacao = 0 Then
            objNF.objTributacaoNF.iTipoTributacao = objItemNF.objTributacaoItemNF.iTipoTributacao
            objNF.objTributacaoNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
        End If
            
        objItemNF.objTributacaoItemNF.dDescontoGrid = objItemNF.dValorDesconto
        objItemNF.objTributacaoItemNF.dPrecoTotal = objItemNF.dValorTotal
        
        objItemNF.objTributacaoItemNF.dValorFreteItem = dFrete
        objItemNF.objTributacaoItemNF.dValorDescontoItem = dDesconto
        objItemNF.objTributacaoItemNF.dValorOutrasDespesasItem = dDespesas
        objItemNF.objTributacaoItemNF.dValorSeguroItem = dSeguro
        
        objItemNF.objTributacaoItemNF.iOrigemMercadoria = iOrigemMercadoria
        objItemNF.objTributacaoItemNF.iOrigemMercadoriaManual = VAR_PREENCH_MANUAL
        
        objItemNF.objTributacaoItemNF.sCST = sICMSCST
        objItemNF.objTributacaoItemNF.iCSTManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.sCSOSN = sICMSCSOSN
        objItemNF.objTributacaoItemNF.iICMSBaseModalidade = iICMSBaseModalidade
        objItemNF.objTributacaoItemNF.iICMSBaseModalidadeManual = VAR_PREENCH_MANUAL
        
        If dICMSPercRedBase = 0 Then
            objItemNF.objTributacaoItemNF.dICMSBase = dICMSBase
        ElseIf dICMSPercRedBase = 1 Then
            objItemNF.objTributacaoItemNF.dICMSBase = dICMSBase
        Else
            dBaseICMSAntesReducao = Arredonda_Moeda(dICMSBase / (1 - dICMSPercRedBase))
            If Abs(dBaseICMSAntesReducao - dPrecoTotal) < 1 Then dBaseICMSAntesReducao = dPrecoTotal
            objItemNF.objTributacaoItemNF.dICMSBase = dBaseICMSAntesReducao
        End If
        objItemNF.objTributacaoItemNF.iICMSBaseManual = VAR_PREENCH_MANUAL
        
        objItemNF.objTributacaoItemNF.dICMSPercRedBase = dICMSPercRedBase
        objItemNF.objTributacaoItemNF.iICMSPercRedBaseManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSAliquota = dICMSAliquota + dpFCP
        objItemNF.objTributacaoItemNF.iICMSAliquotaManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSValor = dICMSValor + dvFCP
        objItemNF.objTributacaoItemNF.iICMSValorManual = VAR_PREENCH_MANUAL
        
        objItemNF.objTributacaoItemNF.dICMSSubstBase = dICMSSubstBase
        objItemNF.objTributacaoItemNF.iICMSSubstBaseManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSSubstPercMVA = dICMSSubstPercMVA
        objItemNF.objTributacaoItemNF.iICMSSubstPercMVAManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSSubstPercRedBase = dICMSSubstPercRedBase
        objItemNF.objTributacaoItemNF.iICMSSubstPercRedBaseManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSSubstAliquota = dICMSSubstAliquota + dpFCPST
        objItemNF.objTributacaoItemNF.iICMSSubstAliquotaManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSSubstValor = dICMSSubstValor + dvFCPST
        objItemNF.objTributacaoItemNF.iICMSSubstValorManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSSTCobrAntBase = dICMSSTCobrAntBase
        objItemNF.objTributacaoItemNF.iICMSSTCobrAntBaseManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSSTCobrAntValor = dICMSSTCobrAntValor + dvFCPSTRet
        objItemNF.objTributacaoItemNF.iICMSSTCobrAntValorManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSPercBaseOperacaoPropria = dICMSPercBaseOperacaoPropria
        objItemNF.objTributacaoItemNF.iICMSpercBaseOperacaoPropriaManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.sICMSUFDevidoST = sICMSUFDevidoST
        objItemNF.objTributacaoItemNF.iICMSUFDevidoSTManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSvBCSTRet = dICMSvBCSTRet
        objItemNF.objTributacaoItemNF.iICMSvBCSTRetManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSvICMSSTRet = dICMSvICMSSTRet
        objItemNF.objTributacaoItemNF.iICMSvICMSSTRetManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSvBCSTDest = dICMSvBCSTDest
        objItemNF.objTributacaoItemNF.iICMSvBCSTDestManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSvICMSSTDest = dICMSvICMSSTDest
        objItemNF.objTributacaoItemNF.iICMSvICMSSTDestManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSvCredSN = dICMSvCredSN
        objItemNF.objTributacaoItemNF.iICMSvCredSNManual = VAR_PREENCH_MANUAL
        
        objItemNF.objTributacaoItemNF.iTotTribTipo = 0
        objItemNF.objTributacaoItemNF.iTotTribTipoManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dTotTrib = dTotTribItem
        objItemNF.objTributacaoItemNF.iTotTribManual = VAR_PREENCH_MANUAL
        
        objItemNF.objTributacaoItemNF.sUMTrib = sUMTrib
        objItemNF.objTributacaoItemNF.dQtdTrib = dQtdTrib
        objItemNF.objTributacaoItemNF.dValorUnitTrib = dValorUnitTrib
        
        
        'NFe 4.00
        objItemNF.objTributacaoItemNF.iICMSInterestBCUFDestManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSInterestBCUFDest = dICMSInterestBCUFDest
        objItemNF.objTributacaoItemNF.iICMSInterestPercFCPUFDestManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSInterestPercFCPUFDest = dICMSInterestPercFCPUFDest
        objItemNF.objTributacaoItemNF.iICMSInterestAliqUFDestManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSInterestAliqUFDest = dICMSInterestAliqUFDest
        objItemNF.objTributacaoItemNF.iICMSInterestPercPartilhaManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSInterestPercPartilha = dICMSInterestPercPartilha
        objItemNF.objTributacaoItemNF.iICMSInterestVlrFCPUFDestManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSInterestVlrFCPUFDest = dICMSInterestVlrFCPUFDest
        objItemNF.objTributacaoItemNF.iICMSInterestVlrUFDestManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSInterestVlrUFDest = dICMSInterestVlrUFDest
        objItemNF.objTributacaoItemNF.iICMSInterestVlrUFRemetManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSInterestVlrUFRemet = dICMSInterestVlrUFRemet
        objItemNF.objTributacaoItemNF.iICMSInterestAliqManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSInterestAliq = dICMSInterestAliq
        
        objItemNF.objTributacaoItemNF.sindEscala = sindEscala
        objItemNF.objTributacaoItemNF.sCNPJFab = sCNPJFab
        objItemNF.objTributacaoItemNF.scBenef = scBenef
        
        If dvBCFCP > 0 Or dvFCP = 0 Then
            objItemNF.objTributacaoItemNF.iICMSvBCFCPManual = VAR_PREENCH_MANUAL
        Else
            objItemNF.objTributacaoItemNF.iICMSvBCFCPManual = VAR_PREENCH_AUTOMATICO
        End If
        objItemNF.objTributacaoItemNF.dICMSvBCFCP = dvBCFCP
        objItemNF.objTributacaoItemNF.iICMSpFCPManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSpFCP = dpFCP
        objItemNF.objTributacaoItemNF.iICMSvFCPManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSvFCP = dvFCP
        objItemNF.objTributacaoItemNF.iICMSvBCFCPSTManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSvBCFCPST = dvBCFCPST
        objItemNF.objTributacaoItemNF.iICMSpFCPSTManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSpFCPST = dpFCPST
        objItemNF.objTributacaoItemNF.iICMSvFCPSTManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSvFCPST = dvFCPST
        objItemNF.objTributacaoItemNF.iICMSvBCFCPSTRetManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSvBCFCPSTRet = dvBCFCPSTRet
        objItemNF.objTributacaoItemNF.iICMSpFCPSTRetManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSpFCPSTRet = dpFCPSTRet
        objItemNF.objTributacaoItemNF.iICMSvFCPSTRetManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSvFCPSTRet = dvFCPSTRet
        
        objItemNF.objTributacaoItemNF.iICMSSTCobrAntAliquotaManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSSTCobrAntAliquota = dICMSSTCobrAntAliquota
        
        objItemNF.objTributacaoItemNF.iICMSInterestBCFCPUFDestManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dICMSInterestBCFCPUFDest = dvBCFCPUFDest
        
        If sICMSCSOSN <> "" Then
        
            lErro = Comando_Executar(alComando(8), "SELECT Tipo FROM TiposTribICMSSimples WHERE CSOSN = ?", iICMSTipoSimples, sICMSCSOSN)
            If lErro <> AD_SQL_SUCESSO Then gError 213501
    
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213502
            
            objItemNF.objTributacaoItemNF.iICMSSimplesTipo = iICMSTipoSimples
            objItemNF.objTributacaoItemNF.iICMSSimplesTipoManual = VAR_PREENCH_MANUAL
            
            objItemNF.objTributacaoItemNF.iRegimeTributario = REGIME_TRIBUTARIO_SIMPLES
        
        Else
        
            Select Case sICMSCST
            
                Case "00"
                    objItemNF.objTributacaoItemNF.iICMSTipo = 1
                
                Case "10"
                    objItemNF.objTributacaoItemNF.iICMSTipo = 6
                
                Case "20"
                    objItemNF.objTributacaoItemNF.iICMSTipo = 7
                
                Case "30"
                    objItemNF.objTributacaoItemNF.iICMSTipo = 10
                
                Case "40"
                    objItemNF.objTributacaoItemNF.iICMSTipo = 2
                
                Case "41"
                    objItemNF.objTributacaoItemNF.iICMSTipo = 0
                
                Case "50"
                    objItemNF.objTributacaoItemNF.iICMSTipo = 3
                
                Case "51"
                    objItemNF.objTributacaoItemNF.iICMSTipo = 5
                
                Case "60"
                    objItemNF.objTributacaoItemNF.iICMSTipo = 8
                
                Case "70"
                    objItemNF.objTributacaoItemNF.iICMSTipo = 4
                
                Case "90"
                    objItemNF.objTributacaoItemNF.iICMSTipo = 99
                
                Case Else
                    objItemNF.objTributacaoItemNF.iICMSTipo = 0 'É NF de ISS
                    'gError 212038
                
            End Select
            
            objItemNF.objTributacaoItemNF.iICMSTipoManual = VAR_PREENCH_MANUAL
            
            objItemNF.objTributacaoItemNF.iRegimeTributario = REGIME_TRIBUTARIO_NORMAL
            
        End If
        
        objItemNF.objTributacaoItemNF.iRegimeTributarioManual = VAR_PREENCH_MANUAL
                    
        If sCFOPPrincipal = "" Then sCFOPPrincipal = sCFOPInterno
        
        If sIPICST <> "" Then
        
            Select Case sIPICST
            
                Case "50"
                    objItemNF.objTributacaoItemNF.iIPITipo = 1
                    
                Case "51"
                    objItemNF.objTributacaoItemNF.iIPITipo = 4
                    
                Case "52"
                    objItemNF.objTributacaoItemNF.iIPITipo = 3
                    
                Case "53"
                    objItemNF.objTributacaoItemNF.iIPITipo = 0
                    
                Case "54"
                    objItemNF.objTributacaoItemNF.iIPITipo = 7
                    
                Case "55"
                    objItemNF.objTributacaoItemNF.iIPITipo = 5
                    
                Case "99"
                    objItemNF.objTributacaoItemNF.iIPITipo = 6
                    
                Case Else
                    objItemNF.objTributacaoItemNF.iIPITipo = 6
                    
            End Select
            
            With objItemNF.objTributacaoItemNF
                .iIPITipoManual = VAR_PREENCH_MANUAL
                .dIPIBaseCalculo = dIPIBase
                dNFIPIBase = dNFIPIBase + dIPIBase
                .dIPIAliquota = dIPIAliquota
                .iIPIAliquotaManual = VAR_PREENCH_MANUAL
                .dIPIValor = dIPIValor
                .iIPIValorManual = VAR_PREENCH_MANUAL
                .dIPIUnidadePadraoQtde = dIPIUnidadePadraoQtde
                .iIPIUnidadePadraoQtdeManual = VAR_PREENCH_MANUAL
                .dIPIUnidadePadraoValor = dIPIUnidadePadraoValor
                .iIPIUnidadePadraoValorManual = VAR_PREENCH_MANUAL
            
                .sIPIEnquadramentoClasse = sIPIEnquadramentoClasse
                .sIPIEnquadramentoCodigo = sIPIEnquadramentoCodigo
                .sIPISeloCodigo = sIPISeloCodigo
                If Len(Trim(Replace(sIPICNPJProdutor, "0", ""))) > 0 Then
                    .sIPICNPJProdutor = sIPICNPJProdutor
                Else
                    .sIPICNPJProdutor = ""
                End If
                .lIPISeloQtde = lIPISeloQtde
            
            End With
            
        Else
            objItemNF.objTributacaoItemNF.iIPITipoManual = VAR_PREENCH_AUTOMATICO
        End If
        
        Select Case sPISCST
        
            Case "01"
                objItemNF.objTributacaoItemNF.iPISTipo = 1
                
            Case Else
                objItemNF.objTributacaoItemNF.iPISTipo = StrParaInt(sPISCST)
            
        End Select
        objItemNF.objTributacaoItemNF.iPISTipoManual = VAR_PREENCH_AUTOMATICO
        
        objItemNF.objTributacaoItemNF.dPISBase = dPISBase
        objItemNF.objTributacaoItemNF.iPISBaseManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dPISAliquota = dPISAliquota
        objItemNF.objTributacaoItemNF.iPISAliquotaManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dPISValor = dPISValor
        objItemNF.objTributacaoItemNF.iPISValorManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dPISQtde = dPISQtde
        objItemNF.objTributacaoItemNF.iPISQtdeManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dPISAliquotaValor = dPISAliquotaValor
        objItemNF.objTributacaoItemNF.iPISAliquotaValorManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.iPISSTTipoCalculo = iPISSTTipoCalculo
        objItemNF.objTributacaoItemNF.iPISSTTipoCalculoManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dPISSTBase = dPISSTBase
        objItemNF.objTributacaoItemNF.iPISSTBaseManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dPISSTAliquota = dPISSTAliquota
        objItemNF.objTributacaoItemNF.iPISSTAliquotaManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dPISSTValor = dPISSTValor
        objItemNF.objTributacaoItemNF.iPISSTValorManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dPISSTQtde = dPISSTQtde
        objItemNF.objTributacaoItemNF.iPISSTQtdeManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dPISSTAliquotaValor = dPISSTAliquotaValor
        objItemNF.objTributacaoItemNF.iPISSTAliquotaValorManual = VAR_PREENCH_MANUAL
        
        'If sModelo = "57" Then
        'Deixando automático para todas NFes e CTes
            objItemNF.objTributacaoItemNF.iPISBaseManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iPISAliquotaManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iPISValorManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iPISQtdeManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iPISAliquotaValorManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iPISSTTipoCalculoManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iPISSTBaseManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iPISSTAliquotaManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iPISSTValorManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iPISSTQtdeManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iPISSTAliquotaValorManual = VAR_PREENCH_AUTOMATICO
        'End If

        Select Case sCOFINSCST
        
            Case "01"
                objItemNF.objTributacaoItemNF.iCOFINSTipo = 1
                
            Case Else
                objItemNF.objTributacaoItemNF.iCOFINSTipo = StrParaInt(sCOFINSCST)
            
        End Select
        objItemNF.objTributacaoItemNF.iCOFINSTipoManual = VAR_PREENCH_AUTOMATICO
        
        objItemNF.objTributacaoItemNF.dCOFINSBase = dCOFINSBase
        objItemNF.objTributacaoItemNF.iCOFINSBaseManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dCOFINSAliquota = dCOFINSAliquota
        objItemNF.objTributacaoItemNF.iCOFINSAliquotaManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dCOFINSValor = dCOFINSValor
        objItemNF.objTributacaoItemNF.iCOFINSValorManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dCOFINSQtde = dCOFINSQtde
        objItemNF.objTributacaoItemNF.iCOFINSQtdeManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dCOFINSAliquotaValor = dCOFINSAliquotaValor
        objItemNF.objTributacaoItemNF.iCOFINSAliquotaValorManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.iCOFINSSTTipoCalculo = iCOFINSSTTipoCalculo
        objItemNF.objTributacaoItemNF.iCOFINSSTTipoCalculoManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dCOFINSSTBase = dCOFINSSTBase
        objItemNF.objTributacaoItemNF.iCOFINSSTBaseManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dCOFINSSTAliquota = dCOFINSSTAliquota
        objItemNF.objTributacaoItemNF.iCOFINSSTAliquotaManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dCOFINSSTValor = dCOFINSSTValor
        objItemNF.objTributacaoItemNF.iCOFINSSTValorManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dCOFINSSTQtde = dCOFINSSTQtde
        objItemNF.objTributacaoItemNF.iCOFINSSTQtdeManual = VAR_PREENCH_MANUAL
        objItemNF.objTributacaoItemNF.dCOFINSSTAliquotaValor = dCOFINSSTAliquotaValor
        objItemNF.objTributacaoItemNF.iCOFINSSTAliquotaValorManual = VAR_PREENCH_MANUAL
        
        'If sModelo = "57" Then
        'Deixando automático para todas NFes e CTes
            objItemNF.objTributacaoItemNF.iCOFINSBaseManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iCOFINSAliquotaManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iCOFINSValorManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iCOFINSQtdeManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iCOFINSAliquotaValorManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iCOFINSSTTipoCalculoManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iCOFINSSTBaseManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iCOFINSSTAliquotaManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iCOFINSSTValorManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iCOFINSSTQtdeManual = VAR_PREENCH_AUTOMATICO
            objItemNF.objTributacaoItemNF.iCOFINSSTAliquotaValorManual = VAR_PREENCH_AUTOMATICO
        'End If
        
        If sISSCST <> "" Then
        
            objItemNF.objTributacaoItemNF.sISSCST = sISSCST
            objItemNF.objTributacaoItemNF.iISSCidadeIBGEManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.sISSCidadeIBGE = sISSCidadeIBGE
            objItemNF.objTributacaoItemNF.iISSCidadeIBGEManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.dISSBase = dISSBase
            objItemNF.objTributacaoItemNF.iISSBaseManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.dISSAliquota = dISSAliquota
            objItemNF.objTributacaoItemNF.iISSAliquotaManual = VAR_PREENCH_MANUAL
            objItemNF.objTributacaoItemNF.dISSValor = dISSValor
            objItemNF.objTributacaoItemNF.iISSValorManual = VAR_PREENCH_MANUAL
        
        End If
        
        objItemNF.objTributacaoItemNF.dIIValor = dIIValor
        objItemNF.objTributacaoItemNF.dIIBase = dIIBase
        objItemNF.objTributacaoItemNF.dIIDespAduaneira = dIIDespAduaneira
        objItemNF.objTributacaoItemNF.dIIIOF = dIIIOF
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213503
    
    Loop
        
    objNF.objTributacaoNF.dIPIBase = Arredonda_Moeda(dNFIPIBase)
    objNF.objTributacaoNF.iIPIBaseManual = VAR_PREENCH_MANUAL
        
    If iExisteCobr <> 0 Then
    
        'carregar a cobrança
        lErro = Comando_Executar(alComando(2), "SELECT DataVencimento, Valor FROM ImportNFeCobrXml WHERE ChvNFe = ? ORDER BY Item", _
            dtDataParc, dValorParc, sChvNFe)
        If lErro <> AD_SQL_SUCESSO Then gError 213504

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213505
        
        iIndice = 0
        Do While lErro <> AD_SQL_SEM_DADOS
    
            iIndice = iIndice + 1
            
            Set objParcela = New ClassParcelaPagar
        
            With objParcela
                .iNumParcela = iIndice
                .dtDataVencimento = dtDataParc
                .dValor = dValorParc
                
                objNF.ColParcelaPagar.Add 0, 0, iIndice, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimento, .dValor, .dValor, 0, 1, 0, 0, "", ""
            End With
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213506
        
        Loop

    End If

    objNF.objTributacaoNF.sNaturezaOpInterna = sCFOPPrincipal
    objNF.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL

    If sModelo = "57" And objTela.Name = "NFiscalEntrada" Then
        If bDestinatario Then
            objNF.iTipoNFiscal = DOCINFO_CFECT 'Conhecimento de Frete - Compra ou Transferencia
        Else
            objNF.iTipoNFiscal = DOCINFO_CFEV 'Conhecimento de Frete - Venda ou Transferencia
        End If
    Else
        objNF.iTipoNFiscal = Codigo_Extrai(objTela.TipoNFiscal.Text)
    End If
  
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NFe_Chv_Monta_NF1 = SUCESSO

    Exit Function

Erro_NFe_Chv_Monta_NF1:

    NFe_Chv_Monta_NF1 = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 213475
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 213476, 213477, 213483 To 213500, 213503 To 213506
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTNFEXML", gErr)
            
        Case 213478
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPORTNFEXML_NAO_DISPONIVEL", gErr, sChvNFe)
            
        Case 213479 To 213482
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESA", gErr)
       
        Case 213501, 213502
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSTRIBICMS", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213507)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function Importa_Xml_NFe_Loc_CodProd(sProduto As String, ByVal lComando As Long, ByVal sProdXml As String) As Long

Dim lErro As Long
Dim sProdutoFormatado As String, iProdutoPreenchido As Integer

On Error GoTo Erro_Importa_Xml_NFe_Loc_CodProd

    lErro = CF("Produto_Formata", sProdXml, sProdutoFormatado, iProdutoPreenchido, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    sProduto = String(STRING_MAXIMO, 0)
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Produtos WHERE Codigo = ?", sProduto, sProdutoFormatado)
    If lErro <> AD_SQL_SUCESSO Then gError 209858
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209859
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    Importa_Xml_NFe_Loc_CodProd = SUCESSO
    
    Exit Function
    
Erro_Importa_Xml_NFe_Loc_CodProd:

    Importa_Xml_NFe_Loc_CodProd = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209858, 209859
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201428)

    End Select
    
    Exit Function

End Function

Private Function Importa_Xml_NFe_Loc_FornProdFF(sProduto As String, ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal lForn As Long, ByVal iFilForn As Integer, ByVal sProdXml As String) As Long

Dim lErro As Long
        
On Error GoTo Erro_Importa_Xml_NFe_Loc_FornProdFF

    lErro = Comando_Executar(lComando, "SELECT Produto FROM FornecedorProdutoFF WHERE FilialEmpresa = ? AND Fornecedor = ? AND FilialForn = ? AND ProdutoFornecedor = ?", sProduto, iFilialEmpresa, lForn, iFilForn, sProdXml)
    If lErro <> AD_SQL_SUCESSO Then gError 209860
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209861
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    Importa_Xml_NFe_Loc_FornProdFF = SUCESSO
    
    Exit Function
    
Erro_Importa_Xml_NFe_Loc_FornProdFF:

    Importa_Xml_NFe_Loc_FornProdFF = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209860, 209861
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201427)

    End Select
    
    Exit Function

End Function

Private Function Importa_Xml_NFe_Loc_FornProdFFCGC(sProduto As String, ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal sCGCFilForn As String, ByVal sProdXml As String) As Long

Dim lErro As Long
        
On Error GoTo Erro_Importa_Xml_NFe_Loc_FornProdFFCGC

    lErro = Comando_Executar(lComando, "SELECT Produto FROM FornecedorProdutoFF, FiliaisFornecedores WHERE FiliaisFornecedores.CGC = ? AND FilialEmpresa = ? AND Fornecedor = CodFornecedor AND FilialForn = CodFilial AND ProdutoFornecedor = ?", sProduto, sCGCFilForn, iFilialEmpresa, sProdXml)
    If lErro <> AD_SQL_SUCESSO Then gError 209860
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209861
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    Importa_Xml_NFe_Loc_FornProdFFCGC = SUCESSO
    
    Exit Function
    
Erro_Importa_Xml_NFe_Loc_FornProdFFCGC:

    Importa_Xml_NFe_Loc_FornProdFFCGC = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209860, 209861
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201427)

    End Select
    
    Exit Function

End Function

Private Function Importa_Xml_NFe_Loc_EAN(sProduto As String, ByVal lComando As Long, ByVal sEANXML As String) As Long

Dim lErro As Long
        
On Error GoTo Erro_Importa_Xml_NFe_Loc_EAN

    lErro = Comando_Executar(lComando, "SELECT CodProduto FROM ProdutoCodBarras WHERE CodBarras = ? ORDER BY CodProduto", sProduto, sEANXML)
    If lErro <> AD_SQL_SUCESSO Then gError 209862
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209863
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    Importa_Xml_NFe_Loc_EAN = SUCESSO
    
    Exit Function
    
Erro_Importa_Xml_NFe_Loc_EAN:

    Importa_Xml_NFe_Loc_EAN = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209862, 209863
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201426)

    End Select
    
    Exit Function

End Function

Private Function Importa_Xml_NFe_Loc_CodProd2(sProduto As String, ByVal lComando As Long, ByVal sProdXml As String) As Long

Dim lErro As Long
Dim sProdutoFormatado As String

On Error GoTo Erro_Importa_Xml_NFe_Loc_CodProd2

    sProdutoFormatado = Replace(Replace(Replace(Replace(Replace(Replace(sProdXml, " ", ""), ".", ""), "_", ""), "-", ""), "/", ""), "\", "")
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Produtos WHERE REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(Codigo, ' ', ''), '.', ''), '_', ''), '-', ''), '/', ''), '\', '') = ?", sProduto, sProdutoFormatado)
    If lErro <> AD_SQL_SUCESSO Then gError 209862
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209863
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    Importa_Xml_NFe_Loc_CodProd2 = SUCESSO
    
    Exit Function
    
Erro_Importa_Xml_NFe_Loc_CodProd2:

    Importa_Xml_NFe_Loc_CodProd2 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209862, 209863
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201425)

    End Select
    
    Exit Function

End Function

Private Function Importa_Xml_NFe_Loc_ProdNomeRed(sProduto As String, ByVal lComando As Long, ByVal sProdXml As String) As Long

Dim lErro As Long
        
On Error GoTo Erro_Importa_Xml_NFe_Loc_ProdNomeRed

    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Produtos WHERE NomeReduzido = ?", sProduto, sProdXml)
    If lErro <> AD_SQL_SUCESSO Then gError 209890
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209891
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    Importa_Xml_NFe_Loc_ProdNomeRed = SUCESSO
    
    Exit Function
    
Erro_Importa_Xml_NFe_Loc_ProdNomeRed:

    Importa_Xml_NFe_Loc_ProdNomeRed = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209890, 209891
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201424)

    End Select
    
    Exit Function

End Function

Private Function Importa_Xml_NFe_Loc_ProdRef(sProduto As String, ByVal lComando As Long, ByVal sProdXml As String) As Long

Dim lErro As Long
        
On Error GoTo Erro_Importa_Xml_NFe_Loc_ProdRef

    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Produtos WHERE Referencia = ?", sProduto, sProdXml)
    If lErro <> AD_SQL_SUCESSO Then gError 209892
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209893
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Importa_Xml_NFe_Loc_ProdRef = SUCESSO
    
    Exit Function
    
Erro_Importa_Xml_NFe_Loc_ProdRef:

    Importa_Xml_NFe_Loc_ProdRef = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209892, 209893
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201423)

    End Select
    
    Exit Function

End Function

Private Function Importa_Xml_NFe_Loc_ProdCGC(sProduto As String, ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal sCgc As String, ByVal sProdXml As String) As Long

Dim lErro As Long
        
On Error GoTo Erro_Importa_Xml_NFe_Loc_ProdCGC

    lErro = Comando_Executar(lComando, "SELECT IXML.ProdutoCorporator FROM ImportNFeXml NXML, ImportNFeItensXml IXML WHERE IXML.ChvNFe = NXML.ChvNFe AND NXML.CGCEmitente = ? AND IXML.Produto = ? AND IXML.ProdutoCorporator <> '' AND NXML.FilialEmpresa = ? ORDER BY NXML.DataEmissao DESC, NXML.ChvNFe DESC", sProduto, sCgc, sProdXml, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 209860
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209861
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    Importa_Xml_NFe_Loc_ProdCGC = SUCESSO
    
    Exit Function
    
Erro_Importa_Xml_NFe_Loc_ProdCGC:

    Importa_Xml_NFe_Loc_ProdCGC = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209860, 209861
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201427)

    End Select
    
    Exit Function

End Function

Private Function Importa_Xml_NFe_Loc_Prod_CTe(sProduto As String, ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal sProdXml As String, ByVal bDestinatario As Boolean) As Long

Dim lErro As Long
        
On Error GoTo Erro_Importa_Xml_NFe_Loc_Prod_CTe

    If bDestinatario Then
        'COMPRA
        lErro = Comando_Executar(lComando, "SELECT IXML.ProdutoCorporator FROM ImportNFeXml NXML, ImportNFeItensXml IXML WHERE IXML.ChvNFe = NXML.ChvNFe AND NXML.Modelo = '57' AND IXML.Produto = ? AND IXML.ProdutoCorporator <> '' AND NXML.FilialEmpresa = ? AND CGCDestinatario IN (SELECT CGC FROM FiliaisEmpresa WHERE FilialEmpresa = ?) ORDER BY NXML.DataEmissao DESC ", sProduto, sProdXml, iFilialEmpresa, iFilialEmpresa)
    Else
        'VENDA
        lErro = Comando_Executar(lComando, "SELECT IXML.ProdutoCorporator FROM ImportNFeXml NXML, ImportNFeItensXml IXML WHERE IXML.ChvNFe = NXML.ChvNFe AND NXML.Modelo = '57' AND IXML.Produto = ? AND IXML.ProdutoCorporator <> '' AND NXML.FilialEmpresa = ? AND CGCDestinatario NOT IN (SELECT CGC FROM FiliaisEmpresa WHERE FilialEmpresa = ?) ORDER BY NXML.DataEmissao DESC ", sProduto, sProdXml, iFilialEmpresa, iFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 209860
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209861
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    Importa_Xml_NFe_Loc_Prod_CTe = SUCESSO
    
    Exit Function
    
Erro_Importa_Xml_NFe_Loc_Prod_CTe:

    Importa_Xml_NFe_Loc_Prod_CTe = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209860, 209861
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201427)

    End Select
    
    Exit Function

End Function

Private Function Importa_Xml_NFe_Loc_Prod2(ByVal iFilialEmpresa As Integer, ByVal sCGCEmitente As String, ByVal sProdXml As String, ByVal sEANXML As String, sProduto As String, ByVal sUMXml As String, sUM As String, Optional ByVal bCTe As Boolean = False, Optional ByVal bDestinatario As Boolean = False) As Long
'para identificar um produto e unidade de medida no Corporator a partir de informacoes do xml
Dim lErro As Long, iIndice As Integer, iClasse As Integer
Dim alComando(1 To 9) As Long, bAchou As Boolean, sTestes As String
Dim sProdutoAux As String, sProdutoFormatado As String, iProdutoPreenchido As Integer, sUMAux As String

On Error GoTo Erro_Importa_Xml_NFe_Loc_Prod2

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209857
    Next
    
    sProdutoAux = String(STRING_PRODUTO, 0)
    sUMAux = String(STRING_UM_SIGLA, 0)
    
    sTestes = UCase(Trim(gobjCRFAT.sNFImportXmlProd))
    
    If sTestes = "" Then sTestes = "ABCDEF"
    
    bAchou = False
    For iIndice = 1 To Len(sTestes)
    
        Select Case Mid(sTestes, iIndice, 1)
        
            Case "A" 'pelo codigo de barras
                lErro = Importa_Xml_NFe_Loc_EAN(sProdutoAux, alComando(3), sEANXML)
                If lErro = SUCESSO Then
                
                    bAchou = True
                    Exit For
                    
                End If
                
            Case "B" 'pelo codigo de produto utilizado anteriormente para um produto deste fornecedor/cliente
                If bCTe Then
                    'Ignora o emissor, apenas busca o produto usado em outros CTes
                    lErro = Importa_Xml_NFe_Loc_Prod_CTe(sProdutoAux, alComando(7), iFilialEmpresa, sProdXml, bDestinatario)
                Else
                    lErro = Importa_Xml_NFe_Loc_ProdCGC(sProdutoAux, alComando(7), iFilialEmpresa, sCGCEmitente, sProdXml)
                End If
                If lErro = SUCESSO Then

                    bAchou = True
                    Exit For

                End If
        
            Case "C" 'pelo codigo do produto da tela fornecedor x produto
                lErro = Importa_Xml_NFe_Loc_FornProdFFCGC(sProdutoAux, alComando(2), iFilialEmpresa, sCGCEmitente, sProdXml)
                If lErro = SUCESSO Then

                    bAchou = True
                    Exit For

                End If

            Case "D" 'prodxml = codproduto
                lErro = Importa_Xml_NFe_Loc_CodProd(sProdutoAux, alComando(1), sProdXml)
                If lErro = SUCESSO Then
                
                    bAchou = True
                    Exit For
                    
                End If
                
            Case "E" 'pelo codigo sem caracteres de pontuação
                lErro = Importa_Xml_NFe_Loc_CodProd2(sProdutoAux, alComando(4), sProdXml)
                If lErro = SUCESSO Then
                
                    bAchou = True
                    Exit For
                    
                End If
                
            Case "F" 'pelo nome reduzido do produto
                lErro = Importa_Xml_NFe_Loc_ProdNomeRed(sProdutoAux, alComando(5), sProdXml)
                If lErro = SUCESSO Then
                
                    bAchou = True
                    Exit For
                    
                End If
                
            Case "G" 'pela referencia do produto
                lErro = Importa_Xml_NFe_Loc_ProdRef(sProdutoAux, alComando(6), sProdXml)
                If lErro = SUCESSO Then
                
                    bAchou = True
                    Exit For
                    
                End If
    
        End Select
        
    Next
    
    If bAchou = False Then gError ERRO_SEM_MENSAGEM
    
    sProduto = sProdutoAux
    
    lErro = CF("ImportNFeXml_ObterSiglaUM", iFilialEmpresa, sCGCEmitente, sProdXml, sUMXml, sUMAux)
    If lErro <> SUCESSO And lErro <> ERRO_SEM_MENSAGEM Then gError ERRO_SEM_MENSAGEM
    If lErro = SUCESSO Then
    
        lErro = Comando_Executar(alComando(8), "SELECT P.ClasseUM FROM Produtos P, UnidadesDeMedida U WHERE P.Codigo = ? AND P.ClasseUM = U.Classe AND U.Sigla = ?", iClasse, sProduto, sUMAux)
        If lErro <> AD_SQL_SUCESSO Then gError 201662
        
        lErro = Comando_BuscarProximo(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201663
        
        If lErro = AD_SQL_SUCESSO Then
        
            sUM = sUMAux
            
        End If
                    
    Else
        
        lErro = Comando_Executar(alComando(9), "SELECT P.ClasseUM FROM Produtos P, UnidadesDeMedida U WHERE P.Codigo = ? AND P.ClasseUM = U.Classe AND U.Sigla = ?", iClasse, sProduto, sUMXml)
        If lErro <> AD_SQL_SUCESSO Then gError 201664
        
        lErro = Comando_BuscarProximo(alComando(9))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201665
        
        If lErro = AD_SQL_SUCESSO Then sUM = sUMXml
        
    End If
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Importa_Xml_NFe_Loc_Prod2 = SUCESSO

    Exit Function

Erro_Importa_Xml_NFe_Loc_Prod2:

    Importa_Xml_NFe_Loc_Prod2 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209857
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 201662 To 201665
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209865)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function ImportNFeXml_ObterSiglaUM(ByVal iFilialEmpresa As Integer, ByVal sCGCEmitente As String, ByVal sProdXml As String, ByVal sUMXml As String, sUM As String) As Long

Dim lErro As Long, lComando As Long, sUMAux As String

On Error GoTo Erro_ImportNFeXml_ObterSiglaUM

    lComando = Comando_Abrir
    If lComando = 0 Then gError 201465
    
    sUMAux = String(STRING_UM_SIGLA, 0)
    lErro = Comando_Executar(lComando, "SELECT SiglaUMCorporator FROM ImportNFeItensXml IXML, ImportNFeXml NXML WHERE NXML.FilialEmpresa = ? AND NXML.ChvNFe = IXML.ChvNFe AND NXML.CGCEmitente = ? AND IXML.Produto = ? AND IXML.UM = ? AND SiglaUMCorporator <> '' ORDER BY DataEmissao DESC, NXML.ChvNFe DESC", _
        sUMAux, iFilialEmpresa, sCGCEmitente, sProdXml, sUMXml)
    If lErro <> AD_SQL_SUCESSO Then gError 201466
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201467
    
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sUM = sUMAux
    
    Call Comando_Fechar(lComando)
    
    ImportNFeXml_ObterSiglaUM = SUCESSO
    
    Exit Function
    
Erro_ImportNFeXml_ObterSiglaUM:

    ImportNFeXml_ObterSiglaUM = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 201465
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 201466, 201467
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTNFEXML", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201468)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ReqProducao_Rateia_Insumos(ByVal objMovEstoque As ClassMovEstoque, ByVal colItens As ColItensMovEstoque) As Long

Dim lErro As Long, iIndice As Integer
Dim objOP As New ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim objItem As ClassItemMovEstoque
Dim objItemAux As ClassItemMovEstoque
Dim colItensRatear As New Collection
Dim colKits As New Collection, dFatorQtdeOPInsumos As Double
Dim objKit As ClassKit, objProdutoKit As ClassProdutoKit
Dim objProduto As ClassProduto, dFatorUMConv As Double
Dim bAchou As Boolean, dTotalUso As Double, dQtdeDist As Double
Dim dQtdeAjustada As Double, colAux As Collection
Dim objRatroMovto As ClassRastreamentoMovto
Dim objRatroMovtoAux As ClassRastreamentoMovto

On Error GoTo Erro_ReqProducao_Rateia_Insumos
    
    'Obtem os itens sem Produto para poder ratear e verifica se eles são de uma única OP
    For Each objItem In objMovEstoque.colItens
        If Len(Trim(objItem.sOPCodigo)) <> 0 And Len(Trim(objItem.sProdutoOP)) = 0 Then
            If Len(Trim(objOP.sCodigo)) = 0 Then
                objOP.sCodigo = objItem.sOPCodigo
                objOP.iFilialEmpresa = objItem.iFilialEmpresa
                If objOP.iFilialEmpresa = 0 Then objOP.iFilialEmpresa = giFilialEmpresa
            End If
            If objItem.sOPCodigo <> objOP.sCodigo Or (objItem.iFilialEmpresa <> objOP.iFilialEmpresa And objItem.iFilialEmpresa <> 0) Then gError 211681 'Só pode ter produtos de uma mesma OP
            colItensRatear.Add objItem
        Else
            colItens.Add1 objItem
        End If
    Next
    
    If colItensRatear.Count > 0 Then
    
        lErro = CF("OrdemDeProducao_Le_ComItens", objOP)
        If lErro <> SUCESSO And lErro <> 21960 Then gError ERRO_SEM_MENSAGEM
        
        For Each objItemOP In objOP.colItens
        
            Set objKit = New ClassKit
            
            colKits.Add objKit
            
            objKit.sProdutoRaiz = objItemOP.sProduto
            objKit.sVersao = objItemOP.sVersao
        
            'Le o Kit / Versao
            lErro = CF("Kit_Le", objKit)
            If lErro <> SUCESSO And lErro <> 21826 Then gError ERRO_SEM_MENSAGEM
            
            lErro = CF("Kit_Le_Componentes", objKit)
            If lErro <> SUCESSO And lErro <> 21831 Then gError ERRO_SEM_MENSAGEM

            'Para cada ProdutoKit na coleção de componentes
            For Each objProdutoKit In objKit.colComponentes
            
                Set objProduto = New ClassProduto
            
                objProduto.sCodigo = objProdutoKit.sProduto
                
                'Lê o Produto Componente do Kit
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                
                objProdutoKit.iClasseUM = objProduto.iClasseUM
                            
                'Se for o ProdutoRaiz do Kit
                If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                           
                    'Descobre o fator de conversao da UM do Item da OP p/UM do Kit do produto
                    lErro = CF("UM_Conversao_Trans", objProdutoKit.iClasseUM, objItemOP.sSiglaUM, objProdutoKit.sUnidadeMed, dFatorUMConv)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                    dFatorQtdeOPInsumos = (objItemOP.dQuantidade * dFatorUMConv) / objProdutoKit.dQuantidade
                
                    objItemOP.dRisco = dFatorQtdeOPInsumos 'Armazena o fator de ajuste para os insumos

                End If
                
            Next

        Next
        
        For Each objItem In colItensRatear
            dTotalUso = 0
            dQtdeDist = 0
            Set colAux = New Collection
            iIndice = 0
            For Each objItemOP In objOP.colItens
                iIndice = iIndice + 1
                Set objKit = colKits.Item(iIndice)
                For Each objProdutoKit In objKit.colComponentes
                    If objProdutoKit.sProduto = objItem.sProduto Then
                        'Descobre o fator de conversao da UM do Item da OP p/UM do Kit do produto
                        lErro = CF("UM_Conversao_Trans", objProdutoKit.iClasseUM, objProdutoKit.sUnidadeMed, objItem.sSiglaUM, dFatorUMConv)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        objProdutoKit.dPercentualPerda = dFatorUMConv
                        dTotalUso = dTotalUso + (objProdutoKit.dQuantidade * objProdutoKit.dPercentualPerda) * objItemOP.dRisco
                        Exit For
                    End If
                Next
            Next
            iIndice = 0
            For Each objItemOP In objOP.colItens
                iIndice = iIndice + 1
                Set objKit = colKits.Item(iIndice)
                For Each objProdutoKit In objKit.colComponentes
                    If objProdutoKit.sProduto = objItem.sProduto Then
                        Set objItemAux = New ClassItemMovEstoque
                        
                        'Pega a quantidade total e faz a propoção de acordo com a quantidade usada por todos os produto
                        dQtdeAjustada = objItem.dQuantidade * (((objProdutoKit.dQuantidade * objProdutoKit.dPercentualPerda) * objItemOP.dRisco) / dTotalUso)
                        If Abs(dQtdeAjustada - Round(dQtdeAjustada, 0)) < 0.01 Then dQtdeAjustada = Round(dQtdeAjustada, 0)

                        dQtdeDist = dQtdeDist + dQtdeAjustada
                                                
                        objItemAux.dQuantidade = dQtdeAjustada
                        objItemAux.iAlmoxarifado = objItem.iAlmoxarifado
                        objItemAux.iBenef = objItem.iBenef
                        objItemAux.sAlmoxarifadoNomeRed = objItem.sAlmoxarifadoNomeRed
                        objItemAux.sCcl = objItem.sCcl
                        objItemAux.sOPCodigo = objItem.sOPCodigo
                        objItemAux.sProduto = objItem.sProduto
                        objItemAux.sProdutoDesc = objItem.sProdutoDesc
                        objItemAux.sProdutoOP = objItemOP.sProduto
                        objItemAux.sSiglaUM = objItem.sSiglaUM
                        
                        If Not (objItem.colRastreamentoMovto Is Nothing) Then
                            If objItem.colRastreamentoMovto.Count = 1 Then
                                Set objRatroMovto = objItem.colRastreamentoMovto.Item(1)
                                Set objRatroMovtoAux = New ClassRastreamentoMovto
                                objRatroMovtoAux.sLote = objRatroMovto.sLote
                                objRatroMovtoAux.iFilialOP = objRatroMovto.iFilialOP
                                Set objItemAux.colRastreamentoMovto = New Collection
                                objItemAux.colRastreamentoMovto.Add objRatroMovtoAux
                            End If
                        End If
                        
                        colItens.Add1 objItemAux
                        colAux.Add objItemAux
                        
                        Exit For
                    End If
                Next
            Next
            'Ajusta as quantidade para bater exato somente no que não é inteiro
            If Abs(objItem.dQuantidade - dQtdeDist) > QTDE_ESTOQUE_DELTA Then
                For Each objItemAux In colAux
                    If Abs(dQtdeAjustada - Round(dQtdeAjustada, 0)) > QTDE_ESTOQUE_DELTA Then
                        objItemAux.dQuantidade = objItemAux.dQuantidade + (objItem.dQuantidade - dQtdeDist)
                        dQtdeDist = objItem.dQuantidade
                        Exit For
                    End If
                Next
            End If
            'Ajusta as quantidade para bater exato, caso ainda exista diferença
            If Abs(objItem.dQuantidade - dQtdeDist) > QTDE_ESTOQUE_DELTA Then
                For Each objItemAux In colAux
                    objItemAux.dQuantidade = objItemAux.dQuantidade + (objItem.dQuantidade - dQtdeDist)
                    dQtdeDist = objItem.dQuantidade
                    Exit For
                Next
            End If
        Next
    
    End If
    
    ReqProducao_Rateia_Insumos = SUCESSO

    Exit Function

Erro_ReqProducao_Rateia_Insumos:

    ReqProducao_Rateia_Insumos = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 211681
            Call Rotina_Erro(vbOKOnly, "ERRO_REQPROD_RATEIO_VARIAS_OPS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211682)

    End Select

    Exit Function
    
End Function

Public Function ICMSSTCobrAnt_Obter(ByVal iFilialEmpresa As Integer, ByVal dtDataRef As Date, ByVal sProduto As String, ByVal dQtde As Double, ByVal sUM As String, dBCA As Double, dVCA As Double) As Long
'obtem base e valor de ICMS st cobrados anteriormente

Dim lErro As Long, lComando As Long, sUMNF As String, dQtdeNF As Double, dICMSSubstBaseNF As Double, dICMSSubstValor As Double

On Error GoTo Erro_ICMSSTCobrAnt_Obter

    lComando = Comando_Abrir
    If lComando = 0 Then gError 201501
    
    sUMNF = String(STRING_UM_SIGLA, 0)
    lErro = Comando_Executar(lComando, "SELECT TDOCI.UnidadeMed, TDOCI.Quantidade, TDOCI.ICMSSubstBase, TDOCI.ICMSSubstValor FROM NFiscal NF, TiposdocInfo TDI, TributacaoDoc TDOC, TributacaoDocItem TDOCI WHERE NF.Status <> 7 AND NF.FilialEmpresa = ? AND " & _
        "NF.TipoNFiscal = TDI.Codigo AND NF.NumIntDoc = TDOC.NumIntDoc AND TDOC.TipoDoc = 0 AND TDOCI.TipoDoc = 0 AND " & _
        "TDOC.NumIntDoc = TDOCI.NumIntDoc AND TDOCI.Produto = ? AND NF.DataEntrada <= ? AND TDOCI.ICMSSubstBase <> 0 AND TDOCI.Quantidade <> 0 AND TDI.Emitente = ? ORDER BY NF.DataEntrada DESC, NF.NumNotaFiscal DESC", _
        sUMNF, dQtdeNF, dICMSSubstBaseNF, dICMSSubstValor, iFilialEmpresa, sProduto, dtDataRef, EMITENTE_FORNECEDOR)
    If lErro <> AD_SQL_SUCESSO Then gError 201502
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201503
    
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    dBCA = Arredonda_Moeda((dICMSSubstBaseNF / dQtdeNF) * dQtde)
    dVCA = Arredonda_Moeda((dICMSSubstValor / dQtdeNF) * dQtde)

    Call Comando_Fechar(lComando)
    
    ICMSSTCobrAnt_Obter = SUCESSO
    
    Exit Function
    
Erro_ICMSSTCobrAnt_Obter:

    ICMSSTCobrAnt_Obter = gErr

    Select Case gErr

        Case 201501
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 201502, 201503
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ICMSSTANT", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201500)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function NBS_Le(ByVal objNBS As ClassNBS) As Long

Dim lErro As Long
Dim lComando As Long
Dim tNBS As typeNBS

On Error GoTo Erro_NBS_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213570

    'Alocação de espaço no buffer
    tNBS.sCodigo = String(STRING_NBS, 0)
    tNBS.sDescricao = String(STRING_MAXIMO, 0)

    'Le a tabelaNBS
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao FROM NBS WHERE Codigo= ? ", _
                tNBS.sCodigo, tNBS.sDescricao, objNBS.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 213571

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213572

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objNBS.sCodigo = tNBS.sCodigo
    objNBS.sDescricao = tNBS.sDescricao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    NBS_Le = SUCESSO

    Exit Function

Erro_NBS_Le:

    NBS_Le = gErr

    Select Case gErr

        Case 213570
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213571, 213572
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NBS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213573)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Servico_ObterCodTribMun(ByVal sProduto As String, ByVal sCodIBGECidade As String, sCodTribMun As String) As Long
'retorna o cod de tributacao do municipio para o serviço

Dim lErro As Long
Dim lComando As Long, sCodTribMunAux As String

On Error GoTo Erro_Servico_ObterCodTribMun

    sCodTribMun = ""
    
    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 213570

    sCodTribMunAux = String(STRING_CODTRIBMUN, 0)
    lErro = Comando_Executar(lComando, "SELECT CodTribMun FROM CodTribMun CTM, Cidades C WHERE CTM.Produto = ? AND CTM.Cidade = C.Codigo AND C.CodIBGE = ?", _
        sCodTribMunAux, sProduto, sCodIBGECidade)
    If lErro <> AD_SQL_SUCESSO Then gError 213571

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213572

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    sCodTribMun = sCodTribMunAux

    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Servico_ObterCodTribMun = SUCESSO
    
    Exit Function
    
Erro_Servico_ObterCodTribMun:

    Servico_ObterCodTribMun = gErr

    Select Case gErr

        Case 213570
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213571, 213572
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NBS", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201548)

    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ImportXMLCadProd_Le(ByVal colProdutos As Collection) As Long

Dim lErro As Long, iIndice As Integer
Dim alComando(0 To 1) As Long
Dim objProd As ClassProduto
Dim objTipoProd As ClassTipoDeProduto
Dim sSQL As String
Dim sProduto As String, sCGCEmitente As String, sDescProd As String, sEAN As String
Dim sNCM As String, sUM As String, iOrigemMercadoria As Integer
Dim sProdutoBD As String, sUMBD As String, sProdAux As String

On Error GoTo Erro_ImportXMLCadProd_Le

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213830
    Next
    
    sProduto = String(STRING_MAXIMO, 0)
    sCGCEmitente = String(STRING_MAXIMO, 0)
    sDescProd = String(STRING_MAXIMO, 0)
    sEAN = String(STRING_MAXIMO, 0)
    sNCM = String(STRING_MAXIMO, 0)
    sUM = String(STRING_MAXIMO, 0)
    
    sSQL = "SELECT DISTINCT I.Produto, N.CGCEmitente, LEFT(I.DescProd,250), I.EAN, I.NCM, I.UM, I.OrigemMercadoria FROM ImportNFeXml As N, ImportNFeItensXml As I"
    sSQL = sSQL & " WHERE N.ChvNFe = I.ChvNFe AND N.NumIntNF = 0 AND"
    sSQL = sSQL & " NOT EXISTS (SELECT * FROM ImportNFeXml As X, ImportNFeItensXml As Y"
    sSQL = sSQL & " WHERE X.ChvNFe = Y.ChvNFe AND N.CGCEmitente = X.CGCEmitente AND (I.Produto = Y.Produto OR (I.EAN = Y.EAN AND I.EAN<> ''))"
    sSQL = sSQL & " AND X.NumIntNF <> 0) ORDER BY I.Produto"
    
    lErro = Comando_Executar(alComando(0), sSQL, sProduto, sCGCEmitente, sDescProd, sEAN, sNCM, sUM, iOrigemMercadoria)
    If lErro <> AD_SQL_SUCESSO Then gError 213831
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213832
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        sProdutoBD = ""
        sUMBD = ""
        
        'GL_lMascProd
        sProdAux = Trim(sProduto)
    
        lErro = Importa_Xml_NFe_Loc_Prod2(giFilialEmpresa, sCGCEmitente, sProdAux, sEAN, sProdutoBD, sUM, sUMBD)
        'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If Len(Trim(sProdutoBD)) = 0 Then
        'Se não encontrou o produto cadastrado
        
            Set objProd = New ClassProduto
            
            objProd.sCodigo = sProdAux
            objProd.sReferencia = sProdAux
            objProd.sNomeReduzido = sProdAux
            objProd.sDescricao = left(sDescProd, STRING_PRODUTO_DESCRICAO_TELA)
            objProd.sSiglaUMTrib = sUM
            objProd.sCodigoBarras = sEAN
            objProd.sIPICodigo = sNCM
            objProd.iOrigemMercadoria = iOrigemMercadoria
            
            Set objTipoProd = New ClassTipoDeProduto
            
            objTipoProd.iTipo = 1
        
            lErro = CF("TipoDeProduto_Le", objTipoProd)
            If lErro <> SUCESSO And lErro <> 22531 Then gError ERRO_SEM_MENSAGEM
            
            If lErro = SUCESSO Then
            
                objProd.iTipo = 1
                objProd.iClasseUM = objTipoProd.iClasseUM
                objProd.sSiglaUMVenda = objTipoProd.sSiglaUMVenda
            
            End If
        
            colProdutos.Add objProd
        
        End If
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213482
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ImportXMLCadProd_Le = SUCESSO

    Exit Function

Erro_ImportXMLCadProd_Le:

    ImportXMLCadProd_Le = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 213830
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 213831 To 213832
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 213833)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Public Function DIInfo_Le_XML(ByVal objDI As ClassDIInfo) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long, iIndice As Integer
Dim alComandoAux(1 To 3) As Long
Dim sNumDI As String
Dim tImportDIXml As typeImportDIXml
Dim tImportADIXml As typeImportADIXml
Dim tImportMADIXml As typeImportMADIXml
Dim objADI As ClassAdicaoDI
Dim objIADI As ClassItemAdicaoDI
Dim objProd As ClassProduto
Dim sTabela As String
Dim colMoedas As New Collection, iMoeda As Integer
Dim objMoeda1 As New ClassMoedas, objMoeda2 As New ClassMoedas
Dim dCambioMoeda1 As Double, dCambioMoeda2 As Double
Dim objIPICodigo As ClassClassificacaoFiscal
Dim dFator As Double

On Error GoTo Erro_DIInfo_Le_XML

    Call Formata_String_Numero(objDI.sNumero, sNumDI)
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 213867
    Next
    
    For iIndice = LBound(alComandoAux) To UBound(alComandoAux)
        alComandoAux(iIndice) = Comando_Abrir()
        If alComandoAux(iIndice) = 0 Then gError 213868
    Next
    
    'Alocação de espaço no buffer
    tImportDIXml.sNumeroDI = String(STRING_MAXIMO, 0)
    tImportDIXml.sOperacaoDesc = String(STRING_MAXIMO, 0)
    tImportDIXml.sPaisProcedenciaNome = String(STRING_MAXIMO, 0)
    tImportDIXml.sUrfEntradaNome = String(STRING_MAXIMO, 0)
    tImportDIXml.sEmbarqueLocal = String(STRING_MAXIMO, 0)
    tImportDIXml.sFreteMoedaNome = String(STRING_MAXIMO, 0)
    tImportDIXml.sModDespachoNome = String(STRING_MAXIMO, 0)
    tImportDIXml.sSeguroMoedaNome = String(STRING_MAXIMO, 0)
    tImportDIXml.sUrfDespachoNome = String(STRING_MAXIMO, 0)
    tImportDIXml.sViaTransporteMultimodal = String(STRING_MAXIMO, 0)
    tImportDIXml.sViaTransporteNome = String(STRING_MAXIMO, 0)
    tImportDIXml.sNomeTransportador = String(STRING_MAXIMO, 0)
    tImportDIXml.sPaisTransportadorNome = String(STRING_MAXIMO, 0)
    tImportDIXml.sRecintoAduaneiroNome = String(STRING_MAXIMO, 0)

    sTabela = "ImportDIXml"
    'Le a tabela ImportDIXml
    lErro = Comando_Executar(alComando(0), "SELECT NumeroDI, PesoBruto, PesoLiquido, DataDesembaraco, DataRegistro, DataChegada, " & _
                "OperacaoCod, OperacaoDesc, PaisProcedenciaCod, PaisProcedenciaNome, UrfEntradaCod, UrfEntradaNome, DataEmbarque, " & _
                "EmbarqueLocal, FreteCollect, FreteEmTerritorioNacional, FreteMoeda, FreteMoedaNome, FretePrepaid, FreteTotalDolares, " & _
                "FreteTotalMoeda, FreteTotalReais, DescargaTotalDolares, DescargaTotalReais, EmbarqueTotalDolares, EmbarqueTotalReais, ModDespachoCod, " & _
                "ModDespachoNome, SeguroMoeda, SeguroMoedaNome, SeguroTotalDolares, SeguroTotalMoeda, SeguroTotalReais, TotalAdicoes, " & _
                "UrfDespachoCod, UrfDespachoNome, ViaTransporteCod, ViaTransporteMultimodal, ViaTransporteNome, NomeTransportador, PaisTransportadorCod, " & _
                "PaisTransportadorNome, RecintoAduaneiroCod, RecintoAduaneiroNome FROM ImportDIXml WHERE NumeroDI= ? ", _
                tImportDIXml.sNumeroDI, tImportDIXml.dPesoBruto, tImportDIXml.dPesoLiquido, tImportDIXml.dtDataDesembaraco, _
                tImportDIXml.dtDataRegistro, tImportDIXml.dtDataChegada, tImportDIXml.lOperacaoCod, tImportDIXml.sOperacaoDesc, tImportDIXml.lPaisProcedenciaCod, _
                tImportDIXml.sPaisProcedenciaNome, tImportDIXml.lUrfEntradaCod, tImportDIXml.sUrfEntradaNome, tImportDIXml.dtDataEmbarque, tImportDIXml.sEmbarqueLocal, _
                tImportDIXml.dFreteCollect, tImportDIXml.dFreteEmTerritorioNacional, tImportDIXml.lFreteMoeda, tImportDIXml.sFreteMoedaNome, tImportDIXml.dFretePrepaid, _
                tImportDIXml.dFreteTotalDolares, tImportDIXml.dFreteTotalMoeda, tImportDIXml.dFreteTotalReais, tImportDIXml.dDescargaTotalDolares, tImportDIXml.dDescargaTotalReais, _
                tImportDIXml.dEmbarqueTotalDolares, tImportDIXml.dEmbarqueTotalReais, tImportDIXml.lModDespachoCod, tImportDIXml.sModDespachoNome, tImportDIXml.lSeguroMoeda, _
                tImportDIXml.sSeguroMoedaNome, tImportDIXml.dSeguroTotalDolares, tImportDIXml.dSeguroTotalMoeda, tImportDIXml.dSeguroTotalReais, tImportDIXml.lTotalAdicoes, _
                tImportDIXml.lUrfDespachoCod, tImportDIXml.sUrfDespachoNome, tImportDIXml.lViaTransporteCod, tImportDIXml.sViaTransporteMultimodal, tImportDIXml.sViaTransporteNome, _
                tImportDIXml.sNomeTransportador, tImportDIXml.lPaisTransportadorCod, tImportDIXml.sPaisTransportadorNome, _
                tImportDIXml.lRecintoAduaneiroCod, tImportDIXml.sRecintoAduaneiroNome, _
                sNumDI)
    If lErro <> AD_SQL_SUCESSO Then gError 213869

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213870

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    
    lErro = DIInfo_Le_XML_Obtem_Moeda(tImportDIXml.lFreteMoeda, colMoedas, iMoeda, objMoeda1, objMoeda2)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objDI.dtData = tImportDIXml.dtDataRegistro
    objDI.iMoedaFrete = iMoeda
    objDI.dValorFreteInternacMoeda = tImportDIXml.dFreteTotalMoeda
    objDI.dValorFreteInternacEmReal = tImportDIXml.dFreteTotalReais
    
    If objDI.dValorFreteInternacMoeda > 0 Then
        If iMoeda = 1 Then
            dCambioMoeda1 = objDI.dValorFreteInternacEmReal / objDI.dValorFreteInternacMoeda
        ElseIf iMoeda = 2 Then
            dCambioMoeda2 = objDI.dValorFreteInternacEmReal / objDI.dValorFreteInternacMoeda
        End If
    End If
    
    lErro = DIInfo_Le_XML_Obtem_Moeda(tImportDIXml.lSeguroMoeda, colMoedas, iMoeda, objMoeda1, objMoeda2)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objDI.iMoedaSeguro = iMoeda
    objDI.dValorSeguroInternacMoeda = tImportDIXml.dSeguroTotalMoeda
    objDI.dValorSeguroInternacEmReal = tImportDIXml.dSeguroTotalReais
    
    If objDI.dValorSeguroInternacMoeda > 0 Then
        If iMoeda = 1 And dCambioMoeda1 = 0 Then
            dCambioMoeda1 = objDI.dValorSeguroInternacEmReal / objDI.dValorSeguroInternacMoeda
        ElseIf iMoeda = 2 And dCambioMoeda2 = 0 Then
            dCambioMoeda2 = objDI.dValorSeguroInternacEmReal / objDI.dValorSeguroInternacMoeda
        End If
    End If
    
    objDI.dPesoBrutoKG = tImportDIXml.dPesoBruto
    objDI.dPesoLiqKG = tImportDIXml.dPesoLiquido
    
    objDI.dtDataDesembaraco = tImportDIXml.dtDataDesembaraco
    objDI.sLocalDesembaraco = tImportDIXml.sRecintoAduaneiroNome
    objDI.iViaTransp = tImportDIXml.lViaTransporteCod
    objDI.iIntermedio = tImportDIXml.lOperacaoCod
        
    'Alocação de espaço no buffer
    tImportADIXml.sCofinsRTNome = String(STRING_MAXIMO, 0)
    tImportADIXml.sPisRTNome = String(STRING_MAXIMO, 0)
    tImportADIXml.sCondVendMoedaCodNome = String(STRING_MAXIMO, 0)
    tImportADIXml.sAplicacao = String(STRING_MAXIMO, 0)
    tImportADIXml.sNCM = String(STRING_MAXIMO, 0)
    tImportADIXml.sNCMNome = String(STRING_MAXIMO, 0)
    tImportADIXml.sDCRID = String(STRING_MAXIMO, 0)
    tImportADIXml.sNomeFornecedor = String(STRING_MAXIMO, 0)
    tImportADIXml.sFreteMoedaNome = String(STRING_MAXIMO, 0)
    tImportADIXml.sIIRTNome = String(STRING_MAXIMO, 0)
    tImportADIXml.sIPIRTNome = String(STRING_MAXIMO, 0)

    sTabela = "ImportADIXml"
    'Le a tabela ImportADIXml
    lErro = Comando_Executar(alComando(1), "SELECT NumeroAdicao, NumeroLI, CofinsPercRed, CofinsBC, CofinsRT, " & _
                "CofinsRTNome, CofinsAliq, CofinsAliqQtde, CofinsAliqVlr, CofinsReduzida, CofinsValorDevido, CofinsValorRecolher, " & _
                "PisPercRed, PisBC, PisRT, PisRTNome, PisAliq, PisAliqQtde, PisAliqVlr, " & _
                "PisReduzida, PisValorDevido, PisValorRecolher, CondVendMoedaCod, CondVendMoedaCodNome, CondVendValorMoeda, CondVendValorReais, " & _
                "Aplicacao, NCM, NCMNome, PesoLiquido, DCRReducao, DCRID, DCRValorDevido, " & _
                "DCRValorDolar, DCRValorReal, DCRValorRecolher, NomeFornecedor, FreteMoedaCod, FreteMoedaNome, FreteValorMoeda, " & _
                "FreteValorReal, IIBC, IIRT, IIRTNome, IIAliq, IIPercReducao, IIAliqReduzida, " & _
                "IIValorCalc, IIValorDevido, IIValorRecolher, IIValorReduzido, IPIRT, IPIRTNome, IPIAliq, " & _
                "IPIAliqEsp, IPIAliqEspVlr, IPIAliqRed, IPIValorDevido, IPIValorRecolher, SeguroMoeda, SeguroValorMoeda, " & _
                "SeguroValorReais, FreteInternacionalValorReais, SeguroInternacionalValorReais, CondVendaValorTotal FROM ImportADIXml WHERE NumeroDI= ? ORDER BY NumeroAdicao ", _
                tImportADIXml.lNumeroAdicao, tImportADIXml.lNumeroLI, tImportADIXml.dCofinsPercRed, _
                tImportADIXml.dCofinsBC, tImportADIXml.lCofinsRT, tImportADIXml.sCofinsRTNome, tImportADIXml.dCofinsAliq, tImportADIXml.dCofinsAliqQtde, _
                tImportADIXml.dCofinsAliqVlr, tImportADIXml.dCofinsReduzida, tImportADIXml.dCofinsValorDevido, tImportADIXml.dCofinsValorRecolher, tImportADIXml.dPisPercRed, _
                tImportADIXml.dPisBC, tImportADIXml.lPisRT, tImportADIXml.sPisRTNome, tImportADIXml.dPisAliq, tImportADIXml.dPisAliqQtde, _
                tImportADIXml.dPisAliqVlr, tImportADIXml.dPisReduzida, tImportADIXml.dPisValorDevido, tImportADIXml.dPisValorRecolher, tImportADIXml.lCondVendMoedaCod, _
                tImportADIXml.sCondVendMoedaCodNome, tImportADIXml.dCondVendValorMoeda, tImportADIXml.dCondVendValorReais, tImportADIXml.sAplicacao, tImportADIXml.sNCM, _
                tImportADIXml.sNCMNome, tImportADIXml.dPesoLiquido, tImportADIXml.dDCRReducao, tImportADIXml.sDCRID, tImportADIXml.dDCRValorDevido, _
                tImportADIXml.dDCRValorDolar, tImportADIXml.dDCRValorReal, tImportADIXml.dDCRValorRecolher, tImportADIXml.sNomeFornecedor, tImportADIXml.lFreteMoedaCod, _
                tImportADIXml.sFreteMoedaNome, tImportADIXml.dFreteValorMoeda, tImportADIXml.dFreteValorReal, tImportADIXml.dIIBC, tImportADIXml.lIIRT, _
                tImportADIXml.sIIRTNome, tImportADIXml.dIIAliq, tImportADIXml.dIIPercReducao, tImportADIXml.dIIAliqReduzida, tImportADIXml.dIIValorCalc, _
                tImportADIXml.dIIValorDevido, tImportADIXml.dIIValorRecolher, tImportADIXml.dIIValorReduzido, tImportADIXml.lIPIRT, tImportADIXml.sIPIRTNome, _
                tImportADIXml.dIPIAliq, tImportADIXml.dIPIAliqEsp, tImportADIXml.dIPIAliqEspVlr, tImportADIXml.dIPIAliqRed, tImportADIXml.dIPIValorDevido, _
                tImportADIXml.dIPIValorRecolher, tImportADIXml.lSeguroMoeda, tImportADIXml.dSeguroValorMoeda, tImportADIXml.dSeguroValorReais, tImportADIXml.dFreteInternacionalValorReais, _
                tImportADIXml.dSeguroInternacionalValorReais, tImportADIXml.dCondVendaValorTotal, _
                sNumDI)
    If lErro <> AD_SQL_SUCESSO Then gError 213871

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213872

    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objADI = New ClassAdicaoDI
           
        objADI.sIPICodigo = tImportADIXml.sNCM
        objADI.sDescricao = tImportADIXml.sNCMNome
        objADI.dCOFINSBase = tImportADIXml.dCofinsBC
        objADI.dCOFINSAliquota = tImportADIXml.dCofinsAliq / 100
        objADI.dCOFINSValor = tImportADIXml.dCofinsValorDevido
        objADI.dIIAliquota = tImportADIXml.dIIAliq / 100
        objADI.dIIValor = tImportADIXml.dIIValorDevido
        objADI.dIPIAliquota = tImportADIXml.dIPIAliq / 100
        objADI.dIPIValor = tImportADIXml.dIPIValorDevido
        objADI.dPISBase = tImportADIXml.dPisBC
        objADI.dPISAliquota = tImportADIXml.dPisAliq / 100
        objADI.dPISValor = tImportADIXml.dPisValorDevido
        objADI.iSeq = tImportADIXml.lNumeroAdicao
        
        Set objIPICodigo = New ClassClassificacaoFiscal
        
        objIPICodigo.sCodigo = objADI.sIPICodigo
        
        lErro = CF("ClassificacaoFiscal_Le", objIPICodigo)
        If lErro <> SUCESSO And lErro <> 123494 Then gError ERRO_SEM_MENSAGEM
    
        If lErro = 123494 Then gError 196692
        
        objADI.dICMSAliquota = objIPICodigo.dICMSAliquota
        
        lErro = DIInfo_Le_XML_Obtem_Moeda(tImportADIXml.lCondVendMoedaCod, colMoedas, iMoeda, objMoeda1, objMoeda2)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objDI.iMoedaItens = iMoeda
        objDI.iMoedaMercadoria = iMoeda
        
        If tImportADIXml.dCondVendValorMoeda > 0 Then
            If iMoeda = 1 And dCambioMoeda1 = 0 Then
                dCambioMoeda1 = tImportADIXml.dCondVendValorReais / tImportADIXml.dCondVendValorMoeda
            ElseIf iMoeda = 2 And dCambioMoeda2 = 0 Then
                dCambioMoeda2 = tImportADIXml.dCondVendValorReais / tImportADIXml.dCondVendValorMoeda
            End If
        End If
        
        'Alocação de espaço no buffer
        tImportMADIXml.sDescricao = String(STRING_MAXIMO, 0)
        tImportMADIXml.sUM = String(STRING_MAXIMO, 0)
    
        sTabela = "ImportMADIXml"
        'Le a tabela ImportMADIXml
        lErro = Comando_Executar(alComando(2), "SELECT Seq, Descricao, UM, Quantidade, ValorUnitario FROM ImportMADIXml WHERE NumeroDI= ? AND NumeroAdicao = ? ORDER BY Seq", _
                    tImportMADIXml.lSeq, tImportMADIXml.sDescricao, tImportMADIXml.sUM, tImportMADIXml.dQuantidade, tImportMADIXml.dValorUnitario, _
                    sNumDI, tImportADIXml.lNumeroAdicao)
        If lErro <> AD_SQL_SUCESSO Then gError 213873
    
        'Busca Primeiro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213874
    
        Do While lErro <> AD_SQL_SEM_DADOS
        
            Set objIADI = New ClassItemAdicaoDI
            Set objProd = New ClassProduto
            
            objIADI.sDescricao = tImportMADIXml.sDescricao
            objIADI.iAdicao = tImportADIXml.lNumeroAdicao
            objIADI.iSeq = tImportMADIXml.lSeq
            objIADI.dQuantidade = tImportMADIXml.dQuantidade
            objIADI.dValorUnitFOBNaMoeda = tImportMADIXml.dValorUnitario
            
            lErro = DIInfo_Le_XML_Loc_Prod(tImportADIXml.sNCM, tImportMADIXml.sDescricao, objProd, alComandoAux)
            If lErro = SUCESSO Then
                objIADI.sProduto = objProd.sCodigo
                objIADI.sUM = objProd.sSiglaUMCompra
                
                lErro = CF("UM_Conversao_Trans", objProd.iClasseUM, objIADI.sUM, objProd.sSiglaUMVenda, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                            
                objIADI.dPesoBruto = objProd.dPesoBruto * objIADI.dQuantidade * dFator
                objIADI.dPesoLiq = objProd.dPesoLiq * objIADI.dQuantidade * dFator
                
            End If

            objADI.colItensAdicaoDI.Add objIADI
        
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213875
        
        Loop
    
        objDI.colAdicoesDI.Add objADI
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213876
    
    Loop
    
    objDI.iMoeda1 = objMoeda1.iCodigo
    objDI.iMoeda2 = objMoeda2.iCodigo
    objDI.dTaxaMoeda1 = Arredonda_Moeda(dCambioMoeda1, 6)
    objDI.dTaxaMoeda2 = Arredonda_Moeda(dCambioMoeda2, 6)
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    For iIndice = LBound(alComandoAux) To UBound(alComandoAux)
        Call Comando_Fechar(alComandoAux(iIndice))
    Next
    
    DIInfo_Le_XML = SUCESSO

    Exit Function

Erro_DIInfo_Le_XML:

    DIInfo_Le_XML = gErr

    Select Case gErr

        Case 213867, 213868
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 213869 To 213876
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, sTabela)
            
        Case ERRO_LEITURA_SEM_DADOS
        
        Case ERRO_SEM_MENSAGEM

        Case 196692
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSIFICACAOFISCAL_NAO_EXISTENTE", gErr, objIPICodigo.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213877)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    For iIndice = LBound(alComandoAux) To UBound(alComandoAux)
        Call Comando_Fechar(alComandoAux(iIndice))
    Next
    
    Exit Function

End Function

Private Function DIInfo_Le_XML_Loc_Prod(ByVal sNCM As String, ByVal sDescProd As String, ByVal objProd As ClassProduto, alComando() As Long) As Long

Dim lErro As Long, bAchou As Boolean
Dim sProduto As String, sTextoAux As String
Dim iPos As Integer, iIndice1 As Integer, iIndice2 As Integer
Dim sCampo As String, lComando As Long, sProdBD As String, sProdBDAux As String

On Error GoTo Erro_DIInfo_Le_XML_Loc_Prod

    bAchou = False

    iPos = InStr(1, sDescProd, "REF:")
    If iPos <> 0 Then
        sTextoAux = Trim(Mid(sDescProd, iPos + 4))
        iPos = InStr(1, sTextoAux, " ")
        If iPos <> 0 Then
            sProduto = left(sTextoAux, iPos - 1)
        Else
            sProduto = sTextoAux
        End If
    Else
        iPos = InStr(1, sDescProd, " ")
        If iPos <> 0 Then
            sProduto = left(sDescProd, iPos - 1)
        End If
    End If
    
    'Se conseguiu extrair o código do produto
    If Len(Trim(sProduto)) > 0 Then
    
        '1o testa a igualdade depois com like %
        For iIndice1 = 1 To 2
        
            If iIndice1 = 2 Then sProduto = sProduto & "%"
        
            'Procura 1o pelo código, depois pela referência e por último pelo nomereduzido
            For iIndice2 = 1 To 3
            
                Select Case iIndice2
                    Case 1
                        sCampo = "Codigo"
                    Case 2
                        sCampo = "Referencia"
                    Case 3
                        sCampo = "NomeReduzido"
                End Select
                lComando = alComando(iIndice2)
                
                sProdBDAux = String(STRING_PRODUTO, 0)
        
                lErro = Comando_Executar(lComando, "SELECT Codigo FROM Produtos WHERE IPICodigo = ? AND " & sCampo & " LIKE ? ", sProdBDAux, sNCM, sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 213878
            
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213879
                
                'Se achou
                If lErro = AD_SQL_SUCESSO Then
                
                    sProdBD = sProdBDAux
                    
                    lErro = Comando_BuscarProximo(lComando)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213880
                    
                    'Se só um produto atendeu ao critério então é esse
                    If lErro <> AD_SQL_SUCESSO Then
                        bAchou = True
                        Exit For
                    End If
                
                End If
        
        
            Next
            
            If bAchou Then Exit For
        
        Next
        
    End If
    
    If Not bAchou Then gError ERRO_SEM_MENSAGEM
    
    objProd.sCodigo = sProdBD
    
    lErro = CF("Produto_Le", objProd)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    DIInfo_Le_XML_Loc_Prod = SUCESSO

    Exit Function

Erro_DIInfo_Le_XML_Loc_Prod:

    DIInfo_Le_XML_Loc_Prod = gErr

    Select Case gErr

        Case 213878 To 213880
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213881)

    End Select
        
    Exit Function
    
End Function

Private Function DIInfo_Le_XML_Obtem_Moeda(ByVal iCodMoedaBacen As Integer, colMoedas As Collection, iMoeda As Integer, objMoeda1 As ClassMoedas, objMoeda2 As ClassMoedas) As Long

Dim bAchou As Boolean
Dim objMoeda As ClassMoedas

On Error GoTo Erro_DIInfo_Le_XML_Obtem_Moeda

    bAchou = False
    For Each objMoeda In colMoedas
        If objMoeda.iCodBacen = iCodMoedaBacen Then
            bAchou = True
            Exit For
        End If
    Next
    If bAchou Then
        If objMoeda1.iCodigo = 0 Then
            iMoeda = 1
            objMoeda1.iCodBacen = iCodMoedaBacen
            objMoeda1.iCodigo = objMoeda.iCodigo
        ElseIf objMoeda2.iCodigo = 0 Then
            iMoeda = 2
            objMoeda2.iCodBacen = iCodMoedaBacen
            objMoeda2.iCodigo = objMoeda.iCodigo
        End If
    End If

    DIInfo_Le_XML_Obtem_Moeda = SUCESSO

    Exit Function

Erro_DIInfo_Le_XML_Obtem_Moeda:

    DIInfo_Le_XML_Obtem_Moeda = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213881)

    End Select
        
    Exit Function
    
End Function
