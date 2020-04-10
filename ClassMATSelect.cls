VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function MovimentoEstoque_Le_SaldoInicial(dQuantidadeTotal As Double, dtDataInventario As Date, dtDataInicial As Date, iFilialEmpresa As Integer, iTipoTerc As Integer, lCodTerc As Long, iFilialTerc As Integer, sProduto As String, sCampoescaninho As String, iCodEscaninho As Integer) As Long
'Esse função calcula o saldo inicial somando-se a quantidade de InventarioTercProd com o somatório
'de movimentos posteriores a esse inventário e anteriores a data inicial passada, se for passado como
'data do inventário a DATA_NULA, então soma-se todos movimentos anteriores a data inicial

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim sSelect1 As String
Dim sSelect2 As String
Dim sWhere1 As String
Dim sWhere2 As String
Dim dQuantidadeInventario As Double
Dim dQuantidade As Double

On Error GoTo Erro_MovimentoEstoque_Le_SaldoInicial

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 110263
    Next

    sSelect2 = "SELECT  SUM(M.Quantidade) FROM MovimentoEstoque AS M, NFiscal AS N, TiposMovimentoEstoque AS T "
    sWhere2 = "WHERE M.TipoNumIntDocOrigem = 2 AND M.Data < ? AND M.FilialEmpresa = ? AND M.Produto = ? AND T." & sCampoescaninho & " <> 0 AND M.TipoMov = T.Codigo AND M.NumIntDocOrigem = N.NumIntDoc "
    
    sSelect1 = sSelect2 & ", ItensNFiscal AS I "
    sWhere1 = "WHERE M.TipoNumIntDocOrigem = 1 AND M.Data < ? AND M.FilialEmpresa = ? AND M.Produto = ? AND T." & sCampoescaninho & " <> 0 AND M.TipoMov = T.Codigo AND I.NumIntDoc = M.NumIntDocOrigem AND N.NumIntDoc = I.NumIntNF "
      
    'Verifica se existe inventário válido
    If dtDataInventario <> DATA_NULA Then
    
        'Busca quantidade no inventario
        lErro = Comando_Executar(alComando(1), "SELECT QuantTotal FROM InventarioTercProd WHERE Data = ? AND FilialEmpresa = ? AND TipoTerc = ? AND CodTerc = ? AND FilialTerc = ? AND Produto = ? AND CodEscaninho = ? ", _
                            dQuantidadeInventario, dtDataInventario, iFilialEmpresa, iTipoTerc, lCodTerc, iFilialTerc, sProduto, iCodEscaninho)
        If lErro <> AD_SQL_SUCESSO Then gError 110264
    
        'Busca primeira tupla
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110265

        If lErro = AD_SQL_SEM_DADOS Then Error 110266
        
        dQuantidadeTotal = dQuantidadeInventario
    
        'Verifica se é cliente ou fornecedor
        If iTipoTerc = TIPO_TERC_CLIENTE Then
        
            'Adiciona filtro por cliente e data maior ou igual que a data do inventário
            sWhere1 = sWhere1 & " AND N.Cliente = ? AND N.FilialCli = ? AND Data >= ? "
            sWhere2 = sWhere2 & " AND N.Cliente = ? AND N.FilialCli = ? AND Data >= ? "
        Else
        
            'Adiciona filtro por fornecedor e data maior ou igual que a data do inventário
            sWhere1 = sWhere1 & " AND N.Fornecedor = ? AND N.FilialForn = ? AND Data >= ? "
            sWhere2 = sWhere2 & " AND N.Fornecedor = ? AND N.FilialForn = ? AND Data >= ? "
        End If
        
        lErro = Comando_Executar(alComando(2), sSelect1 & sWhere1 & " UNION " & sSelect2 & sWhere2, dQuantidade, dtDataInicial, iFilialEmpresa, sProduto, lCodTerc, iFilialTerc, dtDataInventario, dtDataInicial, iFilialEmpresa, sProduto, lCodTerc, iFilialTerc, dtDataInventario)
    
    Else
                
        dQuantidadeTotal = 0
    
        'Verifica se é cliente ou fornecedor
        If iTipoTerc = TIPO_TERC_CLIENTE Then
        
            'Adiciona filtro por cliente
            sWhere1 = sWhere1 & " AND N.Cliente = ? AND N.FilialCli = ? "
            sWhere2 = sWhere2 & " AND N.Cliente = ? AND N.FilialCli = ? "
        Else
        
            'Adiciona filtro por fornecedor
            sWhere1 = sWhere1 & " AND N.Fornecedor = ? AND N.FilialForn = ? "
            sWhere2 = sWhere2 & " AND N.Fornecedor = ? AND N.FilialForn = ? "
        End If
        
        lErro = Comando_Executar(alComando(2), sSelect1 & sWhere1 & " UNION " & sSelect2 & sWhere2, dQuantidade, dtDataInicial, iFilialEmpresa, sProduto, lCodTerc, iFilialTerc, dtDataInicial, iFilialEmpresa, sProduto, lCodTerc, iFilialTerc)
                            
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 110267
        
    'Busca primeira tupla
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110268
        
    Do While lErro <> AD_SQL_SEM_DADOS

        'Soma a quantidade das tuplas
        dQuantidadeTotal = dQuantidadeTotal + dQuantidade

        'Busca próxima tupla
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110269

    Loop
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    MovimentoEstoque_Le_SaldoInicial = SUCESSO
    
    Exit Function
    
Erro_MovimentoEstoque_Le_SaldoInicial:

    MovimentoEstoque_Le_SaldoInicial = gErr
    
        Select Case gErr
    
        Case 110263
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 110264 To 110266
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIOTERCPROD", gErr)
            
        Case 110267 To 110269
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
        End Select
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function InventarioTercProd_Le_MaiorDataAnterior(dtDataMaxima As Date, dtDataInicial As Date, iFilialEmpresa As Integer, iTipoTerc As Integer, lCodTerc As Long, iFilialTerc As Integer, sProduto As String, iCodEscaninho As Integer) As Long
'Essa função recebe a chave primária da tabela InventarioTercProd menos a data e uma data inicial
'e retorna o mais recente Inventário feito antes daquela data, se não encontrar retorna DATA_NULA

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_InventarioTercProd_Le_MaiorDataAnterior

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 110270

    lErro = Comando_Executar(lComando, "SELECT MAX(Data) FROM InventarioTercProd WHERE Data < ? AND FilialEmpresa = ? AND TipoTerc = ? AND CodTerc = ? AND FilialTerc = ? AND Produto = ? AND CodEscaninho = ? ", _
                            dtDataMaxima, dtDataInicial, iFilialEmpresa, iTipoTerc, lCodTerc, iFilialTerc, sProduto, iCodEscaninho)
    If lErro <> AD_SQL_SUCESSO Then gError 110271
    
    'Busca primeira tupla
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110272

    If lErro = AD_SQL_SEM_DADOS Then Error 110273
    
    Call Comando_Fechar(lComando)

    InventarioTercProd_Le_MaiorDataAnterior = SUCESSO
    
    Exit Function
    
Erro_InventarioTercProd_Le_MaiorDataAnterior:

    InventarioTercProd_Le_MaiorDataAnterior = gErr

    Select Case gErr
    
        Case 110270
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 110271 To 110273
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIOTERCPROD", gErr)
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select

    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

'Rotina criada por Daniel
'Revisão: Luiz G.F.Nogueira
'Data da inclusão no MATSelect: 18/09/01
'Pendências: não
Function ProdutoEmbalagem_Testa_Relacionamentos(ByVal objProdutoEmbalagem As ClassProdutoEmbalagem) As Long
'Testa se o Produto e a Embalagem passadas no ObjProdutoEmbalagem existem em EmbalagensItem

Dim lErro As Long
Dim lComando As Long
Dim iEmbalagem As Integer

On Error GoTo Erro_ProdutoEmbalagem_Testa_Relacionamentos

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 95136
        
    'Lê na view ProdutoEmbalagem_Integridade a embalagem e o produto
    lErro = Comando_Executar(lComando, "SELECT Embalagem FROM EmbalagensItem WHERE Produto = ? AND Embalagem = ?", iEmbalagem, objProdutoEmbalagem.sProduto, objProdutoEmbalagem.iEmbalagem)
    If lErro <> AD_SQL_SUCESSO Then gError 95133
        
    'Buscar o Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 95134

    'Se encontrou => erro
    If lErro = AD_SQL_SUCESSO Then gError 95135

    Call Comando_Fechar(lComando)
        
    ProdutoEmbalagem_Testa_Relacionamentos = SUCESSO
    
    Exit Function

Erro_ProdutoEmbalagem_Testa_Relacionamentos:

    ProdutoEmbalagem_Testa_Relacionamentos = gErr
    
    Select Case gErr
    
        Case 95136
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 95133, 95134
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGENSITEM", gErr)
            
        Case 95135
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOEMBALAGEM_RELACIONADO_EMBALAGENSITEM", gErr, objProdutoEmbalagem.iEmbalagem, objProdutoEmbalagem.sProduto)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151954)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Produto_Le_Todos(ByVal colProduto As Collection) As Long
'Lê todos os produtos e coloca-os em colProduto

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto
Dim objProduto As ClassProduto

On Error GoTo Erro_Produto_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28031

    With tProduto
    
        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

        'Lê produtos ordenados por Codigo (apenas os atributos que interessam)
        lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Nivel, Gerencial FROM Produtos ORDER BY Codigo", .sCodigo, .sDescricao, .iNivel, .iGerencial)
        If lErro <> AD_SQL_SUCESSO Then Error 28032

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28033
    
        Do While lErro = AD_SQL_SUCESSO
    
            Set objProduto = New ClassProduto
    
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
            objProduto.iNivel = .iNivel
            objProduto.iGerencial = .iGerencial
        
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28034

        Loop

    End With
    
    lErro = Comando_Fechar(lComando)

    Produto_Le_Todos = SUCESSO

    Exit Function

Erro_Produto_Le_Todos:

    Produto_Le_Todos = Err

    Select Case Err

        Case 28031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28032, 28033, 28034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151955)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
Function Produtos_Le_Reserva(ByVal colProduto As Collection) As Long
'Lê Codigo e Descricao de todos os produtos ATIVOS que fazem RESERVA e coloca-os em colProduto

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto
Dim objProduto As ClassProduto

On Error GoTo Erro_Produtos_Le_Reserva

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 25044

    With tProduto
    
        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
 
        'Lê Codigo e Descricao de produtos ordenados por Codigo
        lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao FROM Produtos, ProdutosFilial WHERE Produtos.Codigo = ProdutosFilial.Produto AND ProdutosFilial.FilialEmpresa = ? AND (ControleEstoque=? OR Gerencial=?) AND Ativo=? ORDER BY Codigo", .sCodigo, .sDescricao, giFilialEmpresa, PRODUTO_CONTROLE_RESERVA, GERENCIAL, Ativo)
        If lErro <> AD_SQL_SUCESSO Then Error 25045

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25046
    
        Do While lErro = AD_SQL_SUCESSO
        
            Set objProduto = New ClassProduto
    
            objProduto.sCodigo = .sCodigo
            objProduto.sDescricao = .sDescricao
        
            colProduto.Add objProduto

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25047

        Loop

    End With
    
    lErro = Comando_Fechar(lComando)

    Produtos_Le_Reserva = SUCESSO

    Exit Function

Erro_Produtos_Le_Reserva:

    Produtos_Le_Reserva = Err

    Select Case Err

        Case 25044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25045, 25046, 25047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151956)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function UnidadesDeMedidas_Le_ClasseUM(ByVal objClasseUM As ClassClasseUM, ByVal colSiglas As Collection) As Long
'Carrega a colSiglas com objetos do tipo ClassUnidadeDeMedida correspondentes a objClasseUM

Dim lErro As Long
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim iIndice As Integer
Dim lComando As Long
Dim iClasse As Integer
Dim sSigla As String
Dim sNome As String
Dim dQuantidade As Double
Dim sSiglaUMBase As String

On Error GoTo Erro_UnidadesDeMedidas_Le_ClasseUM

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22407

    sSigla = String(STRING_UM_SIGLA, 0)
    sNome = String(STRING_CLASSEUM_NOME, 0)
    sSiglaUMBase = String(STRING_UM_SIGLA, 0)

    'Executar comando SQL
    lErro = Comando_Executar(lComando, "SELECT Classe, Sigla, Nome, Quantidade, SiglaUMBase FROM UnidadesDeMedida WHERE Classe = ? ORDER BY Sigla", iClasse, sSigla, sNome, dQuantidade, sSiglaUMBase, objClasseUM.iClasse)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22408

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22409

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 22539

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objUnidadeDeMedida = New ClassUnidadeDeMedida

        objUnidadeDeMedida.iClasse = iClasse
        objUnidadeDeMedida.sSigla = sSigla
        objUnidadeDeMedida.sNome = sNome
        objUnidadeDeMedida.dQuantidade = dQuantidade
        objUnidadeDeMedida.sSiglaUMBase = sSiglaUMBase

        colSiglas.Add objUnidadeDeMedida

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22410

    Loop

    UnidadesDeMedidas_Le_ClasseUM = SUCESSO

    Call Comando_Fechar(lComando)

    Exit Function

Erro_UnidadesDeMedidas_Le_ClasseUM:

    UnidadesDeMedidas_Le_ClasseUM = Err

    Select Case Err

        Case 22539

        Case 22407
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22408, 22409, 22410
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_UNIDADESDEMEDIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151957)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function EstoqueProduto_Le(ByVal objEstoqueProduto As ClassEstoqueProduto) As Long
'Le as informacoes sobre o produto passado em objEstoqueProduto na tabela EstoqueProduto

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueProduto As typeEstoqueProduto

On Error GoTo Erro_EstoqueProduto_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 21307

    tEstoqueProduto.sLocalizacaoFisica = String(STRING_ESTOQUEPRODUTO_LOCALIZACAOFISICA, 0)
    tEstoqueProduto.sContaContabil = String(STRING_CONTA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT ContaContabil, QuantDispNossa, LocalizacaoFisica, QuantReservada, DataInventario, QuantidadeInicial, SaldoInicial, DataInicial, QuantEmpenhada, QuantPedida, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantOP, QuantBenef3, QuantBenef, ValorConsig3, ValorConsig, ValorDemo3, ValorDemo, ValorConserto3, ValorConserto, ValorOutras3, ValorOutras, ValorBenef3, ValorBenef, QuantReservadaConsig, QuantInicialConsig3, QuantInicialConsig, QuantInicialDemo3, QuantInicialDemo, QuantInicialConserto3, QuantInicialConserto, QuantInicialOutras3, QuantInicialOutras, QuantInicialBenef, QuantInicialBenef3, ValorInicialConsig3, ValorInicialConsig, ValorInicialDemo3, ValorInicialDemo, ValorInicialConserto3, ValorInicialConserto, ValorInicialOutras3, ValorInicialOutras, ValorInicialBenef, ValorInicialBenef3 FROM EstoqueProduto WHERE Produto = ? AND Almoxarifado = ?", _
            tEstoqueProduto.sContaContabil, tEstoqueProduto.dQuantDispNossa, tEstoqueProduto.sLocalizacaoFisica, tEstoqueProduto.dQuantReservada, tEstoqueProduto.dtDataInventario, tEstoqueProduto.dQuantidadeInicial, tEstoqueProduto.dSaldoInicial, tEstoqueProduto.dtDataInicial, tEstoqueProduto.dQuantEmpenhada, tEstoqueProduto.dQuantPedido, tEstoqueProduto.dQuantRecIndl, tEstoqueProduto.dQuantInd, tEstoqueProduto.dQuantDefeituosa, tEstoqueProduto.dQuantConsig3, tEstoqueProduto.dQuantConsig, tEstoqueProduto.dQuantDemo3, tEstoqueProduto.dQuantDemo, tEstoqueProduto.dQuantConserto3, tEstoqueProduto.dQuantConserto, tEstoqueProduto.dQuantOutras3, tEstoqueProduto.dQuantOutras, tEstoqueProduto.dQuantOP, tEstoqueProduto.dQuantBenef3, tEstoqueProduto.dQuantBenef, tEstoqueProduto.dValorConsig3, tEstoqueProduto.dValorConsig, tEstoqueProduto.dValorDemo3, tEstoqueProduto.dValorDemo, tEstoqueProduto.dValorConserto3, tEstoqueProduto.dValorConserto, tEstoqueProduto.dValorOutras3, _
            tEstoqueProduto.dValorOutras, tEstoqueProduto.dValorBenef3, tEstoqueProduto.dValorBenef, tEstoqueProduto.dQuantReservadaConsig, tEstoqueProduto.dQuantInicialConsig3, tEstoqueProduto.dQuantInicialConsig, tEstoqueProduto.dQuantInicialDemo3, tEstoqueProduto.dQuantInicialDemo, tEstoqueProduto.dQuantInicialConserto3, tEstoqueProduto.dQuantInicialConserto, tEstoqueProduto.dQuantInicialOutras3, tEstoqueProduto.dQuantInicialOutras, tEstoqueProduto.dQuantInicialBenef, tEstoqueProduto.dQuantInicialBenef3, tEstoqueProduto.dValorInicialConsig3, tEstoqueProduto.dValorInicialConsig, tEstoqueProduto.dValorInicialDemo3, tEstoqueProduto.dValorInicialDemo, tEstoqueProduto.dValorInicialConserto3, tEstoqueProduto.dValorInicialConserto, tEstoqueProduto.dValorInicialOutras3, tEstoqueProduto.dValorInicialOutras, tEstoqueProduto.dValorInicialBenef, tEstoqueProduto.dValorInicialBenef3, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then Error 21304

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21305

    If lErro = AD_SQL_SEM_DADOS Then Error 21306
    
    objEstoqueProduto.sContaContabil = tEstoqueProduto.sContaContabil
    objEstoqueProduto.dQuantDispNossa = tEstoqueProduto.dQuantDispNossa
    objEstoqueProduto.sLocalizacaoFisica = tEstoqueProduto.sLocalizacaoFisica
    objEstoqueProduto.dQuantReservada = tEstoqueProduto.dQuantReservada
    objEstoqueProduto.dtDataInventario = tEstoqueProduto.dtDataInventario
    objEstoqueProduto.dQuantidadeInicial = tEstoqueProduto.dQuantidadeInicial
    objEstoqueProduto.dSaldoInicial = tEstoqueProduto.dSaldoInicial
    objEstoqueProduto.dtDataInicial = tEstoqueProduto.dtDataInicial
    objEstoqueProduto.dQuantEmpenhada = tEstoqueProduto.dQuantEmpenhada
    objEstoqueProduto.dQuantPedido = tEstoqueProduto.dQuantPedido
    objEstoqueProduto.dQuantRecIndl = tEstoqueProduto.dQuantRecIndl
    objEstoqueProduto.dQuantInd = tEstoqueProduto.dQuantInd
    objEstoqueProduto.dQuantDefeituosa = tEstoqueProduto.dQuantDefeituosa
    objEstoqueProduto.dQuantConsig3 = tEstoqueProduto.dQuantConsig3
    objEstoqueProduto.dQuantConsig = tEstoqueProduto.dQuantConsig
    objEstoqueProduto.dQuantDemo3 = tEstoqueProduto.dQuantDemo3
    objEstoqueProduto.dQuantDemo = tEstoqueProduto.dQuantDemo
    objEstoqueProduto.dQuantConserto3 = tEstoqueProduto.dQuantConserto3
    objEstoqueProduto.dQuantConserto = tEstoqueProduto.dQuantConserto
    objEstoqueProduto.dQuantOutras3 = tEstoqueProduto.dQuantOutras3
    objEstoqueProduto.dQuantOutras = tEstoqueProduto.dQuantOutras
    objEstoqueProduto.dQuantOP = tEstoqueProduto.dQuantOP
    objEstoqueProduto.dQuantBenef = tEstoqueProduto.dQuantBenef
    objEstoqueProduto.dQuantBenef3 = tEstoqueProduto.dQuantBenef3
    objEstoqueProduto.dQuantReservadaConsig = tEstoqueProduto.dQuantReservadaConsig
    objEstoqueProduto.dValorConsig3 = tEstoqueProduto.dValorConsig3
    objEstoqueProduto.dValorConsig = tEstoqueProduto.dValorConsig
    objEstoqueProduto.dValorDemo3 = tEstoqueProduto.dValorDemo3
    objEstoqueProduto.dValorDemo = tEstoqueProduto.dValorDemo
    objEstoqueProduto.dValorConserto3 = tEstoqueProduto.dValorConserto3
    objEstoqueProduto.dValorConserto = tEstoqueProduto.dValorConserto
    objEstoqueProduto.dValorOutras3 = tEstoqueProduto.dValorOutras3
    objEstoqueProduto.dValorOutras = tEstoqueProduto.dValorOutras
    objEstoqueProduto.dValorBenef3 = tEstoqueProduto.dValorBenef3
    objEstoqueProduto.dValorBenef = tEstoqueProduto.dValorBenef
    objEstoqueProduto.dQuantInicialConsig3 = tEstoqueProduto.dQuantInicialConsig3
    objEstoqueProduto.dQuantInicialConsig = tEstoqueProduto.dQuantInicialConsig
    objEstoqueProduto.dQuantInicialDemo3 = tEstoqueProduto.dQuantInicialDemo3
    objEstoqueProduto.dQuantInicialDemo = tEstoqueProduto.dQuantInicialDemo
    objEstoqueProduto.dQuantInicialConserto3 = tEstoqueProduto.dQuantInicialConserto3
    objEstoqueProduto.dQuantInicialConserto = tEstoqueProduto.dQuantInicialConserto
    objEstoqueProduto.dQuantInicialOutras3 = tEstoqueProduto.dQuantInicialOutras3
    objEstoqueProduto.dQuantInicialOutras = tEstoqueProduto.dQuantInicialOutras
    objEstoqueProduto.dQuantInicialBenef = tEstoqueProduto.dQuantInicialBenef
    objEstoqueProduto.dQuantInicialBenef3 = tEstoqueProduto.dQuantInicialBenef3
    objEstoqueProduto.dValorInicialConsig3 = tEstoqueProduto.dValorInicialConsig3
    objEstoqueProduto.dValorInicialConsig = tEstoqueProduto.dValorInicialConsig
    objEstoqueProduto.dValorInicialDemo3 = tEstoqueProduto.dValorInicialDemo3
    objEstoqueProduto.dValorInicialDemo = tEstoqueProduto.dValorInicialDemo
    objEstoqueProduto.dValorInicialConserto3 = tEstoqueProduto.dValorInicialConserto3
    objEstoqueProduto.dValorInicialConserto = tEstoqueProduto.dValorInicialConserto
    objEstoqueProduto.dValorInicialOutras3 = tEstoqueProduto.dValorInicialOutras3
    objEstoqueProduto.dValorInicialOutras = tEstoqueProduto.dValorInicialOutras
    objEstoqueProduto.dValorInicialBenef = tEstoqueProduto.dValorInicialBenef
    objEstoqueProduto.dValorInicialBenef3 = tEstoqueProduto.dValorInicialBenef3

    'Fechamento comando
    Call Comando_Fechar(lComando)

    EstoqueProduto_Le = SUCESSO

    Exit Function

Erro_EstoqueProduto_Le:

    EstoqueProduto_Le = Err

    Select Case Err

        Case 21304, 21305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO1", Err, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 21306

        Case 21307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151958)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Almoxarifado_Le(ByVal objAlmoxarifado As ClassAlmoxarifado) As Long
'Le no BD os dados do Almoxarifado cujo código foi passado como parâmetro

Dim lComando As Long
Dim lErro As Long
Dim tAlmoxarifado As typeAlmoxarifado

On Error GoTo Erro_Almoxarifado_Le

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 25053

    With tAlmoxarifado

        'Prepara variáveis string para receber conteúdo do BD
        .sNomeReduzido = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sDescricao = String(STRING_ALMOXARIFADO_DESCRICAO, 0)
        .sContaContabil = String(STRING_CONTA, 0)
    
        'Seleciona os dados da tabela Almoxarifado
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, NomeReduzido, Descricao, Endereco, ContaContabil FROM Almoxarifado WHERE Codigo = ? ", _
            .iFilialEmpresa, .sNomeReduzido, .sDescricao, .lEndereco, .sContaContabil, objAlmoxarifado.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 25054
    
        'Tenta ler
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25055
    
        'Almoxarifado não existe ---> erro
        If lErro = AD_SQL_SEM_DADOS Then Error 25056
    
        'Passa dados lidos para objAlmoxarifado
        objAlmoxarifado.iFilialEmpresa = .iFilialEmpresa
        objAlmoxarifado.sNomeReduzido = .sNomeReduzido
        objAlmoxarifado.sDescricao = .sDescricao
        objAlmoxarifado.lEndereco = .lEndereco
        objAlmoxarifado.sContaContabil = .sContaContabil

    End With

    'Fecha comando
    Call Comando_Fechar(lComando)

    Almoxarifado_Le = SUCESSO

    Exit Function

Erro_Almoxarifado_Le:

    Almoxarifado_Le = Err

    Select Case Err

        Case 25053
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25054, 25055
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO", Err, objAlmoxarifado.iCodigo)

        Case 25056 'Não existe o Almoxarifado (sem dados)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151959)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Function Almoxarifado_Le_NomeReduzido(ByVal objAlmoxarifado As ClassAlmoxarifado) As Long
'Le todos no BD os dados do Almoxarifado cujo NomeReduzido foi passado por parâmetro

Dim lComando As Long
Dim lErro As Long
Dim tAlmoxarifado As typeAlmoxarifado

On Error GoTo Erro_Almoxarifado_Le_NomeReduzido

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 25057

    With tAlmoxarifado

        'Prepara variáveis string para receber conteúdo do BD
        .sNomeReduzido = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sDescricao = String(STRING_ALMOXARIFADO_DESCRICAO, 0)
        .sContaContabil = String(STRING_CONTA, 0)
    
        'Seleciona os dados da tabela Almoxarifado
        lErro = Comando_Executar(lComando, "SELECT Codigo, NomeReduzido, FilialEmpresa, Descricao, Endereco, ContaContabil FROM Almoxarifado WHERE NomeReduzido = ? ", _
        .iCodigo, .sNomeReduzido, .iFilialEmpresa, .sDescricao, .lEndereco, .sContaContabil, objAlmoxarifado.sNomeReduzido)
    
        If lErro <> AD_SQL_SUCESSO Then Error 25058
    
        'Tenta ler
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25059
    
        'Almoxarifado não existe ---> erro
        If lErro = AD_SQL_SEM_DADOS Then Error 25060
    
        'Passa dados lidos para objAlmoxarifado
        objAlmoxarifado.iCodigo = .iCodigo
        objAlmoxarifado.sNomeReduzido = .sNomeReduzido
        objAlmoxarifado.iFilialEmpresa = .iFilialEmpresa
        objAlmoxarifado.sDescricao = .sDescricao
        objAlmoxarifado.lEndereco = .lEndereco
        objAlmoxarifado.sContaContabil = .sContaContabil

    End With

    'Fecha comando
    lErro = Comando_Fechar(lComando)

    Almoxarifado_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_Almoxarifado_Le_NomeReduzido:

    Almoxarifado_Le_NomeReduzido = Err

    Select Case Err

        Case 25057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25058, 25059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO1", Err, objAlmoxarifado.sNomeReduzido)

        Case 25060 'Não existe o Almoxarifado (sem dados)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151960)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function AlmoxarifadoPadrao_Le(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, iAlmoxarifadoPadrao As Integer) As Long
'Lê Almoxarifado Padrão do Produto da iFilialEmpresa

Dim lErro As Long
Dim iAlmoxarifado As Integer
Dim lComando As Long
Dim lComando2 As Long
Dim iFilialEmpresa1 As Integer

On Error GoTo Erro_AlmoxarifadoPadrao_Le

    iAlmoxarifado = 0

    lComando = 0
    lComando2 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23793

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 30087

    iFilialEmpresa1 = iFilialEmpresa
    lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa1)
    If lErro <> SUCESSO Then Error 30899
                
    'Tenta selecinar a partir do ProdutoFilial
    lErro = Comando_Executar(lComando, "SELECT Almoxarifado FROM ProdutosFilial WHERE FilialEmpresa=? AND Produto=?", iAlmoxarifado, iFilialEmpresa1, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 23794

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23795

    'Caso não obtenha sucesso
    If lErro = AD_SQL_SEM_DADOS Then Error 23796

    iAlmoxarifadoPadrao = iAlmoxarifado

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    AlmoxarifadoPadrao_Le = SUCESSO

    Exit Function

Erro_AlmoxarifadoPadrao_Le:

    AlmoxarifadoPadrao_Le = Err

    Select Case Err

        Case 23793, 30087
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23794, 23795
            lErro = Rotina_Erro(vbOKOnly, "0", Err)

        Case 23796, 30899
                'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151961)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function


Function EstoquesProduto_Le(ByVal sProduto As String, ByVal colEstoque As colEstoqueProduto) As Long
'Lê nas tabelas de EstoqueProduto e Almoxarifado as posições do produto
'Retorna os Almoxarifados com os produtos em questão em colEstoque
'Filtra apenas os Estoques que têm quantidade positiva do produto

Dim lComando As Long
Dim lComando2 As Long
Dim lErro As Long
Dim tEstoqueProduto As typeEstoqueProduto
Dim sFiltro As String

On Error GoTo Erro_EstoquesProduto_Le

    lComando = 0
    lComando2 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23713

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 23714
        
    With tEstoqueProduto
    
        .sAlmoxarifadoNomeReduzido = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sLocalizacaoFisica = String(STRING_LOCALIZACAO_FISICA, 0)
        .sContaContabil = String(STRING_CONTA, 0)
        
        'Verifica EstoqueProduto da própria filial
        lErro = Comando_Executar(lComando, "SELECT Almoxarifado, NomeReduzido, LocalizacaoFisica, EstoqueProduto.ContaContabil, QuantDispNossa, SaldoInicial, DataInicial, QuantReservada, DataInventario, QuantidadeInicial, QuantEmpenhada, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig, QuantConsig3, QuantDemo, QuantDemo3, QuantConserto, QuantConserto3, QuantOutras, QuantOutras3, QuantBenef, QuantBenef3 FROM Almoxarifado, EstoqueProduto WHERE Produto=?  AND FilialEmpresa=? AND Codigo=Almoxarifado AND ( QuantDispNossa > 0 OR QuantConsig3 > 0 OR QuantReservada > 0 OR QuantReservadaConsig > 0 ) ORDER BY NomeReduzido", _
        .iAlmoxarifado, .sAlmoxarifadoNomeReduzido, .sLocalizacaoFisica, .sContaContabil, .dQuantDispNossa, .dSaldoInicial, .dtDataInicial, .dQuantReservada, .dtDataInventario, .dQuantidadeInicial, .dQuantEmpenhada, .dQuantRecIndl, .dQuantInd, .dQuantDefeituosa, .dQuantConsig, .dQuantConsig3, .dQuantDemo, .dQuantDemo3, .dQuantConserto, .dQuantConserto3, .dQuantOutras, .dQuantOutras3, .dQuantBenef, .dQuantBenef3, sProduto, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 23715
        
        Call CF("FilialEmpresa_Le_Filtro_Filial2", sFiltro)
    
        'Verifica EstoqueProduto do restante da empresa
        lErro = Comando_Executar(lComando2, "SELECT Almoxarifado, NomeReduzido, LocalizacaoFisica, EstoqueProduto.ContaContabil, QuantDispNossa, SaldoInicial, DataInicial, QuantReservada, DataInventario, QuantidadeInicial, QuantEmpenhada, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig, QuantConsig3, QuantDemo, QuantDemo3, QuantConserto, QuantConserto3, QuantOutras, QuantOutras3, QuantBenef, QuantBenef3 FROM Almoxarifado, EstoqueProduto WHERE Produto=?  AND FilialEmpresa<>? AND Codigo=Almoxarifado AND ( QuantDispNossa > 0 OR QuantConsig3 > 0 OR QuantReservada > 0 OR QuantReservadaConsig > 0) " & sFiltro & " ORDER BY FilialEmpresa, NomeReduzido", _
            .iAlmoxarifado, .sAlmoxarifadoNomeReduzido, .sLocalizacaoFisica, .sContaContabil, .dQuantDispNossa, .dSaldoInicial, .dtDataInicial, .dQuantReservada, .dtDataInventario, .dQuantidadeInicial, .dQuantEmpenhada, .dQuantRecIndl, .dQuantInd, .dQuantDefeituosa, .dQuantConsig, .dQuantConsig3, .dQuantDemo, .dQuantDemo3, .dQuantConserto, .dQuantConserto3, .dQuantOutras, .dQuantOutras3, .dQuantBenef, .dQuantBenef3, sProduto, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 23716
    
        'Tenta ler EstoqueProduto da própria filial
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23717
    
        'Coloca na coleção caso não esteja vazio
        Do While lErro <> AD_SQL_SEM_DADOS
            
            colEstoque.Add .sProduto, .iAlmoxarifado, .sAlmoxarifadoNomeReduzido, .dQuantDispNossa, .dSaldo, .sLocalizacaoFisica, .sContaContabil, .dQuantReservada, .dtDataInventario, .dQuantidadeInicial, .dSaldoInicial, .dtDataInicial, .dQuantEmpenhada, .dQuantRecIndl, .dQuantInd, .dQuantDefeituosa, .dQuantConsig, .dQuantConsig3, .dQuantDemo, .dQuantDemo3, .dQuantConserto3, .dQuantConserto, .dQuantOutras, .dQuantOutras3, .dQuantBenef, .dQuantBenef3
    
            'le o proximo registro da tabela
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23719
    
        Loop
    
        'Tenta ler EstoqueProduto do restante da empresa
        lErro = Comando_BuscarPrimeiro(lComando2)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23718
    
        Do While lErro <> AD_SQL_SEM_DADOS
    
            colEstoque.Add .sProduto, .iAlmoxarifado, .sAlmoxarifadoNomeReduzido, .dQuantDispNossa, .dSaldo, .sLocalizacaoFisica, .sContaContabil, .dQuantReservada, .dtDataInventario, .dQuantidadeInicial, .dSaldoInicial, .dtDataInicial, .dQuantEmpenhada, .dQuantRecIndl, .dQuantInd, .dQuantDefeituosa, .dQuantConsig, .dQuantConsig3, .dQuantDemo, .dQuantDemo3, .dQuantConserto3, .dQuantConserto, .dQuantOutras, .dQuantOutras3, .dQuantBenef, .dQuantBenef3
    
            'le o proximo registro da tabela
            lErro = Comando_BuscarProximo(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23728
    
        Loop
        
    End With
        
    If colEstoque.Count = 0 Then Error 30100
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    EstoquesProduto_Le = SUCESSO

    Exit Function

Erro_EstoquesProduto_Le:

    EstoquesProduto_Le = Err

    Select Case Err

        Case 23713, 23714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23715, 23716, 23717, 23718, 23719, 23728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUESPRODUTO", Err)

        Case 30100
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151962)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function AlmoxarifadosProduto_Le(ByVal sProduto As String, ByVal colAlmoxarifado As Collection) As Long
'Lê os códigos de todos os Almoxarifados que tem o Produto cadastrado,
'mesmo que a quantidade disponível seja zero.
'Entra com código do produto em sProduto, retorna colAlmoxarifado

Dim lComando As Long
Dim lErro As Long
Dim iAlmoxarifado As Integer

On Error GoTo Erro_AlmoxarifadosProduto_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 25123
       
    'Seleciona registros de EstoqueProduto associados ao Produto
    lErro = Comando_Executar(lComando, "SELECT Almoxarifado FROM EstoqueProduto, Almoxarifado WHERE Produto=? AND EstoqueProduto.Almoxarifado= Almoxarifado.Codigo", iAlmoxarifado, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 25124

    'Tenta ler EstoqueProduto
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25125

    'Coloca na coleção caso não esteja vazio
    Do While lErro <> AD_SQL_SEM_DADOS
        
        colAlmoxarifado.Add iAlmoxarifado

        'Lê o próximo registro da tabela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25126

    Loop

    If colAlmoxarifado.Count = 0 Then Error 25127
    
    Call Comando_Fechar(lComando)

    AlmoxarifadosProduto_Le = SUCESSO

    Exit Function

Erro_AlmoxarifadosProduto_Le:

    AlmoxarifadosProduto_Le = Err

    Select Case Err

        Case 25123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25124, 25125, 25126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUESPRODUTO", Err)

        Case 25127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_EXISTE_ALMOXARIFADO", Err, sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151963)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoCategorias_Le_Produto(ByVal sProduto As String, ByVal colProdutoCategoria As Collection) As Long
'preenche uma colecao com objProdutoCategoria(s) associados a sProduto
Dim lErro As Long, objProdutoCategoria As ClassProdutoCategoria
Dim lComando As Long, sCategoria As String, sItem As String
On Error GoTo Erro_ProdutoCategorias_Le_Produto

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27264
    
    sCategoria = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Categoria, Item FROM ProdutoCategoria WHERE Produto = ?", sCategoria, sItem, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 27265
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27266
        
    Do While lErro <> AD_SQL_SEM_DADOS
        
        Set objProdutoCategoria = New ClassProdutoCategoria
        
        objProdutoCategoria.sProduto = sProduto
        objProdutoCategoria.sCategoria = sCategoria
        objProdutoCategoria.sItem = sItem
        
        colProdutoCategoria.Add objProdutoCategoria
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27267
    
    Loop
    
    lErro = Comando_Fechar(lComando)
    
    ProdutoCategorias_Le_Produto = SUCESSO

    Exit Function

Erro_ProdutoCategorias_Le_Produto:

    ProdutoCategorias_Le_Produto = Err

    Select Case Err

        Case 27264
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 27265, 27266, 27267
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGS_PROD", Err, sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151964)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function EstoqueMesProduto_Le(ByVal iAno As Integer, ByVal iMes As Integer, ByVal colEstoqueMesProduto As Collection) As Long
'preenche uma colecao com objProduto(s) associados a um item de categoria

Dim lErro As Long, objEstoqueMesProduto As ClassEstoqueMesProduto
Dim lComando As Long, sProduto As String, sDescricao As String
Dim iFilialEmpresa As Integer
Dim iCustoProdApurado As Integer
Dim dGasto As Double
Dim dQuantidade As Double

On Error GoTo Erro_EstoqueMesProduto_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92865
    
    sProduto = String(STRING_PRODUTO, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Produto, CustoProdApurado, Gasto, Descricao, Quantidade FROM EstoqueMesProduto, Produtos WHERE Produto = Codigo AND Ano = ? AND Mes = ? ORDER BY FilialEmpresa, Produto", iFilialEmpresa, sProduto, iCustoProdApurado, dGasto, sDescricao, dQuantidade, iAno, iMes)
    If lErro <> AD_SQL_SUCESSO Then gError 92866
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92867
        
    Do While lErro = AD_SQL_SUCESSO
        
        Set objEstoqueMesProduto = New ClassEstoqueMesProduto
        
        objEstoqueMesProduto.iFilialEmpresa = iFilialEmpresa
        objEstoqueMesProduto.sProduto = sProduto
        objEstoqueMesProduto.sDescricao = sDescricao
        objEstoqueMesProduto.iCustoProdApurado = iCustoProdApurado
        objEstoqueMesProduto.dGasto = dGasto
        objEstoqueMesProduto.dQuantidade = dQuantidade
        objEstoqueMesProduto.iMes = iMes
        objEstoqueMesProduto.iAno = iAno
        
        colEstoqueMesProduto.Add objEstoqueMesProduto
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92868
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    EstoqueMesProduto_Le = SUCESSO

    Exit Function

Erro_EstoqueMesProduto_Le:

    EstoqueMesProduto_Le = gErr

    Select Case gErr

        Case 92865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 92866, 92867, 92868
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMESPRODUTO", gErr, iAno, iMes)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151965)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Produto_Le_NomeReduzido(ByVal objProduto As ClassProduto) As Long
'Lê o Produto com o Nome Reduzido passado

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto

On Error GoTo Erro_Produto_Le_NomeReduzido

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26924

    'Inicializa as variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sReferencia = String(STRING_PRODUTO_REFERENCIA, 0)
        .sFigura = String(STRING_NOME_ARQ_COMPLETO, 0)
        .sSerieProx = String(STRING_LOTE_RASTREAMENTO, 0) 'Inserido por Wagner 08/03/2006
        .sFCI = String(STRING_FCI, 0)
        .sCEST = String(STRING_CEST, 0)
        .scProdANVISA = String(STRING_MAXIMO, 0)

        '##################################################
        'Alterado por Wagner 08/03/2006
        'Tenta ler
        lErro = Comando_Executar(lComando, "SELECT  Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, Codigo, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao,TempoProducao,Rastro,HorasMaquina,PesoEspecifico, Embalagem, Referencia, Figura, SerieProx, SerieParteNum, Natureza, QtdeEmbBase, FatorAjuste, FCI, PercComissao, MetaComissao, CEST, ProdEmEscalaRelev, cProdANVISA FROM Produtos WHERE NomeReduzido = ? ", .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sCodigo, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .iTempoProducao, .iRastro, .lHorasMaquina, .dPesoEspecifico, .iEmbalagem, .sReferencia, .sFigura, .sSerieProx, .iSerieParteNum, .iNatureza, .dQtdeEmbBase, .dFatorAjuste, .sFCI, .dPercComissao, .dMetaComissao, .sCEST, .iProdEmEscalaRelev, .scProdANVISA, objProduto.sNomeReduzido)
        If lErro <> AD_SQL_SUCESSO Then Error 26925
        '##################################################

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26926
        'Não encontra
        If lErro = AD_SQL_SEM_DADOS Then Error 26927

        'Guarda os dados no objProduto
        objProduto.dComprimento = .dComprimento
        objProduto.dEspessura = .dEspessura
        objProduto.dIPIAliquota = .dIPIAliquota
        objProduto.dLargura = .dLargura
        objProduto.dPesoBruto = .dPesoBruto
        objProduto.dPesoLiq = .dPesoLiq
        objProduto.iAtivo = .iAtivo
        objProduto.iClasseUM = .iClasseUM
        objProduto.iCompras = .iCompras
        objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
        objProduto.iFaturamento = .iFaturamento
        objProduto.iGerencial = .iGerencial
        objProduto.iKitBasico = .iKitBasico
        objProduto.iKitInt = .iKitInt
        objProduto.iNivel = .iNivel
        objProduto.iPCP = .iPCP
        objProduto.iPrazoValidade = .iPrazoValidade
        objProduto.iTipo = .iTipo
        objProduto.sCodigo = .sCodigo
        objProduto.sCodigoBarras = .sCodigoBarras
        objProduto.sCor = .sCor
        objProduto.sDescricao = .sDescricao
        objProduto.sIPICodDIPI = .sIPICodDIPI
        objProduto.sIPICodigo = .sIPICodigo
        objProduto.sModelo = .sModelo
        objProduto.sObsFisica = .sObsFisica
        objProduto.sSiglaUMCompra = .sSiglaUMCompra
        objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
        objProduto.sSiglaUMVenda = .sSiglaUMVenda
        objProduto.sSubstituto1 = .sSubstituto1
        objProduto.sSubstituto2 = .sSubstituto2
        objProduto.iControleEstoque = .iControleEstoque
        objProduto.iApropriacaoCusto = .iApropriacaoCusto
        objProduto.iTempoProducao = .iTempoProducao
        objProduto.iRastro = .iRastro
        objProduto.lHorasMaquina = .lHorasMaquina
        objProduto.dPesoEspecifico = .dPesoEspecifico
        objProduto.iEmbalagem = .iEmbalagem
        objProduto.sReferencia = .sReferencia
        objProduto.sFigura = .sFigura
        objProduto.iNatureza = .iNatureza
        objProduto.dQtdeEmbBase = .dQtdeEmbBase
        objProduto.dFatorAjuste = .dFatorAjuste
        objProduto.sFCI = .sFCI
        objProduto.dPercComissao = .dPercComissao
        objProduto.dMetaComissao = .dMetaComissao
        objProduto.sSerieProx = .sSerieProx
        objProduto.iSerieParteNum = .iSerieParteNum
        
        objProduto.sCEST = .sCEST
        objProduto.iProdEmEscalaRelev = .iProdEmEscalaRelev
        objProduto.scProdANVISA = .scProdANVISA
        
    End With

    'fecha o comando
    Call Comando_Fechar(lComando)

    Produto_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_Produto_Le_NomeReduzido:

    Produto_Le_NomeReduzido = Err

    Select Case Err

        Case 26924
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 26925, 26926
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS2", Err, objProduto.sNomeReduzido)

        Case 26927

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151966)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CategoriaClienteItem_Le(colCategoria As Collection, sCategoria As String)
'Lê os itens contidos na tabela CategoriaClienteItem, onde categoria vale sCategoria, e coloca-os em colCategoria

'Sendo Criada em EXCECOES ICMS.frm (TelasFAT)

End Function

Function CategoriaClienteItem_Le_Todos(colCategoria As Collection)
'Lê todos os itens contidos na tabela CategoriaClienteItem e coloca-os em colCategoria

'Sendo Criada em EXCECOES ICMS.frm (TelasFAT)

End Function

Function CategoriaProdutoItem_Le_Todos(colCategoria As Collection)
'Lê todos os itens contidos na tabela CategoriaProdutoItem e coloca-os em colCategoria

'Sendo Criada em EXCECOES ICMS.frm (TelasFAT)

End Function

Function CategoriaProdutoItem_Le(colCategoria As Collection, sCategoria As String)
'Lê os itens contidos na tabela CategoriaProdutoItem, onde categoria vale sCategoria, e coloca-os em colCategoria

'Sendo Criada em EXCECOES ICMS.frm (TelasFAT)

End Function

Function Produto_Le_TabelaPrecoItem(ByVal objProduto As ClassProduto, ByVal colTabelaPrecoItem As Collection, ByVal iFilialEmpresa As Integer) As Long
'Lê os preços de um produto para a iFilialEmpresa (filial corrente ou EMPRESA_TODA). Retorna os dados em colTabelaPrecoItem

Dim lErro As Long
Dim lComando As Long
Dim tTabelaPrecoItem As typeTabelaPrecoItem
Dim objTabelaPrecoItem As ClassTabelaPrecoItem
Dim iTabelaAnterior As Integer

On Error GoTo Erro_Produto_Le_TabelaPrecoItem

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28011

    tTabelaPrecoItem.sDescricaoTabela = String(STRING_TABELA_DESCRICAO, 0)
    
    'Pesquisa no BD Ítens de Tabela de Preço
    lErro = Comando_Executar(lComando, "SELECT TabelasDePrecoItens.CodTabela, TabelasDePrecoItens.Preco, TabelasDePrecoItens.DataVigencia, TabelasDePreco.Descricao FROM TabelasDePrecoItens, TabelasDePreco WHERE TabelasDePrecoItens.DataVigencia <= ? And TabelasDePrecoItens.FilialEmpresa = ? And TabelasDePrecoItens.CodProduto = ? AND TabelasDePreco.Codigo = TabelasDePrecoItens.CodTabela ORDER BY TabelasDePrecoItens.CodTabela, TabelasDePrecoItens.DataVigencia DESC", tTabelaPrecoItem.iCodTabela, tTabelaPrecoItem.dPreco, tTabelaPrecoItem.dtDataVigencia, tTabelaPrecoItem.sDescricaoTabela, Date, iFilialEmpresa, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 28012

    'Verifica se Ítem existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28013
    
    iTabelaAnterior = 0
    
    Do While lErro = AD_SQL_SUCESSO
        
        'Preenche com uma Tabela (a que está na data de Vigencia)
        If iTabelaAnterior <> tTabelaPrecoItem.iCodTabela Then
                    
            Set objTabelaPrecoItem = New ClassTabelaPrecoItem
        
            'Preenche objTabelaPrecoItem
            objTabelaPrecoItem.iCodTabela = tTabelaPrecoItem.iCodTabela
            objTabelaPrecoItem.dPreco = tTabelaPrecoItem.dPreco
            objTabelaPrecoItem.dtDataVigencia = tTabelaPrecoItem.dtDataVigencia
            objTabelaPrecoItem.sDescricaoTabela = tTabelaPrecoItem.sDescricaoTabela

            'Preenche colTabelaPrecoItem
            colTabelaPrecoItem.Add objTabelaPrecoItem
        
        End If
        
        'Atualiza a última Tabela da flag
        iTabelaAnterior = tTabelaPrecoItem.iCodTabela
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26856
    
    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Produto_Le_TabelaPrecoItem = SUCESSO

    Exit Function

Erro_Produto_Le_TabelaPrecoItem:

    Produto_Le_TabelaPrecoItem = Err

    Select Case Err

        Case 28011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28012, 28013, 26856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS", Err)

        Case 28014
            'Não encontrou preco para o produto. Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151967)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function UM_Le_Lock(ByVal objUnidadeMedida As ClassUnidadeDeMedida) As Long
'Lê e "locka" a unidade de medida a partir da Classe e da SiglaUMBase -->
'e devolve dados no Objeto

Dim lErro As Long
Dim lComando As Long
Dim tUnidMed As typeUnidadeMedida

On Error GoTo Erro_UM_Le_Lock

    'Inicializa as variáveis do tipo String
    tUnidMed.sNome = String(STRING_UM_NOME, 0)
    tUnidMed.sSiglaUMBase = String(STRING_UM_SIGLA_UM_BASE, 0)

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26939

    'Lê a UM com a Sigla e a Classe passadas
    lErro = Comando_ExecutarLockado(lComando, "SELECT Nome, SiglaUMBase, Quantidade FROM UnidadesdeMedida WHERE Classe=? AND Sigla=?", tUnidMed.sNome, tUnidMed.sSiglaUMBase, tUnidMed.dQuantidade, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 26940

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26941
    'não encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 26942

    'faz o "lock" da UM
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 26943

    With tUnidMed
        'Passa os dados para o obj
        objUnidadeMedida.dQuantidade = .dQuantidade
        objUnidadeMedida.sNome = .sNome
        objUnidadeMedida.sSigla = .sSigla

    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)

    UM_Le_Lock = SUCESSO

    Exit Function

Erro_UM_Le_Lock:

    UM_Le_Lock = Err

    Select Case Err

        Case 26939
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 26940, 26941
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_UNIDADESDEMEDIDA", Err, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)

        Case 26942
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA", Err, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)

        Case 26943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_UNIDADESDEMEDIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151968)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoDeProduto_Lock(ByVal objTipoDeProduto As ClassTipoDeProduto) As Long
'Lê e faz "lock" na tabela TiposDeProduto, a partir do Tipo De Produto

Dim lErro As Long
Dim lComando As Long
Dim objTipoDeProdutoCategoria As New ClassTipoDeProdutoCategoria
Dim iControleEstoque As Integer

On Error GoTo Erro_TipoDeProduto_Lock

    'Inicializa o coando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26947

    'Lê a tabela TipoDeProduto a partir do Tipo De Produto
    lErro = Comando_ExecutarLockado(lComando, "SELECT ControleEstoque FROM TiposDeProduto WHERE TipoDeProduto = ?", iControleEstoque, objTipoDeProduto.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 26948

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26949

    'Verifica se o Tipo de Produto foi encontrado
    If lErro = AD_SQL_SEM_DADOS Then Error 26946

    'Faz "lock" no Tipo de Produto
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 26951

    objTipoDeProduto.iControleEstoque = iControleEstoque

    'Fecha o comando
    Call Comando_Fechar(lComando)

    TipoDeProduto_Lock = SUCESSO

    Exit Function

Erro_TipoDeProduto_Lock:

    TipoDeProduto_Lock = Err

    Select Case Err

        Case 26946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", Err, objTipoDeProduto.iTipo)

        Case 26947
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 26948, 26949
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEPRODUTO", Err)

        Case 26950
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODEPRODUTOCATEGORIAS", Err)

        Case 26951
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOSDEPRODUTO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151969)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoKit_Le(objProdutoKit As ClassProdutoKit, colComponentes As Collection) As Long
'Le a tabela de Kits de Produtos a partir do produto e da versão passados em objProdutoKit

'Sendo Criada em Kit.frm (TelasEst)

End Function

Function ClasseUM_Le(ByVal objClasseUM As ClassClasseUM) As Long
'Lê na tabela ClasseUM a Descrição e a Sigla da classe, a partir do código da classe

Dim lErro As Long
Dim iClasse As Integer
Dim sDescricao As String
Dim sSigla As String
Dim lComando As Long

On Error GoTo Erro_ClasseUM_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22364

    sDescricao = String(STRING_CLASSEUM_DESCRICAO, 0)
    sSigla = String(STRING_UM_SIGLA, 0)

    'Lê todas as classes
    lErro = Comando_Executar(lComando, "SELECT Classe, Descricao, Sigla FROM ClasseUM WHERE Classe = ?", iClasse, sDescricao, sSigla, objClasseUM.iClasse)
    If lErro <> AD_SQL_SUCESSO Then Error 22365

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22366

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 22537

    objClasseUM.sDescricao = sDescricao
    objClasseUM.sSiglaUMBase = sSigla

    ClasseUM_Le = SUCESSO

    Call Comando_Fechar(lComando)

    Exit Function

Erro_ClasseUM_Le:

    ClasseUM_Le = Err

    Select Case Err

        Case 22537

        Case 22364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22365, 22366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSEUM", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151970)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ClasseUM_Le_Todas(ByVal colClasse As Collection) As Long

Dim lErro As Long
Dim iClasse As Integer
Dim sDescricao As String
Dim sSigla As String
Dim lComando As Long
Dim objClasseUM As ClassClasseUM

On Error GoTo Erro_ClasseUM_Le_Todas

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22364

    sDescricao = String(STRING_CLASSEUM_DESCRICAO, 0)
    sSigla = String(STRING_UM_SIGLA, 0)

    'Lê todas as classes
    lErro = Comando_Executar(lComando, "SELECT Classe, Descricao, Sigla FROM ClasseUM ORDER BY Classe ", iClasse, sDescricao, sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 22365

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22366

    'Se não encontrou
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objClasseUM = New ClassClasseUM

        objClasseUM.iClasse = iClasse
        objClasseUM.sDescricao = sDescricao
        objClasseUM.sSiglaUMBase = sSigla
        
        colClasse.Add objClasseUM
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22366
        
    Loop

    ClasseUM_Le_Todas = SUCESSO

    Call Comando_Fechar(lComando)

    Exit Function

Erro_ClasseUM_Le_Todas:

    ClasseUM_Le_Todas = Err

    Select Case Err

        Case 22537

        Case 22364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22365, 22366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSEUM", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151970)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_ProdutoMaior(ByVal sProduto As String, ByVal objProduto As ClassProduto) As Long
'Lê um Produto imediatamente maior que o produto passada como parametro. Devolve os dados em objProduto

Dim lComando As Long
Dim lErro As Long
Dim tProduto As typeProduto

On Error GoTo Erro_Produto_Le_ProdutoMaior

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 31005

    'Inicializa as variáveis do tipo String
    With tProduto
    
        .sCodigo = String(STRING_PRODUTO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)

        'Lê o Produto
        lErro = Comando_Executar(lComando, "SELECT  Codigo, Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao,TempoProducao FROM Produtos WHERE Codigo > ? ORDER BY Produtos.Codigo", .sCodigo, .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .iTempoProducao, sProduto)
        If lErro <> AD_SQL_SUCESSO Then Error 31006

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31007

        'não encontrou
        If lErro = AD_SQL_SEM_DADOS Then Error 31008

        'Carrega os dados encontrados em objProduto
        Call Produto_MoverCamposObj(tProduto, objProduto)
        
    End With

    'Fechar o comando
    lErro = Comando_Fechar(lComando)

    Produto_Le_ProdutoMaior = SUCESSO

    Exit Function

Erro_Produto_Le_ProdutoMaior:

    Produto_Le_ProdutoMaior = Err

    Select Case Err

        Case 31005
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 31006, 31007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", Err)

        Case 31008

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151971)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Tem_Filho(ByVal sProduto As String, iTemFilho As Integer) As Long
'testa se o Produto em questão tem Produtos abaixo dele, ou seja tem Produtos filho.

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iNivel As Integer
Dim iNivel1 As Integer

On Error GoTo Erro_Produto_Tem_Filho

    'lê o Produto com número imediatamente maior que o Produto em questão
    lErro = CF("Produto_Le_ProdutoMaior", sProduto, objProduto)
    If lErro <> SUCESSO And lErro <> 31008 Then Error 31002

    If lErro = SUCESSO Then

        'obtem o nivel do Produto em questão
        lErro = Mascara_Produto_ObterNivel(sProduto, iNivel)
        If lErro <> SUCESSO Then Error 31003

        'obtem o nivel do Produto lido
        lErro = Mascara_Produto_ObterNivel(objProduto.sCodigo, iNivel1)
        If lErro <> SUCESSO Then Error 31004

        'se o nivel do Produto em questão for menor que o Produto lido ==> tem filho
        If iNivel < iNivel1 Then
            iTemFilho = True
        Else
            iTemFilho = False
        End If

    Else
         iTemFilho = False
    End If

    Produto_Tem_Filho = SUCESSO

    Exit Function

Erro_Produto_Tem_Filho:

    Produto_Tem_Filho = Err

    Select Case Err

        Case 31002

        Case 31003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_PRODUTO_OBTERNIVEL", Err, sProduto)

        Case 31004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_PRODUTO_OBTERNIVEL", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151972)

    End Select

    Exit Function

End Function

Function CategoriaProduto_Le(ByVal objCategoriaProduto As ClassCategoriaProduto) As Long
'Lê na tabela CategoriaProduto a descrição da categoria

Dim lErro As Long
Dim sCategoria As String
Dim sDescricao As String
Dim sSigla As String
Dim lComando As Long

On Error GoTo Erro_CategoriaProduto_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22364

    sDescricao = String(STRING_CATEGORIAPRODUTO_DESCRICAO, 0)
    sSigla = String(STRING_CATEGORIAPRODUTO_SIGLA, 0)

    'Lê todas as categorias
    lErro = Comando_Executar(lComando, "SELECT Descricao, Sigla FROM CategoriaProduto WHERE Categoria = ?", sDescricao, sSigla, objCategoriaProduto.sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 22365

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22366

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 22540

    objCategoriaProduto.sDescricao = sDescricao
    objCategoriaProduto.sSigla = sSigla

    'Libera comando
    Call Comando_Fechar(lComando)

    CategoriaProduto_Le = SUCESSO

    Exit Function

Erro_CategoriaProduto_Le:

    CategoriaProduto_Le = Err

    Select Case Err

        Case 22540

        Case 22364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22365, 22366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151973)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CategoriasProduto_Le_Todas(ByVal colCategorias As Collection) As Long
'Lê na tabela CategoriaProduto todas as Categorias, passando para a coleção colCategorias

Dim lErro As Long
Dim objCategoriaProduto As ClassCategoriaProduto
Dim sCategoria As String
Dim sDescricao As String
Dim sSigla As String
Dim lComando As Long

On Error GoTo Erro_CategoriasProduto_Le_Todas

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22338

    sCategoria = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    sDescricao = String(STRING_CATEGORIAPRODUTO_DESCRICAO, 0)
    sSigla = String(STRING_CATEGORIAPRODUTO_SIGLA, 0)

    'Lê todas as categorias
    lErro = Comando_Executar(lComando, "SELECT Categoria, Descricao, Sigla FROM CategoriaProduto", sCategoria, sDescricao, sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 22339

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22340

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 22542

    'Passa para a collection
    Do While lErro = AD_SQL_SUCESSO

        Set objCategoriaProduto = New ClassCategoriaProduto

        objCategoriaProduto.sCategoria = sCategoria
        objCategoriaProduto.sDescricao = sDescricao
        objCategoriaProduto.sSigla = sSigla

        colCategorias.Add objCategoriaProduto

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22341

    Loop

    'Libera comando
    Call Comando_Fechar(lComando)

    CategoriasProduto_Le_Todas = SUCESSO

    Exit Function

Erro_CategoriasProduto_Le_Todas:

    CategoriasProduto_Le_Todas = Err

    Select Case Err

        Case 22542

        Case 22338
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22339, 22340, 22341
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151974)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'alterado por Ivan 3/4/03
Function CategoriaProduto_Le_Itens(ByVal objCategoriaProduto As ClassCategoriaProduto, ByVal colItensCategoria As Collection) As Long
'Le na tabela de CategoriaProdutoItem todos os itens de uma Categoria e os retorna na coleção colItensCategoria

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem
Dim tCategoriaItem As typeCategoriaProdutoItem

On Error GoTo Erro_CategoriaProduto_Le_Itens

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22354

    With tCategoriaItem
        .sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
        .sDescricao = String(STRING_CATEGORIAPRODUTOITEM_DESCRICAO, 0)
        .sCategoria = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)

        lErro = Comando_Executar(lComando, "SELECT Categoria, Item, Ordem, Descricao, Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Valor7, Valor8 FROM CategoriaProdutoItem WHERE Categoria = ? ORDER BY Ordem", _
            .sCategoria, .sItem, .iOrdem, .sDescricao, .dvalor1, .dvalor2, .dvalor3, .dvalor4, .dvalor5, .dvalor6, .dvalor7, .dvalor8, objCategoriaProduto.sCategoria)
    End With
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22355

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22356

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 22541

    'Passa todos para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objCategoriaProdutoItem = New ClassCategoriaProdutoItem

        With tCategoriaItem

            objCategoriaProdutoItem.sCategoria = .sCategoria
            objCategoriaProdutoItem.sItem = .sItem
            objCategoriaProdutoItem.iOrdem = .iOrdem
            objCategoriaProdutoItem.sDescricao = .sDescricao
            objCategoriaProdutoItem.dvalor1 = .dvalor1
            objCategoriaProdutoItem.dvalor2 = .dvalor2
            objCategoriaProdutoItem.dvalor3 = .dvalor3
            objCategoriaProdutoItem.dvalor4 = .dvalor4
            objCategoriaProdutoItem.dvalor5 = .dvalor5
            objCategoriaProdutoItem.dvalor6 = .dvalor6
            objCategoriaProdutoItem.dvalor7 = .dvalor7
            objCategoriaProdutoItem.dvalor8 = .dvalor8

        End With
        
        colItensCategoria.Add objCategoriaProdutoItem

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22357

    Loop

    Call Comando_Fechar(lComando)

    CategoriaProduto_Le_Itens = SUCESSO

    Exit Function

Erro_CategoriaProduto_Le_Itens:

    CategoriaProduto_Le_Itens = Err

    Select Case Err

        Case 22541

        Case 22354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22355, 22356, 22357
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTOITEM", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151975)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoDeProduto_Le_Categorias(ByVal objTipoDeProduto As ClassTipoDeProduto, ByVal colCategoriaItem As Collection) As Long
'Lê na tabela de TiposDeProdutoCategoria, todas as categorias de um Tipo de Produto, guardando na coleção colCategoriaItem

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim sCategoria As String
Dim sItem As String
Dim objTipoDeProdutoCategoria As ClassTipoDeProdutoCategoria

On Error GoTo Erro_TipoDeProduto_Le_Categorias

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22522

    sCategoria = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0) '??? os 2 defines estao errados

    'Executar comando SQL
    lErro = Comando_Executar(lComando, "SELECT Categoria, Item FROM TipoDeProdutoCategorias WHERE TipoDeProduto = ? ", sCategoria, sItem, objTipoDeProduto.iTipo)
    If lErro <> AD_SQL_SUCESSO Then Error 22523

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22524

    'Passa para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objTipoDeProdutoCategoria = New ClassTipoDeProdutoCategoria

        objTipoDeProdutoCategoria.iTipoDeProduto = objTipoDeProduto.iTipo
        objTipoDeProdutoCategoria.sCategoria = sCategoria
        objTipoDeProdutoCategoria.sItem = sItem

        colCategoriaItem.Add objTipoDeProdutoCategoria

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22525

    Loop

    Call Comando_Fechar(lComando)

    TipoDeProduto_Le_Categorias = SUCESSO

    Exit Function

Erro_TipoDeProduto_Le_Categorias:

    TipoDeProduto_Le_Categorias = Err

    Select Case Err

        Case 22522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22523, 22524, 22525
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODEPRODUTOCATEGORIAS", Err, objTipoDeProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151976)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Alterado por Ivan 25/4/03
'inclusão de leitura de novos valores
Public Function CategoriaProduto_Le_Item(ByVal objCategoriaProdutoItem As ClassCategoriaProdutoItem) As Long
'Le na tabela de CategoriaProdutoItem a Categoria e o Item de uma deteminada Categoria de Produto
'Alterada por tulio em 22/04

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim sCategoria As String
Dim sItem As String, iOrdem As Integer
Dim sDescricaoItem As String
Dim adValor(0 To 7) As Double

On Error GoTo Erro_CategoriaProduto_Le_Item

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22600

    sCategoria = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    sDescricaoItem = String(STRING_CATEGORIAPRODUTO_DESCRICAO, 0)

    'Executar comando SQL
    lErro = Comando_Executar(lComando, "SELECT Categoria, Item, Ordem, Descricao, Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Valor7, Valor8 FROM CategoriaProdutoItem WHERE Categoria = ? AND Item = ?", _
        sCategoria, sItem, iOrdem, sDescricaoItem, adValor(0), adValor(1), adValor(2), adValor(3), adValor(4), adValor(5), adValor(6), adValor(7), objCategoriaProdutoItem.sCategoria, objCategoriaProdutoItem.sItem)
    If lErro <> AD_SQL_SUCESSO Then Error 22601

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22602

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then Error 22603

    'Atribui ao obj os dados lidos do bd
    objCategoriaProdutoItem.sCategoria = sCategoria
    objCategoriaProdutoItem.sItem = sItem
    objCategoriaProdutoItem.iOrdem = iOrdem
    objCategoriaProdutoItem.sDescricao = sDescricaoItem
    objCategoriaProdutoItem.dvalor1 = adValor(0)
    objCategoriaProdutoItem.dvalor2 = adValor(1)
    objCategoriaProdutoItem.dvalor3 = adValor(2)
    objCategoriaProdutoItem.dvalor4 = adValor(3)
    objCategoriaProdutoItem.dvalor5 = adValor(4)
    objCategoriaProdutoItem.dvalor6 = adValor(5)
    objCategoriaProdutoItem.dvalor7 = adValor(6)
    objCategoriaProdutoItem.dvalor8 = adValor(7)

    Call Comando_Fechar(lComando)

    CategoriaProduto_Le_Item = SUCESSO

    Exit Function

Erro_CategoriaProduto_Le_Item:

    CategoriaProduto_Le_Item = Err

    Select Case Err

        Case 22600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22601, 22602
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTOITEM2", Err, objCategoriaProdutoItem.sCategoria, objCategoriaProdutoItem.sItem)

        Case 22603

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151977)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Kit_Le_Componentes(ByVal objKit As ClassKit) As Long
'Preenche objKit.colComponentes a partir do produto e da versão passados em objKit

Dim lErro As Long
Dim lComando As Long
Dim tProdutoKit As typeProdutoKit
Dim objProdutoKit As ClassProdutoKit

On Error GoTo Erro_Kit_Le_Componentes

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 21828

    tProdutoKit.sProdutoRaiz = String(STRING_PRODUTO, 0)
    tProdutoKit.sProduto = String(STRING_PRODUTO, 0)
    tProdutoKit.sUnidadeMed = String(STRING_UM_SIGLA_UM_BASE, 0)
    tProdutoKit.sVersao = String(STRING_KIT_VERSAO, 0)
    
    '################################################
    'Inserido por Wagner 23/11/04
    tProdutoKit.sVersaoKitComp = String(STRING_KIT_VERSAO, 0)
    '################################################

    '################################################
    'Alterado por Wagner 23/11/04
    lErro = Comando_Executar(lComando, "SELECT ProdutoRaiz, Versao, Nivel, Seq, Produto, SeqPai, Quantidade, UnidadeMed, Composicao, PosicaoArvore, PercentualPerda, CustoStandard, VersaoKitComp FROM ProdutoKit WHERE ProdutoRaiz = ? AND Versao = ? ORDER BY Nivel, Seq ", tProdutoKit.sProdutoRaiz, tProdutoKit.sVersao, tProdutoKit.iNivel, tProdutoKit.iSeq, tProdutoKit.sProduto, tProdutoKit.iSeqPai, tProdutoKit.dQuantidade, tProdutoKit.sUnidadeMed, tProdutoKit.iComposicao, tProdutoKit.iPosicaoArvore, tProdutoKit.dPercentualPerda, tProdutoKit.dCustoStandard, tProdutoKit.sVersaoKitComp, objKit.sProdutoRaiz, objKit.sVersao)
    If lErro <> AD_SQL_SUCESSO Then Error 21829
    '################################################

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21830

    If lErro = AD_SQL_SEM_DADOS Then Error 21831

    Do While lErro <> AD_SQL_SEM_DADOS

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
        objProdutoKit.iPosicaoArvore = tProdutoKit.iPosicaoArvore
        objProdutoKit.dPercentualPerda = tProdutoKit.dPercentualPerda
        objProdutoKit.dCustoStandard = tProdutoKit.dCustoStandard
        
        '################################################
        'Inserido por Wagner 23/11/04
        objProdutoKit.sVersaoKitComp = tProdutoKit.sVersaoKitComp
        '################################################
        
        lErro = CF("ProdutoKit_Le_InfoUsu", objProdutoKit)
        If lErro <> SUCESSO Then Error 21831
        
        objKit.colComponentes.Add objProdutoKit

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21832

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Kit_Le_Componentes = SUCESSO

    Exit Function

Erro_Kit_Le_Componentes:

    Kit_Le_Componentes = Err

    Select Case Err

        Case 21828
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 21829, 21830, 21832
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", Err)
        
        Case 21831
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151978)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Kit_Le_MaisRecente(ByVal objKit As ClassKit) As Long
'Le uma versão na tabela de Kits de Produtos a partir do produto passado em objKit e devolve os dados do Kit em objKit

Dim lErro As Long
Dim lComando As Long
Dim tKit As typeKit

On Error GoTo Erro_Kit_Le_MaisRecente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 21855

    With tKit
        .sVersao = String(STRING_KIT_VERSAO, 0)
        .sObservacao = String(STRING_KIT_OBSERVACAO, 0)

        lErro = Comando_Executar(lComando, "SELECT Versao, Observacao, Data, PesoFator1, PesoFator2, PesoFator3, PesoFator4, PesoFator5, PesoFator6 FROM Kit WHERE ProdutoRaiz = ? ORDER BY Data DESC", .sVersao, .sObservacao, .dtData, .dPesoFator1, .dPesoFator2, .dPesoFator3, .dPesoFator4, .dPesoFator5, .dPesoFator6, objKit.sProdutoRaiz)
    End With
    If lErro <> AD_SQL_SUCESSO Then Error 21856

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 21857

    If lErro = AD_SQL_SEM_DADOS Then Error 21858

    With tKit
        objKit.sVersao = .sVersao
        objKit.dtData = .dtData
        objKit.sObservacao = .sObservacao
        objKit.dPesoFator1 = .dPesoFator1
        objKit.dPesoFator2 = .dPesoFator2
        objKit.dPesoFator3 = .dPesoFator3
        objKit.dPesoFator4 = .dPesoFator4
        objKit.dPesoFator5 = .dPesoFator5
        objKit.dPesoFator6 = .dPesoFator6
    End With

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Kit_Le_MaisRecente = SUCESSO

    Exit Function

Erro_Kit_Le_MaisRecente:

    Kit_Le_MaisRecente = Err

    Select Case Err

        Case 21855
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 21856, 21857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_KIT", Err)

        Case 21858

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151979)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Kit_Le(ByVal objKit As ClassKit) As Long
'Le a tabela de Kits a partir do produto e da versão passados em objKit

Dim lErro As Long
Dim lComando As Long
Dim tKit As typeKit

On Error GoTo Erro_Kit_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 21823

    With tKit
        .sObservacao = String(STRING_KIT_OBSERVACAO, 0)

        lErro = Comando_Executar(lComando, "SELECT Data, Observacao, Situacao, VersaoFormPreco, PesoFator1, PesoFator2, PesoFator3, PesoFator4, PesoFator5, PesoFator6 FROM Kit WHERE ProdutoRaiz = ? AND Versao = ?", .dtData, .sObservacao, .iSituacao, .iVersaoFormPreco, .dPesoFator1, .dPesoFator2, .dPesoFator3, .dPesoFator4, .dPesoFator5, .dPesoFator6, objKit.sProdutoRaiz, objKit.sVersao)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 21824

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 21825

    If lErro = AD_SQL_SEM_DADOS Then gError 21826

    With tKit
        objKit.dtData = .dtData
        objKit.sObservacao = .sObservacao
        objKit.iSituacao = .iSituacao
        objKit.iVersaoFormPreco = .iVersaoFormPreco
        objKit.dPesoFator1 = .dPesoFator1
        objKit.dPesoFator2 = .dPesoFator2
        objKit.dPesoFator3 = .dPesoFator3
        objKit.dPesoFator4 = .dPesoFator4
        objKit.dPesoFator5 = .dPesoFator5
        objKit.dPesoFator6 = .dPesoFator6
    End With
    
    '??? para ler os itens chame Kit_Le_Componentes !!!

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Kit_Le = SUCESSO

    Exit Function

Erro_Kit_Le:

    Kit_Le = gErr

    Select Case gErr

        Case 21823
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 21824, 21825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_KIT", gErr)

        Case 21826

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151980)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function SldMesEst_Le_Status(ByVal lComando As Long, ByVal objSldMesEst As ClassSldMesEst) As Long
'le os dados de SldMesEst relativos ao Ano, FilialEmpresa e Produto passados como parametro
    
Dim tSldMesEst As typeSldMesEst
Dim lErro As Long
Dim iIndice As Integer
    
On Error GoTo Erro_SldMesEst_Le_Status
    
    'seleciona o registro de SldMesEst relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_ExecutarLockado(lComando, "SELECT Status1, Status2, Status3, Status4, Status5, Status6, Status7, Status8, Status9, Status10, Status11, Status12  FROM SldMesEst WHERE Ano=? AND FilialEmpresa=? AND Produto=?", tSldMesEst.aiStatus(1), tSldMesEst.aiStatus(2), tSldMesEst.aiStatus(3), tSldMesEst.aiStatus(4), tSldMesEst.aiStatus(5), tSldMesEst.aiStatus(6), tSldMesEst.aiStatus(7), tSldMesEst.aiStatus(8), tSldMesEst.aiStatus(9), tSldMesEst.aiStatus(10), tSldMesEst.aiStatus(11), tSldMesEst.aiStatus(12), objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 20380
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 20381
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 20382
    
    For iIndice = 1 To 12
        objSldMesEst.iStatus(iIndice) = tSldMesEst.aiStatus(iIndice)
    Next
    
    SldMesEst_Le_Status = SUCESSO

    Exit Function

Erro_SldMesEst_Le_Status:
    
    SldMesEst_Le_Status = Err
    
    Select Case Err
    
        Case 20380, 20381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", Err, objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
            
        Case 20382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDMESEST", Err, objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151981)
        
    End Select
    
    Exit Function

End Function

Function TiposMovEst_Le1(ByVal lComando As Long, ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'Le a tabela de tipos de movimentos de estoque a partir do codigo fornecido em objTipoMovEstoque e devolve os dados neste mesmo obj.

Dim lErro As Long
Dim tTipoMovEst As typeTipoMovEst
Dim iCodigo As Integer
Dim iAtualizaMovEstoque As Integer

On Error GoTo Erro_TiposMovEst_Le1

    Call objTipoMovEstoque.Copiar(gobjMAT.colTiposMovEst(CStr(objTipoMovEstoque.iCodigo)))
    
    If objTipoMovEstoque.iCodigoOrig <> 0 Then
        'se for um movimento que possui o codigo original preenchido (normalmente estorno)
        'guarda o codigo do estorno e a indicacao de que é um estorno, carrega o registro do movimento origninal e recoloca o codigo do movimento e seu tipo
        'isto permitirá que o movimento seja processado de modo contrário ao tipo original já que a quantidade (no caso de estorno) terá seu sinal invertido
        iAtualizaMovEstoque = objTipoMovEstoque.iAtualizaMovEstoque
        iCodigo = objTipoMovEstoque.iCodigo
        Call objTipoMovEstoque.Copiar(gobjMAT.colTiposMovEst(CStr(objTipoMovEstoque.iCodigoOrig)))
        objTipoMovEstoque.iCodigo = iCodigo
        objTipoMovEstoque.iAtualizaMovEstoque = iAtualizaMovEstoque
    End If
    
    TiposMovEst_Le1 = SUCESSO

    Exit Function

Erro_TiposMovEst_Le1:

    TiposMovEst_Le1 = Err

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151982)

    End Select

    Exit Function

End Function

Function MovimentoEstoque_Le_Estorno(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque) As Long
'Carrega os dados do movimento a ser estornado
    
Dim lErro As Long
Dim lNumIntDocEst As Long
Dim dCusto As Double
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovimentoEstoque_Le_Estorno

    tItemMovEstoque.sProduto = String(STRING_PRODUTO, 0)
    tItemMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemMovEstoque.sCcl = String(STRING_CCL, 0)
    tItemMovEstoque.sOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
    tItemMovEstoque.sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)
    tItemMovEstoque.sContaContabilAplic = String(STRING_CONTA, 0)
    tItemMovEstoque.sContaContabilEst = String(STRING_CONTA, 0)
    
    'le o numintdoc de estorno do movimento de estoque passado como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa, Codigo, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao, Hora FROM MovimentoEstoque WHERE NumIntDoc = ?", 0, _
    tItemMovEstoque.iFilialEmpresa, tItemMovEstoque.lCodigo, tItemMovEstoque.dCusto, tItemMovEstoque.iApropriacao, tItemMovEstoque.sProduto, tItemMovEstoque.sSiglaUM, tItemMovEstoque.dQuantidade, tItemMovEstoque.iAlmoxarifado, tItemMovEstoque.iTipoMov, tItemMovEstoque.lNumIntDocOrigem, tItemMovEstoque.iTipoNumIntDocOrigem, tItemMovEstoque.dtData, tItemMovEstoque.sCcl, tItemMovEstoque.lNumIntDocEst, tItemMovEstoque.lCliente, tItemMovEstoque.lFornecedor, tItemMovEstoque.sOPCodigo, tItemMovEstoque.sDocOrigem, tItemMovEstoque.sContaContabilEst, tItemMovEstoque.sContaContabilAplic, tItemMovEstoque.lHorasMaquina, tItemMovEstoque.dtDataInicioProducao, tItemMovEstoque.dHora, objItemMovEst.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 20411

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 20412

    'se o movimento não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 20428

    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 20413

    objItemMovEst.iFilialEmpresa = tItemMovEstoque.iFilialEmpresa
    objItemMovEst.lCodigo = tItemMovEstoque.lCodigo
    objItemMovEst.dCusto = tItemMovEstoque.dCusto
    objItemMovEst.iApropriacao = tItemMovEstoque.iApropriacao
    objItemMovEst.sProduto = tItemMovEstoque.sProduto
    objItemMovEst.sSiglaUM = tItemMovEstoque.sSiglaUM
    objItemMovEst.dQuantidade = tItemMovEstoque.dQuantidade
    objItemMovEst.iAlmoxarifado = tItemMovEstoque.iAlmoxarifado
    objItemMovEst.iTipoMov = tItemMovEstoque.iTipoMov
    objItemMovEst.lNumIntDocOrigem = tItemMovEstoque.lNumIntDocOrigem
    objItemMovEst.iTipoNumIntDocOrigem = tItemMovEstoque.iTipoNumIntDocOrigem
    objItemMovEst.dtData = tItemMovEstoque.dtData
    objItemMovEst.sCcl = tItemMovEstoque.sCcl
    objItemMovEst.lNumIntDocEst = tItemMovEstoque.lNumIntDocEst
    objItemMovEst.lCliente = tItemMovEstoque.lCliente
    objItemMovEst.lFornecedor = tItemMovEstoque.lFornecedor
    objItemMovEst.sOPCodigo = tItemMovEstoque.sOPCodigo
    objItemMovEst.sDocOrigem = tItemMovEstoque.sDocOrigem
    objItemMovEst.sContaContabilEst = tItemMovEstoque.sContaContabilEst
    objItemMovEst.sContaContabilAplic = tItemMovEstoque.sContaContabilAplic
    objItemMovEst.lHorasMaquina = tItemMovEstoque.lHorasMaquina
    objItemMovEst.dtDataInicioProducao = tItemMovEstoque.dtDataInicioProducao
    objItemMovEst.dtHora = tItemMovEstoque.dHora
        
    MovimentoEstoque_Le_Estorno = SUCESSO
    
    Exit Function

Erro_MovimentoEstoque_Le_Estorno:

    MovimentoEstoque_Le_Estorno = Err

    Select Case Err

        Case 20411, 20412
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE6", Err, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)

        Case 20413
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MOVIMENTOESTOQUE", Err, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)

        Case 20428
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151983)

    End Select

    Exit Function

End Function

Function Empenho_Le(ByVal objEmpenho As ClassEmpenho) As Long
'Lê os dados de Empenho à partir da chave fornecida em objEmpenho
'objEmpenho.iFilialEmpresa e objEmpenho.lCodigo

Dim lErro As Long
Dim lComando As Long
Dim tEmpenho As typeEmpenho

On Error GoTo Erro_Empenho_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33111

    tEmpenho.sProduto = String(STRING_PRODUTO, 0)

    'Pesquisa no BD o Empenho passado como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Produto, Almoxarifado, Quantidade, QuantidadeRequisitada, NumIntDocItemOP, Data FROM Empenho WHERE FilialEmpresa = ? AND Codigo = ?", _
        tEmpenho.sProduto, tEmpenho.iAlmoxarifado, tEmpenho.dQuantidade, tEmpenho.dQuantidadeRequisitada, tEmpenho.lNumIntDocItemOP, tEmpenho.dtData, objEmpenho.iFilialEmpresa, objEmpenho.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 33112

    'Verifica se o Empenho existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33113

    'Não encontrou o Empenho ==> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 33114

    'Preenche objEmpenho
    objEmpenho.sProduto = tEmpenho.sProduto
    objEmpenho.iAlmoxarifado = tEmpenho.iAlmoxarifado
    objEmpenho.dQuantidade = tEmpenho.dQuantidade
    objEmpenho.dQuantidadeRequisitada = tEmpenho.dQuantidadeRequisitada
    objEmpenho.lNumIntDocItemOP = tEmpenho.lNumIntDocItemOP
    objEmpenho.dtData = tEmpenho.dtData
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    Empenho_Le = SUCESSO

    Exit Function

Erro_Empenho_Le:

    Empenho_Le = Err

    Select Case Err

        Case 33111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33112, 33113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPENHO", Err, objEmpenho.iFilialEmpresa, objEmpenho.lCodigo)

        Case 33114
            'Não encontrou o Empenho.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151984)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemOP_Le_Item(ByVal objItemOP As ClassItemOP) As Long
'Lê os dados do Item da Ordem de Produção à partir do Código da OP e Código do Item de uma Filial da Empresa

Dim lErro As Long
Dim lComando As Long
Dim tItemOP As typeItemOP

On Error GoTo Erro_ItemOP_Le_Item

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 33212

    tItemOP.sProduto = String(STRING_PRODUTO, 0)
    tItemOP.sCcl = String(STRING_CCL, 0)
    tItemOP.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemOP.sVersao = String(STRING_KIT_VERSAO, 0)

    '########################################################
    'INSERIDO POR WAGNER
    tItemOP.sJustificativaBloqueio = String(STRING_MAX_NOME_ARQUIVO, 0)
    '########################################################

    'Pesquisa no BD o Item da OP passado como parâmetro
    '########################################################
    'ALTERADO POR WAGNER
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Produto, FilialPedido, CodPedido, NumIntOrigem, Almoxarifado, Ccl, SiglaUM, Quantidade, DataInicioProd, DataFimProd, Prioridade, QuantidadeProd, Situacao, Destinacao, NumIntEquipamento, Versao, NumIntDocPai, JustificativaBloqueio FROM ItensOrdemProducao WHERE FilialEmpresa = ? AND Codigo = ? AND Item = ?", _
        tItemOP.lNumIntDoc, tItemOP.sProduto, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.lNumIntOrigem, tItemOP.iAlmoxarifado, tItemOP.sCcl, tItemOP.sSiglaUM, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.dQuantidadeProd, tItemOP.iSituacao, tItemOP.iDestinacao, tItemOP.lNumIntEquipamento, tItemOP.sVersao, tItemOP.lNumIntDocPai, tItemOP.sJustificativaBloqueio, objItemOP.iFilialEmpresa, objItemOP.sCodigo, objItemOP.iItem)
    If lErro <> AD_SQL_SUCESSO Then gError 33213
    '########################################################

    'Verifica se o item da OP existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33214

    'Não encontrou o item da OP ==> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 33215

    'Preenche objItemOP
    objItemOP.lNumIntDoc = tItemOP.lNumIntDoc
    objItemOP.sProduto = tItemOP.sProduto
    objItemOP.iFilialPedido = tItemOP.iFilialPedido
    objItemOP.lCodPedido = tItemOP.lCodPedido
    objItemOP.lNumIntOrigem = tItemOP.lNumIntOrigem
    objItemOP.iAlmoxarifado = tItemOP.iAlmoxarifado
    objItemOP.sCcl = tItemOP.sCcl
    objItemOP.sSiglaUM = tItemOP.sSiglaUM
    objItemOP.dQuantidade = tItemOP.dQuantidade
    objItemOP.dtDataInicioProd = tItemOP.dtDataInicioProd
    objItemOP.dtDataFimProd = tItemOP.dtDataFimProd
    objItemOP.iPrioridade = tItemOP.iPrioridade
    objItemOP.dQuantidadeProd = tItemOP.dQuantidadeProd
    objItemOP.iSituacao = tItemOP.iSituacao
    objItemOP.iDestinacao = tItemOP.iDestinacao
    
    objItemOP.lNumIntEquipamento = tItemOP.lNumIntEquipamento
    objItemOP.sVersao = tItemOP.sVersao

    '########################################################
    'INSERIDO POR WAGNER
    objItemOP.lNumIntDocPai = tItemOP.lNumIntDocPai
    objItemOP.sJustificativaBloqueio = tItemOP.sJustificativaBloqueio
    '########################################################

    'Fechamento comando
    Call Comando_Fechar(lComando)

    ItemOP_Le_Item = SUCESSO

    Exit Function

Erro_ItemOP_Le_Item:

    ItemOP_Le_Item = gErr

    Select Case gErr

        Case 33212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 33213, 33214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO", gErr, objItemOP.iFilialEmpresa, objItemOP.sCodigo, objItemOP.iItem)

        Case 33215
            'Não encontrou o item da OP.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151985)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Estoque_Grava(ByVal objProdutoFilial As ClassProdutoFilial, iControleEstoque As Integer) As Long
'Grava o Estoque do Produto da Filial

    'SENDO CRIADA EM Estoque.Frm

End Function

Function MovEstoqueItens_Le(objMovEstoque As ClassMovEstoque) As Long
'Lê todos os Itens correspondentes a partir da FilialEmpresa, Código e do NumIntDocEst=0 e retorna no colItens do objeto
'Ordena pelo Item.

Dim lErro As Long
Dim lComando As Long
Dim dtData As Date
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovEstoqueItens_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 30112

    With tItemMovEstoque

        .sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
        .sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sProdutoOP = String(STRING_PRODUTO, 0)
        .sOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        .sContaContabilEst = String(STRING_CONTA, 0)
        .sContaContabilAplic = String(STRING_CONTA, 0)
        .sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)

        '##########################################
        'Inserido por Wagner
        .sObservacao = String(STRING_MOVESTOQUE_OBSERVACAO, 0)
        '##########################################

        '##########################################
        'Alterado por Wagner
        lErro = Comando_Executar(lComando, "SELECT MovimentoEstoque.ContaContabilEst, MovimentoEstoque.ContaContabilAplic, MovimentoEstoque.Data, MovimentoEstoque.NumIntDoc, MovimentoEstoque.Custo, MovimentoEstoque.Apropriacao, MovimentoEstoque.Produto, Produtos.Descricao, MovimentoEstoque.SiglaUM, MovimentoEstoque.Quantidade, MovimentoEstoque.Almoxarifado, Almoxarifado.NomeReduzido, MovimentoEstoque.NumIntDocOrigem, MovimentoEstoque.Ccl, MovimentoEstoque.NumIntDocEst, MovimentoEstoque.TipoMov,MovimentoEstoque.DocOrigem,MovimentoEstoque.Cliente,MovimentoEstoque.Fornecedor, MovimentoEstoque.Observacao, MovimentoEstoque.Requisitante FROM MovimentoEstoque, Produtos, Almoxarifado WHERE Almoxarifado.Codigo = MovimentoEstoque.Almoxarifado AND Produtos.Codigo = MovimentoEstoque.Produto AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.Codigo = ? AND NumIntDocEst = 0 ORDER BY NumIntDoc", _
        .sContaContabilEst, .sContaContabilAplic, dtData, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .iTipoMov, .sDocOrigem, .lCliente, .lFornecedor, .sObservacao, .lRequisitante, objMovEstoque.iFilialEmpresa, objMovEstoque.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 30113
        '##########################################

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30114

        Do While lErro <> AD_SQL_SEM_DADOS

            objMovEstoque.dtData = dtData

            '##########################################
            'Inserido por Wagner
            objMovEstoque.sObservacao = .sObservacao
            objMovEstoque.lRequisitante = .lRequisitante
            '##########################################

            objMovEstoque.colItens.Add .lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, 0, Nothing, Nothing, DATA_NULA
            objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).sDocOrigem = .sDocOrigem
            objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lCliente = .lCliente
            objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lFornecedor = .lFornecedor

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30115

        Loop

    End With

    If objMovEstoque.colItens.Count = 0 Then Error 30116

    Call Comando_Fechar(lComando)

    MovEstoqueItens_Le = SUCESSO

    Exit Function

Erro_MovEstoqueItens_Le:

    MovEstoqueItens_Le = Err

    Select Case Err

        Case 30112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30113, 30114, 30115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)

        Case 30116 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151986)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovEstoqueInternoItens_Le(ByVal objMovEstoque As ClassMovEstoque) As Long
'Lê todos os Itens correspondentes a partir da FilialEmpresa, Código e do NumIntDocEst=0 e retorna no colItens do objeto
'Ordena pelo Item.

Dim lErro As Long
Dim lComando As Long
Dim dtData As Date
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovEstoqueInternoItens_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 30561

    With tItemMovEstoque

        .sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
        .sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)

        'Usa STRING_BUFFER_MAX_TEXTO para todos os campos da Query MovEstoqueInterno_OP
        .sProdutoOP = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sOPCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUM = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCcl = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilEst = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilAplic = String(STRING_BUFFER_MAX_TEXTO, 0)

        '##################################
        'Inserido por Wagner
        .sObservacao = String(STRING_MOVESTOQUE_OBSERVACAO, 0)
        '##################################
        
        '##################################
        'Alterado por Wagner
        lErro = Comando_Executar(lComando, "SELECT MovEstoqueInterno_OP.ContaContabilEst, MovEstoqueInterno_OP.ContaContabilAplic, MovEstoqueInterno_OP.Data, MovEstoqueInterno_OP.NumIntDoc, MovEstoqueInterno_OP.Custo, MovEstoqueInterno_OP.Apropriacao, MovEstoqueInterno_OP.Produto, Produtos.Descricao, MovEstoqueInterno_OP.SiglaUM, MovEstoqueInterno_OP.Quantidade, MovEstoqueInterno_OP.Almoxarifado, Almoxarifado.NomeReduzido, MovEstoqueInterno_OP.NumIntDocOrigem, MovEstoqueInterno_OP.Ccl, MovEstoqueInterno_OP.NumIntDocEst, MovEstoqueInterno_OP.TipoMov, MovEstoqueInterno_OP.ProdutoOP, MovEstoqueInterno_OP.OPCodigo, MovEstoqueInterno_OP.Observacao, MovEstoqueInterno_OP.Requisitante FROM  MovEstoqueInterno_OP, Produtos, Almoxarifado WHERE Almoxarifado.Codigo = MovEstoqueInterno_OP.Almoxarifado AND Produtos.Codigo = MovEstoqueInterno_OP.Produto AND MovEstoqueInterno_OP.FilialEmpresa = ? AND MovEstoqueInterno_OP.Codigo = ? AND NumIntDocEst = 0 ORDER BY NumIntDoc", _
        .sContaContabilEst, .sContaContabilAplic, dtData, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .iTipoMov, .sProdutoOP, .sOPCodigo, .sObservacao, .lRequisitante, giFilialEmpresa, objMovEstoque.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 30562
        '##################################

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30563

        Do While lErro <> AD_SQL_SEM_DADOS

            '##################################
            'Inserido por Wagner
            objMovEstoque.sObservacao = .sObservacao
            objMovEstoque.lRequisitante = .lRequisitante
            '##################################
            
            objMovEstoque.dtData = dtData

            objMovEstoque.colItens.Add .lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, 0, Nothing, Nothing, DATA_NULA

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30564

        Loop

    End With

    If objMovEstoque.colItens.Count = 0 Then Error 30565

    Call Comando_Fechar(lComando)

    MovEstoqueInternoItens_Le = SUCESSO

    Exit Function

Erro_MovEstoqueInternoItens_Le:

    MovEstoqueInternoItens_Le = Err

    Select Case Err

        Case 30561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30562, 30563, 30564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)

        Case 30565 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151987)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovEstoqueInterno2Itens_Le(ByVal objMovEstoque As ClassMovEstoque) As Long
'Lê todos os Itens correspondentes a partir da FilialEmpresa, Código e do NumIntDocEst=0 e retorna no colItens do objeto
'Ordena pelo Item.

Dim lErro As Long
Dim lComando As Long
Dim dtData As Date
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovEstoqueInterno2Itens_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 30561

    With tItemMovEstoque

        .sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
        .sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)

        'Usa STRING_BUFFER_MAX_TEXTO para todos os campos da Query MovEstoqueInterno_OS
        .sProdutoOP = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sOPCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUM = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCcl = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilEst = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilAplic = String(STRING_BUFFER_MAX_TEXTO, 0)

        '##################################
        'Inserido por Wagner
        .sObservacao = String(STRING_MOVESTOQUE_OBSERVACAO, 0)
        '##################################
        
        '##################################
        'Alterado por Wagner
        lErro = Comando_Executar(lComando, "SELECT MovEstoqueInterno_OS.ContaContabilEst, MovEstoqueInterno_OS.ContaContabilAplic, MovEstoqueInterno_OS.Data, MovEstoqueInterno_OS.NumIntDoc, MovEstoqueInterno_OS.Custo, MovEstoqueInterno_OS.Apropriacao, MovEstoqueInterno_OS.Produto, Produtos.Descricao, MovEstoqueInterno_OS.SiglaUM, MovEstoqueInterno_OS.Quantidade, MovEstoqueInterno_OS.Almoxarifado, Almoxarifado.NomeReduzido, MovEstoqueInterno_OS.NumIntDocOrigem, MovEstoqueInterno_OS.Ccl, MovEstoqueInterno_OS.NumIntDocEst, MovEstoqueInterno_OS.TipoMov, MovEstoqueInterno_OS.ServicoOS, MovEstoqueInterno_OS.OSCodigo, MovEstoqueInterno_OS.Observacao, MovEstoqueInterno_OS.Requisitante FROM  MovEstoqueInterno_OS, Produtos, Almoxarifado WHERE Almoxarifado.Codigo = MovEstoqueInterno_OS.Almoxarifado AND Produtos.Codigo = MovEstoqueInterno_OS.Produto AND MovEstoqueInterno_OS.FilialEmpresa = ? AND MovEstoqueInterno_OS.Codigo = ? AND NumIntDocEst = 0 ORDER BY NumIntDoc", _
        .sContaContabilEst, .sContaContabilAplic, dtData, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .iTipoMov, .sProdutoOP, .sOPCodigo, .sObservacao, .lRequisitante, giFilialEmpresa, objMovEstoque.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 30562
        '##################################

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30563

        Do While lErro <> AD_SQL_SEM_DADOS

            '##################################
            'Inserido por Wagner
            objMovEstoque.sObservacao = .sObservacao
            objMovEstoque.lRequisitante = .lRequisitante
            '##################################
            
            objMovEstoque.dtData = dtData

            objMovEstoque.colItens.Add .lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, 0, Nothing, Nothing, DATA_NULA

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30564

        Loop

    End With

    If objMovEstoque.colItens.Count = 0 Then Error 30565

    Call Comando_Fechar(lComando)

    MovEstoqueInterno2Itens_Le = SUCESSO

    Exit Function

Erro_MovEstoqueInterno2Itens_Le:

    MovEstoqueInterno2Itens_Le = Err

    Select Case Err

        Case 30561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30562, 30563, 30564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)

        Case 30565 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151987)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoMovEstoque_Le(ByVal objTipoMovEstoque As ClassTipoMovEst) As Long
'Lê o Tipo de MovEstoque a partir do Código

Dim lErro As Long

On Error GoTo Erro_TipoMovEstoque_Le

    Call objTipoMovEstoque.Copiar(gobjMAT.colTiposMovEst(CStr(objTipoMovEstoque.iCodigo)))
    
    TipoMovEstoque_Le = SUCESSO

    Exit Function

Erro_TipoMovEstoque_Le:

    TipoMovEstoque_Le = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151988)

    End Select

    Exit Function

End Function

Function MovimentoEstoque_Le_Codigo(ByVal iFilialEmpresa As Integer, ByVal lCodigo As Long, dtData As Date) As Long
'Verifica se o filial/codigo já está cadastrado
    
Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_MovimentoEstoque_Le_Codigo

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36012

    'le o codigo do movimento de estoque passado como parametro
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Data FROM MovimentoEstoque WHERE FilialEmpresa = ? AND Codigo = ?", lNumIntDoc, dtData, iFilialEmpresa, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 36013

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36014

    'se o movimento não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 36015

    Call Comando_Fechar(lComando)

    MovimentoEstoque_Le_Codigo = SUCESSO
    
    Exit Function

Erro_MovimentoEstoque_Le_Codigo:

    MovimentoEstoque_Le_Codigo = Err

    Select Case Err

        Case 36012
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 36013, 36014
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE1", Err, iFilialEmpresa, lCodigo)

        Case 36015
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151989)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovimentoEstoque_Le_Ultima_Data(ByVal iFilialEmpresa As Integer, ByVal objItemMovEstoque As ClassItemMovEstoque) As Long
'le a ultima data de movimento cadastrada e a coloca em objItemMovEstoque

' *** Função alterada em 03/08/2001 por Luiz Gustavo de Freitas Nogueira ***
' *** Além de retornar a última data de movimento cadastrada, a função passa a retornar ***
' *** a maior hora de movimento cadastrado para essa data ***

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovimentoEstoque_Le_Ultima_Data

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 36026

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 90590

    'le a ultima data cadastrada
    lErro = Comando_Executar(lComando, "SELECT MAX(Data) FROM MovimentoEstoque WHERE FilialEmpresa = ?", tItemMovEstoque.dtData, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 36027

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 36028

    'se o movimento não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Or tItemMovEstoque.dtData = DATA_NULA Then gError 36029
    
    objItemMovEstoque.dtData = tItemMovEstoque.dtData
    
    ' *** Incluído por Luiz Gustavo de Freitas Nogueira em 03/08/2001 ***
    'le a última hora em que foi gravado um movimento na última data lida
    lErro = Comando_Executar(lComando1, "SELECT MAX(Hora) FROM MovimentoEstoque WHERE FilialEmpresa = ? and Data = ?", tItemMovEstoque.dHora, iFilialEmpresa, objItemMovEstoque.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 90591
    
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90592

    'se o movimento não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 90593
    
    'Guarda no obj a hora encontrada
    objItemMovEstoque.dtHora = tItemMovEstoque.dHora
    ' ********************************************************************
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    MovimentoEstoque_Le_Ultima_Data = SUCESSO
    
    Exit Function

Erro_MovimentoEstoque_Le_Ultima_Data:

    MovimentoEstoque_Le_Ultima_Data = gErr

    Select Case gErr

        Case 36026, 90590
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 36027, 36028, 90591, 90592
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", Err, iFilialEmpresa)

        Case 36029, 90593
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 151990)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Produto_Le_ControleEstoque(ByVal lComando As Long, ByVal objProduto As ClassProduto) As Long
'Lê o atributo controle de estoque do produto, cujo codigo foi passado como parâmetro

Dim lErro As Long
Dim tProduto As typeProduto

On Error GoTo Erro_Produto_Le_ControleEstoque

    tProduto.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)

    lErro = Comando_ExecutarLockado(lComando, "SELECT ControleEstoque, SiglaUMEstoque, ClasseUM, Apropriacao,TempoProducao, PrazoValidade, Rastro, Compras FROM Produtos WHERE Codigo = ? ", tProduto.iControleEstoque, tProduto.sSiglaUMEstoque, tProduto.iClasseUM, tProduto.iApropriacaoCusto, tProduto.iTempoProducao, tProduto.iPrazoValidade, tProduto.iRastro, tProduto.iCompras, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 36033

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then
        Error 36034
    End If

    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36572

    objProduto.iControleEstoque = tProduto.iControleEstoque
    objProduto.sSiglaUMEstoque = tProduto.sSiglaUMEstoque
    objProduto.iClasseUM = tProduto.iClasseUM
    objProduto.iApropriacaoCusto = tProduto.iApropriacaoCusto
    objProduto.iTempoProducao = tProduto.iTempoProducao
    objProduto.iPrazoValidade = tProduto.iPrazoValidade
    objProduto.iRastro = tProduto.iRastro
    objProduto.iCompras = tProduto.iCompras
    
    Produto_Le_ControleEstoque = SUCESSO

    Exit Function

Erro_Produto_Le_ControleEstoque:

    Produto_Le_ControleEstoque = Err

    Select Case Err

        Case 36033, 36034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", Err, objProduto.sCodigo)

        Case 36572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151991)

    End Select

    Exit Function

End Function

Function UM_Le1(ByVal lComando As Long, ByVal objUnidadeMedida As ClassUnidadeDeMedida) As Long
'Lê a unidade de medida a partir da Classe e da Sigla -->
'e devolve dados no Objeto

Dim lErro As Long
Dim tUnidMed As typeUnidadeMedida

On Error GoTo Erro_UM_Le1

    tUnidMed.sNome = String(STRING_UM_NOME, 0)
    tUnidMed.sSiglaUMBase = String(STRING_UM_SIGLA_UM_BASE, 0)

    lErro = Comando_Executar(lComando, "SELECT Nome, SiglaUMBase, Quantidade FROM UnidadesdeMedida WHERE Classe=? AND Sigla=?", tUnidMed.sNome, tUnidMed.sSiglaUMBase, tUnidMed.dQuantidade, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 36109

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO Then Error 36110

    With tUnidMed

        objUnidadeMedida.dQuantidade = .dQuantidade
        objUnidadeMedida.sNome = .sNome
        objUnidadeMedida.sSiglaUMBase = .sSiglaUMBase
        
    End With

    UM_Le1 = SUCESSO

    Exit Function

Erro_UM_Le1:

    UM_Le1 = Err

    Select Case Err

        Case 36109, 36110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_UNIDADESDEMEDIDA", Err, objUnidadeMedida.iClasse, objUnidadeMedida.sSigla)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151992)

    End Select

    Exit Function

End Function

Function ProdutoFilial_Le1(ByVal lComando As Long, ByVal objProdutoFilial As ClassProdutoFilial) As Long
'Lê os dados do Produto da Filial à partir da chave fornecida em objProdutoFilial. Retorna os dados em objProdutoFilial

Dim lErro As Long
Dim tProdutoFilial As typeProdutoFilial

On Error GoTo Erro_ProdutoFilial_Le1

    tProdutoFilial.sClasseABC = String(STRING_PRODUTOFILIAL_CLASSEABC, 0)
    tProdutoFilial.sSituacaoTribECF = String(STRING_PRODUTOFILIAL_SITUACAOTRIBECF, 0)
    tProdutoFilial.sICMSAliquota = String(STRING_PRODUTOFILIAL_ICMS, 0)

    'Pesquisa no BD ProdutoFilial
    With tProdutoFilial
        lErro = Comando_Executar(lComando, "SELECT Almoxarifado, Fornecedor, VisibilidadeAlmoxarifados, EstoqueSeguranca, EstoqueMaximo, PontoPedido, ClasseABC, LoteEconomico, IntRessup, SituacaoTribECF, ICMSAliquota FROM ProdutosFilial WHERE FilialEmpresa= ? AND Produto = ?", _
            .iAlmoxarifado, .lFornecedor, .iVisibilidadeAlmoxarifados, .dEstoqueSeguranca, .dEstoqueMaximo, .dPontoPedido, .sClasseABC, .dLoteEconomico, .iIntRessup, .sSituacaoTribECF, .sICMSAliquota, objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)
        If lErro <> AD_SQL_SUCESSO Then Error 36222
    End With

    'Tenta selecionar ProdutoFilial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36223

    'Preenche objProdutoFilial
    With tProdutoFilial
        objProdutoFilial.iAlmoxarifado = .iAlmoxarifado
        objProdutoFilial.lFornecedor = .lFornecedor
        objProdutoFilial.iVisibilidadeAlmoxarifados = .iVisibilidadeAlmoxarifados
        objProdutoFilial.dEstoqueSeguranca = .dEstoqueSeguranca
        objProdutoFilial.dEstoqueMaximo = .dEstoqueMaximo
        objProdutoFilial.dPontoPedido = .dPontoPedido
        objProdutoFilial.sClasseABC = .sClasseABC
        objProdutoFilial.dLoteEconomico = .dLoteEconomico
        objProdutoFilial.iIntRessup = .iIntRessup
        objProdutoFilial.sSituacaoTribECF = .sSituacaoTribECF
        objProdutoFilial.sICMSAliquota = .sICMSAliquota
    End With

    ProdutoFilial_Le1 = SUCESSO

    Exit Function

Erro_ProdutoFilial_Le1:

    ProdutoFilial_Le1 = Err

    Select Case Err

        Case 36222, 36223
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err, objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151993)

    End Select

    Exit Function

End Function

Function MovEstoque_Le(ByVal objMovEstoque As ClassMovEstoque) As Long
'Lê o MovEstoque a partir do Código.
'A função só pesquisa os MovEstoque com FilialEmpresa=giFilialEmpresa e NumIntDocEst=0

Dim lErro As Long
Dim lComando As Long
Dim tMovEstoque As typeMovEst
Dim dHora As Double

On Error GoTo Erro_MovEstoque_Le

    'Abertura de comando
    lComando = Comando_Abrir
    If lComando = 0 Then Error 30125

    With tMovEstoque

        '#####################################
        'Inserido por Wagner
        .sObservacao = String(STRING_MOVESTOQUE_OBSERVACAO, 0)
        '#####################################

        '#####################################
        'Alterado por Wagner
        lErro = Comando_Executar(lComando, "SELECT TipoMov, Data, NumIntDoc, Hora, Observacao, Requisitante FROM MovimentoEstoque WHERE FilialEmpresa = ? AND Codigo = ? AND NumIntDocEst = 0 ORDER BY NumIntDoc", .iTipoMov, .dtData, .lNumIntDoc, .dHora, .sObservacao, .lRequisitante, objMovEstoque.iFilialEmpresa, objMovEstoque.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 30126
        '#####################################

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30127

        If lErro = AD_SQL_SEM_DADOS Then Error 30128

        objMovEstoque.iTipoMov = .iTipoMov
        objMovEstoque.dtData = .dtData
        objMovEstoque.dtHora = .dHora
        objMovEstoque.lNumIntDoc = .lNumIntDoc

        '#####################################
        'Inserido por Wagner
        objMovEstoque.sObservacao = .sObservacao
        objMovEstoque.lRequisitante = .lRequisitante
        '#####################################

    End With

    Call Comando_Fechar(lComando)

    MovEstoque_Le = SUCESSO

    Exit Function

Erro_MovEstoque_Le:

    MovEstoque_Le = Err

    Select Case Err

        Case 30125
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30126, 30127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)

        Case 30128 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151994)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TiposMovEstoque_Le(ByVal colTipoMovEst As colTipoMovEst) As Long
'Lê os Tipos de MovEstoque

Dim lErro As Long
Dim lComando As Long
Dim tTipoMovEst As typeTipoMovEst
Dim objTipoMovEst As ClassTipoMovEst

On Error GoTo Erro_TiposMovEstoque_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 30348

    With tTipoMovEst
        .sDescricao = String(STRING_DESCRICAO_TIPOMOVEST, 0)
        .sEntradaOuSaida = String(STRING_ENTRADAOUSAIDA, 0)
        .sSigla = String(STRING_SIGLA_TIPOMOVEST, 0)
        .sNomeTela = String(STRING_NOME_TELA, 0)
        .sDescrNumIntDocOrigem = String(STRING_DESCR_NUMINTDOCORIG_TIPOMOVEST, 0)
        .sEntradaSaidaCMP = String(STRING_ENTRADAOUSAIDA, 0)
        
        lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, EntradaOuSaida, Inativo, AtualizaConsumo, AtualizaVenda, AtualizaVendaConsig3, Editavel, ValidoMovInt, AtualizaCusto, AtualizaConsig, AtualizaDemo, AtualizaConserto, ProdutodeTerc, AtualizaOutras, AtualizaIndOutras, AtualizaNossaDisp, AtualizaDefeituosa, AtualizaRecebIndisp, AtualizaCMRProd, AtualizaBenef, CustoInformado, ValidaDataUltMov, DescrNumIntDocOrigem, TipoNumIntDocOrigem, AtualizaOP, AtualizaCompra, Kardex, Sigla, NomeTela, LivroMod3, AtualizaMovEstoque, AtualizaReserva, Transferencia, AtualizaSaldoCusto, CustoMedio, EntradaSaidaCMP, " & _
                                  " AtualizaCustoConsig, AtualizaCustoDemo, AtualizaCustoConserto, AtualizaCustoOutros, AtualizaCustoBenef, AtualizaCustoConsig3, AtualizaCustoDemo3, AtualizaCustoConserto3, AtualizaCustoOutros3, AtualizaCustoBenef3, CodigoOrig, AtualizaSoLote, Inventario, NFDevolucao FROM TiposMovimentoEstoque ORDER BY Codigo", _
                                .iCodigo, .sDescricao, .sEntradaOuSaida, .iInativo, .iAtualizaConsumo, .iAtualizaVenda, .iAtualizaVendaConsig3, .iEditavel, .iValidoMovInt, .iAtualizaCusto, .iAtualizaConsig, .iAtualizaDemo, .iAtualizaConserto, .iProdutoDeTerc, .iAtualizaOutras, .iAtualizaIndOutras, .iAtualizaNossaDisp, .iAtualizaDefeituosa, .iAtualizaRecebIndisp, .iAtualizaCMRProd, .iAtualizaBenef, .iCustoInformado, .iValidaDataUltMov, .sDescrNumIntDocOrigem, .iTipoNumIntDocOrigem, .iAtualizaOP, .iAtualizaCompra, .iKardex, .sSigla, .sNomeTela, .iLivroMod3, .iAtualizaMovEstoque, .iAtualizaReserva, .iTransferencia, .iAtualizaSaldoCusto, .iCustoMedio, .sEntradaSaidaCMP, .iAtualizaCustoConsig, .iAtualizaCustoDemo, .iAtualizaCustoConserto, .iAtualizaCustoOutros, .iAtualizaCustoBenef, .iAtualizaCustoConsig3, .iAtualizaCustoDemo3, .iAtualizaCustoConserto3, .iAtualizaCustoOutros3, .iAtualizaCustoBenef3, .iCodigoOrig, .iAtualizaSoLote, .iInventario, .iNFDevolucao)
        If lErro <> AD_SQL_SUCESSO Then Error 30349

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30350

        While lErro <> AD_SQL_SEM_DADOS

            Set objTipoMovEst = colTipoMovEst.Add(.iCodigo, .sDescricao, .sEntradaOuSaida, .iInativo, .iAtualizaConsumo, .iAtualizaVenda, .iEditavel, .iValidoMovInt, .iAtualizaCusto, .iAtualizaConsig, .iAtualizaDemo, .iAtualizaConserto, .iProdutoDeTerc, .iAtualizaOutras, .iAtualizaIndOutras, .iAtualizaNossaDisp, .iAtualizaDefeituosa, .iAtualizaRecebIndisp, .iTransferencia, .iAtualizaCMRProd, .iAtualizaBenef, .iCustoInformado, .iValidaDataUltMov, .iAtualizaMovEstoque, .iAtualizaReserva, .iAtualizaCompra, .sDescrNumIntDocOrigem, .iTipoNumIntDocOrigem, .iAtualizaOP, .iKardex, .sSigla, .sNomeTela, .iLivroMod3, .sEntradaSaidaCMP, .iNFDevolucao, CStr(.iCodigo))

            objTipoMovEst.iAtualizaVendaConsig3 = .iAtualizaVendaConsig3
            objTipoMovEst.iAtualizaSaldoCusto = .iAtualizaSaldoCusto
            objTipoMovEst.iCustoMedio = .iCustoMedio
            objTipoMovEst.iAtualizaCustoConsig = .iAtualizaCustoConsig
            objTipoMovEst.iAtualizaCustoDemo = .iAtualizaCustoDemo
            objTipoMovEst.iAtualizaCustoConserto = .iAtualizaCustoConserto
            objTipoMovEst.iAtualizaCustoOutros = .iAtualizaCustoOutros
            objTipoMovEst.iAtualizaCustoBenef = .iAtualizaCustoBenef
            objTipoMovEst.iAtualizaCustoConsig3 = .iAtualizaCustoConsig3
            objTipoMovEst.iAtualizaCustoDemo3 = .iAtualizaCustoDemo3
            objTipoMovEst.iAtualizaCustoConserto3 = .iAtualizaCustoConserto3
            objTipoMovEst.iAtualizaCustoOutros3 = .iAtualizaCustoOutros3
            objTipoMovEst.iAtualizaCustoBenef3 = .iAtualizaCustoBenef3
            objTipoMovEst.iCodigoOrig = .iCodigoOrig
            objTipoMovEst.iAtualizaSoLote = .iAtualizaSoLote
            objTipoMovEst.iInventario = .iInventario
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30351

        Wend

    End With

    'Verifica se há Algum tipo de Movimento de Estoque cadastrado
    If colTipoMovEst.Count = 0 Then Error 30352

    Call Comando_Fechar(lComando)

    TiposMovEstoque_Le = SUCESSO

    Exit Function

Erro_TiposMovEstoque_Le:

    TiposMovEstoque_Le = Err

    Select Case Err

        Case 30348
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30349, 30350, 30351
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOMOVEST", Err)

        Case 30352 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151995)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OrdemProducao_Le(ByVal objOrdemProducao As ClassOrdemDeProducao) As Long
'Lê a Ordem de Produção pelo Código e pela FilialEmpresa

Dim lErro As Long
Dim lComando As Long
Dim tOrdemProducao As typeOrdemProducao

On Error GoTo Erro_OrdemProducao_Le

    If objOrdemProducao.iFilialEmpresa = 0 Then objOrdemProducao.iFilialEmpresa = giFilialEmpresa
    
    'Abertura de comando
    lComando = Comando_Abrir
    If lComando = 0 Then Error 30365

    With tOrdemProducao

        lErro = Comando_Executar(lComando, "SELECT DataEmissao, NumItens, NumItensBaixados, CodPrestador, Tipo, TipoTerc, CodTerc, FilialTerc, EscaninhoTerc, IgnoraEst FROM OrdensDeProducao WHERE FilialEmpresa = ? AND Codigo = ?", .dtDataEmissao, .iNumItens, .iNumItensBaixados, .lCodPrestador, .iTipo, .iTipoTerc, .lCodTerc, .iFilialTerc, .iEscaninhoTerc, .iIgnoraEst, objOrdemProducao.iFilialEmpresa, objOrdemProducao.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 30366

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30367

        If lErro = AD_SQL_SEM_DADOS Then
        
            lErro = Comando_Executar(lComando, "SELECT DataEmissao, NumItens, NumItensBaixados, Tipo, TipoTerc, CodTerc, FilialTerc, EscaninhoTerc, IgnoraEst FROM OrdensDeProducaoBaixadas WHERE FilialEmpresa = ? AND Codigo = ?", .dtDataEmissao, .iNumItens, .iNumItensBaixados, .iTipo, .iTipoTerc, .lCodTerc, .iFilialTerc, .iEscaninhoTerc, .iIgnoraEst, objOrdemProducao.iFilialEmpresa, objOrdemProducao.sCodigo)
            If lErro <> AD_SQL_SUCESSO Then Error 55314
    
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55315
        
            If lErro = AD_SQL_SUCESSO Then Error 55316
        
            Error 30368
            
        End If

        objOrdemProducao.iNumItens = .iNumItens
        objOrdemProducao.dtDataEmissao = .dtDataEmissao
        objOrdemProducao.iNumItensBaixados = .iNumItensBaixados
        objOrdemProducao.lCodPrestador = .lCodPrestador
        objOrdemProducao.iTipo = .iTipo
        objOrdemProducao.iTipoTerc = .iTipoTerc
        objOrdemProducao.lCodTerc = .lCodTerc
        objOrdemProducao.iFilialTerc = .iFilialTerc
        objOrdemProducao.iEscaninhoTerc = .iEscaninhoTerc
        objOrdemProducao.iIgnoraEst = .iIgnoraEst
        
    End With

    Call Comando_Fechar(lComando)

    OrdemProducao_Le = SUCESSO

    Exit Function

Erro_OrdemProducao_Le:

    OrdemProducao_Le = Err

    Select Case Err

        Case 30365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 30366, 30367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAO", Err)

        Case 30368, 55316 'Tratado na Rotina Chamadora

        Case 55314, 55315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAOBAIXADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 151996)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensOrdemProducao_Le(ByVal objOrdemProducao As ClassOrdemDeProducao) As Long
'Lê os Itens Ordem de Produção a partir do OPCódigo e da FilialEmpresa
'Retorna os Itens na coleção

Dim lErro As Long
Dim lComando As Long
Dim tOrdemProducao As typeOrdemProducao
Dim objItemOP As ClassItemOP
Dim tItemOP As typeItemOP

On Error GoTo Erro_ItensOrdemProducao_Le

    'Abertura de comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 30397

    With tItemOP

        .sCcl = String(STRING_CCL, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)
        .sVersao = String(STRING_KIT_VERSAO, 0)

        '##############################################################
        'INSERIDO POR WAGNER
        .sJustificativaBloqueio = String(STRING_MAX_NOME_ARQUIVO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        '##############################################################

        '##############################################################
        'ALTERADO POR WAGNER
        lErro = Comando_Executar(lComando, "SELECT I.Quantidade ,I.QuantidadeProd ,I.DataFimProd ,I.DataInicioProd, I.Almoxarifado, I.Destinacao, I.FilialPedido, I.Item, I.NumIntOrigem, I.Prioridade, I.Situacao, I.CodPedido, I.NumIntDoc, I.Ccl, I.Produto, I.SiglaUM, I.NumIntEquipamento, I.Versao, I.NumIntDocPai, I.JustificativaBloqueio, P.Descricao, I.Beneficiamento FROM ItensOrdemProducao AS I, Produtos AS P WHERE I.Produto=P.Codigo AND I.FilialEmpresa = ? AND I.Codigo = ? ", _
        .dQuantidade, .dQuantidadeProd, .dtDataFimProd, .dtDataInicioProd, .iAlmoxarifado, .iDestinacao, .iFilialPedido, .iItem, .lNumIntOrigem, .iPrioridade, .iSituacao, .lCodPedido, .lNumIntDoc, .sCcl, .sProduto, .sSiglaUM, .lNumIntEquipamento, .sVersao, .lNumIntDocPai, .sJustificativaBloqueio, .sDescricao, .iBeneficiamento, objOrdemProducao.iFilialEmpresa, objOrdemProducao.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 30398
        '##############################################################

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 30399

        While lErro <> AD_SQL_SEM_DADOS

            Set objItemOP = New ClassItemOP
            
            objItemOP.iFilialEmpresa = objOrdemProducao.iFilialEmpresa
            objItemOP.dQuantidade = .dQuantidade
            objItemOP.dQuantidadeProd = .dQuantidadeProd
            objItemOP.dtDataFimProd = .dtDataFimProd
            objItemOP.dtDataInicioProd = .dtDataInicioProd
            objItemOP.iAlmoxarifado = .iAlmoxarifado
            objItemOP.iDestinacao = .iDestinacao
            objItemOP.iFilialPedido = .iFilialPedido
            objItemOP.iItem = .iItem
            objItemOP.lNumIntOrigem = .lNumIntOrigem
            objItemOP.iPrioridade = .iPrioridade
            objItemOP.iSituacao = .iSituacao
            objItemOP.lCodPedido = .lCodPedido
            objItemOP.lNumIntDoc = .lNumIntDoc
            objItemOP.sCcl = .sCcl
            objItemOP.sProduto = .sProduto
            objItemOP.sSiglaUM = .sSiglaUM
            objItemOP.iBeneficiamento = .iBeneficiamento
            
            objItemOP.lNumIntEquipamento = .lNumIntEquipamento
            objItemOP.sVersao = .sVersao
            
            '###################################################
            'INSERIDO POR WAGNER
            objItemOP.lNumIntDocPai = .lNumIntDocPai
            objItemOP.sJustificativaBloqueio = .sJustificativaBloqueio
            objItemOP.sDescricao = .sDescricao
            '###################################################

            'Inserido por Jorge Specian - 14/06/2005
            '---------------------------------------
            lErro = CF("OrdemDeProducao_Le_Operacoes", objItemOP)
            If lErro <> SUCESSO And lErro <> 137016 Then gError 137185
            '---------------------------------------

            objOrdemProducao.colItens.Add objItemOP

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 30400

        Wend

    End With

    'Verifica se há algum Item para nessa Ordem de Produção
    If objOrdemProducao.colItens.Count = 0 Then gError 30401

    lErro = Comando_Fechar(lComando)

    ItensOrdemProducao_Le = SUCESSO

    Exit Function

Erro_ItensOrdemProducao_Le:

    ItensOrdemProducao_Le = gErr

    Select Case gErr

        Case 30397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 30398, 30399, 30400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOP", gErr)

        Case 30401 'Tratado na Rotina Chamadora

        Case 137185 'Tratado na Rotina Chamada - Inserido por Jorge Specian - 14/06/2005

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 151997)

    End Select

    Call Comando_Fechar(lComando)

End Function

Public Function ItemOP_Le(ByVal objItemOP As ClassItemOP) As Long
'Le de ItemOP com filtros objItemOP.iFilialEmpresa , objItemOP.sCodigo , objItemOP.sProduto

Dim lErro As Long
Dim lComando As Long
Dim sComando_SQL As String
Dim tItemOP As typeItemOP

On Error GoTo Erro_ItemOP_Le

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 34708

    With tItemOP

        .sCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sCcl = String(STRING_CCL, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)
        .sVersao = String(STRING_KIT_VERSAO, 0)
        
        '########################################################
        'INSERIDO POR WAGNER
        .sJustificativaBloqueio = String(STRING_MAX_NOME_ARQUIVO, 0)
        '########################################################
        

        'Pesquisa no BD o Item da OP passado como parâmetro
        '########################################################
        'ALTERADO POR WAGNER
        sComando_SQL = "SELECT NumIntDoc, FilialEmpresa, Codigo, Item, Produto, FilialPedido, CodPedido, NumIntOrigem, Almoxarifado, Ccl, SiglaUM, Quantidade, DataInicioProd, DataFimProd, Prioridade, QuantidadeProd, Situacao, Destinacao, NumIntEquipamento, Versao, NumIntDocPai, JustificativaBloqueio, Beneficiamento FROM ItensOrdemProducao WHERE Codigo = ? AND Produto = ? AND FilialEmpresa = ?"
        lErro = Comando_Executar(lComando, sComando_SQL, .lNumIntDoc, .iFilialEmpresa, .sCodigo, .iItem, .sProduto, .iFilialPedido, .lCodPedido, .lNumIntOrigem, .iAlmoxarifado, .sCcl, .sSiglaUM, .dQuantidade, .dtDataInicioProd, .dtDataFimProd, .iPrioridade, .dQuantidadeProd, .iSituacao, .iDestinacao, .lNumIntEquipamento, .sVersao, .lNumIntDocPai, .sJustificativaBloqueio, .iBeneficiamento, objItemOP.sCodigo, objItemOP.sProduto, objItemOP.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 34709
        '########################################################

        'Verifica se o item existe
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 34710

        'Não encontrou o item ==> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 34711

        'Preenche objItemOP
        objItemOP.lNumIntDoc = .lNumIntDoc
        objItemOP.iFilialEmpresa = .iFilialEmpresa
        objItemOP.sCodigo = .sCodigo
        objItemOP.iItem = .iItem
        objItemOP.sProduto = .sProduto
        objItemOP.iFilialPedido = .iFilialPedido
        objItemOP.lCodPedido = .lCodPedido
        objItemOP.lNumIntOrigem = .lNumIntOrigem
        objItemOP.iAlmoxarifado = .iAlmoxarifado
        objItemOP.sCcl = .sCcl
        objItemOP.sSiglaUM = .sSiglaUM
        objItemOP.dQuantidade = .dQuantidade
        objItemOP.dtDataInicioProd = .dtDataInicioProd
        objItemOP.dtDataFimProd = .dtDataFimProd
        objItemOP.iPrioridade = .iPrioridade
        objItemOP.dQuantidadeProd = .dQuantidadeProd
        objItemOP.iSituacao = .iSituacao
        objItemOP.iDestinacao = .iDestinacao
        objItemOP.iBeneficiamento = .iBeneficiamento
        
        objItemOP.lNumIntEquipamento = .lNumIntEquipamento
        objItemOP.sVersao = .sVersao
        
        '########################################################
        'INSERIDO POR WAGNER
        objItemOP.lNumIntDocPai = .lNumIntDocPai
        objItemOP.sJustificativaBloqueio = .sJustificativaBloqueio
        '########################################################


    End With

    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ItemOP_Le = SUCESSO

    Exit Function

Erro_ItemOP_Le:

    ItemOP_Le = gErr

    Select Case gErr

        Case 34708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 34709, 34710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO1", gErr, objItemOP.sCodigo, objItemOP.sProduto, objItemOP.iFilialEmpresa)

        Case 34711
            'Não encontrou o item

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151998)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemOP_Le_NumIntDoc(ByVal objItemOP As ClassItemOP, Optional ByVal bTodas As Boolean = False) As Long
'Lê os dados do item do empenho à partir de NumIntDoc

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim tItemOP As typeItemOP

On Error GoTo Erro_ItemOP_Le_NumIntDoc

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 33223

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 33223

    tItemOP.sCodigo = String(STRING_OPCODIGO, 0)
    tItemOP.sProduto = String(STRING_PRODUTO, 0)
    tItemOP.sCcl = String(STRING_CCL, 0)
    tItemOP.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemOP.sVersao = String(STRING_KIT_VERSAO, 0)
    tItemOP.sJustificativaBloqueio = String(STRING_MAX_NOME_ARQUIVO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, Item, Produto, FilialPedido, CodPedido, NumIntOrigem, Almoxarifado, Ccl, SiglaUM, Quantidade, DataInicioProd, DataFimProd, Prioridade, QuantidadeProd, Situacao, Destinacao, NumIntEquipamento, Versao, NumIntDocPai, JustificativaBloqueio FROM ItensOrdemProducao WHERE NumIntDoc = ?", _
        tItemOP.sCodigo, tItemOP.iItem, tItemOP.sProduto, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.lNumIntOrigem, tItemOP.iAlmoxarifado, tItemOP.sCcl, tItemOP.sSiglaUM, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.dQuantidadeProd, tItemOP.iSituacao, tItemOP.iDestinacao, tItemOP.lNumIntEquipamento, tItemOP.sVersao, tItemOP.lNumIntDocPai, tItemOP.sJustificativaBloqueio, objItemOP.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 33224

    'Verifica se o item do empenho existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33225

    'Não encontrou o item do empenho ==> Erro
    If lErro = AD_SQL_SEM_DADOS And Not bTodas Then gError 33226
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        tItemOP.sCodigo = String(STRING_OPCODIGO, 0)
        tItemOP.sProduto = String(STRING_PRODUTO, 0)
        tItemOP.sCcl = String(STRING_CCL, 0)
        tItemOP.sSiglaUM = String(STRING_UM_SIGLA, 0)
        tItemOP.sVersao = String(STRING_KIT_VERSAO, 0)
        tItemOP.sJustificativaBloqueio = String(STRING_MAX_NOME_ARQUIVO, 0)
        
        lErro = Comando_Executar(lComando, "SELECT Codigo, Item, Produto, FilialPedido, CodPedido, NumIntOrigem, Almoxarifado, Ccl, SiglaUM, Quantidade, DataInicioProd, DataFimProd, Prioridade, QuantidadeProd, Situacao, Destinacao, NumIntEquipamento, Versao, NumIntDocPai, JustificativaBloqueio FROM ItensOrdemProducaoBaixadas WHERE NumIntDoc = ?", _
            tItemOP.sCodigo, tItemOP.iItem, tItemOP.sProduto, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.lNumIntOrigem, tItemOP.iAlmoxarifado, tItemOP.sCcl, tItemOP.sSiglaUM, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.dQuantidadeProd, tItemOP.iSituacao, tItemOP.iDestinacao, tItemOP.lNumIntEquipamento, tItemOP.sVersao, tItemOP.lNumIntDocPai, tItemOP.sJustificativaBloqueio, objItemOP.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 33224
    
        'Verifica se o item do empenho existe
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 33225
    
        'Não encontrou o item  ==> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 33226
    
    End If

    'Preenche objItemOP
    objItemOP.sCodigo = tItemOP.sCodigo
    objItemOP.iItem = tItemOP.iItem
    objItemOP.sProduto = tItemOP.sProduto
    objItemOP.iFilialPedido = tItemOP.iFilialPedido
    objItemOP.lCodPedido = tItemOP.lCodPedido
    objItemOP.lNumIntOrigem = tItemOP.lNumIntOrigem
    objItemOP.iAlmoxarifado = tItemOP.iAlmoxarifado
    objItemOP.sCcl = tItemOP.sCcl
    objItemOP.sSiglaUM = tItemOP.sSiglaUM
    objItemOP.dQuantidade = tItemOP.dQuantidade
    objItemOP.dtDataInicioProd = tItemOP.dtDataInicioProd
    objItemOP.dtDataFimProd = tItemOP.dtDataFimProd
    objItemOP.iPrioridade = tItemOP.iPrioridade
    objItemOP.dQuantidadeProd = tItemOP.dQuantidadeProd
    objItemOP.iSituacao = tItemOP.iSituacao
    objItemOP.iDestinacao = tItemOP.iDestinacao
    objItemOP.lNumIntEquipamento = tItemOP.lNumIntEquipamento
    objItemOP.sVersao = tItemOP.sVersao
    objItemOP.lNumIntDocPai = tItemOP.lNumIntDocPai
    objItemOP.sJustificativaBloqueio = tItemOP.sJustificativaBloqueio

    'Fechamento comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    ItemOP_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_ItemOP_Le_NumIntDoc:

    ItemOP_Le_NumIntDoc = gErr

    Select Case gErr

        Case 33223
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 33224, 33225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO2", gErr, objItemOP.lNumIntDoc, objItemOP.iFilialEmpresa)

        Case 33226
            'Não encontrou o item do Empenho.
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151999)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Produto_Le_Categorias(ByVal objProduto As ClassProduto, ByVal colCategoriaItem As Collection) As Long
'Lê na tabela de ProdutoCategoria, todas as categorias de um  Produto, guardando na coleção colCategoriaItem

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long
Dim sCategoria As String
Dim sItem As String
Dim objProdutoCategoria As ClassProdutoCategoria

On Error GoTo Erro_Produto_Le_Categorias

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 31233

    sCategoria = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)

    'Executar comando SQL
    lErro = Comando_Executar(lComando, "SELECT Categoria, Item FROM ProdutoCategoria WHERE Produto = ? ", sCategoria, sItem, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 31234

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31235

    'Passa para a collection
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objProdutoCategoria = New ClassProdutoCategoria

        objProdutoCategoria.sProduto = objProduto.sCodigo
        objProdutoCategoria.sCategoria = sCategoria
        objProdutoCategoria.sItem = sItem

        'Adicicona na coleção
        colCategoriaItem.Add objProdutoCategoria

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 31236

    Loop

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Produto_Le_Categorias = SUCESSO

    Exit Function

Erro_Produto_Le_Categorias:

    Produto_Le_Categorias = Err

    Select Case Err

        Case 31233
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 31234, 31235, 31236
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOCATEGORIA", Err, objProdutoCategoria.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152000)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Lock(ByVal objProduto As ClassProduto) As Long
'Lê e faz o "lock" no Produto passado por parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto

On Error GoTo Erro_Produto_Lock

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26959
    
    'Inicializa as Variáveis do tipo String
    With tProduto

        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sGrade = String(STRING_GRADE_CODIGO, 0)
        
        'Lê o Produto
        lErro = Comando_ExecutarLockado(lComando, "SELECT  Descricao, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, NomeReduzido, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao,TempoProducao,Rastro,HorasMaquina,PesoEspecifico,KitVendaComp, Grade FROM Produtos WHERE Codigo = ?", .sDescricao, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sNomeReduzido, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .iTempoProducao, .iRastro, .lHorasMaquina, .dPesoEspecifico, .iKitVendaComp, .sGrade, objProduto.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 26960
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26961
        
        'não encontru --> Erro
        If lErro = AD_SQL_SEM_DADOS Then Error 26962
        
        'Faz o "lock"
        lErro = Comando_LockShared(lComando)
        If lErro <> SUCESSO Then Error 26963
        
        'Guarda so dados lidos em objProduto
        objProduto.sGrade = .sGrade
        objProduto.sDescricao = .sDescricao
        objProduto.iGerencial = .iGerencial
        objProduto.iAtivo = .iAtivo
        objProduto.iFaturamento = .iFaturamento
        objProduto.iClasseUM = .iClasseUM
        objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
        objProduto.sSiglaUMCompra = .sSiglaUMCompra
        objProduto.sSiglaUMVenda = .sSiglaUMVenda
        objProduto.sSubstituto1 = .sSubstituto1
        objProduto.sSubstituto2 = .sSubstituto2
        objProduto.iTipo = .iTipo
        objProduto.sNomeReduzido = .sNomeReduzido
        objProduto.sModelo = .sModelo
        objProduto.iNivel = .iNivel
        objProduto.iPrazoValidade = .iPrazoValidade
        objProduto.sCodigoBarras = .sCodigoBarras
        objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
        objProduto.dPesoLiq = .dPesoLiq
        objProduto.dPesoBruto = .dPesoBruto
        objProduto.dComprimento = .dComprimento
        objProduto.dEspessura = .dEspessura
        objProduto.dLargura = .dLargura
        objProduto.sCor = .sCor
        objProduto.sObsFisica = .sObsFisica
        objProduto.iCompras = .iCompras
        objProduto.iPCP = .iPCP
        objProduto.iKitBasico = .iKitBasico
        objProduto.iKitInt = .iKitInt
        objProduto.dIPIAliquota = .dIPIAliquota
        objProduto.sIPICodigo = .sIPICodigo
        objProduto.sIPICodDIPI = .sIPICodDIPI
        objProduto.iControleEstoque = .iControleEstoque
        objProduto.iApropriacaoCusto = .iApropriacaoCusto
        objProduto.iTempoProducao = .iTempoProducao
        objProduto.iRastro = .iRastro
        objProduto.lHorasMaquina = .lHorasMaquina
        objProduto.dPesoEspecifico = .dPesoEspecifico
        objProduto.iKitVendaComp = .iKitVendaComp
        
    End With
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Produto_Lock = SUCESSO
    
    Exit Function
    
Erro_Produto_Lock:

    Produto_Lock = Err
        
    Select Case Err
    
        Case 26959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 26960, 26961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS2", Err, objProduto.sNomeReduzido)
        
        Case 26962
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
        
        Case 26963
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOS", Err, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152001)
        
    End Select
        
    Call Comando_Fechar(lComando)
        
    Exit Function
    
End Function

Function Produto_Critica_ProdutoPai(ByVal sProduto As String, ByVal sSiglaModulo As String) As Long
'verifica se o Produto tem um pai sintetico

Dim sProduto1 As String
Dim iNivel As Integer
Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Produto_Critica_ProdutoPai

    'retorna o nivel do Produto
    lErro = Mascara_Produto_ObterNivel(sProduto, iNivel)
    If lErro <> SUCESSO Then Error 26989

    'se a Produto tem um "pai"
    If iNivel > 1 Then

        sProduto1 = String(STRING_PRODUTO, 0)

        'descobre a Produto "pai"
        lErro = Mascara_RetornaProdutoNoNivel(iNivel - 1, sProduto, sProduto1)
        If lErro <> SUCESSO Then Error 26990

        'verifica se a Produto "pai" está cadastrado
        objProduto.sCodigo = sProduto1

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 26991

        'nao encontrou a Produto
        If lErro = 28030 Then Error 26992

        'verifica se a Produto pai é Gerencial. Se não for é erro.
        If objProduto.iGerencial <> GERENCIAL Then Error 26993

    End If

    Produto_Critica_ProdutoPai = SUCESSO

    Exit Function

Erro_Produto_Critica_ProdutoPai:

    Produto_Critica_ProdutoPai = Err

    Select Case Err

        Case 26989

            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_PRODUTO_OBTERNIVEL", Err, sProduto)

        Case 26990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTONONIVEL", Err, sProduto, iNivel - 1)

        Case 26991

        Case 26992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_NAO_CADASTRADO", Err, objProduto.sCodigo)

        Case 26993
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_FINAL", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152002)

    End Select

    Exit Function

End Function


Function EstoqueMes_Le(ByVal objEstoqueMes As ClassEstoqueMes) As Long
'Le as informacoes sobre o EstoqueMes através do parametro Filial/Ano/Mes passado em objEstoqueMes

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueMes As typeEstoqueMes

On Error GoTo Erro_EstoqueMes_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 36510
    
    '28/11/01 Marcelo inclusão da leituta dos campos HorasMaquina, CustoMatPrim, GastosDiretos, GastosIndiretos
    lErro = Comando_Executar(lComando, "SELECT Fechamento, Alterado, CustoProdApurado, HorasMaquina, CustoMatPrim, GastosDiretos, GastosIndiretos, QuantFator1, CustoFator1, QuantFator2, CustoFator2, QuantFator3, CustoFator3, QuantFator4, CustoFator4, QuantFator5, CustoFator5, QuantFator6, CustoFator6 FROM EstoqueMes WHERE FilialEmpresa = ? AND Ano = ? AND Mes = ?", tEstoqueMes.iFechamento, tEstoqueMes.iAlterado, tEstoqueMes.iCustoProdApurado, tEstoqueMes.lHorasMaquina, tEstoqueMes.dCustoMatPrim, tEstoqueMes.dGastosDiretos, tEstoqueMes.dGastosIndiretos, _
        tEstoqueMes.dQuantFator1, tEstoqueMes.dCustoFator1, tEstoqueMes.dQuantFator2, tEstoqueMes.dCustoFator2, tEstoqueMes.dQuantFator3, tEstoqueMes.dCustoFator3, tEstoqueMes.dQuantFator4, tEstoqueMes.dCustoFator4, tEstoqueMes.dQuantFator5, tEstoqueMes.dCustoFator5, tEstoqueMes.dQuantFator6, tEstoqueMes.dCustoFator6, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
    If lErro <> AD_SQL_SUCESSO Then Error 36511

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36512

    If lErro = AD_SQL_SEM_DADOS Then Error 36513

    objEstoqueMes.iFechamento = tEstoqueMes.iFechamento
    objEstoqueMes.iAlterado = tEstoqueMes.iAlterado
    objEstoqueMes.iCustoProdApurado = tEstoqueMes.iCustoProdApurado
    
    '28/11/01 Marcelo atribuição dos campos acima para o obj
    objEstoqueMes.lHorasMaquina = tEstoqueMes.lHorasMaquina
    objEstoqueMes.dCustoMatPrim = tEstoqueMes.dCustoMatPrim
    objEstoqueMes.dGastosDiretos = tEstoqueMes.dGastosDiretos
    objEstoqueMes.dGastosIndiretos = tEstoqueMes.dGastosIndiretos
    objEstoqueMes.dQuantFator1 = tEstoqueMes.dQuantFator1
    objEstoqueMes.dCustoFator1 = tEstoqueMes.dCustoFator1
    objEstoqueMes.dQuantFator2 = tEstoqueMes.dQuantFator2
    objEstoqueMes.dCustoFator2 = tEstoqueMes.dCustoFator2
    objEstoqueMes.dQuantFator3 = tEstoqueMes.dQuantFator3
    objEstoqueMes.dCustoFator3 = tEstoqueMes.dCustoFator3
    objEstoqueMes.dQuantFator4 = tEstoqueMes.dQuantFator4
    objEstoqueMes.dCustoFator4 = tEstoqueMes.dCustoFator4
    objEstoqueMes.dQuantFator5 = tEstoqueMes.dQuantFator5
    objEstoqueMes.dCustoFator5 = tEstoqueMes.dCustoFator5
    objEstoqueMes.dQuantFator6 = tEstoqueMes.dQuantFator6
    objEstoqueMes.dCustoFator6 = tEstoqueMes.dCustoFator6
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    EstoqueMes_Le = SUCESSO

    Exit Function

Erro_EstoqueMes_Le:

    EstoqueMes_Le = Err

    Select Case Err

        Case 36510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 36511, 36512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES", Err, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)

        Case 36513

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152003)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function
 
Private Sub Produto_MoverCamposObj(tProduto As typeProduto, ByVal objProduto As ClassProduto)

    With tProduto

        'Carrega os dados encontrados em objProduto
        objProduto.sCodigo = .sCodigo
        objProduto.sDescricao = .sDescricao
        objProduto.iGerencial = .iGerencial
        objProduto.iAtivo = .iAtivo
        objProduto.iFaturamento = .iFaturamento
        objProduto.iClasseUM = .iClasseUM
        objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
        objProduto.sSiglaUMCompra = .sSiglaUMCompra
        objProduto.sSiglaUMVenda = .sSiglaUMVenda
        objProduto.sSubstituto1 = .sSubstituto1
        objProduto.sSubstituto2 = .sSubstituto2
        objProduto.iTipo = .iTipo
        objProduto.sNomeReduzido = .sNomeReduzido
        objProduto.sModelo = .sModelo
        objProduto.iNivel = .iNivel
        objProduto.iPrazoValidade = .iPrazoValidade
        objProduto.sCodigoBarras = .sCodigoBarras
        objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
        objProduto.dPesoLiq = .dPesoLiq
        objProduto.dPesoBruto = .dPesoBruto
        objProduto.dComprimento = .dComprimento
        objProduto.dEspessura = .dEspessura
        objProduto.dLargura = .dLargura
        objProduto.sCor = .sCor
        objProduto.sObsFisica = .sObsFisica
        objProduto.iCompras = .iCompras
        objProduto.iPCP = .iPCP
        objProduto.iKitBasico = .iKitBasico
        objProduto.iKitInt = .iKitInt
        objProduto.dIPIAliquota = .dIPIAliquota
        objProduto.sIPICodigo = .sIPICodigo
        objProduto.sIPICodDIPI = .sIPICodDIPI
        objProduto.iControleEstoque = .iControleEstoque
        objProduto.iApropriacaoCusto = .iApropriacaoCusto
        objProduto.iTempoProducao = .iTempoProducao
        
    End With

End Sub

Function EstoqueProduto_Le1(ByVal lComando As Long, ByVal objEstoqueProduto As ClassEstoqueProduto) As Long
'Le as informacoes sobre o produto/almoxarifado passado em objEstoqueProduto na tabela EstoqueProduto

Dim lErro As Long
Dim tEstoqueProduto As typeEstoqueProduto

On Error GoTo Erro_EstoqueProduto_Le1

    tEstoqueProduto.sLocalizacaoFisica = String(STRING_ESTOQUEPRODUTO_LOCALIZACAOFISICA, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT QuantDispNossa, LocalizacaoFisica, QuantReservada, QuantReservadaConsig, DataInventario, QuantidadeInicial, SaldoInicial, DataInicial, QuantEmpenhada, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantBenef3, QuantBenef, QuantInicialConsig3, QuantInicialConsig, QuantInicialDemo3, QuantInicialDemo, QuantInicialConserto3, QuantInicialConserto, QuantInicialOutras3, QuantInicialOutras, QuantInicialBenef3, QuantInicialBenef, " & _
                                           "ValorInicialConsig3, ValorInicialConsig, ValorInicialDemo3, ValorInicialDemo, ValorInicialConserto3, ValorInicialConserto, ValorInicialOutras3, ValorInicialOutras, ValorInicialBenef3, ValorInicialBenef, ValorConsig3, ValorConsig, ValorDemo3, ValorDemo, ValorConserto3, ValorConserto, ValorOutras3, ValorOutras, ValorBenef3, ValorBenef FROM EstoqueProduto WHERE Produto = ? AND Almoxarifado = ?", 0, _
        tEstoqueProduto.dQuantDispNossa, tEstoqueProduto.sLocalizacaoFisica, tEstoqueProduto.dQuantReservada, tEstoqueProduto.dQuantReservadaConsig, tEstoqueProduto.dtDataInventario, tEstoqueProduto.dQuantidadeInicial, tEstoqueProduto.dSaldoInicial, tEstoqueProduto.dtDataInicial, tEstoqueProduto.dQuantEmpenhada, tEstoqueProduto.dQuantRecIndl, tEstoqueProduto.dQuantInd, tEstoqueProduto.dQuantDefeituosa, tEstoqueProduto.dQuantConsig3, tEstoqueProduto.dQuantConsig, tEstoqueProduto.dQuantDemo3, tEstoqueProduto.dQuantDemo, tEstoqueProduto.dQuantConserto3, tEstoqueProduto.dQuantConserto, tEstoqueProduto.dQuantOutras3, tEstoqueProduto.dQuantOutras, tEstoqueProduto.dQuantBenef3, tEstoqueProduto.dQuantBenef, _
        tEstoqueProduto.dQuantInicialConsig3, tEstoqueProduto.dQuantInicialConsig, tEstoqueProduto.dQuantInicialDemo3, tEstoqueProduto.dQuantInicialDemo, tEstoqueProduto.dQuantInicialConserto3, tEstoqueProduto.dQuantInicialConserto, tEstoqueProduto.dQuantInicialOutras3, tEstoqueProduto.dQuantInicialOutras, tEstoqueProduto.dQuantInicialBenef3, tEstoqueProduto.dQuantInicialBenef, _
        tEstoqueProduto.dValorInicialConsig3, tEstoqueProduto.dValorInicialConsig, tEstoqueProduto.dValorInicialDemo3, tEstoqueProduto.dValorInicialDemo, tEstoqueProduto.dValorInicialConserto3, tEstoqueProduto.dValorInicialConserto, tEstoqueProduto.dValorInicialOutras3, tEstoqueProduto.dValorInicialOutras, tEstoqueProduto.dValorInicialBenef3, tEstoqueProduto.dValorInicialBenef, _
        tEstoqueProduto.dValorConsig3, tEstoqueProduto.dValorConsig, tEstoqueProduto.dValorDemo3, tEstoqueProduto.dValorDemo, tEstoqueProduto.dValorConserto3, tEstoqueProduto.dValorConserto, tEstoqueProduto.dValorOutras3, tEstoqueProduto.dValorOutras, tEstoqueProduto.dValorBenef3, tEstoqueProduto.dValorBenef, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then Error 36662

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36663

    If lErro = AD_SQL_SEM_DADOS Then Error 36664

    objEstoqueProduto.dQuantDispNossa = tEstoqueProduto.dQuantDispNossa
    objEstoqueProduto.sLocalizacaoFisica = tEstoqueProduto.sLocalizacaoFisica
    objEstoqueProduto.dQuantReservada = tEstoqueProduto.dQuantReservada
    objEstoqueProduto.dQuantReservadaConsig = tEstoqueProduto.dQuantReservadaConsig
    objEstoqueProduto.dtDataInventario = tEstoqueProduto.dtDataInventario
    objEstoqueProduto.dQuantidadeInicial = tEstoqueProduto.dQuantidadeInicial
    objEstoqueProduto.dSaldoInicial = tEstoqueProduto.dSaldoInicial
    objEstoqueProduto.dtDataInicial = tEstoqueProduto.dtDataInicial
    objEstoqueProduto.dQuantEmpenhada = tEstoqueProduto.dQuantEmpenhada
    objEstoqueProduto.dQuantRecIndl = tEstoqueProduto.dQuantRecIndl
    objEstoqueProduto.dQuantInd = tEstoqueProduto.dQuantInd
    objEstoqueProduto.dQuantDefeituosa = tEstoqueProduto.dQuantDefeituosa
    objEstoqueProduto.dQuantConsig3 = tEstoqueProduto.dQuantConsig3
    objEstoqueProduto.dQuantConsig = tEstoqueProduto.dQuantConsig
    objEstoqueProduto.dQuantDemo3 = tEstoqueProduto.dQuantDemo3
    objEstoqueProduto.dQuantDemo = tEstoqueProduto.dQuantDemo
    objEstoqueProduto.dQuantConserto3 = tEstoqueProduto.dQuantConserto3
    objEstoqueProduto.dQuantConserto = tEstoqueProduto.dQuantConserto
    objEstoqueProduto.dQuantOutras3 = tEstoqueProduto.dQuantOutras3
    objEstoqueProduto.dQuantOutras = tEstoqueProduto.dQuantOutras
    objEstoqueProduto.dQuantBenef3 = tEstoqueProduto.dQuantBenef3
    objEstoqueProduto.dQuantBenef = tEstoqueProduto.dQuantBenef
    
    objEstoqueProduto.dQuantInicialConsig3 = tEstoqueProduto.dQuantInicialConsig3
    objEstoqueProduto.dQuantInicialConsig = tEstoqueProduto.dQuantInicialConsig
    objEstoqueProduto.dQuantInicialDemo3 = tEstoqueProduto.dQuantInicialDemo3
    objEstoqueProduto.dQuantInicialDemo = tEstoqueProduto.dQuantInicialDemo
    objEstoqueProduto.dQuantInicialConserto3 = tEstoqueProduto.dQuantInicialConserto3
    objEstoqueProduto.dQuantInicialConserto = tEstoqueProduto.dQuantInicialConserto
    objEstoqueProduto.dQuantInicialOutras3 = tEstoqueProduto.dQuantInicialOutras3
    objEstoqueProduto.dQuantInicialOutras = tEstoqueProduto.dQuantInicialOutras
    objEstoqueProduto.dQuantInicialBenef3 = tEstoqueProduto.dQuantInicialBenef3
    objEstoqueProduto.dQuantInicialBenef = tEstoqueProduto.dQuantInicialBenef

    objEstoqueProduto.dValorInicialConsig3 = tEstoqueProduto.dValorInicialConsig3
    objEstoqueProduto.dValorInicialConsig = tEstoqueProduto.dValorInicialConsig
    objEstoqueProduto.dValorInicialDemo3 = tEstoqueProduto.dValorInicialDemo3
    objEstoqueProduto.dValorInicialDemo = tEstoqueProduto.dValorInicialDemo
    objEstoqueProduto.dValorInicialConserto3 = tEstoqueProduto.dValorInicialConserto3
    objEstoqueProduto.dValorInicialConserto = tEstoqueProduto.dValorInicialConserto
    objEstoqueProduto.dValorInicialOutras3 = tEstoqueProduto.dValorInicialOutras3
    objEstoqueProduto.dValorInicialOutras = tEstoqueProduto.dValorInicialOutras
    objEstoqueProduto.dValorInicialBenef3 = tEstoqueProduto.dValorInicialBenef3
    objEstoqueProduto.dValorInicialBenef = tEstoqueProduto.dValorInicialBenef

    objEstoqueProduto.dValorConsig3 = tEstoqueProduto.dValorConsig3
    objEstoqueProduto.dValorConsig = tEstoqueProduto.dValorConsig
    objEstoqueProduto.dValorDemo3 = tEstoqueProduto.dValorDemo3
    objEstoqueProduto.dValorDemo = tEstoqueProduto.dValorDemo
    objEstoqueProduto.dValorConserto3 = tEstoqueProduto.dValorConserto3
    objEstoqueProduto.dValorConserto = tEstoqueProduto.dValorConserto
    objEstoqueProduto.dValorOutras3 = tEstoqueProduto.dValorOutras3
    objEstoqueProduto.dValorOutras = tEstoqueProduto.dValorOutras
    objEstoqueProduto.dValorBenef3 = tEstoqueProduto.dValorBenef3
    objEstoqueProduto.dValorBenef = tEstoqueProduto.dValorBenef


    EstoqueProduto_Le1 = SUCESSO

    Exit Function

Erro_EstoqueProduto_Le1:

    EstoqueProduto_Le1 = Err

    Select Case Err

        Case 36662, 36663
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO1", Err, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 36664

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152004)

    End Select

    Exit Function

End Function

Function EstoqueMes_Le_Mes(ByVal objEstoqueMes As ClassEstoqueMes) As Long
'Le mes e ano em EstoqueMes sendo passado como parametro filial empresa e o fechamento como ABERTO

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueMes As typeEstoqueMes

On Error GoTo Erro_EstoqueMes_Le_Mes
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 40670
    
    'Le Ano e Mes
    lErro = Comando_ExecutarLockado(lComando, "SELECT Ano, Mes  FROM EstoqueMes WHERE FilialEmpresa = ? AND Fechamento = ? ORDER BY Ano, Mes", tEstoqueMes.iAno, tEstoqueMes.iMes, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iFechamento)
    If lErro <> AD_SQL_SUCESSO Then Error 40671

    'Busca o primeiro ano e mes que estiverem aberto ( so podera haver um mes aberto
    lErro = Comando_BuscarUltimo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 40672

    If lErro = AD_SQL_SEM_DADOS Then Error 40673

    '"Lock" do Mes
    lErro = Comando_LockExclusive(lComando)
    If lErro <> SUCESSO Then Error 40675
    
    'preenche o objeto com o mes e o ano
    objEstoqueMes.iMes = tEstoqueMes.iMes
    objEstoqueMes.iAno = tEstoqueMes.iAno

    'Fechamento comando
    Call Comando_Fechar(lComando)

    EstoqueMes_Le_Mes = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_Mes:

    EstoqueMes_Le_Mes = Err

    Select Case Err

        Case 40670
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 40671, 40672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", Err)

        Case 40673

        Case 40675
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ESTOQUEMES", Err, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152005)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Almoxarifados_Le_FilialEmpresa(ByVal iFilialEmpresa As Integer, ByVal colAlmoxFilial As Collection) As Long
'Le todos os almoxaridados da Filial iFilialEmpresa e coloca-os em colAlmoxFilial

Dim lComando As Long
Dim lErro As Long
Dim tAlmoxarifado As typeAlmoxarifado
Dim objAlmoxarifado As ClassAlmoxarifado

On Error GoTo Erro_Almoxarifados_Le_FilialEmpresa

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44167

    With tAlmoxarifado

        'Prepara variáveis string para receber conteúdo do BD
        .sDescricao = String(STRING_ALMOXARIFADO_DESCRICAO, 0)
        .sNomeReduzido = String(STRING_ALMOXARIFADO_NOMEREDUZIDO, 0)
    
        'Seleciona os dados da tabela Almoxarifado para a Filial em questão
        lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, Descricao, Endereco, NomeReduzido FROM Almoxarifado WHERE FilialEmpresa = ? ", _
        .iCodigo, .iFilialEmpresa, .sDescricao, .lEndereco, .sNomeReduzido, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 44168
    
        'Le o primeiro almoxarifado
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44169
    
        Do While lErro = SUCESSO
    
            Set objAlmoxarifado = New ClassAlmoxarifado
    
            'Passa dados lidos para objAlmoxarifado
            objAlmoxarifado.iCodigo = .iCodigo
            objAlmoxarifado.iFilialEmpresa = .iFilialEmpresa
            objAlmoxarifado.sDescricao = .sDescricao
            objAlmoxarifado.lEndereco = .lEndereco
            objAlmoxarifado.sNomeReduzido = .sNomeReduzido
            
            colAlmoxFilial.Add objAlmoxarifado

            'Ler o proximo almoxarifado
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44170

        Loop

    End With

    'Fecha comando
    Call Comando_Fechar(lComando)

    Almoxarifados_Le_FilialEmpresa = SUCESSO

    Exit Function

Erro_Almoxarifados_Le_FilialEmpresa:

    Almoxarifados_Le_FilialEmpresa = Err

    Select Case Err

        Case 44167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44168, 44169, 44170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADOS", Err, iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152006)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoFilial_Le2(ByVal lComando As Long, ByVal objProdutoFilial As ClassProdutoFilial) As Long
'Lê os dados do Produto da Filial à partir da chave fornecida em objProdutoFilial. Retorna os dados em objProdutoFilial

Dim lErro As Long
Dim tProdutoFilial As typeProdutoFilial

On Error GoTo Erro_ProdutoFilial_Le2

    tProdutoFilial.sClasseABC = String(STRING_PRODUTOFILIAL_CLASSEABC, 0)
    tProdutoFilial.sSituacaoTribECF = String(STRING_PRODUTOFILIAL_SITUACAOTRIBECF, 0)
    tProdutoFilial.sICMSAliquota = String(STRING_PRODUTOFILIAL_ICMS, 0)

    'Pesquisa no BD ProdutoFilial
    With tProdutoFilial
        lErro = Comando_Executar(lComando, "SELECT Almoxarifado, Fornecedor, VisibilidadeAlmoxarifados, EstoqueSeguranca, EstoqueMaximo, PontoPedido, ClasseABC, LoteEconomico, IntRessup, SituacaoTribECF, ICMSAliquota FROM ProdutosFilial WHERE FilialEmpresa= ? AND Produto = ?", _
            .iAlmoxarifado, .lFornecedor, .iVisibilidadeAlmoxarifados, .dEstoqueSeguranca, .dEstoqueMaximo, .dPontoPedido, .sClasseABC, .dLoteEconomico, .iIntRessup, .sSituacaoTribECF, .sICMSAliquota, objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)
        If lErro <> AD_SQL_SUCESSO Then Error 44206
    End With

    'Tenta selecionar ProdutoFilial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44207

    If lErro = AD_SQL_SEM_DADOS Then Error 44208

    'Preenche objProdutoFilial
    With tProdutoFilial
        objProdutoFilial.iAlmoxarifado = .iAlmoxarifado
        objProdutoFilial.lFornecedor = .lFornecedor
        objProdutoFilial.iVisibilidadeAlmoxarifados = .iVisibilidadeAlmoxarifados
        objProdutoFilial.dEstoqueSeguranca = .dEstoqueSeguranca
        objProdutoFilial.dEstoqueMaximo = .dEstoqueMaximo
        objProdutoFilial.dPontoPedido = .dPontoPedido
        objProdutoFilial.sClasseABC = .sClasseABC
        objProdutoFilial.dLoteEconomico = .dLoteEconomico
        objProdutoFilial.iIntRessup = .iIntRessup
        objProdutoFilial.sSituacaoTribECF = .sSituacaoTribECF
        objProdutoFilial.sICMSAliquota = .sICMSAliquota
    End With

    ProdutoFilial_Le2 = SUCESSO

    Exit Function

Erro_ProdutoFilial_Le2:

    ProdutoFilial_Le2 = Err

    Select Case Err

        Case 44206, 44207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err, objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)

        Case 44208

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152007)

    End Select

    Exit Function

End Function


'mario
Function EstoqueProduto_Le_Exclusivo(ByVal lComando As Long, ByVal objEstoqueProduto As ClassEstoqueProduto) As Long
'le o registro em EstoqueProduto cuja chave (Produto, Almoxarifado) foi passada como parametro em objEstoqueProduto.
'Devolve os dados lidos em objEstoqueProduto.

Dim lErro As Long
Dim tEstoqueProduto As typeEstoqueProduto

On Error GoTo Erro_EstoqueProduto_Le_Exclusivo

    tEstoqueProduto.sLocalizacaoFisica = String(STRING_ESTOQUEPRODUTO_LOCALIZACAOFISICA, 0)

    lErro = Comando_ExecutarLockado(lComando, "SELECT QuantDispNossa, LocalizacaoFisica, QuantReservada, DataInventario, QuantidadeInicial, SaldoInicial, DataInicial, QuantEmpenhada, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantBenef, QuantBenef3 FROM EstoqueProduto WHERE Produto = ? AND Almoxarifado = ?", _
        tEstoqueProduto.dQuantDispNossa, tEstoqueProduto.sLocalizacaoFisica, tEstoqueProduto.dQuantReservada, tEstoqueProduto.dtDataInventario, tEstoqueProduto.dQuantidadeInicial, tEstoqueProduto.dSaldoInicial, tEstoqueProduto.dtDataInicial, tEstoqueProduto.dQuantEmpenhada, tEstoqueProduto.dQuantRecIndl, tEstoqueProduto.dQuantInd, tEstoqueProduto.dQuantDefeituosa, tEstoqueProduto.dQuantConsig3, tEstoqueProduto.dQuantConsig, tEstoqueProduto.dQuantDemo3, tEstoqueProduto.dQuantDemo, tEstoqueProduto.dQuantConserto3, tEstoqueProduto.dQuantConserto, tEstoqueProduto.dQuantOutras3, tEstoqueProduto.dQuantOutras, tEstoqueProduto.dQuantBenef, tEstoqueProduto.dQuantBenef3, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then Error 44227

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44228

    If lErro = AD_SQL_SEM_DADOS Then Error 44229

    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 44230

    With objEstoqueProduto

        .dQuantDispNossa = tEstoqueProduto.dQuantDispNossa
        .sLocalizacaoFisica = tEstoqueProduto.sLocalizacaoFisica
        .dQuantReservada = tEstoqueProduto.dQuantReservada
        .dtDataInventario = tEstoqueProduto.dtDataInventario
        .dQuantidadeInicial = tEstoqueProduto.dQuantidadeInicial
        .dSaldoInicial = tEstoqueProduto.dSaldoInicial
        .dtDataInicial = tEstoqueProduto.dtDataInicial
        .dQuantEmpenhada = tEstoqueProduto.dQuantEmpenhada
        .dQuantRecIndl = tEstoqueProduto.dQuantRecIndl
        .dQuantInd = tEstoqueProduto.dQuantInd
        .dQuantDefeituosa = tEstoqueProduto.dQuantDefeituosa
        .dQuantConsig3 = tEstoqueProduto.dQuantConsig3
        .dQuantConsig = tEstoqueProduto.dQuantConsig
        .dQuantDemo3 = tEstoqueProduto.dQuantDemo3
        .dQuantDemo = tEstoqueProduto.dQuantDemo
        .dQuantConserto3 = tEstoqueProduto.dQuantConserto
        .dQuantConserto = tEstoqueProduto.dQuantConserto
        .dQuantOutras3 = tEstoqueProduto.dQuantOutras3
        .dQuantOutras = tEstoqueProduto.dQuantOutras
        .dQuantBenef = tEstoqueProduto.dQuantBenef
        .dQuantBenef3 = tEstoqueProduto.dQuantBenef3
    End With

    EstoqueProduto_Le_Exclusivo = SUCESSO

    Exit Function

Erro_EstoqueProduto_Le_Exclusivo:

    EstoqueProduto_Le_Exclusivo = Err

    Select Case Err

        Case 44227, 44228
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO1", Err, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 44229 'nao encontrou

        Case 44230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ESTOQUEPRODUTO", Err, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152008)

    End Select

    Exit Function

End Function

Function Produto_Lock1(ByVal lComando As Long, ByVal objProduto As ClassProduto) As Long
'Lê o faz o "lock" no Produto passado por parâmetro

Dim lErro As Long
Dim tProduto As typeProduto

On Error GoTo Erro_Produto_Lock1

    tProduto.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    tProduto.sSiglaUMVenda = String(STRING_UM_SIGLA, 0)
    tProduto.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    
    'Lê o Produto
    lErro = Comando_ExecutarPos(lComando, "SELECT ControleEstoque, SiglaUMEstoque, SiglaUMVenda, Apropriacao,TabelaPreco, Descricao  FROM Produtos WHERE Codigo = ? ", 0, tProduto.iControleEstoque, tProduto.sSiglaUMEstoque, tProduto.sSiglaUMVenda, tProduto.iApropriacaoCusto, tProduto.iTabelaPreco, tProduto.sDescricao, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 55169

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55170

    'não encontru --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 55172

    'Faz o "lock"
    lErro = Comando_LockExclusive(lComando)
    If lErro <> SUCESSO Then Error 55173

    objProduto.iControleEstoque = tProduto.iControleEstoque
    objProduto.sSiglaUMEstoque = tProduto.sSiglaUMEstoque
    objProduto.sSiglaUMVenda = tProduto.sSiglaUMVenda
    objProduto.iApropriacaoCusto = tProduto.iApropriacaoCusto
    objProduto.iTabelaPreco = tProduto.iTabelaPreco
    objProduto.sDescricao = tProduto.sDescricao
    
    Produto_Lock1 = SUCESSO

    Exit Function

Erro_Produto_Lock1:

    Produto_Lock1 = Err

    Select Case Err

        Case 55169, 55170
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS2", Err, objProduto.sNomeReduzido)

        Case 55172
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 55173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOS", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152009)

    End Select

    Exit Function

End Function

Function CustoMedioProducaoApurado_Le(ByVal sProduto As String, dCMPAtual As Double, ByVal iFilialEmpresa As Integer) As Long
'Calcula último Custo Médio de Produção Apurado do Produto passado

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_CustoMedioProducaoApurado_Le

    objEstoqueMes.iFilialEmpresa = iFilialEmpresa
    
    lErro = CF("EstoqueMes_Le_Apurado", objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 46225 Then gError 25430
    
    'Se não foi apurado nenhum mês ==> descobre o ano em que o estoque foi
    If lErro = 46225 Then
        objEstoqueMes.iMes = 0
        objEstoqueMes.iAno = 0
    End If
    
    'calcula o custo medio de produção do produto em questão
    lErro = CF("Calcula_CustoMedioProducao", iFilialEmpresa, sProduto, objEstoqueMes.iAno, objEstoqueMes.iMes, dCMPAtual)
    If lErro <> SUCESSO And lErro <> 55052 And lErro <> 25433 Then gError 55046

    'Se não encontrou o SldMesEst
    If lErro = 25433 Then gError 69867
            
    'Se não encontrou o SldMesEst
    If lErro = 55052 Then gError 69868

    CustoMedioProducaoApurado_Le = SUCESSO

    Exit Function

Erro_CustoMedioProducaoApurado_Le:

    CustoMedioProducaoApurado_Le = gErr

    Select Case gErr

        Case 25430, 55030, 55046 'Tratado na rotina chamada
        
        Case 69867 'Não encontrou SldMesEst
        
        Case 69868 'Não encontrou SldMesEst
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152010)

    End Select

    Exit Function

End Function

Function EstoquesProduto_Le_Filial(ByVal sProduto As String, ByVal colEstoqueProduto As colEstoqueProduto) As Long
'??? documentar

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueProduto As typeEstoqueProduto
Dim iFilialEmpresa1 As Integer

On Error GoTo Erro_EstoquesProduto_Le_Filial

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 39344
    
    With tEstoqueProduto
    
        .sAlmoxarifadoNomeReduzido = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sLocalizacaoFisica = String(STRING_ESTOQUEPRODUTO_LOCALIZACAOFISICA, 0)

        iFilialEmpresa1 = giFilialEmpresa
        lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa1)
        If lErro <> SUCESSO Then Error 30899
                
        lErro = Comando_Executar(lComando, "SELECT EstoqueProduto.Almoxarifado, EstoqueProduto.LocalizacaoFisica, EstoqueProduto.DataInventario,EstoqueProduto.SaldoInicial, EstoqueProduto.DataInicial, EstoqueProduto.QuantidadeInicial, EstoqueProduto.QuantDispNossa, EstoqueProduto.QuantReservada, EstoqueProduto.QuantEmpenhada, EstoqueProduto.QuantRecIndl, EstoqueProduto.QuantIndOutras, EstoqueProduto.QuantDefeituosa, EstoqueProduto.QuantConsig, EstoqueProduto.QuantDemo3, EstoqueProduto.QuantConsig3, EstoqueProduto.QuantDemo, EstoqueProduto.QuantConserto3, EstoqueProduto.QuantConserto, EstoqueProduto.QuantOutras3, EstoqueProduto.QuantOutras, QuantBenef, QuantBenef3, Almoxarifado.NomeReduzido FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.FilialEmpresa = ? ORDER BY Almoxarifado.NomeReduzido", _
            .iAlmoxarifado, .sLocalizacaoFisica, .dtDataInventario, .dSaldoInicial, .dtDataInicial, .dQuantidadeInicial, .dQuantDispNossa, .dQuantReservada, .dQuantEmpenhada, .dQuantRecIndl, .dQuantInd, .dQuantDefeituosa, .dQuantConsig, .dQuantDemo3, .dQuantConsig3, .dQuantDemo, .dQuantConserto3, .dQuantConserto, .dQuantOutras3, .dQuantOutras, .dQuantBenef, .dQuantBenef3, .sAlmoxarifadoNomeReduzido, sProduto, iFilialEmpresa1)
        If lErro <> AD_SQL_SUCESSO Then Error 39345
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 39346
       
        Do While lErro = AD_SQL_SUCESSO
        
            colEstoqueProduto.Add sProduto, .iAlmoxarifado, .sAlmoxarifadoNomeReduzido, .dQuantDispNossa, .dSaldo, .sLocalizacaoFisica, .sContaContabil, .dQuantReservada, .dtDataInventario, .dQuantidadeInicial, .dSaldoInicial, .dtDataInicial, .dQuantEmpenhada, .dQuantRecIndl, .dQuantInd, .dQuantDefeituosa, .dQuantConsig, .dQuantConsig3, .dQuantDemo, .dQuantDemo3, .dQuantConserto3, .dQuantConserto, .dQuantOutras, .dQuantOutras3, .dQuantBenef, .dQuantBenef3
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 39348
            
        Loop
        
    End With
            
    Call Comando_Fechar(lComando)
    
    EstoquesProduto_Le_Filial = SUCESSO
    
    Exit Function
            
Erro_EstoquesProduto_Le_Filial:

    EstoquesProduto_Le_Filial = Err
    
    Select Case Err
    
        Case 30899
        
        Case 39344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 39345, 39346, 39348
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUESPRODUTO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152011)
            
    End Select
        
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function CurvaABC_LePontos(ByVal objClassABC As ClassClassificacaoABC, ByVal objCurvaABC As ClassCurvaABC) As Long
'Lê dados da curva ABC no Banco de Dados e guarda os Pontos em objCurvaABC.colPontos
    
Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim lNumProdutos As Long
Dim tItemABC As typeItemClassifABC
Dim iIndice As Integer
Dim lIndice As Long
Dim dDemandaAcumulada As Double
Dim dDemandaAcumuladaPerc As Double
Dim objPonto As New ClassPonto
Dim iNumPontos As Integer

On Error GoTo Erro_CurvaABC_LePontos

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 25387
    Next
    
    objCurvaABC.lClassifABC = objClassABC.lNumInt
    objCurvaABC.dDemandaTotal = objClassABC.dDemandaTotal
    objCurvaABC.iFaixaA = objClassABC.iFaixaA
    objCurvaABC.iFaixaB = objClassABC.iFaixaB
    
    'Lê o número total de ítens da ClassifABC
    lErro = Comando_Executar(alComando(1), "SELECT COUNT(*) FROM ItensClassifABC WHERE ClassifABC = ?", lNumProdutos, objClassABC.lNumInt)
    If lErro <> AD_SQL_SUCESSO Then Error 25388
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 25389
    
    If lNumProdutos <= CURVA_ABC_MAX_PONTOS Then
        objCurvaABC.lPasso = 1 'passo que será usado no aproveitamento dos dados do BD
        iNumPontos = lNumProdutos
    Else
        objCurvaABC.lPasso = Int(lNumProdutos / 100) 'passo que será usado no aproveitamento dos dados do BD
        iNumPontos = Int(lNumProdutos / objCurvaABC.lPasso)
    End If
        
    With tItemABC
        .sClasseABC = String(STRING_CLASSEABC, 0)
        lErro = Comando_Executar(alComando(2), "SELECT Demanda, ClasseABC FROM ItensClassifABC WHERE ClassifABC = ? ORDER BY Demanda DESC, Produto", .dDemanda, .sClasseABC, objClassABC.lNumInt)
    End With
    If lErro <> AD_SQL_SUCESSO Then Error 25390
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25391
    If lErro = AD_SQL_SEM_DADOS Then Error 25392
   
    objCurvaABC.ColPontos.Add 0, 0
   
    'Geração de pontos
    For iIndice = 1 To iNumPontos
        
        For lIndice = 1 To objCurvaABC.lPasso
            
            'Armazena mudança de A para B
            If tItemABC.sClasseABC = "B" And objCurvaABC.objPontoClasseA.dX = 0 Then
            
                objCurvaABC.lItensA = (iIndice - 1) * objCurvaABC.lPasso + lIndice - 1
                objCurvaABC.objPontoClasseA.dX = (objCurvaABC.lItensA / lNumProdutos) * 100
                objCurvaABC.objPontoClasseA.dY = dDemandaAcumuladaPerc
   
                If iIndice = 1 Or lIndice > 1 Then 'Ponto não faz parte da coleção da curva
                    'Adiciona o ponto à coleção
                    objCurvaABC.ColPontos.Add objCurvaABC.objPontoClasseA.dX, objCurvaABC.objPontoClasseA.dY
                
                End If

            End If
            
            'Armazena mudança de B para C
            If tItemABC.sClasseABC = "C" And objCurvaABC.objPontoClasseB.dX = 0 Then
            
                objCurvaABC.lItensB = (iIndice - 1) * objCurvaABC.lPasso + lIndice - 1 - objCurvaABC.lItensA
                objCurvaABC.lItensC = lNumProdutos - objCurvaABC.lItensA - objCurvaABC.lItensB
                objCurvaABC.objPontoClasseB.dX = ((objCurvaABC.lItensA + objCurvaABC.lItensB) / lNumProdutos) * 100
                objCurvaABC.objPontoClasseB.dY = dDemandaAcumuladaPerc
                
                If iIndice = 1 Or lIndice > 1 Then 'Ponto não faz parte da coleção da curva
                    'Adiciona o ponto à coleção
                    objCurvaABC.ColPontos.Add objCurvaABC.objPontoClasseB.dX, objCurvaABC.objPontoClasseB.dY
                
                End If
            
            End If
            
            'Calcula demanda acumulada até esse Produto
            dDemandaAcumulada = dDemandaAcumulada + tItemABC.dDemanda
            dDemandaAcumuladaPerc = (dDemandaAcumulada / objClassABC.dDemandaTotal) * 100
        
            'Se nao estiver no ultimo Produto a ser lido, lê o próximo
            If iIndice < iNumPontos Or lIndice < objCurvaABC.lPasso Then
                lErro = Comando_BuscarProximo(alComando(2))
                If lErro <> AD_SQL_SUCESSO Then Error 25393
            End If
          
        Next
        
        'Coordenadas do ponto
        objPonto.dX = ((iIndice * objCurvaABC.lPasso) / lNumProdutos) * 100
        objPonto.dY = dDemandaAcumuladaPerc
        
        'Adiciona o ponto à coleção
        objCurvaABC.ColPontos.Add objPonto.dX, objPonto.dY
        
    Next
    
    'Se último ponto não está incluído, inclui
    If objPonto.dX < 100 Then
        
        objCurvaABC.ColPontos.Add 100, 100
        
    End If
    
    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CurvaABC_LePontos = SUCESSO
    
    Exit Function

Erro_CurvaABC_LePontos:
 
    CurvaABC_LePontos = Err
    
    Select Case Err
            
        Case 25387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 25388, 25389, 25390, 25391, 25393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMCLASSIFABC", Err, objClassABC.lNumInt)
            
        Case 25392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ITEMCLASSABC_INEXISTENTE", Err, objClassABC.lNumInt)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152012)

    End Select

    'Fecha comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function ClassificacoesABC_Le(ByVal colNumIntCodigo As Collection) As Long
'Lê todas as Classificações coloca-as em ColNumIntCodigo

Dim lErro As Long
Dim lComando As Long
Dim tClassABC As typeClassificacaoABC
Dim objClassABC As ClassClassificacaoABC

On Error GoTo Erro_ClassificacoesABC_Le

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43456
    
    tClassABC.sCodigo = String(STRING_CLASSABC_CODIGO, 0)
    tClassABC.sDescricao = String(STRING_CLASSABC_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NumInt, Codigo,Descricao FROM ClassificacaoABC WHERE FilialEmpresa = ?", tClassABC.lNumInt, tClassABC.sCodigo, tClassABC.sDescricao, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 43457
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43458
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objClassABC = New ClassClassificacaoABC
    
        objClassABC.lNumInt = tClassABC.lNumInt
        objClassABC.sCodigo = tClassABC.sCodigo
        objClassABC.sDescricao = tClassABC.sDescricao
        
        colNumIntCodigo.Add objClassABC
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43459
        
    Loop
        
    Call Comando_Fechar(lComando)
    
    ClassificacoesABC_Le = SUCESSO
    
    Exit Function
        
Erro_ClassificacoesABC_Le:

    ClassificacoesABC_Le = Err
    
    Select Case Err
    
        Case 43456
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 43456, 43457, 43458
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSIFICACAOABC1", Err, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152013)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ClassificacaoABC_Le(ByVal objClassABC As ClassClassificacaoABC) As Long
'Lê a ClassificacaoABC no BD à partir do Código e Da FilialEmpresa passados em objClassABC

Dim lErro As Long
Dim lComando As Long
Dim tClassABC As typeClassificacaoABC

On Error GoTo Erro_ClassificacaoABC_Le

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43462
    
    tClassABC.sDescricao = String(STRING_CLASSABC_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT NumInt, Descricao, Data, MesInicial, AnoInicial, MesFinal, AnoFinal, FaixaA, FaixaB, TipoProduto, DemandaTotal, AtualizaProdutosFilial FROM ClassificacaoABC WHERE Codigo = ? AND FilialEmpresa = ?", _
        tClassABC.lNumInt, tClassABC.sDescricao, tClassABC.dtData, tClassABC.iMesInicial, tClassABC.iAnoInicial, tClassABC.iMesFinal, tClassABC.iAnoFinal, tClassABC.iFaixaA, tClassABC.iFaixaB, tClassABC.iTipoProduto, tClassABC.dDemandaTotal, tClassABC.iAtualizaProdutosFilial, objClassABC.sCodigo, objClassABC.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 43463
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43464
    
    'Se não encontrou a ClassificacaoABC --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 43465
    
    objClassABC.lNumInt = tClassABC.lNumInt
    objClassABC.sDescricao = tClassABC.sDescricao
    objClassABC.dtData = tClassABC.dtData
    objClassABC.iMesInicial = tClassABC.iMesInicial
    objClassABC.iAnoInicial = tClassABC.iAnoInicial
    objClassABC.iMesFinal = tClassABC.iMesFinal
    objClassABC.iAnoFinal = tClassABC.iAnoFinal
    objClassABC.iFaixaA = tClassABC.iFaixaA
    objClassABC.iFaixaB = tClassABC.iFaixaB
    objClassABC.iTipoProduto = tClassABC.iTipoProduto
    objClassABC.dDemandaTotal = tClassABC.dDemandaTotal
    objClassABC.iAtualizaProdutosFilial = tClassABC.iAtualizaProdutosFilial
        
    Call Comando_Fechar(lComando)
    
    ClassificacaoABC_Le = SUCESSO
    
    Exit Function
        
Erro_ClassificacaoABC_Le:

    ClassificacaoABC_Le = Err
    
    Select Case Err
    
        Case 43462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 43463, 43464
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSIFICACAOABC", Err, objClassABC.iFilialEmpresa, objClassABC.sCodigo)
            
        Case 43465  'Não encontrou a classificação ABC. A ser tratado na rotina chamadora.
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152014)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ClassificacaoABC_Le_NumInt(ByVal objClassABC As ClassClassificacaoABC) As Long
'Lê a ClassificacaoABC no BD à partir do NumInt passado em objClassABC

Dim lErro As Long
Dim lComando As Long
Dim tClassABC As typeClassificacaoABC

On Error GoTo Erro_ClassificacaoABC_Le_NumInt

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43497
    
    tClassABC.sCodigo = String(STRING_CLASSABC_CODIGO, 0)
    tClassABC.sDescricao = String(STRING_CLASSABC_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo, Descricao, Data, MesInicial, AnoInicial, MesFinal, AnoFinal, FaixaA, FaixaB, TipoProduto, DemandaTotal, AtualizaProdutosFilial FROM ClassificacaoABC WHERE NumInt = ?", _
        tClassABC.iFilialEmpresa, tClassABC.sCodigo, tClassABC.sDescricao, tClassABC.dtData, tClassABC.iMesInicial, tClassABC.iAnoInicial, tClassABC.iMesFinal, tClassABC.iAnoFinal, tClassABC.iFaixaA, tClassABC.iFaixaB, tClassABC.iTipoProduto, tClassABC.dDemandaTotal, tClassABC.iAtualizaProdutosFilial, objClassABC.lNumInt)
    If lErro <> AD_SQL_SUCESSO Then Error 43498
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43499
    
    'Se não encontrou a ClassificacaoABC --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 43500
    
    objClassABC.iFilialEmpresa = tClassABC.iFilialEmpresa
    objClassABC.sCodigo = tClassABC.sCodigo
    objClassABC.sDescricao = tClassABC.sDescricao
    objClassABC.dtData = tClassABC.dtData
    objClassABC.iMesInicial = tClassABC.iMesInicial
    objClassABC.iAnoInicial = tClassABC.iAnoInicial
    objClassABC.iMesFinal = tClassABC.iMesFinal
    objClassABC.iAnoFinal = tClassABC.iAnoFinal
    objClassABC.iFaixaA = tClassABC.iFaixaA
    objClassABC.iFaixaB = tClassABC.iFaixaB
    objClassABC.iTipoProduto = tClassABC.iTipoProduto
    objClassABC.dDemandaTotal = tClassABC.dDemandaTotal
    objClassABC.iAtualizaProdutosFilial = tClassABC.iAtualizaProdutosFilial
        
    Call Comando_Fechar(lComando)
    
    ClassificacaoABC_Le_NumInt = SUCESSO
    
    Exit Function
        
Erro_ClassificacaoABC_Le_NumInt:

    ClassificacaoABC_Le_NumInt = Err
    
    Select Case Err
    
        Case 43497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 43498, 43499
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSIFICACAOABC2", Err, objClassABC.lNumInt)
            
        Case 43500 'Não encontrou a Classificação ABC. A ser tratado na rotina chamadora.
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152015)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function EstoqueProduto_Le_Almoxarifados(ByVal sProduto As String, ByVal colAlmoxarifados As Collection) As Long
'Lê todos os Almoxarifados do Produto para todas as filiais

Dim lErro As Long
Dim lComando As Long
Dim objAlmoxarifado As ClassAlmoxarifado
Dim iAlmoxarifado As Integer

On Error GoTo Erro_EstoqueProduto_Le_Almoxarifados

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33626

    'Lê a tabela Estoque Produto
    lErro = Comando_Executar(lComando, "SELECT Almoxarifado FROM EstoqueProduto WHERE Produto = ?", iAlmoxarifado, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 33627

    'Lê o primeiro Almoxarifado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33628

    'Enquanto a tabela tiver dados
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objAlmoxarifado = New ClassAlmoxarifado

        objAlmoxarifado.iCodigo = iAlmoxarifado

        'Coloca o Almoxarifado lido na coleção
        colAlmoxarifados.Add objAlmoxarifado

        'Lê o próximo Almoxarifado
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33629

    Loop

    Call Comando_Fechar(lComando)

    EstoqueProduto_Le_Almoxarifados = SUCESSO

    Exit Function

Erro_EstoqueProduto_Le_Almoxarifados:

    EstoqueProduto_Le_Almoxarifados = Err

    Select Case Err

        Case 33626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33627, 33628, 33629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO3", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152016)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Function FornecedorProduto_Le(objFornecedorProduto As ClassFornecedorProduto) As Long
''Lê dados da associacao Fornecedor-Produto à partir da chave fornecida em objFornecedorProduto. Retorna os dados em objFornecedorProduto
'
'Dim lErro As Long
'Dim lComando As Long
'Dim tFornecedorProduto As typeFornecedorProduto
'
'On Error GoTo Erro_FornecedorProduto_Le
'
'    'Abertura comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then Error 28237
'
'    tFornecedorProduto.sProdutoFornecedor = String(STRING_PRODUTO_FORNECEDOR, 0)
'
'    'Pesquisa no BD o Fornecedor do Produto
'    lErro = Comando_Executar(lComando, "SELECT ProdutoFornecedor, LoteMinimo, LoteEconomico, QuantPedAbertos, TempoMedio, QuantPedida, QuantRecebida, Valor, DataPedido, DataReceb FROM FornecedorProduto WHERE Fornecedor = ? AND Produto = ?", _
'        tFornecedorProduto.sProdutoFornecedor, tFornecedorProduto.dLoteMinimo, tFornecedorProduto.dLoteEconomico, tFornecedorProduto.dQuantPedAbertos, tFornecedorProduto.iTempoMedio, tFornecedorProduto.dQuantPedida, tFornecedorProduto.dQuantRecebida, tFornecedorProduto.dValor, tFornecedorProduto.dtDataPedido, tFornecedorProduto.dtDataReceb, objFornecedorProduto.lFornecedor, objFornecedorProduto.sProduto)
'    If lErro <> AD_SQL_SUCESSO Then Error 28238
'
'    'Verifica se o Fornecedor do Produto existe
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28239
'
'    'Não encontrou o FornecedorProduto
'    If lErro = AD_SQL_SEM_DADOS Then Error 28240
'
'    'Preenche objFornecedorProduto
'    With objFornecedorProduto
'        .sProdutoFornecedor = tFornecedorProduto.sProdutoFornecedor
'        .dLoteMinimo = tFornecedorProduto.dLoteMinimo
'        .dLoteEconomico = tFornecedorProduto.dLoteEconomico
'        .dQuantPedAbertos = tFornecedorProduto.dQuantPedAbertos
'        .iTempoMedio = tFornecedorProduto.iTempoMedio
'        .dQuantPedida = tFornecedorProduto.dQuantPedida
'        .dQuantRecebida = tFornecedorProduto.dQuantRecebida
'        .dValor = tFornecedorProduto.dValor
'        .dtDataPedido = tFornecedorProduto.dtDataPedido
'        .dtDataReceb = tFornecedorProduto.dtDataReceb
'    End With
'
'    'Fechamento comando
'    Call Comando_Fechar(lComando)
'
'    FornecedorProduto_Le = SUCESSO
'
'    Exit Function
'
'Erro_FornecedorProduto_Le:
'
'    FornecedorProduto_Le = Err
'
'    Select Case Err
'
'        Case 28237
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 28238, 28239
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTO", Err, objFornecedorProduto.lFornecedor, objFornecedorProduto.sProduto)
'
'        Case 28240
'            'Não encontrou o Fornecedor do Produto. Erro tratado na rotina chamadora.
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152017)
'
'    End Select
'
'    'Fechamento comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Function FornecedorProduto_Le_Fornecedores(colFornecedores As Collection, objProduto As ClassProduto) As Long
''Lê todos nomes reduzidos de Fornecedores associados a sProduto e coloca-os em colFornecedores
'
'Dim lErro As Long
'Dim lComando As Long
'Dim sNomeReduzido As String
'
'On Error GoTo Erro_FornecedorProduto_Le_Fornecedores
'
'    'Abertura comandos
'    lComando = Comando_Abrir()
'    If lComando = 0 Then Error 28270
'
'    sNomeReduzido = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)
'
'    'Lê a tabela de FornecedorProduto e Fornecedores
'    lErro = Comando_Executar(lComando, "SELECT Fornecedores.NomeReduzido FROM FornecedorProduto, Fornecedores WHERE FornecedorProduto.Produto = ? AND FornecedorProduto.Fornecedor = Fornecedores.Codigo ORDER BY FornecedorProduto.Fornecedor", sNomeReduzido, objProduto.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then Error 28271
'
'    'Lê o primeiro Fornecedor de Produto
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28272
'
'    Do While lErro <> AD_SQL_SEM_DADOS
'
'        'Coloca o código lido na coleção
'         colFornecedores.Add sNomeReduzido
'
'        'Lê o próximo código
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28273
'
'    Loop
'
'    'Fechamento comando
'    Call Comando_Fechar(lComando)
'
'    FornecedorProduto_Le_Fornecedores = SUCESSO
'
'    Exit Function
'
'Erro_FornecedorProduto_Le_Fornecedores:
'
'    FornecedorProduto_Le_Fornecedores = Err
'
'    Select Case Err
'
'        Case 28270
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 28271, 28272, 28273
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTO_FORNECEDOR", Err, objProduto.sCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152018)
'
'    End Select
'
'    'Fechamento comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Function InventarioItens_Le(ByVal objInventario As ClassInventario) As Long
'preenche a colecao de itens do inventario passado e a data de objInventario
'os itens sao recuperados por ordem de numintdoc

Dim lErro As Long
Dim lComando As Long
Dim tItemInventario As typeItemInventario
Dim sComando_SQL As String
Dim dtData As Date
Dim objItemInventario As ClassItemInventario

On Error GoTo Erro_InventarioItens_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41003

    tItemInventario.sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
    tItemInventario.sEtiqueta = String(STRING_INVENTARIO_ETIQUETA, 0)
    tItemInventario.sProduto = String(STRING_PRODUTO, 0)
    tItemInventario.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemInventario.sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemInventario.sContaContabilEst = String(STRING_CONTA, 0)
    tItemInventario.sContaContabilInv = String(STRING_CONTA, 0)

    'Maristela: Inclusão do AtualizaSoLote
    sComando_SQL = "SELECT Inventario.ContaContabilEst, Inventario.ContaContabilInv, Inventario.NumIntDoc, Inventario.Data, Inventario.Produto, Produtos.Descricao, " & _
    "Inventario.SiglaUM, Inventario.Quantidade, Inventario.QuantEst, Inventario.Custo, Inventario.Almoxarifado, Almoxarifado.NomeReduzido, Inventario.Etiqueta, Inventario.Tipo, Inventario.AtualizaSoLote " & _
    "FROM Inventario, Produtos, Almoxarifado WHERE Inventario.Produto=Produtos.Codigo AND Inventario.Almoxarifado=Almoxarifado.Codigo AND " & _
    "Inventario.Codigo = ? AND Inventario.FilialEmpresa = ? ORDER BY Inventario.NumIntDoc"

    lErro = Comando_Executar(lComando, sComando_SQL, tItemInventario.sContaContabilEst, tItemInventario.sContaContabilInv, tItemInventario.lNumIntDoc, dtData, tItemInventario.sProduto, tItemInventario.sProdutoDesc, tItemInventario.sSiglaUM, tItemInventario.dQuantidade, tItemInventario.dQuantEst, tItemInventario.dCusto, tItemInventario.iAlmoxarifado, tItemInventario.sAlmoxarifadoNomeRed, tItemInventario.sEtiqueta, tItemInventario.iTipo, tItemInventario.iAtualizaSoLote, objInventario.sCodigo, giFilialEmpresa)
    If lErro <> SUCESSO Then Error 41004

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41005

    If lErro = AD_SQL_SEM_DADOS Then Error 41006
        
    objInventario.dtData = dtData
    objInventario.iFilialEmpresa = giFilialEmpresa

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItemInventario = objInventario.colItens.Add(tItemInventario.lNumIntDoc, tItemInventario.sProduto, tItemInventario.sProdutoDesc, tItemInventario.sSiglaUM, tItemInventario.dQuantidade, tItemInventario.dCusto, tItemInventario.iAlmoxarifado, tItemInventario.sAlmoxarifadoNomeRed, tItemInventario.iTipo, tItemInventario.sEtiqueta, tItemInventario.sContaContabilEst, tItemInventario.sContaContabilInv, 0, 0)
        objItemInventario.dQuantEst = tItemInventario.dQuantEst
        
        'Maristela: Inclusão do AtualizaSoLote
        objItemInventario.iAtualizaSoLote = tItemInventario.iAtualizaSoLote

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41007

    Loop

    Call Comando_Fechar(lComando)

    InventarioItens_Le = SUCESSO

    Exit Function

Erro_InventarioItens_Le:

    InventarioItens_Le = Err

    Select Case Err

        Case 41003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41004, 41005, 41007
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIO", Err, objInventario.sCodigo)

        Case 41006 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152019)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Inventario_Le(ByVal objInventario As ClassInventario) As Long
'Le atributos de objInventario EXCETO colItens
'Faz Select DISTINCT

Dim lErro As Long
Dim lComando As Long
Dim tInventario As typeInventario
Dim sComando_SQL As String
Dim iFilialEmpresa As Integer

On Error GoTo Erro_Inventario_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41008

    tInventario.sCodigo = String(STRING_INVENTARIO_CODIGO, 0)
    
    If objInventario.iFilialEmpresa = 0 Then
        iFilialEmpresa = giFilialEmpresa
    Else
        iFilialEmpresa = objInventario.iFilialEmpresa
    End If

    sComando_SQL = "SELECT DISTINCT Codigo, FilialEmpresa, Data, Hora FROM Inventario WHERE Inventario.Codigo = ? AND Inventario.FilialEmpresa = ?"

    lErro = Comando_Executar(lComando, sComando_SQL, tInventario.sCodigo, tInventario.iFilialEmpresa, tInventario.dtData, tInventario.dHora, objInventario.sCodigo, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 41009

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41010

    If lErro = AD_SQL_SEM_DADOS Then Error 41011

    objInventario.sCodigo = tInventario.sCodigo
    objInventario.iFilialEmpresa = tInventario.iFilialEmpresa
    objInventario.dtData = tInventario.dtData
    objInventario.dtHora = tInventario.dHora

    Call Comando_Fechar(lComando)

    Inventario_Le = SUCESSO

    Exit Function

Erro_Inventario_Le:

    Inventario_Le = Err

    Select Case Err

        Case 41008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41009, 41010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIO", Err, objInventario.sCodigo)

        Case 41011 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152020)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function SldMesEst_Le_Custos(ByVal iMes As Integer, ByVal objSldMesEst As ClassSldMesEst) As Long
'Le na tabela SldMesEst os custos relativos aos parametros passados : iMes e objSldMesEst

Dim lErro As Long
Dim lComando As Long
Dim sComando_SQL As String
Dim sMes As String
Dim dCustoProducao As Double, dCustoStandard As Double
Dim dCustoMedio As Double, dCustoReposicao As Double

On Error GoTo Erro_SldMesEst_Le_Custos

    sMes = CStr(iMes)

    sComando_SQL = "SELECT CustoProducao" & sMes & " , CustoStandard" & sMes & " , CustoMedio" & sMes & " , CustoReposicao" & sMes & " FROM SldMesEst WHERE  Ano = ? AND FilialEmpresa = ? AND Produto = ?"

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 41752

    lErro = Comando_Executar(lComando, sComando_SQL, dCustoProducao, dCustoStandard, dCustoMedio, dCustoReposicao, objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 41753

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 41754

    If lErro = AD_SQL_SEM_DADOS Then gError 41755
    
    '######################################################
    'Inserido por Wagner 10/02/2006
    lErro = CF("SldMesEst_Le_CustoMedio2", lComando, objSldMesEst.iFilialEmpresa, objSldMesEst.iAno, objSldMesEst.sProduto, iMes, dCustoMedio)
    If lErro <> AD_SQL_SUCESSO Then gError 141765
    '######################################################

    objSldMesEst.dCustoProducao(iMes) = dCustoProducao
    objSldMesEst.dCustoStandard(iMes) = dCustoStandard
    objSldMesEst.dCustoMedio(iMes) = dCustoMedio
    objSldMesEst.dCustoReposicao(iMes) = dCustoReposicao

    Call Comando_Fechar(lComando)

    SldMesEst_Le_Custos = SUCESSO

    Exit Function

Erro_SldMesEst_Le_Custos:

    SldMesEst_Le_Custos = gErr

    Select Case gErr

        Case 41752
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 41753, 41754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", gErr, objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
        
        Case 41755, 141765
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152021)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CustoMedioAtual_Le(ByVal sProduto As String, dCustoMedioAtual As Double, ByVal iFilialEmpresa As Integer) As Long
'Parametros a serem passados: sProduto
'Le o Ultimo CustoMedio do Ano-Mes aberto e retorna em dCustoMedioAtual

Dim objEstoqueMes As New ClassEstoqueMes
Dim lErro As Long

On Error GoTo Erro_CustoMedioAtual_Le

    objEstoqueMes.iFilialEmpresa = iFilialEmpresa

    lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
    If lErro <> SUCESSO Then Error 49603
    
    'Le o CustoMedio do Ano/Mes passado como parametro e retorna em dCustoMedioAtual
    lErro = Calcula_CustoMedioAtual(objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes, sProduto, dCustoMedioAtual)
    If lErro <> SUCESSO And lErro <> 55040 Then Error 55044
    
    CustoMedioAtual_Le = SUCESSO
    
    Exit Function
    
Erro_CustoMedioAtual_Le:

    CustoMedioAtual_Le = Err
    
    Select Case Err
        
        Case 49603, 55044
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152022)
        
    End Select
    
    Exit Function
    
End Function

Function InvLotePendente_Le(ByVal objInvLote As ClassInvLote) As Long

Dim lErro As Long
Dim sComando_SQL As String
Dim lComando As Long
Dim tInvLote As typeInvLote

On Error GoTo Erro_InvLotePendente_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41178

    
    tInvLote.sDescricao = String(STRING_INVLOTE_DESCRICAO, 0)
    
    sComando_SQL = "SELECT FilialEmpresa, Lote, Descricao, NumItensAtual, NumItensInf, IdAtualizacao FROM InvLotePendente WHERE FilialEmpresa = ? AND Lote = ?"
    lErro = Comando_Executar(lComando, sComando_SQL, tInvLote.iFilialEmpresa, tInvLote.iLote, tInvLote.sDescricao, tInvLote.iNumItensAtual, tInvLote.iNumItensInf, tInvLote.iIdAtualizacao, objInvLote.iFilialEmpresa, objInvLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 41179

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41180

    If lErro = AD_SQL_SEM_DADOS Then Error 41181
    
    objInvLote.iFilialEmpresa = tInvLote.iFilialEmpresa
    objInvLote.iLote = tInvLote.iLote
    objInvLote.sDescricao = tInvLote.sDescricao
    objInvLote.iNumItensAtual = tInvLote.iNumItensAtual
    objInvLote.iNumItensInf = tInvLote.iNumItensInf
    objInvLote.iIdAtualizacao = tInvLote.iIdAtualizacao

    Call Comando_Fechar(lComando)

    InvLotePendente_Le = SUCESSO

    Exit Function

Erro_InvLotePendente_Le:

    InvLotePendente_Le = Err

    Select Case Err

        Case 41178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41179, 41180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVLOTE", Err, objInvLote.iLote)

        Case 41181 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152023)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function InvLote_Le(ByVal objInvLote As ClassInvLote) As Long

Dim lErro As Long
Dim sComando_SQL As String
Dim lComando As Long
Dim tInvLote As typeInvLote

On Error GoTo Erro_InvLote_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41174

    tInvLote.sDescricao = String(STRING_INVLOTE_DESCRICAO, 0)
    
    sComando_SQL = "SELECT FilialEmpresa, Lote, Descricao, NumItensAtual, NumItensInf, IdAtualizacao FROM InvLote WHERE FilialEmpresa = ? AND Lote = ?"
    lErro = Comando_Executar(lComando, sComando_SQL, tInvLote.iFilialEmpresa, tInvLote.iLote, tInvLote.sDescricao, tInvLote.iNumItensAtual, tInvLote.iNumItensInf, tInvLote.iIdAtualizacao, objInvLote.iFilialEmpresa, objInvLote.iLote)
    If lErro <> AD_SQL_SUCESSO Then Error 41175

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41176

    If lErro = AD_SQL_SEM_DADOS Then Error 41177
    
    objInvLote.iFilialEmpresa = tInvLote.iFilialEmpresa
    objInvLote.iLote = tInvLote.iLote
    objInvLote.sDescricao = tInvLote.sDescricao
    objInvLote.iNumItensAtual = tInvLote.iNumItensAtual
    objInvLote.iNumItensInf = tInvLote.iNumItensInf
    objInvLote.iIdAtualizacao = tInvLote.iIdAtualizacao

    Call Comando_Fechar(lComando)

    InvLote_Le = SUCESSO

    Exit Function

Erro_InvLote_Le:

    InvLote_Le = Err

    Select Case Err

        Case 41174
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41175, 41176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVLOTE", Err, objInvLote.iLote)

        Case 41177 'Tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152024)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function InventarioPendenteItens_Le(ByVal objInventario As ClassInventario) As Long
'preenche a colecao de itens do inventario pendente passado e a data de objInventario
'os itens sao recuperados por ordem de numintdoc

Dim lErro As Long
Dim lComando As Long
Dim tItemInventario As typeItemInventario
Dim sComando_SQL As String
Dim dtData As Date
Dim iLote As Integer
Dim objItemInventario As ClassItemInventario

On Error GoTo Erro_InventarioPendenteItens_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41149

    tItemInventario.sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
    tItemInventario.sEtiqueta = String(STRING_INVENTARIO_ETIQUETA, 0)
    tItemInventario.sProduto = String(STRING_PRODUTO, 0)
    tItemInventario.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemInventario.sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemInventario.sContaContabilEst = String(STRING_CONTA, 0)
    tItemInventario.sContaContabilInv = String(STRING_CONTA, 0)
    tItemInventario.sLoteProduto = String(STRING_OPCODIGO, 0)
    
    'Maristela:Inclusão do Campo AtualizaSoLote
    sComando_SQL = "SELECT InventarioPendente.ContaContabilEst, InventarioPendente.ContaContabilInv, InventarioPendente.NumIntDoc, InventarioPendente.Data, InventarioPendente.Produto, Produtos.Descricao, " & _
        "InventarioPendente.SiglaUM, InventarioPendente.Quantidade, InventarioPendente.Custo, InventarioPendente.Almoxarifado, Almoxarifado.NomeReduzido, InventarioPendente.Etiqueta, InventarioPendente.Lote, " & _
        "InventarioPendente.Tipo, InventarioPendente.LoteProduto, InventarioPendente.FilialOP, InventarioPendente.AtualizaSoLote FROM InventarioPendente, Produtos, Almoxarifado WHERE InventarioPendente.Produto=Produtos.Codigo AND InventarioPendente.Almoxarifado=Almoxarifado.Codigo AND " & _
        "InventarioPendente.Codigo = ? AND InventarioPendente.FilialEmpresa = ? ORDER BY InventarioPendente.NumIntDoc"

    With tItemInventario
        lErro = Comando_Executar(lComando, sComando_SQL, .sContaContabilEst, .sContaContabilInv, .lNumIntDoc, dtData, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .dCusto, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .sEtiqueta, iLote, .iTipo, .sLoteProduto, .iFilialOP, .iAtualizaSoLote, objInventario.sCodigo, giFilialEmpresa)
    End With
    If lErro <> SUCESSO Then Error 41150

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41151

    If lErro <> AD_SQL_SUCESSO Then Error 41152

    objInventario.dtData = dtData
    objInventario.iFilialEmpresa = giFilialEmpresa
    objInventario.iLote = iLote

    Do While lErro <> AD_SQL_SEM_DADOS

        With tItemInventario
            'Maristela
            'objInventario.colItens.Add .lNumIntDoc, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .dCusto, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .iTipo, .sEtiqueta, .sContaContabilEst, .sContaContabilInv, .sLoteProduto, .iFilialOP
            Set objItemInventario = objInventario.colItens.Add(.lNumIntDoc, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .dCusto, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .iTipo, .sEtiqueta, .sContaContabilEst, .sContaContabilInv, .sLoteProduto, .iFilialOP)
        End With
        
        'Maristela: Inclusão do AtualizaSoLote
        objItemInventario.iAtualizaSoLote = tItemInventario.iAtualizaSoLote
       
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41153

    Loop

    Call Comando_Fechar(lComando)

    InventarioPendenteItens_Le = SUCESSO

    Exit Function

Erro_InventarioPendenteItens_Le:

    InventarioPendenteItens_Le = Err

    Select Case Err

        Case 41149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41150, 41151, 41153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIO", Err, objInventario.sCodigo)

        Case 41152 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152025)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function InventarioPendente_Le(ByVal objInventario As ClassInventario) As Long
'Le atributos de objInventario EXCETO colItens
'Faz Select DISTINCT

Dim lErro As Long
Dim lComando As Long
Dim tInventario As typeInventario
Dim sComando_SQL As String

On Error GoTo Erro_InventarioPendente_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41154

    tInventario.sCodigo = String(STRING_INVENTARIO_CODIGO, 0)

    '???? Para que a Qualificacao dos campos da Tabela se só está usando uma tabela??
    '???? O DISTINCT é desnecessário.
    sComando_SQL = "SELECT DISTINCT Codigo, FilialEmpresa, Data, Lote, Hora FROM InventarioPendente WHERE Codigo = ? AND FilialEmpresa = ?"

    lErro = Comando_Executar(lComando, sComando_SQL, tInventario.sCodigo, tInventario.iFilialEmpresa, tInventario.dtData, tInventario.iLote, tInventario.dHora, objInventario.sCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 41155

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41156

    If lErro = AD_SQL_SEM_DADOS Then Error 41157

    objInventario.sCodigo = tInventario.sCodigo
    objInventario.iFilialEmpresa = tInventario.iFilialEmpresa
    objInventario.dtData = tInventario.dtData
    objInventario.iLote = tInventario.iLote
    objInventario.dtHora = tInventario.dHora

    Call Comando_Fechar(lComando)

    InventarioPendente_Le = SUCESSO

    Exit Function

Erro_InventarioPendente_Le:

    InventarioPendente_Le = Err

    Select Case Err

        Case 41154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41155, 41156
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVENTARIO", Err, objInventario.sCodigo)

        Case 41157 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152026)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function InvLotePendente_Le_Desatualizados(ByVal iFilialEmpresa As Integer, ByVal colInvLote As Collection, ByVal iIncluiLotesAtualizando As Integer) As Long
'Le todos os lotes com IdAtualizacao = LOTE_DESATUALIZADO ou todos pendentes

Dim lComando As Long
Dim lErro As Long
Dim objInvLote As ClassInvLote
Dim tInvLote As typeInvLote
Dim sComando_SQL As String

On Error GoTo Erro_InvLotePendente_Le_Desatualizados

    tInvLote.sDescricao = String(STRING_INVLOTE_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41045

    If iIncluiLotesAtualizando = LOTE_DESATUALIZADO Then
        sComando_SQL = "SELECT FilialEmpresa, Lote, Descricao, NumItensAtual, NumItensInf, IdAtualizacao FROM InvLotePendente WHERE FilialEmpresa = ? AND IdAtualizacao = ?"
        lErro = Comando_Executar(lComando, sComando_SQL, tInvLote.iFilialEmpresa, tInvLote.iLote, tInvLote.sDescricao, tInvLote.iNumItensAtual, tInvLote.iNumItensInf, tInvLote.iIdAtualizacao, iFilialEmpresa, LOTE_DESATUALIZADO)
        If lErro <> AD_SQL_SUCESSO Then Error 41046
    Else
        sComando_SQL = "SELECT FilialEmpresa, Lote, Descricao, NumItensAtual, NumItensInf, IdAtualizacao FROM InvLotePendente WHERE FilialEmpresa = ?"
        lErro = Comando_Executar(lComando, sComando_SQL, tInvLote.iFilialEmpresa, tInvLote.iLote, tInvLote.sDescricao, tInvLote.iNumItensAtual, tInvLote.iNumItensInf, tInvLote.iIdAtualizacao, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 41239
    End If

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41047

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objInvLote = New ClassInvLote

            objInvLote.iFilialEmpresa = iFilialEmpresa
            objInvLote.iLote = tInvLote.iLote
            objInvLote.sDescricao = tInvLote.sDescricao
            objInvLote.iNumItensAtual = tInvLote.iNumItensAtual
            objInvLote.iNumItensInf = tInvLote.iNumItensInf
            objInvLote.iIdAtualizacao = tInvLote.iIdAtualizacao

            colInvLote.Add objInvLote

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41048

    Loop

    Call Comando_Fechar(lComando)

    InvLotePendente_Le_Desatualizados = SUCESSO

Exit Function

Erro_InvLotePendente_Le_Desatualizados:

    InvLotePendente_Le_Desatualizados = Err

    Select Case Err

        Case 41045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41046, 41047, 41048, 41239
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INVLOTEPENDENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152027)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function InvLotePendente_Le1(ByVal iFilialEmpresa As Integer) As Long
'Lê todos os lotes de uma FilialEmpresa

Dim lErro As Long
Dim lComando As Long
Dim iLote As Integer

On Error GoTo Erro_InvLotePendente_Le1


    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41240

    lErro = Comando_Executar(lComando, "SELECT Lote FROM InvLotePendente WHERE FilialEmpresa = ? ORDER BY Lote", iLote, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 41241

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41242

    'Não encontrou nenhum lote
    If lErro = AD_SQL_SEM_DADOS Then Error 41243

    'Libera comando
    Call Comando_Fechar(lComando)

    InvLotePendente_Le1 = SUCESSO

    Exit Function

Erro_InvLotePendente_Le1:

    InvLotePendente_Le1 = Err

    Select Case Err

        Case 41240
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41241, 41242
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOTEPENDENTE1", Err)

        Case 41243

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152028)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function OrdemDeProducao_Le_ComItens(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'carrega do bd a ordem de producao inclusive os seus itens

Dim lErro As Long, lCodPrestador As Long
Dim iIndice As Integer
Dim dtDataEmissao As Date
Dim iNumItens As Integer
Dim iNumItensBaixados As Integer
Dim tItemOP As typeItemOP
Dim objItemOP As ClassItemOP
Dim alComando(0 To 2) As Long
Dim tItemOPGrade As typeItemOP
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim iTipo As Integer
Dim iTipoTerc As Integer, lCodTerc As Long, iFilialTerc As Integer, iEscaninhoTerc As Integer
Dim iIgnoraEst As Integer

On Error GoTo Erro_OrdemDeProducao_Le_ComItens

    For iIndice = LBound(alComando) To UBound(alComando)

        'Abertura comando
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 21955

    Next

    'Leitura da tabela de Ordens de Produção
    lErro = Comando_Executar(alComando(0), "SELECT DataEmissao, NumItens, NumItensBaixados, CodPrestador, Tipo, TipoTerc, CodTerc, FilialTerc, EscaninhoTerc, IgnoraEst FROM OrdensDeProducao WHERE FilialEmpresa = ? AND Codigo = ? ", dtDataEmissao, iNumItens, iNumItensBaixados, lCodPrestador, iTipo, iTipoTerc, lCodTerc, iFilialTerc, iEscaninhoTerc, iIgnoraEst, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 21956

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 21957

    If lErro <> AD_SQL_SUCESSO Then gError 21960

    objOrdemDeProducao.dtDataEmissao = dtDataEmissao
    objOrdemDeProducao.iNumItens = iNumItens
    objOrdemDeProducao.iNumItensBaixados = iNumItensBaixados
    objOrdemDeProducao.lCodPrestador = lCodPrestador
    objOrdemDeProducao.iTipo = iTipo
    objOrdemDeProducao.iTipoTerc = iTipoTerc
    objOrdemDeProducao.lCodTerc = lCodTerc
    objOrdemDeProducao.iFilialTerc = iFilialTerc
    objOrdemDeProducao.iEscaninhoTerc = iEscaninhoTerc
    objOrdemDeProducao.iIgnoraEst = iIgnoraEst

    tItemOP.sProduto = String(STRING_PRODUTO, 0)
    tItemOP.sCcl = String(STRING_CCL, 0)
    tItemOP.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemOP.sVersao = String(STRING_KIT_VERSAO, 0)
    tItemOP.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)


    'Leitura da tabela de Itens de Ordem de Produção
    lErro = Comando_Executar(alComando(1), "SELECT ItensOPGrade.NumIntDoc, ItensOPGrade.Item, ItensOPGrade.Produto, Descricao, ItensOPGrade.FilialPedido, " & _
    "ItensOPGrade.CodPedido, ItensOPGrade.Almoxarifado, ItensOPGrade.Beneficiamento, ItensOPGrade.CCL, ItensOPGrade.SiglaUM, ItensOPGrade.Quantidade, ItensOPGrade.DataInicioProd, " & _
    "ItensOPGrade.DataFimProd,  ItensOPGrade.Prioridade, ItensOPGrade.QuantidadeProd, ItensOPGrade.Situacao, ItensOPGrade.Destinacao, ItensOPGrade.NumIntEquipamento, ItensOPGrade.Versao, ItensOPGrade.Tipo FROM ItensOPGrade, Produtos WHERE ItensOPGrade.FilialEmpresa = ? AND ItensOPGrade.Codigo = ? AND ItensOPGrade.Produto = Produtos.Codigo ORDER BY Item", _
    tItemOP.lNumIntDoc, tItemOP.iItem, tItemOP.sProduto, tItemOP.sDescricao, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.iAlmoxarifado, tItemOP.iBeneficiamento, tItemOP.sCcl, tItemOP.sSiglaUM, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.dQuantidadeProd, tItemOP.iSituacao, tItemOP.iDestinacao, tItemOP.lNumIntEquipamento, tItemOP.sVersao, tItemOP.iTipo, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126553

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126554

    'Enquanto existirem itens para uma determinada ordem de produção
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItemOP = New ClassItemOP

        objItemOP.lNumIntDoc = tItemOP.lNumIntDoc
        objItemOP.iItem = tItemOP.iItem
        objItemOP.sProduto = tItemOP.sProduto
        objItemOP.sDescricao = tItemOP.sDescricao
        objItemOP.iAlmoxarifado = tItemOP.iAlmoxarifado
        objItemOP.iBeneficiamento = tItemOP.iBeneficiamento
        objItemOP.sCcl = tItemOP.sCcl
        objItemOP.sSiglaUM = tItemOP.sSiglaUM
        objItemOP.dQuantidade = tItemOP.dQuantidade
        objItemOP.dtDataInicioProd = tItemOP.dtDataInicioProd
        objItemOP.dtDataFimProd = tItemOP.dtDataFimProd
        objItemOP.iPrioridade = tItemOP.iPrioridade
        objItemOP.dQuantidadeProd = tItemOP.dQuantidadeProd
        objItemOP.iSituacao = tItemOP.iSituacao
        objItemOP.iDestinacao = tItemOP.iDestinacao
        objItemOP.iFilialPedido = tItemOP.iFilialPedido
        objItemOP.lCodPedido = tItemOP.lCodPedido
        objItemOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
        objItemOP.sCodigo = objOrdemDeProducao.sCodigo
        objItemOP.iPossuiGrade = MARCADO
        
        objItemOP.lNumIntEquipamento = tItemOP.lNumIntEquipamento
        objItemOP.sVersao = tItemOP.sVersao
        objItemOP.iTipo = tItemOP.iTipo
        
        'Inserido por Jorge Specian - 11/05/2005
        '---------------------------------------
        lErro = CF("OrdemDeProducao_Le_Operacoes", objItemOP)
        If lErro <> SUCESSO And lErro <> 137016 Then gError 137009
        '---------------------------------------
        
        objOrdemDeProducao.colItens.Add objItemOP

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126557

    Loop
    

    tItemOP.sProduto = String(STRING_PRODUTO, 0)
    tItemOP.sCcl = String(STRING_CCL, 0)
    tItemOP.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemOP.sVersao = String(STRING_KIT_VERSAO, 0)
    tItemOP.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

    '##############################################################
    'INSERIDO POR WAGNER
    tItemOP.sJustificativaBloqueio = String(STRING_MAX_NOME_ARQUIVO, 0)
    '##############################################################

    'Leitura da tabela de Itens de Ordem de Produção
    '##############################################################
    'ALTERADO POR WAGNER
    lErro = Comando_Executar(alComando(1), "SELECT ItensOrdemProducao.NumIntItemOPGrade, ItensOrdemProducao.NumIntDoc, ItensOrdemProducao.Item, ItensOrdemProducao.Produto, Descricao, ItensOrdemProducao.FilialPedido, " & _
    "ItensOrdemProducao.CodPedido, ItensOrdemProducao.Almoxarifado, ItensOrdemProducao.Beneficiamento, ItensOrdemProducao.CCL, ItensOrdemProducao.SiglaUM, ItensOrdemProducao.Quantidade, ItensOrdemProducao.DataInicioProd, " & _
    "ItensOrdemProducao.DataFimProd,  ItensOrdemProducao.Prioridade, ItensOrdemProducao.QuantidadeProd, ItensOrdemProducao.Situacao, ItensOrdemProducao.Destinacao, ItensOrdemProducao.NumIntEquipamento, ItensOrdemProducao.NumIntEquipamento2, ItensOrdemProducao.Versao, ItensOrdemProducao.Tipo, ItensOrdemProducao.NumIntDocPai, ItensOrdemProducao.JustificativaBloqueio, ItensOrdemProducao.ProduzLogo FROM ItensOrdemProducao, Produtos WHERE ItensOrdemProducao.FilialEmpresa = ? AND ItensOrdemProducao.Codigo = ? AND ItensOrdemProducao.Produto = Produtos.Codigo ORDER BY Item", _
    tItemOP.lNumIntItemOP, tItemOP.lNumIntDoc, tItemOP.iItem, tItemOP.sProduto, tItemOP.sDescricao, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.iAlmoxarifado, tItemOP.iBeneficiamento, tItemOP.sCcl, tItemOP.sSiglaUM, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.dQuantidadeProd, tItemOP.iSituacao, tItemOP.iDestinacao, tItemOP.lNumIntEquipamento, tItemOP.lNumIntEquipamento2, tItemOP.sVersao, tItemOP.iTipo, tItemOP.lNumIntDocPai, tItemOP.sJustificativaBloqueio, tItemOP.iProduzLogo, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 21964
    '#############################################################

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 21965

    'Enquanto existirem itens para uma determinada ordem de produção
    Do While lErro <> AD_SQL_SEM_DADOS

        If tItemOP.lNumIntItemOP <> 0 Then

            Set objItemRomaneioGrade = New ClassItemRomaneioGrade
            
            objItemRomaneioGrade.sProduto = tItemOP.sProduto
            objItemRomaneioGrade.sDescricao = tItemOP.sDescricao
            objItemRomaneioGrade.iAlmoxarifado = tItemOP.iAlmoxarifado
            objItemRomaneioGrade.sUMEstoque = tItemOP.sSiglaUMEstoque
            objItemRomaneioGrade.dQuantidade = tItemOP.dQuantidade
            objItemRomaneioGrade.sVersao = tItemOP.sVersao
            objItemRomaneioGrade.lNumIntDoc = tItemOP.lNumIntDoc

            For Each objItemOP In objOrdemDeProducao.colItens
            
                If objItemOP.lNumIntDoc = tItemOP.lNumIntItemOP Then

                    objItemOP.colItensRomaneioGrade.Add objItemRomaneioGrade
                    Exit For
                    
                End If
                
            Next

        Else

            Set objItemOP = New ClassItemOP

            objItemOP.lNumIntDoc = tItemOP.lNumIntDoc
            objItemOP.iItem = tItemOP.iItem
            objItemOP.sProduto = tItemOP.sProduto
            objItemOP.sDescricao = tItemOP.sDescricao
            objItemOP.iAlmoxarifado = tItemOP.iAlmoxarifado
            objItemOP.iBeneficiamento = tItemOP.iBeneficiamento
            objItemOP.sCcl = tItemOP.sCcl
            objItemOP.sSiglaUM = tItemOP.sSiglaUM
            objItemOP.dQuantidade = tItemOP.dQuantidade
            objItemOP.dtDataInicioProd = tItemOP.dtDataInicioProd
            objItemOP.dtDataFimProd = tItemOP.dtDataFimProd
            objItemOP.iPrioridade = tItemOP.iPrioridade
            objItemOP.dQuantidadeProd = tItemOP.dQuantidadeProd
            objItemOP.iSituacao = tItemOP.iSituacao
            objItemOP.iDestinacao = tItemOP.iDestinacao
            objItemOP.iFilialPedido = tItemOP.iFilialPedido
            objItemOP.lCodPedido = tItemOP.lCodPedido
            objItemOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
            objItemOP.sCodigo = objOrdemDeProducao.sCodigo
            objItemOP.iPossuiGrade = DESMARCADO
            
            objItemOP.lNumIntEquipamento = tItemOP.lNumIntEquipamento
            objItemOP.lNumIntEquipamento2 = tItemOP.lNumIntEquipamento2
            objItemOP.sVersao = tItemOP.sVersao
            objItemOP.iTipo = tItemOP.iTipo
            
            '###################################################
            'INSERIDO POR WAGNER
            objItemOP.lNumIntDocPai = tItemOP.lNumIntDocPai
            objItemOP.sJustificativaBloqueio = tItemOP.sJustificativaBloqueio
            objItemOP.iProduzLogo = tItemOP.iProduzLogo
            '###################################################
    
            'Inserido por Jorge Specian - 11/05/2005
            '---------------------------------------
            lErro = CF("OrdemDeProducao_Le_Operacoes", objItemOP)
            If lErro <> SUCESSO And lErro <> 137016 Then gError 137010
            '---------------------------------------
        
            objOrdemDeProducao.colItens.Add objItemOP

        End If

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 21967

    Loop

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducao_Le_ComItens = SUCESSO

    Exit Function

Erro_OrdemDeProducao_Le_ComItens:

    OrdemDeProducao_Le_ComItens = gErr

    Select Case gErr

        Case 21955
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 21956, 21957
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", gErr)

        Case 21958, 21959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAOBAIXADAS", gErr)

        Case 21960 'Ordem de Producao nao encontrada

        Case 21961, 21962
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAOBAIXADAS", gErr)

        Case 21964, 21965, 21967
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAO", gErr)
        
        Case 126553, 126554, 126557
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOPGRADE1", gErr, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)

        Case 137009, 137010
            'erros tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152029)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function EstoqueProdutoCC_Le(ByVal objEstoqueProduto As ClassEstoqueProduto) As Long
'Le a conta contabil de almoxarifado-produto a partir de codigos de almoxarifado e produto.

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_EstoqueProdutoCC_Le
    
    'le a ContaContabil em estoque produto
    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then Error 49987
    
    'se nao encontrou --- ERRO
    If lErro = 21306 Then Error 49988
    
    'se nao a encontrou a conta
    If Trim(objEstoqueProduto.sContaContabil) = "" Then
    
        objAlmoxarifado.iCodigo = objEstoqueProduto.iAlmoxarifado
        
        'procura em Almoxarifado passando o codigo com paremetro
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 49989
        
        'se nao encontrou o Almoxarifado -----> ERRO
        If lErro = 25056 Then Error 49990
            
        If Trim(objAlmoxarifado.sContaContabil) = "" Then Error 49991
        
        objEstoqueProduto.sContaContabil = objAlmoxarifado.sContaContabil
        
    End If
        
    EstoqueProdutoCC_Le = SUCESSO
        
    Exit Function
    
Erro_EstoqueProdutoCC_Le:
    
    EstoqueProdutoCC_Le = Err
    
    Select Case Err
        
        Case 49987, 49989
        
        Case 49988
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", Err, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 49990
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", Err, objEstoqueProduto.iAlmoxarifado)
        
        Case 49991 'Conta não encontrada. A ser tratado pela rotina chamadora
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152030)
    
    End Select
    
    Exit Function

End Function

Function TabelaPrecoItem_Le(ByVal objTabelaPrecoItem As ClassTabelaPrecoItem, Optional dtDataLeitura As Date = DATA_NULA) As Long
'Lê a tabela de preço a partir da chave fornecida em objTabelaPreco. Retorna os dados em objTabelaPreco

'*** ALTERADO POR LUIZ G.F.NOGUEIRA EM 02/04/2002 ***
'dtDataLeitura indica que a leitura deve ser feita utilizando a data passada na variável. Seu valor padrão é DATA_NULA. Isso significa
'que se o programador não passar essa variável, a leitura será feita utilizando a data atual (Date)
'****************************************************

Dim lErro As Long
Dim lComando As Long
Dim tTabelaPrecoItem As typeTabelaPrecoItem

On Error GoTo Erro_TabelaPrecoItem_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28011

    'Se não foi passada uma data específica para a leitura = > o select será feito com a data de hoje
    If dtDataLeitura = DATA_NULA Then dtDataLeitura = gdtDataHoje
    
    tTabelaPrecoItem.sObservacao = String(STRING_TABELA_OBSERVACAO, 0)
    tTabelaPrecoItem.sTextoGrade = String(STRING_TABELA_TEXTOGRADE, 0)
    
    'Pesquisa no BD Ítens de Tabela de Preço
    lErro = Comando_Executar(lComando, "SELECT Preco, DataVigencia, Observacao, TextoGrade, PercDesconto,PrecoComDesconto FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto = ? AND DataVigencia <= ? ORDER BY DataVigencia DESC", tTabelaPrecoItem.dPreco, tTabelaPrecoItem.dtDataVigencia, tTabelaPrecoItem.sObservacao, tTabelaPrecoItem.sTextoGrade, tTabelaPrecoItem.dPercDesconto, tTabelaPrecoItem.dPrecoComDesconto, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto, dtDataLeitura)
    If lErro <> AD_SQL_SUCESSO Then Error 28012

    'Verifica se Ítem existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28013

    'Não encontrou o Ítem
    If lErro = AD_SQL_SEM_DADOS Then Error 28014

    'Preenche objTabelaPrecoItem
    objTabelaPrecoItem.dPreco = tTabelaPrecoItem.dPreco
    objTabelaPrecoItem.dtDataVigencia = tTabelaPrecoItem.dtDataVigencia
    objTabelaPrecoItem.sObservacao = tTabelaPrecoItem.sObservacao
    objTabelaPrecoItem.sTextoGrade = tTabelaPrecoItem.sTextoGrade
    objTabelaPrecoItem.dPercDesconto = tTabelaPrecoItem.dPercDesconto
    objTabelaPrecoItem.dPrecoComDesconto = tTabelaPrecoItem.dPrecoComDesconto

    'Fechamento comando
    Call Comando_Fechar(lComando)

    TabelaPrecoItem_Le = SUCESSO

    Exit Function

Erro_TabelaPrecoItem_Le:

    TabelaPrecoItem_Le = Err

    Select Case Err

        Case 28011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28012, 28013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS1", Err, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto)

        Case 28014
            'Não encontrou a Ítem da Tabela de Preço. Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152031)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function EstoqueProduto_TestaAssociacao(ByVal sProduto As String, ByVal Almoxarifado As Object) As Long
'Verifica associação entre Produto e almoxarifado na filial corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_EstoqueProduto_TestaAssociacao

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 41234

    lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then Error 41235

    EstoqueProduto_TestaAssociacao = SUCESSO

    Exit Function

Erro_EstoqueProduto_TestaAssociacao:

    EstoqueProduto_TestaAssociacao = Err

    Select Case Err

        Case 41234, 41235

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152032)

    End Select

    Exit Function

End Function

Function EstoqueMesNaoApurado_Le(ByVal objEstoqueMes As ClassEstoqueMes) As Long

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueMes As typeEstoqueMes

On Error GoTo Erro_EstoqueMesNaoApurado_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 25218

    'Lê EstoqueMes não apurado
    With tEstoqueMes
        lErro = Comando_Executar(lComando, "SELECT Ano, Mes, Fechamento, Alterado FROM EstoqueMes WHERE FilialEmpresa = ? AND Fechamento=? AND CustoProdApurado = ? ORDER BY Ano, Mes", .iAno, .iMes, .iFechamento, .iAlterado, giFilialEmpresa, ESTOQUEMES_FECHAMENTO_FECHADO, CUSTO_NAO_APURADO)
    End With
    If lErro <> AD_SQL_SUCESSO Then Error 25219

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25220
    If lErro = AD_SQL_SEM_DADOS Then Error 25221 'Não encontrou

    With tEstoqueMes

        'Coloca os dados lidos no objProduto
        objEstoqueMes.iAlterado = .iAlterado
        objEstoqueMes.iAno = .iAno
        objEstoqueMes.iFechamento = ESTOQUEMES_FECHAMENTO_FECHADO
        objEstoqueMes.iMes = .iMes
        objEstoqueMes.iCustoProdApurado = CUSTO_NAO_APURADO
        objEstoqueMes.iFilialEmpresa = giFilialEmpresa

    End With

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    EstoqueMesNaoApurado_Le = SUCESSO

    Exit Function

Erro_EstoqueMesNaoApurado_Le:

    EstoqueMesNaoApurado_Le = Err

    Select Case Err

        Case 25218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25219, 25220

        Case 25221
            'EstoqueMes não encontrado. Será tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152033)

    End Select

    'Fecha o comando
    lErro = Comando_Fechar(lComando)

    Exit Function

End Function

Function EstoqueMes_Le_Anos(ByVal colAnos As Collection) As Long
'Le todos os anos da tabela EstoqueMes com FilialEmpresa = giFilialEmpresa

Dim lErro As Long
Dim sComando_SQL As String
Dim lComando As Long
Dim iAno As Integer

On Error GoTo Erro_EstoqueMes_Le_Anos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41742

    sComando_SQL = "SELECT DISTINCT Ano FROM EstoqueMes WHERE FilialEmpresa = ? ORDER BY Ano"

    lErro = Comando_Executar(lComando, sComando_SQL, iAno, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 41743

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41744

    If lErro = AD_SQL_SEM_DADOS Then Error 41745

    Do While lErro <> AD_SQL_SEM_DADOS

        colAnos.Add iAno

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41746

    Loop

    Call Comando_Fechar(lComando)

    EstoqueMes_Le_Anos = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_Anos:

    EstoqueMes_Le_Anos = Err

    Select Case Err

        Case 41742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 41743, 41744, 41746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", Err)

        Case 41745
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152034)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function EstoqueMes_Le_Meses(ByVal iAno As Integer, ByVal colMeses As Collection) As Long
'Le todos os meses da tabela EstoqueMes com Ano = iAno e FilialEmpresa = giFilialEmpresa

Dim lErro As Long
Dim sComando_SQL As String
Dim lComando As Long
Dim iMes As Integer

On Error GoTo Erro_EstoqueMes_Le_Meses

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41747

    sComando_SQL = "SELECT Mes FROM EstoqueMes WHERE Ano = ? AND FilialEmpresa = ? ORDER BY Mes"

    lErro = Comando_Executar(lComando, sComando_SQL, iMes, iAno, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 41748

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41749

    If lErro = AD_SQL_SEM_DADOS Then Error 41750

    Do While lErro <> AD_SQL_SEM_DADOS

        colMeses.Add iMes

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41751

    Loop

    Call Comando_Fechar(lComando)

    EstoqueMes_Le_Meses = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_Meses:

    EstoqueMes_Le_Meses = Err

    Select Case Err

        Case 41747
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41748, 41749, 41751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", Err)
        
        Case 41750
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152035)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoFilial_Le_Lock(ByVal lComando As Long, ByVal objProdutoFilial As ClassProdutoFilial) As Long

Dim lErro As Long
Dim tProdutoFilial As typeProdutoFilial

On Error GoTo Erro_ProdutoFilial_Le_Lock

    tProdutoFilial.sClasseABC = String(STRING_PRODUTOFILIAL_CLASSEABC, 0)
    tProdutoFilial.sSituacaoTribECF = String(STRING_PRODUTOFILIAL_SITUACAOTRIBECF, 0)
    tProdutoFilial.sICMSAliquota = String(STRING_PRODUTOFILIAL_ICMS, 0)
    
    'Pesquisa no BD ProdutoFilial
    With tProdutoFilial
        lErro = Comando_ExecutarLockado(lComando, "SELECT Almoxarifado, Fornecedor, VisibilidadeAlmoxarifados, EstoqueSeguranca, EstoqueMaximo, PontoPedido, ClasseABC, LoteEconomico, IntRessup,SituacaoTribECF, ICMSAliquota FROM ProdutosFilial WHERE FilialEmpresa= ? AND Produto = ?", _
            .iAlmoxarifado, .lFornecedor, .iVisibilidadeAlmoxarifados, .dEstoqueSeguranca, .dEstoqueMaximo, .dPontoPedido, .sClasseABC, .dLoteEconomico, .iIntRessup, .sSituacaoTribECF, .sICMSAliquota, objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)
        If lErro <> AD_SQL_SUCESSO Then Error 41768
    End With

    'Tenta selecionar ProdutoFilial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 41769

    If lErro = AD_SQL_SEM_DADOS Then Error 41770

    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 41771

    'Preenche objProdutoFilial
    With tProdutoFilial
        objProdutoFilial.iAlmoxarifado = .iAlmoxarifado
        objProdutoFilial.lFornecedor = .lFornecedor
        objProdutoFilial.iVisibilidadeAlmoxarifados = .iVisibilidadeAlmoxarifados
        objProdutoFilial.dEstoqueSeguranca = .dEstoqueSeguranca
        objProdutoFilial.dEstoqueMaximo = .dEstoqueMaximo
        objProdutoFilial.dPontoPedido = .dPontoPedido
        objProdutoFilial.sClasseABC = .sClasseABC
        objProdutoFilial.dLoteEconomico = .dLoteEconomico
        objProdutoFilial.iIntRessup = .iIntRessup
        objProdutoFilial.sSituacaoTribECF = .sSituacaoTribECF
        objProdutoFilial.sICMSAliquota = .sICMSAliquota
        
               
    End With

    ProdutoFilial_Le_Lock = SUCESSO

    Exit Function

Erro_ProdutoFilial_Le_Lock:

    ProdutoFilial_Le_Lock = Err

    Select Case Err

        Case 41768, 41769
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err, objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)

        Case 41770
        
        Case 41771
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOSFILIAL", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152036)

    End Select

    Exit Function

End Function

Function EstoqueMes_Le_Lock(ByVal lComando As Long, ByVal objEstoqueMes As ClassEstoqueMes) As Long
'Le a tabela EstoqueMes e loca o registro

Dim lErro As Long
Dim sComando_SQL As String
Dim tEstoqueMes As typeEstoqueMes

On Error GoTo Erro_EstoqueMes_Le_Lock

    sComando_SQL = "SELECT Fechamento, Alterado, CustoProdApurado, HorasMaquina, CustoMatPrim, GastosDiretos, GastosIndiretos, QuantFator1, CustoFator1, QuantFator2, CustoFator2, QuantFator3, CustoFator3, QuantFator4, CustoFator4, QuantFator5, CustoFator5, QuantFator6, CustoFator6 FROM EstoqueMes WHERE FilialEmpresa = ? AND Ano = ? AND Mes = ?"

    lErro = Comando_ExecutarLockado(lComando, sComando_SQL, tEstoqueMes.iFechamento, tEstoqueMes.iAlterado, tEstoqueMes.iCustoProdApurado, tEstoqueMes.lHorasMaquina, tEstoqueMes.dCustoMatPrim, tEstoqueMes.dGastosDiretos, tEstoqueMes.dGastosIndiretos, _
        tEstoqueMes.dQuantFator1, tEstoqueMes.dCustoFator1, tEstoqueMes.dQuantFator2, tEstoqueMes.dCustoFator2, tEstoqueMes.dQuantFator3, tEstoqueMes.dCustoFator3, tEstoqueMes.dQuantFator4, tEstoqueMes.dCustoFator4, tEstoqueMes.dQuantFator5, tEstoqueMes.dCustoFator5, tEstoqueMes.dQuantFator6, tEstoqueMes.dCustoFator6, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
    If lErro <> AD_SQL_SUCESSO Then Error 41772

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41773

    If lErro = AD_SQL_SEM_DADOS Then Error 41774

    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 41775

    objEstoqueMes.iFechamento = tEstoqueMes.iFechamento
    objEstoqueMes.iAlterado = tEstoqueMes.iAlterado
    objEstoqueMes.iCustoProdApurado = tEstoqueMes.iCustoProdApurado
    objEstoqueMes.lHorasMaquina = tEstoqueMes.lHorasMaquina
    objEstoqueMes.dCustoMatPrim = tEstoqueMes.dCustoMatPrim
    objEstoqueMes.dGastosDiretos = tEstoqueMes.dGastosDiretos
    objEstoqueMes.dGastosIndiretos = tEstoqueMes.dGastosIndiretos
    objEstoqueMes.dQuantFator1 = tEstoqueMes.dQuantFator1
    objEstoqueMes.dCustoFator1 = tEstoqueMes.dCustoFator1
    objEstoqueMes.dQuantFator2 = tEstoqueMes.dQuantFator2
    objEstoqueMes.dCustoFator2 = tEstoqueMes.dCustoFator2
    objEstoqueMes.dQuantFator3 = tEstoqueMes.dQuantFator3
    objEstoqueMes.dCustoFator3 = tEstoqueMes.dCustoFator3
    objEstoqueMes.dQuantFator4 = tEstoqueMes.dQuantFator4
    objEstoqueMes.dCustoFator4 = tEstoqueMes.dCustoFator4
    objEstoqueMes.dQuantFator5 = tEstoqueMes.dQuantFator5
    objEstoqueMes.dCustoFator5 = tEstoqueMes.dCustoFator5
    objEstoqueMes.dQuantFator6 = tEstoqueMes.dQuantFator6
    objEstoqueMes.dCustoFator6 = tEstoqueMes.dCustoFator6

    EstoqueMes_Le_Lock = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_Lock:

    EstoqueMes_Le_Lock = Err

    Select Case Err

        Case 41772, 41773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES", Err, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
            
        Case 41774
        
        Case 41775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ESTOQUEMES", Err, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152037)

    End Select

    Exit Function

End Function

Function MovEstoque_Le_ProdutoMes(ByVal iMes As Integer, ByVal iAno As Integer, ByVal objItemMovEstoque As ClassItemMovEstoque) As Long
'Le o primeiro Movimento de Estoque do Mes-Ano em questão
'com parametros e o sProduto passado dentro do objMovEstoque

Dim lErro As Long
Dim lComando As Long
Dim tItemMovEstoque As typeItemMovEstoque
Dim dtDataFinal As Date
Dim dtDataInicial As Date

On Error GoTo Erro_MovEstoque_Le_ProdutoMes

    'Abertura de comando
    lComando = Comando_Abrir
    If lComando = 0 Then Error 49968
    
    'Preenche as datas inicial e final
    dtDataInicial = CDate(1 & "/" & iMes & "/" & iAno)
    dtDataFinal = CDate(Dias_Mes(iMes, iAno) & "/" & iMes & "/" & iAno)
    
    lErro = Comando_Executar(lComando, "SELECT MovimentoEstoque.Codigo, Data,TipoMov, Custo FROM MovimentoEstoque, TiposMovimentoEstoque WHERE MovimentoEstoque.Produto = ? AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.Data >=? AND MovimentoEstoque.Data <=? AND MovimentoEstoque.TipoMov = TiposMovimentoEstoque.Codigo AND (TiposMovimentoEstoque.EntradaOuSaida = 'E' Or TiposMovimentoEstoque.EntradaOuSaida = 'S')", tItemMovEstoque.lCodigo, tItemMovEstoque.dtData, tItemMovEstoque.iTipoMov, tItemMovEstoque.dCusto, objItemMovEstoque.sProduto, giFilialEmpresa, dtDataInicial, dtDataFinal)
    If lErro <> AD_SQL_SUCESSO Then Error 49969

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 49970

    If lErro = AD_SQL_SEM_DADOS Then Error 49971
    
    objItemMovEstoque.lCodigo = tItemMovEstoque.lCodigo
    objItemMovEstoque.iFilialEmpresa = giFilialEmpresa
    objItemMovEstoque.iTipoMov = tItemMovEstoque.iTipoMov
    objItemMovEstoque.dtData = tItemMovEstoque.dtData
    objItemMovEstoque.dCusto = tItemMovEstoque.dCusto
    
    Call Comando_Fechar(lComando)

    MovEstoque_Le_ProdutoMes = SUCESSO

    Exit Function

Erro_MovEstoque_Le_ProdutoMes:

    MovEstoque_Le_ProdutoMes = Err

    Select Case Err

        Case 49968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 49969, 49970
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)

        Case 49971 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 152038)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Almoxarifados_Filial(ByVal sProduto As String, ByVal colAlmoxarifados As Collection) As Long
'Lê todos os Almoxarifados do Produto para a filial corrente ou EMPRESA_TODA

Dim lErro As Long
Dim lComando As Long, tAlmoxarifado As typeAlmoxarifado
Dim objAlmoxarifado As ClassAlmoxarifado

On Error GoTo Erro_Produto_Le_Almoxarifados_Filial

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 19312

    With tAlmoxarifado
        .sNomeReduzido = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sDescricao = String(STRING_ALMOXARIFADO_DESCRICAO, 0)
        'Lê a tabela Estoque Produto
        If giFilialEmpresa = EMPRESA_TODA Then
            lErro = Comando_Executar(lComando, "SELECT EstoqueProduto.Almoxarifado, Almoxarifado.FilialEmpresa, Almoxarifado.NomeReduzido, Almoxarifado.Descricao, Almoxarifado.Endereco FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo", .iCodigo, .iFilialEmpresa, .sNomeReduzido, .sDescricao, .lEndereco, sProduto)
        Else
            .iFilialEmpresa = giFilialEmpresa
            lErro = Comando_Executar(lComando, "SELECT EstoqueProduto.Almoxarifado, Almoxarifado.NomeReduzido, Almoxarifado.Descricao, Almoxarifado.Endereco FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.FilialEmpresa = ?", .iCodigo, .sNomeReduzido, .sDescricao, .lEndereco, sProduto, giFilialEmpresa)
        End If
    End With
    
    If lErro <> AD_SQL_SUCESSO Then Error 19313

    'Lê o primeiro Almoxarifado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19314

    'Enquanto a tabela tiver dados
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objAlmoxarifado = New ClassAlmoxarifado

        With objAlmoxarifado
            .iCodigo = tAlmoxarifado.iCodigo
            .iFilialEmpresa = tAlmoxarifado.iFilialEmpresa
            .lEndereco = tAlmoxarifado.lEndereco
            .sDescricao = tAlmoxarifado.sDescricao
            .sNomeReduzido = tAlmoxarifado.sNomeReduzido
        End With
        
        'Coloca o Almoxarifado lido na coleção
        colAlmoxarifados.Add objAlmoxarifado

        'Lê o próximo Almoxarifado
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19315

    Loop

    Call Comando_Fechar(lComando)

    Produto_Le_Almoxarifados_Filial = SUCESSO

    Exit Function

Erro_Produto_Le_Almoxarifados_Filial:

    Produto_Le_Almoxarifados_Filial = Err

    Select Case Err

        Case 19312
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 19313, 19314, 19315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXPRODFILIAL", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152039)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_ExisteProduto(iExiste As Integer) As Long
'verifica se tem algum produto cadastrado.
'se existir iExiste = POSSUI_PRODUTO , caso contrario iExiste = NAO_POSSUI_PRODUTO

Dim lComando As Long
Dim lErro As Long
Dim sDescricao As String

On Error GoTo Erro_Produto_ExisteProduto

    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44840

    'Verifica se o Produto existe
    lErro = Comando_Executar(lComando, "SELECT Descricao FROM Produtos", sDescricao)
    If lErro <> AD_SQL_SUCESSO Then Error 44841

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44842
        
    'Ccl está cadastrado
    If lErro = SUCESSO Then
        iExiste = POSSUI_PRODUTO
    Else
        iExiste = NAO_POSSUI_PRODUTO
    End If

    Call Comando_Fechar(lComando)

    Produto_ExisteProduto = SUCESSO

    Exit Function

Erro_Produto_ExisteProduto:

    Produto_ExisteProduto = Err
    
    Select Case Err

        Case 44840
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44841, 44842
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152040)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Empenho_Le_SemCodigo(ByVal objEmpenho As ClassEmpenho) As Long
'Lê os dados de Empenho à partir da chave candidata fornecida em objEmpenho, com os seguintes campos preenchidos:
'objEmpenho.iFilialEmpresa, objEmpenho.sCodigoOP, objEmpenho.iItemOP, objEmpenho.sProduto, objEmpenho.iAlmoxarifado

Dim lErro As Long
Dim lComando As Long
Dim tEmpenho As typeEmpenho

On Error GoTo Erro_Empenho_Le_SemCodigo

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 41886

    'Pesquisa no BD o Empenho passado como parâmetro
    lErro = Comando_Executar(lComando, "SELECT Empenho.Codigo, Empenho.Quantidade, Empenho.QuantidadeRequisitada, Empenho.Data, Empenho.NumIntDocItemOP FROM Empenho, ItensOrdemProducao WHERE Empenho.NumIntDocItemOP = ItensOrdemProducao.NumIntDoc AND ItensOrdemProducao.FilialEmpresa = ? AND ItensOrdemProducao.Codigo = ? AND ItensOrdemProducao.Item = ? AND Empenho.Produto = ? AND Empenho.Almoxarifado = ?", _
        tEmpenho.lCodigo, tEmpenho.dQuantidade, tEmpenho.dQuantidadeRequisitada, tEmpenho.dtData, tEmpenho.lNumIntDocItemOP, objEmpenho.iFilialEmpresa, objEmpenho.sCodigoOP, objEmpenho.iItemOP, objEmpenho.sProduto, objEmpenho.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then Error 41887

    'Verifica se o Empenho existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41888

    'Não encontrou o Empenho ==> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 41889

    'Preenche objEmpenho
    objEmpenho.lCodigo = tEmpenho.lCodigo
    objEmpenho.dQuantidade = tEmpenho.dQuantidade
    objEmpenho.dQuantidadeRequisitada = tEmpenho.dQuantidadeRequisitada
    objEmpenho.dtData = tEmpenho.dtData
    objEmpenho.lNumIntDocItemOP = tEmpenho.lNumIntDocItemOP
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    Empenho_Le_SemCodigo = SUCESSO

    Exit Function

Erro_Empenho_Le_SemCodigo:

    Empenho_Le_SemCodigo = Err

    Select Case Err

        Case 41886
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41887, 41888
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPENHO2", Err)

        Case 41889
            'Não encontrou o Empenho.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152041)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Kit_Le_Empenho(ByVal objKit As ClassKit) As Long
'Lê uma versão mais recente na tabela de Kits de Produtos à partir do produto passado em objKit
'e devolve os dados do Kit em objKit

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sVersao As String
Dim sObservacao As String
Dim dtData As Date

On Error GoTo Erro_Kit_Le_Empenho

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43811

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 43819
    
    sVersao = String(STRING_KIT_VERSAO, 0)
    sObservacao = String(STRING_KIT_OBSERVACAO, 0)

    lErro = Comando_Executar(lComando, "SELECT Versao, Observacao FROM Kit AS Kit1 WHERE Data = (SELECT Max(Data) FROM Kit As Kit2 WHERE  ProdutoRaiz = ? AND Data <= ? AND Kit1.ProdutoRaiz = Kit2.ProdutoRaiz)", sVersao, sObservacao, objKit.sProdutoRaiz, objKit.dtData)
    If lErro <> AD_SQL_SUCESSO Then Error 43812

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43813

    'Se não encontrou o Kit --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 43814

'    lErro = Comando_Executar(lComando1, "SELECT Versao, Observacao FROM Kit WHERE ProdutoRaiz = ? AND Data = ?", sVersao, sObservacao, objKit.sProdutoRaiz, dtData)
'    If lErro <> AD_SQL_SUCESSO Then Error 43820
'
'    lErro = Comando_BuscarPrimeiro(lComando1)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43821

    objKit.sVersao = sVersao
    objKit.sObservacao = sObservacao

    'Libera comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Kit_Le_Empenho = SUCESSO

    Exit Function

Erro_Kit_Le_Empenho:

    Kit_Le_Empenho = Err

    Select Case Err

        Case 43811, 43819
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 43812, 43813, 43820, 43821
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_KIT", Err, objKit.sProdutoRaiz)

        Case 43814
            'Não encontrou o Kit

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152042)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function AlmoxarifadoPadrao_Le_NomeReduzido(ByVal sProduto As String, sNomeReduzidoAlmox As String) As Long
'Retorna o Nome reduzido do Almoxarifado Padrão do Produto em Questão com a Filial Empresa

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iAlmoxarifadoPadrao As Integer

On Error GoTo Erro_AlmoxarifadoPadrao_Le_NomeReduzido

    'se o almoxarifado padrão não foi sugerido na tela - - > Busca o padrão do produto
    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, sProduto, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 23796 Then Error 52218
    
    'se não encontrou o Registro então  -- -  - - > Erro
    If lErro = 23796 Then Error 52219
    
    'se o Almoxarifado Padrão estiver Preenchido le seu nome reduzido
    If iAlmoxarifadoPadrao <> 0 Then
        
        objAlmoxarifado.iCodigo = iAlmoxarifadoPadrao
        
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 52220
        
        'se não encontrou o registro  - - - -> ERRO
        If lErro = 25056 Then Error 52221
            
        sNomeReduzidoAlmox = objAlmoxarifado.sNomeReduzido
    
    End If
            
    AlmoxarifadoPadrao_Le_NomeReduzido = SUCESSO
        
    Exit Function
        
Erro_AlmoxarifadoPadrao_Le_NomeReduzido:
    
    AlmoxarifadoPadrao_Le_NomeReduzido = Err
    
    Select Case Err
         
        Case 52218, 52220
            
        Case 52219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", Err, sProduto)
            
        Case 52221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", Err, iAlmoxarifadoPadrao)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 152043)
        
    End Select
    
    Exit Function
        
End Function

Function EstoqueMes_Le_Aberto(ByVal objEstoqueMes As ClassEstoqueMes) As Long
'Le o mes em aberto do Estoque para a filial em questão (giFilialEmpresa).

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueMes As typeEstoqueMes

On Error GoTo Erro_EstoqueMes_Le_Aberto
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55025
    
    'Le Ano e Mes
    lErro = Comando_Executar(lComando, "SELECT Ano, Mes FROM EstoqueMes WHERE FilialEmpresa = ? AND Fechamento = ?", tEstoqueMes.iAno, tEstoqueMes.iMes, objEstoqueMes.iFilialEmpresa, ESTOQUEMES_FECHAMENTO_ABERTO)
    If lErro <> AD_SQL_SUCESSO Then Error 55026

    'Le o mes em aberto do Estoque para a filial em questão (giFilialEmpresa).
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55027

    If lErro = AD_SQL_SEM_DADOS Then Error 55028

    'preenche o objeto com o mes e o ano
    objEstoqueMes.iMes = tEstoqueMes.iMes
    objEstoqueMes.iAno = tEstoqueMes.iAno
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    EstoqueMes_Le_Aberto = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_Aberto:

    EstoqueMes_Le_Aberto = Err

    Select Case Err

        Case 55025
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55026, 55027
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", Err)

        Case 55028
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_ABERTO_INEXISTENTE", Err, giFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152044)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function EstoqueMes_Le_PrimeiroAnoMes(ByVal objEstoqueMes As ClassEstoqueMes) As Long
'Le o primeiro ano/mes do Estoque para a filial em questão (giFilialEmpresa).

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueMes As typeEstoqueMes

On Error GoTo Erro_EstoqueMes_Le_PrimeiroAnoMes
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55032
    
    'Le EstoqueMes
    lErro = Comando_Executar(lComando, "SELECT Ano, Mes FROM EstoqueMes WHERE FilialEmpresa = ? ORDER BY Ano, Mes", tEstoqueMes.iAno, tEstoqueMes.iMes, giFilialEmpresa, ESTOQUEMES_FECHAMENTO_ABERTO)
    If lErro <> AD_SQL_SUCESSO Then Error 55033

    'Busca o primeiro ano e mes para a filial em questão
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55034

    If lErro = AD_SQL_SEM_DADOS Then Error 55035

    'preenche o objeto com o mes e o ano
    objEstoqueMes.iMes = tEstoqueMes.iMes
    objEstoqueMes.iAno = tEstoqueMes.iAno
    objEstoqueMes.iFilialEmpresa = giFilialEmpresa

    'Fechamento comando
    Call Comando_Fechar(lComando)

    EstoqueMes_Le_PrimeiroAnoMes = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_PrimeiroAnoMes:

    EstoqueMes_Le_PrimeiroAnoMes = Err

    Select Case Err

        Case 55032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55033, 55034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", Err)

        Case 55035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE1", Err, giFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152045)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CustoStandard_Le(ByVal sProduto As String, dCustoStandard As Double, ByVal iFilialEmpresa As Integer) As Long
'Parametros a serem passados: sProduto
'Le o CustoStandard do Ano-Mes aberto e retorna em dCustoStandard

Dim objEstoqueMes As New ClassEstoqueMes
Dim lErro As Long

On Error GoTo Erro_CustoStandard_Le

    objEstoqueMes.iFilialEmpresa = iFilialEmpresa

    'descobre o mes em aberto do estoque
    lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
    If lErro <> SUCESSO Then gError 55038
    
    'Le o CustoStandard do Ano/Mes passado como parametro e retorna em dCustoStandard
    lErro = Calcula_CustoStandard(iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes, sProduto, dCustoStandard)
    If lErro <> SUCESSO And lErro <> 55041 Then gError 55045
    
    If lErro = 55041 Then gError 69852
    
    CustoStandard_Le = SUCESSO
    
    Exit Function
    
Erro_CustoStandard_Le:

    CustoStandard_Le = gErr
    
    Select Case gErr
        
        Case 55038, 55045
        
        Case 69852 'Produto não cadastrado em SdlMesEst

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152046)
        
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_CustoMedioAtual(ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal iMes As Integer, ByVal sProduto As String, dCustoMedioAtual As Double) As Long
'Le o CustoMedio do Ano/Mes passado como parametro e retorna em dCustoMedioAtual

Dim objSldMesEst As New ClassSldMesEst
Dim lErro As Long

On Error GoTo Erro_Calcula_CustoMedioAtual

    'Preenche o objSldMesEst para a Leitura do custoMedio
    objSldMesEst.iAno = iAno
    objSldMesEst.iFilialEmpresa = iFilialEmpresa
    objSldMesEst.sProduto = sProduto
    
    'Le o CustoMedio do Ano/Mes passados como parametro e retorna em objSldMesEst
    lErro = CF("SldMesEst_Le_Custos", iMes, objSldMesEst)
    If lErro <> SUCESSO And lErro <> 41755 Then Error 49604
    
    'se não estiver encontrado registro com o custo medio
    If lErro = 41755 Then Error 55040
    
    'passa o custoMedio para dCustoMedioAtual
    dCustoMedioAtual = objSldMesEst.dCustoMedio(iMes)
    
    Calcula_CustoMedioAtual = SUCESSO
    
    Exit Function
    
Erro_Calcula_CustoMedioAtual:

    Calcula_CustoMedioAtual = Err
    
    Select Case Err
        
        Case 49604
        
        Case 55040 'Não encontrou
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152047)
        
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_CustoStandard(ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal iMes As Integer, ByVal sProduto As String, dCustoStandard As Double) As Long
'Le o CustoStandard do Ano/Mes passado como parametro e retorna em dCustoStandard

Dim objSldMesEst As New ClassSldMesEst
Dim lErro As Long

On Error GoTo Erro_Calcula_CustoStandard

    'Preenche o objSldMesEst para a Leitura do custoStandard
    objSldMesEst.iAno = iAno
    objSldMesEst.iFilialEmpresa = iFilialEmpresa
    objSldMesEst.sProduto = sProduto
    
    'Le o CustoStandard do Ano-Mes passados como parametros e retorna em objSldMesEst
    lErro = CF("SldMesEst_Le_Custos", iMes, objSldMesEst)
    If lErro <> SUCESSO And lErro <> 41755 Then Error 55039
    
    'se não estiver encontrado registro com o custo standard
    If lErro = 41755 Then Error 55041
    
    dCustoStandard = objSldMesEst.dCustoStandard(iMes)
    
    Calcula_CustoStandard = SUCESSO
    
    Exit Function
    
Erro_Calcula_CustoStandard:

    Calcula_CustoStandard = Err
    
    Select Case Err
        
        Case 55039
        
        Case 55041 'Produto não cadastrado em SdlMesEst
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152048)
        
    End Select
    
    Exit Function
    
End Function

Public Function Calcula_Custo(dCusto As Double, ByVal iFilialEmpresa As Integer, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstAberto As Integer, ByVal iMesEstAberto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal sProduto As String) As Long
'devolve o custo do produto

Dim lErro As Long, lComando As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Calcula_Custo
    
    'O valor da diferenca vai depender se é um produto Produzido ou Compravel
    If iApropriacaoCusto = APROPR_CUSTO_MEDIO Then
        
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 124154
        
        lErro = CF("SldMesEst_Le_CustoMedio2", lComando, iFilialEmpresa, iAnoEstAberto, sProduto, iMesEstAberto, dCusto)
        If lErro <> SUCESSO Then gError 69841
        
        Call Comando_Fechar(lComando)
                
    ElseIf iApropriacaoCusto = APROPR_CUSTO_REAL Then
            
        'Calcula o Custo Médio de Produção do Produto sProduto
        lErro = CF("Calcula_CustoMedioProducao", iFilialEmpresa, sProduto, iAnoEstApurado, iMesEstApurado, dCusto)
        If lErro <> SUCESSO And lErro <> 25433 And lErro <> 55052 Then gError 55054
            
        'Se não encontrou o SldMesEst
        If lErro = 25433 Then gError 69865
            
        'Se não encontrou o SldMesEst
        If lErro = 55052 Then gError 69866
            
    ElseIf iApropriacaoCusto = APROPR_CUSTO_STANDARD Then
    
        'Le o CustoStandard do Ano/Mes passado como parametro e retorna em dCustoStandard
        lErro = Calcula_CustoStandard(iFilialEmpresa, iAnoEstAberto, iMesEstAberto, sProduto, dCusto)
        If lErro <> SUCESSO And lErro <> 55041 Then gError 55055
    
        If lErro = 55041 Then gError 69850
        
    End If
    
    Calcula_Custo = SUCESSO
    
    Exit Function
    
Erro_Calcula_Custo:
            
    Calcula_Custo = gErr
    
    Select Case gErr
    
        Case 55053, 55054, 55055
    
        Case 69841, 69850
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTOS_INEXISTENTES", gErr, iFilialEmpresa, iAnoEstAberto, sProduto)
        
        Case 69865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO", gErr, iFilialEmpresa, iAnoEstAberto, sProduto)
        
        Case 69866
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO1", gErr, iFilialEmpresa, sProduto)
        
        Case 124154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152049)

    End Select

    If lComando <> 0 Then Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Produto_Lock_Shared(ByVal lComando As Long, ByVal objProduto As ClassProduto) As Long
'Lê o faz o "lock compartilhado" no Produto passado por parâmetro

Dim lErro As Long
Dim tProduto As typeProduto

On Error GoTo Erro_Produto_Lock_Shared

    tProduto.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    tProduto.sSiglaUMVenda = String(STRING_UM_SIGLA, 0)
    
    'Lê o Produto
    lErro = Comando_ExecutarLockado(lComando, "SELECT ControleEstoque, SiglaUMEstoque, SiglaUMVenda, Apropriacao, ClasseUM, Rastro, Ativo FROM Produtos WHERE Codigo = ? ", tProduto.iControleEstoque, tProduto.sSiglaUMEstoque, tProduto.sSiglaUMVenda, tProduto.iApropriacaoCusto, tProduto.iClasseUM, tProduto.iRastro, tProduto.iAtivo, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 55233

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55234

    'não encontru --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 55235

    'Faz o "lock"
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then Error 55236

    objProduto.iControleEstoque = tProduto.iControleEstoque
    objProduto.sSiglaUMEstoque = tProduto.sSiglaUMEstoque
    objProduto.sSiglaUMVenda = tProduto.sSiglaUMVenda
    objProduto.iApropriacaoCusto = tProduto.iApropriacaoCusto
    objProduto.iClasseUM = tProduto.iClasseUM
    objProduto.iRastro = tProduto.iRastro
    objProduto.iAtivo = tProduto.iAtivo
    
    Produto_Lock_Shared = SUCESSO

    Exit Function

Erro_Produto_Lock_Shared:

    Produto_Lock_Shared = Err

    Select Case Err

        Case 55233, 55234
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS2", Err, objProduto.sNomeReduzido)

        Case 55235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 55236
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOS", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152050)

    End Select

    Exit Function

End Function

Function Kit_Le_Produto(ByVal objKit As ClassKit, ByVal colKits As Collection) As Long
'Função que le todos os kits, produtos iguais ao codigo passado e diferente da versao passada
'retornado na coleção
Dim lErro As Long
Dim lComando As Long
Dim tKit As typeKit
Dim objKit1 As ClassKit

On Error GoTo Erro_Kit_Le_Produto
    
    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 49946

    tKit.sObservacao = String(STRING_KIT_OBSERVACAO, 0)
    tKit.sVersao = String(STRING_KIT_VERSAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Versao, Data, Observacao, Situacao, VersaoFormPreco FROM Kit WHERE ProdutoRaiz = ? AND Versao <> ?", tKit.sVersao, tKit.dtData, tKit.sObservacao, tKit.iSituacao, tKit.iVersaoFormPreco, objKit.sProdutoRaiz, objKit.sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 49947

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 49948

    If lErro = AD_SQL_SEM_DADOS Then gError 49949
    
    'Passa para a collection
    Do While lErro = AD_SQL_SUCESSO

        Set objKit1 = New ClassKit
        
        objKit1.sProdutoRaiz = objKit.sProdutoRaiz
        objKit1.sVersao = tKit.sVersao
        objKit1.dtData = tKit.dtData
        objKit1.sObservacao = tKit.sObservacao
        objKit1.iSituacao = tKit.iSituacao
        objKit1.iVersaoFormPreco = tKit.iVersaoFormPreco
        
        colKits.Add objKit1

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 49950

    Loop
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    Kit_Le_Produto = SUCESSO

    Exit Function

Erro_Kit_Le_Produto:

    Kit_Le_Produto = gErr

    Select Case gErr

        Case 49946
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 49947, 49948, 49950
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_KIT", gErr)

        Case 49949

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152051)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function EstoqueProduto_Le_Todos_Almoxarifados(ByVal objEstoqueProduto As ClassEstoqueProduto) As Long
'Lê a soma das Quantidades de todos Almoxarifados que pertence ao Produto passado em objEstoqueProduto

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueProduto As typeEstoqueProduto
Dim dOp As Double

On Error GoTo Erro_EstoqueProduto_Le_Todos_Almoxarifados

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 52976
    
    lErro = Comando_Executar(lComando, "SELECT SUM(QuantDispNossa), SUM(QuantReservada), MAX(DataInventario), SUM(QuantidadeInicial), SUM(SaldoInicial), SUM(QuantEmpenhada), SUM(QuantPedida), SUM(QuantRecIndl), SUM(QuantIndOutras), SUM(QuantDefeituosa), SUM(QuantConsig3), SUM(QuantConsig), SUM(QuantDemo3), SUM(QuantDemo), SUM(QuantConserto3), SUM(QuantConserto), SUM(QuantOutras3), SUM(QuantOutras), SUM(QuantBenef3), SUM(QuantBenef), SUM(QuantOP), SUM(ValorConsig3), SUM(ValorConsig), SUM(ValorDemo3), SUM(ValorDemo), SUM(ValorConserto3), SUM(ValorConserto), SUM(ValorOutras3), SUM(ValorOutras), SUM(ValorBenef3), SUM(ValorBenef) FROM EstoqueProduto WHERE Produto = ? ", _
        tEstoqueProduto.dQuantDispNossa, tEstoqueProduto.dQuantReservada, tEstoqueProduto.dtDataInventario, tEstoqueProduto.dQuantidadeInicial, tEstoqueProduto.dSaldoInicial, tEstoqueProduto.dQuantEmpenhada, tEstoqueProduto.dQuantPedido, tEstoqueProduto.dQuantRecIndl, tEstoqueProduto.dQuantInd, tEstoqueProduto.dQuantDefeituosa, tEstoqueProduto.dQuantConsig3, tEstoqueProduto.dQuantConsig, tEstoqueProduto.dQuantDemo3, tEstoqueProduto.dQuantDemo, tEstoqueProduto.dQuantConserto3, tEstoqueProduto.dQuantConserto, tEstoqueProduto.dQuantOutras3, tEstoqueProduto.dQuantOutras, tEstoqueProduto.dQuantBenef3, tEstoqueProduto.dQuantBenef, tEstoqueProduto.dQuantOP, _
        tEstoqueProduto.dValorConsig3, tEstoqueProduto.dValorConsig, tEstoqueProduto.dValorDemo3, tEstoqueProduto.dValorDemo, tEstoqueProduto.dValorConserto3, tEstoqueProduto.dValorConserto, tEstoqueProduto.dValorOutras3, tEstoqueProduto.dValorOutras, tEstoqueProduto.dValorBenef3, tEstoqueProduto.dValorBenef, objEstoqueProduto.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 52977

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52978

    'Não encontrou ---> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 52979
    
    objEstoqueProduto.dQuantDispNossa = tEstoqueProduto.dQuantDispNossa
    objEstoqueProduto.dQuantReservada = tEstoqueProduto.dQuantReservada
    objEstoqueProduto.dtDataInventario = tEstoqueProduto.dtDataInventario
    objEstoqueProduto.dQuantidadeInicial = tEstoqueProduto.dQuantidadeInicial
    objEstoqueProduto.dSaldoInicial = tEstoqueProduto.dSaldoInicial
    objEstoqueProduto.dQuantEmpenhada = tEstoqueProduto.dQuantEmpenhada
    objEstoqueProduto.dQuantPedido = tEstoqueProduto.dQuantPedido
    objEstoqueProduto.dQuantRecIndl = tEstoqueProduto.dQuantRecIndl
    objEstoqueProduto.dQuantInd = tEstoqueProduto.dQuantInd
    objEstoqueProduto.dQuantDefeituosa = tEstoqueProduto.dQuantDefeituosa
    objEstoqueProduto.dQuantConsig3 = tEstoqueProduto.dQuantConsig3
    objEstoqueProduto.dQuantConsig = tEstoqueProduto.dQuantConsig
    objEstoqueProduto.dQuantDemo3 = tEstoqueProduto.dQuantDemo3
    objEstoqueProduto.dQuantDemo = tEstoqueProduto.dQuantDemo
    objEstoqueProduto.dQuantConserto3 = tEstoqueProduto.dQuantConserto3
    objEstoqueProduto.dQuantConserto = tEstoqueProduto.dQuantConserto
    objEstoqueProduto.dQuantOutras3 = tEstoqueProduto.dQuantOutras3
    objEstoqueProduto.dQuantOutras = tEstoqueProduto.dQuantOutras
    objEstoqueProduto.dQuantBenef = tEstoqueProduto.dQuantBenef
    objEstoqueProduto.dQuantBenef3 = tEstoqueProduto.dQuantBenef3
    objEstoqueProduto.dQuantOP = tEstoqueProduto.dQuantOP
    objEstoqueProduto.dValorConsig3 = tEstoqueProduto.dValorConsig3
    objEstoqueProduto.dValorConsig = tEstoqueProduto.dValorConsig
    objEstoqueProduto.dValorDemo3 = tEstoqueProduto.dValorDemo3
    objEstoqueProduto.dValorDemo = tEstoqueProduto.dValorDemo
    objEstoqueProduto.dValorConserto3 = tEstoqueProduto.dValorConserto3
    objEstoqueProduto.dValorConserto = tEstoqueProduto.dValorConserto
    objEstoqueProduto.dValorOutras3 = tEstoqueProduto.dValorOutras3
    objEstoqueProduto.dValorOutras = tEstoqueProduto.dValorOutras
    objEstoqueProduto.dValorBenef3 = tEstoqueProduto.dValorBenef3
    objEstoqueProduto.dValorBenef = tEstoqueProduto.dValorBenef
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    EstoqueProduto_Le_Todos_Almoxarifados = SUCESSO

    Exit Function

Erro_EstoqueProduto_Le_Todos_Almoxarifados:

    EstoqueProduto_Le_Todos_Almoxarifados = Err

    Select Case Err
        
        Case 52976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 52977, 52978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO1", Err, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 52979 'Não encontrou --> Tratar na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152052)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function MovEstoqueItens_Le1(ByVal objMovEstoque As ClassMovEstoque, ByVal iTipoMov As Integer) As Long
'Le os Ítens do MovimentodeEstoque relacionados ao Tipo de Movimento passado como parametro e que possuem o codigo especificado em objMovEstoque

Dim lErro As Long
Dim lComando As Long
Dim tItemMovEstoque As typeItemMovEstoque
Dim sComando_SQL As String
Dim objMovEstGrade As New ClassMovEstoque
Dim objItemMovEst As ClassItemMovEstoque
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemMovEstGrade As ClassItemMovEstoque
Dim iIndice As Integer
Dim colRastreamentoMovto As Collection
Dim objRastreamentoMovto As ClassRastreamentoMovto

On Error GoTo Erro_MovEstoqueItens_Le1

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55383

    With tItemMovEstoque

        .sAlmoxarifadoNomeRed = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCcl = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sOPCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sProdutoDesc = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sProdutoOP = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUM = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilEst = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilAplic = String(STRING_BUFFER_MAX_TEXTO, 0)

        If iTipoMov = MOV_EST_PRODUCAO Or iTipoMov = MOV_EST_REQ_PRODUCAO Or iTipoMov = MOV_EST_REQ_PRODUCAO_BENEF3 Or iTipoMov = MOV_EST_REQ_PRODUCAO_OUTROS Or iTipoMov = MOV_EST_PRODUCAO_BENEF3 Or iTipoMov = MOV_EST_PRODUCAO_OUTROS Or iTipoMov = MOV_EST_DEV_MATERIAL_PRODUCAO Then

            sComando_SQL = "SELECT MovEstoque_OP.NumIntDocGrade, MovEstoque_OP.ContaContabilEst, MovEstoque_OP.ContaContabilAplic, MovEstoque_OP.Data, MovEstoque_OP.NumIntDoc, MovEstoque_OP.FilialEmpresa, MovEstoque_OP.Codigo, MovEstoque_OP.OPCodigo, MovEstoque_OP.Custo, MovEstoque_OP.Apropriacao, MovEstoque_OP.Produto, Produtos.Descricao, MovEstoque_OP.ProdutoOP, MovEstoque_OP.SiglaUM, MovEstoque_OP.Quantidade, MovEstoque_OP.Almoxarifado, Almoxarifado.NomeReduzido, MovEstoque_OP.TipoMov, MovEstoque_OP.NumIntDocOrigem, MovEstoque_OP.Ccl, MovEstoque_OP.NumIntDocEst, MovEstoque_OP.HorasMaquina, MovEstoque_OP.DataInicioProducao FROM MovEstoque_OP, Produtos, Almoxarifado WHERE Produtos.Codigo = MovEstoque_OP.Produto AND Almoxarifado.Codigo = MovEstoque_OP.Almoxarifado AND MovEstoque_OP.FilialEmpresa = ? AND MovEstoque_OP.Codigo = ? AND MovEstoque_OP.NumIntDocEst = 0 AND (MovEstoque_OP.TipoMov = ? Or MovEstoque_OP.TipoMov = ? Or MovEstoque_OP.TipoMov = ?) ORDER BY MovEstoque_OP.NumIntDoc"

            If iTipoMov = MOV_EST_PRODUCAO Or iTipoMov = MOV_EST_PRODUCAO_BENEF3 Or iTipoMov = MOV_EST_PRODUCAO_OUTROS Then

                lErro = Comando_Executar(lComando, sComando_SQL, .lNumIntDocGrade, .sContaContabilEst, .sContaContabilAplic, .dtData, .lNumIntDoc, .iFilialEmpresa, .lCodigo, .sOPCodigo, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sProdutoOP, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .iTipoMov, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .lHorasMaquina, .dtDataInicioProducao, giFilialEmpresa, objMovEstoque.lCodigo, MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3, MOV_EST_PRODUCAO_OUTROS)
                If lErro <> AD_SQL_SUCESSO Then Error 55384
                
            ElseIf iTipoMov = MOV_EST_DEV_MATERIAL_PRODUCAO Then

                lErro = Comando_Executar(lComando, sComando_SQL, .lNumIntDocGrade, .sContaContabilEst, .sContaContabilAplic, .dtData, .lNumIntDoc, .iFilialEmpresa, .lCodigo, .sOPCodigo, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sProdutoOP, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .iTipoMov, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .lHorasMaquina, .dtDataInicioProducao, giFilialEmpresa, objMovEstoque.lCodigo, MOV_EST_DEV_MATERIAL_PRODUCAO, MOV_EST_DEV_MATERIAL_PRODUCAO, MOV_EST_DEV_MATERIAL_PRODUCAO)
                If lErro <> AD_SQL_SUCESSO Then Error 55384

            Else

                lErro = Comando_Executar(lComando, sComando_SQL, .lNumIntDocGrade, .sContaContabilEst, .sContaContabilAplic, .dtData, .lNumIntDoc, .iFilialEmpresa, .lCodigo, .sOPCodigo, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sProdutoOP, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .iTipoMov, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .lHorasMaquina, .dtDataInicioProducao, giFilialEmpresa, objMovEstoque.lCodigo, MOV_EST_REQ_PRODUCAO, MOV_EST_REQ_PRODUCAO_BENEF3, MOV_EST_REQ_PRODUCAO_OUTROS)
                If lErro <> AD_SQL_SUCESSO Then Error 55384

            End If

            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55385

            objMovEstoque.dtData = .dtData
            objMovEstoque.iFilialEmpresa = .iFilialEmpresa

            Do While lErro <> AD_SQL_SEM_DADOS

                Set objItemMovEst = objMovEstoque.colItens.Add(.lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, .lHorasMaquina, Nothing, Nothing, .dtDataInicioProducao)
                
                objItemMovEst.lNumIntDocGrade = .lNumIntDocGrade

                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55386

            Loop
            
        ElseIf iTipoMov = MOV_EST_APONT_SRV Then

            sComando_SQL = "SELECT MovEstoque_OS.NumIntDocGrade, MovEstoque_OS.ContaContabilEst, MovEstoque_OS.ContaContabilAplic, MovEstoque_OS.Data, MovEstoque_OS.NumIntDoc, MovEstoque_OS.FilialEmpresa, MovEstoque_OS.Codigo, MovEstoque_OS.OPCodigo, MovEstoque_OS.Custo, MovEstoque_OS.Apropriacao, MovEstoque_OS.Produto, Produtos.Descricao, MovEstoque_OS.ProdutoOP, MovEstoque_OS.SiglaUM, MovEstoque_OS.Quantidade, MovEstoque_OS.Almoxarifado, Almoxarifado.NomeReduzido, MovEstoque_OS.TipoMov, MovEstoque_OS.NumIntDocOrigem, MovEstoque_OS.Ccl, MovEstoque_OS.NumIntDocEst, MovEstoque_OS.HorasMaquina, MovEstoque_OS.DataInicioProducao FROM MovEstoque_OS, Produtos, Almoxarifado WHERE Produtos.Codigo = MovEstoque_OS.Produto AND Almoxarifado.Codigo = MovEstoque_OS.Almoxarifado AND MovEstoque_OS.FilialEmpresa = ? AND MovEstoque_OS.Codigo = ? AND MovEstoque_OS.NumIntDocEst = 0 ORDER BY MovEstoque_OS.NumIntDoc"
                
            lErro = Comando_Executar(lComando, sComando_SQL, .lNumIntDocGrade, .sContaContabilEst, .sContaContabilAplic, .dtData, .lNumIntDoc, .iFilialEmpresa, .lCodigo, .sOPCodigo, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sProdutoOP, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .iTipoMov, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .lHorasMaquina, .dtDataInicioProducao, giFilialEmpresa, objMovEstoque.lCodigo)
            If lErro <> AD_SQL_SUCESSO Then Error 55384
            
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55385

            objMovEstoque.dtData = .dtData
            objMovEstoque.iFilialEmpresa = .iFilialEmpresa

            Do While lErro <> AD_SQL_SEM_DADOS

                Set objItemMovEst = objMovEstoque.colItens.Add(.lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, .lHorasMaquina, Nothing, Nothing, .dtDataInicioProducao)
                
                objItemMovEst.lNumIntDocGrade = .lNumIntDocGrade
                objItemMovEst.lCodigo = objMovEstoque.lCodigo
                objItemMovEst.iFilialEmpresa = objMovEstoque.iFilialEmpresa

                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55386

            Loop
        
        Else

            '############################################
            'Inserido por Wagner
            .sObservacao = String(STRING_MOVESTOQUE_OBSERVACAO, 0)
            '############################################

            '#####################################
            'Alterado por Wagner
            lErro = Comando_Executar(lComando, "SELECT MovimentoEstoque.NumIntDocGrade, MovimentoEstoque.FilialEmpresa, MovimentoEstoque.ContaContabilEst, MovimentoEstoque.ContaContabilAplic, MovimentoEstoque.Data, MovimentoEstoque.NumIntDoc, MovimentoEstoque.Custo, MovimentoEstoque.Apropriacao, MovimentoEstoque.Produto, Produtos.Descricao, MovimentoEstoque.SiglaUM, MovimentoEstoque.Quantidade, MovimentoEstoque.Almoxarifado, Almoxarifado.NomeReduzido, MovimentoEstoque.NumIntDocOrigem, MovimentoEstoque.Ccl, MovimentoEstoque.NumIntDocEst, MovimentoEstoque.TipoMov, MovimentoEstoque.Observacao, MovimentoEstoque.Requisitante FROM MovimentoEstoque, Produtos, Almoxarifado WHERE Almoxarifado.Codigo = MovimentoEstoque.Almoxarifado AND Produtos.Codigo = MovimentoEstoque.Produto AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.Codigo = ? AND MovimentoEstoque.NumIntDocEst = 0 And MovimentoEstoque.TipoMov =? ORDER BY NumIntDoc", _
            .lNumIntDocGrade, .iFilialEmpresa, .sContaContabilEst, .sContaContabilAplic, .dtData, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .iTipoMov, .sObservacao, .lRequisitante, giFilialEmpresa, objMovEstoque.lCodigo, iTipoMov)
            If lErro <> AD_SQL_SUCESSO Then Error 55398
            '#####################################

            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55399

            Do While lErro <> AD_SQL_SEM_DADOS

                '############################################
                'Inserido por Wagner
                objMovEstoque.sObservacao = .sObservacao
                objMovEstoque.lRequisitante = .lRequisitante
                '############################################
                
                objMovEstoque.dtData = .dtData
                objMovEstoque.iFilialEmpresa = .iFilialEmpresa

                Set objItemMovEst = objMovEstoque.colItens.Add(.lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, 0, Nothing, Nothing, DATA_NULA)
                
                objItemMovEst.lNumIntDocGrade = .lNumIntDocGrade

                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55400

            Loop

        End If

    End With

    If objMovEstoque.colItens.Count = 0 Then Error 55387

    objMovEstGrade.iFilialEmpresa = objMovEstoque.iFilialEmpresa
    objMovEstGrade.lCodigo = objMovEstoque.lCodigo
    objMovEstGrade.iTipoMov = iTipoMov

    'Lê os ítens do Movimento de Estoque
    lErro = CF("MovimentoEstoqueGrade_Le", objMovEstGrade)
    If lErro <> SUCESSO Then gError 117635

    For Each objItemMovEstGrade In objMovEstGrade.colItens
    
        For iIndice = objMovEstoque.colItens.Count To 1 Step -1
        
            Set objItemMovEst = objMovEstoque.colItens(iIndice)
            
            If objItemMovEst.lNumIntDocGrade = objItemMovEstGrade.lNumIntDoc Then
    
                Set objItemRomaneioGrade = New ClassItemRomaneioGrade
        
                objItemRomaneioGrade.sProduto = objItemMovEst.sProduto
                objItemRomaneioGrade.sDescricao = objItemMovEst.sProdutoDesc
                objItemRomaneioGrade.iAlmoxarifado = objItemMovEst.iAlmoxarifado
                objItemRomaneioGrade.sAlmoxarifado = objItemMovEst.sAlmoxarifadoNomeRed
                objItemRomaneioGrade.sUMEstoque = objItemMovEst.sSiglaUM
                objItemRomaneioGrade.dQuantidade = objItemMovEst.dQuantidade
                objItemRomaneioGrade.sCodOP = objItemMovEst.sOPCodigo
                objItemRomaneioGrade.sProdOP = objItemMovEst.sProdutoOP
                objItemRomaneioGrade.lNumIntDoc = objItemMovEst.lNumIntDoc
                
                If iTipoMov = MOV_EST_REQ_PRODUCAO Then
                
                    Set colRastreamentoMovto = New Collection
                    
                    'Le o Rastreamento e preenche o grid com o Número do Lote e o Numero da Filial OP
                    lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEst.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRastreamentoMovto)
                    If lErro <> SUCESSO And lErro <> 78414 Then gError 78422
                    
                    'Se existe rastreamento
                    If colRastreamentoMovto.Count > 0 Then
                                    
                        'Seta o primeiro Lote
                        Set objRastreamentoMovto = colRastreamentoMovto(1)
                        
                        objItemRomaneioGrade.sLote = objRastreamentoMovto.sLote
                        
                        objItemRomaneioGrade.iFilialOP = objRastreamentoMovto.iFilialOP
                        
                    End If
                
                End If
                
                objItemMovEstGrade.colItensRomaneioGrade.Add objItemRomaneioGrade
                
                objItemMovEstGrade.dQuantidade = objItemMovEstGrade.dQuantidade + objItemRomaneioGrade.dQuantidade
                
                objMovEstoque.colItens.Remove (iIndice)
    
            End If
    
        Next
    
        objItemMovEstGrade.iPossuiGrade = MARCADO
        objMovEstoque.colItens.Add1 objItemMovEstGrade
    
    Next

    Call Comando_Fechar(lComando)

    MovEstoqueItens_Le1 = SUCESSO

    Exit Function

Erro_MovEstoqueItens_Le1:

    MovEstoqueItens_Le1 = Err

    Select Case Err

        Case 55383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55384, 55385, 55386, 55398, 55399, 55400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)

        Case 55387, 117635, 205839 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152053)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function MovEstoqueItem_Le_Primeiro(ByVal objItemMovEstoque As ClassItemMovEstoque) As Long
'Lê o primeiro item correspondentes a partir da FilialEmpresa, Código

Dim lErro As Long
Dim lComando As Long
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovEstoqueItem_Le_Primeiro

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 52888

    With tItemMovEstoque

        .sProduto = String(STRING_PRODUTO, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        .sContaContabilEst = String(STRING_CONTA, 0)
        .sContaContabilAplic = String(STRING_CONTA, 0)

        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov , NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, ContaContabilEst, ContaContabilAplic FROM MovimentoEstoque WHERE FilialEmpresa = ? AND Codigo = ? ORDER BY NumIntDoc", _
            .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .iTipoMov, .lNumIntDocOrigem, .iTipoNumIntDocOrigem, .dtData, .sCcl, .lNumIntDocEst, .sContaContabilEst, .sContaContabilAplic, objItemMovEstoque.iFilialEmpresa, objItemMovEstoque.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 52889

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 52890
        
        If lErro = AD_SQL_SEM_DADOS Then Error 52891
        
        objItemMovEstoque.lNumIntDoc = .lNumIntDoc
        objItemMovEstoque.dCusto = .dCusto
        objItemMovEstoque.iApropriacao = .iApropriacao
        objItemMovEstoque.sProduto = .sProduto
        objItemMovEstoque.sSiglaUM = .sSiglaUM
        objItemMovEstoque.dQuantidade = .dQuantidade
        objItemMovEstoque.iAlmoxarifado = .iAlmoxarifado
        objItemMovEstoque.iTipoMov = .iTipoMov
        objItemMovEstoque.lNumIntDocOrigem = .lNumIntDocOrigem
        objItemMovEstoque.iTipoNumIntDocOrigem = .iTipoNumIntDocOrigem
        objItemMovEstoque.dtData = .dtData
        objItemMovEstoque.sCcl = .sCcl
        objItemMovEstoque.lNumIntDocEst = .lNumIntDocEst
        objItemMovEstoque.sContaContabilEst = .sContaContabilEst
        objItemMovEstoque.sContaContabilAplic = .sContaContabilAplic
        
    End With

    Call Comando_Fechar(lComando)

    MovEstoqueItem_Le_Primeiro = SUCESSO

    Exit Function

Erro_MovEstoqueItem_Le_Primeiro:

    MovEstoqueItem_Le_Primeiro = Err

    Select Case Err

        Case 52888
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 52889, 52890
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)

        Case 52891 'Tratado na Rotina Chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 152054)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function MovEstoque_Le_NomeTela(ByVal objMovEstoque As ClassMovEstoque, sTela As String) As Long
'Esta função retorna o nome da tela
'Relacionada com o Movimento Estoque passado no parametro

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer

On Error GoTo Erro_MovEstoque_Le_NomeTela
    
    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 52892
    Next
    
    sTela = String(STRING_NOME_TELA, 0)
    
    'Lê o Nome da Tela relacionado com o Movimento de Estoque
    lErro = Comando_Executar(alComando(1), "SELECT NomeTela FROM MovimentoEstoque, TiposMovimentoEstoque WHERE MovimentoEstoque.Codigo = ? AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.TipoMov = TiposMovimentoEstoque.Codigo", sTela, objMovEstoque.lCodigo, objMovEstoque.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 52893

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 52894
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    MovEstoque_Le_NomeTela = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Le_NomeTela:

    MovEstoque_Le_NomeTela = Err
    
    Select Case Err
        
        Case 52892
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 52893, 52894
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152055)
        
    End Select
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function MovEstoque_Le_NomeTela_Inventario(ByVal objMovEstoque As ClassMovEstoque, sTela As String, sCodigoInv As String) As Long
'Esta função retorna o nome da tela e Codigo do Inventario
'Relacionada com o Movimento Estoque passado no parametro
    
Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 2) As Long

On Error GoTo Erro_MovEstoque_Le_NomeTela_Inventario
    
    'Abre os  comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 52900
    Next
    
    sTela = String(STRING_NOME_TELA, 0)
    sCodigoInv = String(STRING_INVENTARIO_CODIGO, 0)
    
    'Lê o Nome da Tela relacionado com o Movimento de Estoque
    lErro = Comando_Executar(alComando(1), "SELECT Inventario.Codigo, TiposMovimentoEstoque.NomeTela FROM MovimentoEstoque, TiposMovimentoEstoque, Inventario WHERE MovimentoEstoque.Codigo = ? AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.TipoMov = TiposMovimentoEstoque.Codigo AND MovimentoEstoque.NumIntDocOrigem = Inventario.NumIntDoc", sCodigoInv, sTela, objMovEstoque.lCodigo, objMovEstoque.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 52901

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 52902
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    MovEstoque_Le_NomeTela_Inventario = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Le_NomeTela_Inventario:

    MovEstoque_Le_NomeTela_Inventario = Err
    
    Select Case Err
        
        Case 52900
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 52901, 52902
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 152056)
                
    End Select
    
    Exit Function
    
End Function

Public Function Calcula_Custo_Atual(ByVal objProduto As ClassProduto, dCusto As Double, ByVal sTipo As String) As Long
'calcula o custo atual do produto passado como parametro e devolve-o de dCusto. sTipo indica o tipo do escaninho em que o produto se encontra

Dim lErro As Long
Dim lComando As Long
Dim objEstoqueMes As New ClassEstoqueMes

On Error GoTo Erro_Calcula_Custo_Atual

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 55487

    'O valor da diferenca vai depender se é um produto Produzido ou Compravel
    If objProduto.iCompras = PRODUTO_COMPRAVEL Then
    
        'se se tratar de um produto nosso
        If sTipo = STRING_QUANT_DISPONIVEL_NOSSA Or sTipo = STRING_QUANT_RECEB_INDISP Or sTipo = STRING_QUANT_OUTRAS_INDISP Or sTipo = STRING_QUANT_DEFEIT Then

            'Le o ultimo CustoMedio do Mes-Ano aberto
            lErro = CF("CustoMedioAtual_Le", objProduto.sCodigo, dCusto, giFilialEmpresa)
            If lErro <> SUCESSO Then gError 49621
            
        'se se tratar de um produto consignado de terceiros, le o custo médio do escaninho respectivo
        ElseIf sTipo = STRING_QUANT_3_CONSIG Then

            objEstoqueMes.iFilialEmpresa = giFilialEmpresa

            lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
            If lErro <> SUCESSO Then gError 71516
    
            lErro = CF("SldMesEst1_Le_CustoMedioConsig3", lComando, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objProduto.sCodigo, objEstoqueMes.iMes, dCusto)
            If lErro <> SUCESSO Then gError 71517

        'se se tratar de um produto em conserto de terceiros, le o custo médio do escaninho respectivo
        ElseIf sTipo = STRING_QUANT_3_CONSERTO Then

            objEstoqueMes.iFilialEmpresa = giFilialEmpresa

            lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
            If lErro <> SUCESSO Then gError 71518
    
            lErro = CF("SldMesEst1_Le_CustoMedioConserto3", lComando, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objProduto.sCodigo, objEstoqueMes.iMes, dCusto)
            If lErro <> SUCESSO Then gError 71519
            
        'se se tratar de um produto demo de terceiros, le o custo médio do escaninho respectivo
        ElseIf sTipo = STRING_QUANT_3_DEMO Then

            objEstoqueMes.iFilialEmpresa = giFilialEmpresa

            lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
            If lErro <> SUCESSO Then gError 71520
    
            lErro = CF("SldMesEst1_Le_CustoMedioDemo3", lComando, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objProduto.sCodigo, objEstoqueMes.iMes, dCusto)
            If lErro <> SUCESSO Then gError 71521

        'se se tratar de um outro produto de terceiros, le o custo médio do escaninho respectivo
        ElseIf sTipo = STRING_QUANT_3_OUTRAS Then

            objEstoqueMes.iFilialEmpresa = giFilialEmpresa

            lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
            If lErro <> SUCESSO Then gError 71522
    
            lErro = CF("SldMesEst1_Le_CustoMedioOutros3", lComando, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objProduto.sCodigo, objEstoqueMes.iMes, dCusto)
            If lErro <> SUCESSO Then gError 71523

        'se se tratar de um produto em beneficiamento de terceiros, le o custo médio do escaninho respectivo
        ElseIf sTipo = STRING_QUANT_3_BENEF Then

            objEstoqueMes.iFilialEmpresa = giFilialEmpresa

            lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
            If lErro <> SUCESSO Then gError 71524
    
            lErro = CF("SldMesEst1_Le_CustoMedioBenef3", lComando, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objProduto.sCodigo, objEstoqueMes.iMes, dCusto)
            If lErro <> SUCESSO Then gError 71525

        'se se tratar de um produto consignado nosso, le o custo médio do escaninho respectivo
        ElseIf sTipo = STRING_QUANT_NOSSO_CONSIG Or sTipo = STRING_QUANT_NOSSO_CONSIG_CI Then

            objEstoqueMes.iFilialEmpresa = giFilialEmpresa

            lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
            If lErro <> SUCESSO Then gError 71516
    
            lErro = CF("SldMesEst2_Le_CustoMedioConsig", lComando, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objProduto.sCodigo, objEstoqueMes.iMes, dCusto)
            If lErro <> SUCESSO Then gError 71517

        'se se tratar de um produto em conserto nosso, le o custo médio do escaninho respectivo
        ElseIf sTipo = STRING_QUANT_NOSSO_CONSERTO Or sTipo = STRING_QUANT_NOSSO_CONSERTO_CI Then

            objEstoqueMes.iFilialEmpresa = giFilialEmpresa

            lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
            If lErro <> SUCESSO Then gError 71518
    
            lErro = CF("SldMesEst2_Le_CustoMedioConserto", lComando, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objProduto.sCodigo, objEstoqueMes.iMes, dCusto)
            If lErro <> SUCESSO Then gError 71519
            
        'se se tratar de um produto demo nosso, le o custo médio do escaninho respectivo
        ElseIf sTipo = STRING_QUANT_NOSSO_DEMO Or sTipo = STRING_QUANT_NOSSO_DEMO_CI Then

            objEstoqueMes.iFilialEmpresa = giFilialEmpresa

            lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
            If lErro <> SUCESSO Then gError 71520
    
            lErro = CF("SldMesEst2_Le_CustoMedioDemo", lComando, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objProduto.sCodigo, objEstoqueMes.iMes, dCusto)
            If lErro <> SUCESSO Then gError 71521

        'se se tratar de um outro produto nosso, le o custo médio do escaninho respectivo
        ElseIf sTipo = STRING_QUANT_NOSSO_OUTRAS Or sTipo = STRING_QUANT_NOSSO_OUTRAS_CI Then

            objEstoqueMes.iFilialEmpresa = giFilialEmpresa

            lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
            If lErro <> SUCESSO Then gError 71522
    
            lErro = CF("SldMesEst2_Le_CustoMedioOutros", lComando, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objProduto.sCodigo, objEstoqueMes.iMes, dCusto)
            If lErro <> SUCESSO Then gError 71523

        'se se tratar de um produto em beneficiamento de terceiros, le o custo médio do escaninho respectivo
        ElseIf sTipo = STRING_QUANT_NOSSO_BENEF Or sTipo = STRING_QUANT_NOSSO_BENEF_CI Then

            objEstoqueMes.iFilialEmpresa = giFilialEmpresa

            lErro = CF("EstoqueMes_Le_Aberto", objEstoqueMes)
            If lErro <> SUCESSO Then gError 71524
    
            lErro = CF("SldMesEst2_Le_CustoMedioBenef", lComando, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objProduto.sCodigo, objEstoqueMes.iMes, dCusto)
            If lErro <> SUCESSO Then gError 71525

        End If
        
    ElseIf objProduto.iCompras = PRODUTO_PRODUZIVEL And objProduto.iApropriacaoCusto = APROPR_CUSTO_STANDARD Then
    
        'Le o CustoStandard do Ano-Mes aberto
        lErro = CF("CustoStandard_Le", objProduto.sCodigo, dCusto, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 69852 Then gError 55428

        If lErro = 69852 Then gError 69853

    Else
    
        'le o ultimo custo medio producao apurado
        lErro = CF("CustoMedioProducaoApurado_Le", objProduto.sCodigo, dCusto, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 69867 And lErro <> 69868 Then gError 49654
            
        'Se não encontrou o SldMesEst
        If lErro = 69867 Then gError 69869
            
        'Se não encontrou o SldMesEst
        If lErro = 69868 Then gError 69870

    End If
    
    Call Comando_Fechar(lComando)
    
    Calcula_Custo_Atual = SUCESSO
    
    Exit Function

Erro_Calcula_Custo_Atual:

    Calcula_Custo_Atual = gErr
    
    Select Case gErr
        
        Case 49621, 49654, 55428, 71516, 71517, 71518, 71519, 71520, 71521, 71522, 71523, 71524, 71525
                
        Case 69869, 69870, 69853
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SLDMESEST_NAO_CADASTRADO1", gErr, giFilialEmpresa, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152057)
                
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function EstoqueProduto_Le_Almoxarifados1(ByVal sProduto As String, ByVal colAlmoxarifados As Collection) As Long
'Lê todos os Almoxarifados do Produto para todas as filiais e seus nomes reduzidos

Dim lErro As Long
Dim lComando As Long
Dim objAlmoxarifado As ClassAlmoxarifado
Dim iAlmoxarifado As Integer
Dim sNomeReduzido As String

On Error GoTo Erro_EstoqueProduto_Le_Almoxarifados1

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 55487

    sNomeReduzido = String(STRING_ALMOXARIFADO_NOMEREDUZIDO, 0)

    'Lê a tabela Estoque Produto
    lErro = Comando_Executar(lComando, "SELECT Almoxarifado, NomeReduzido FROM EstoqueProduto, Almoxarifado WHERE Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo", iAlmoxarifado, sNomeReduzido, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 55488

    'Lê o primeiro Almoxarifado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55489

    'Enquanto a tabela tiver dados
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objAlmoxarifado = New ClassAlmoxarifado

        objAlmoxarifado.iCodigo = iAlmoxarifado
        objAlmoxarifado.sNomeReduzido = sNomeReduzido

        'Coloca o Almoxarifado lido na coleção
        colAlmoxarifados.Add objAlmoxarifado

        'Lê o próximo Almoxarifado
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55490

    Loop

    Call Comando_Fechar(lComando)

    EstoqueProduto_Le_Almoxarifados1 = SUCESSO

    Exit Function

Erro_EstoqueProduto_Le_Almoxarifados1:

    EstoqueProduto_Le_Almoxarifados1 = Err

    Select Case Err

        Case 55487
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 55488, 55489, 55490
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO2", Err, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152058)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_ObterRelMesMedia12Meses(ByVal sProduto As String, ByVal iFilialEmpresa As Integer, dConsumoMedioQtde As Double, dVendaMediaQtde As Double, alComando() As Long) As Long
'IMPORTANTE: especifica para classe ClassRelMesEst de apoio a execucao de relatorios
'obtem consumo medio de um produto em uma filial considerando os 12 meses anteriores ao atual
'Os meses anteriores a implantacao do produto na filial sao desprezados
'ATENCAO: se o alComando() vier zerado, será aberto nesta funcao e ficará aberto.
'poderia retornar -1 p/indicar que nao houve consumo calculado

Dim lErro As Long, iIndice As Integer
Dim adConsumoMensal(1 To 12) As Double, adConsumo(1 To 12) As Double, dConsumoAcum As Double
Dim adVendaMensal(1 To 12) As Double, adVenda(1 To 12) As Double, dVendaAcum As Double
Dim iMesAtual As Integer, iAno As Integer, iAnoAtual As Integer, dtDataInicialProdFilial As Date, iMesesMedia As Integer, iMesAux As Integer, iMesInicial As Integer

On Error GoTo Erro_Produto_ObterRelMesMedia12Meses

    For iIndice = 0 To 1
        If alComando(iIndice) = 0 Then
            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
            If alComando(iIndice) = 0 Then Error 41952
        End If
    Next
    
    'obter a data de implantacao do produto na filial
    lErro = Comando_Executar(alComando(0), "SELECT MIN(EstoqueProduto.DataInicial) FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.FilialEmpresa = ?", dtDataInicialProdFilial, sProduto, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 41953
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41954
    
    If lErro <> AD_SQL_SUCESSO Then dtDataInicialProdFilial = DATA_NULA
    
    'se o produto foi implantado na filial antes do mes corrente
    If CDate("01/" & Month(gdtDataAtual) & "/" & Year(gdtDataAtual)) > dtDataInicialProdFilial And dtDataInicialProdFilial <> DATA_NULA Then
    
        iAnoAtual = Year(gdtDataAtual)
        iMesAtual = Month(gdtDataAtual)
        iMesInicial = Month(dtDataInicialProdFilial)
        
        If iMesAtual = 1 Then
        
            lErro = Comando_Executar(alComando(1), "SELECT QuantCons1, QuantCons2, QuantCons3, QuantCons4, QuantCons5, QuantCons6, QuantCons7, QuantCons8, QuantCons9, QuantCons10, QuantCons11, QuantCons12, QuantVend1, QuantVend2, QuantVend3, QuantVend4, QuantVend5, QuantVend6, QuantVend7, QuantVend8, QuantVend9, QuantVend10, QuantVend11, QuantVend12 FROM SldMesEst WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ?", adConsumoMensal(1), adConsumoMensal(2), adConsumoMensal(3), adConsumoMensal(4), adConsumoMensal(5), adConsumoMensal(6), adConsumoMensal(7), adConsumoMensal(8), adConsumoMensal(9), adConsumoMensal(10), adConsumoMensal(11), adConsumoMensal(12), adVendaMensal(1), adVendaMensal(2), adVendaMensal(3), adVendaMensal(4), adVendaMensal(5), adVendaMensal(6), adVendaMensal(7), adVendaMensal(8), adVendaMensal(9), adVendaMensal(10), adVendaMensal(11), adVendaMensal(12), iFilialEmpresa, iAnoAtual - 1, sProduto)
            If lErro <> AD_SQL_SUCESSO Then Error 41955
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41956

            If lErro <> AD_SQL_SUCESSO Then
            
                dConsumoMedioQtde = 0
                dVendaMediaQtde = 0
                
            Else
            
                'se o produto foi implantado no ano anterior ao atual
                If Year(dtDataInicialProdFilial) = iAnoAtual - 1 Then
                
                    For iMesAux = Month(dtDataInicialProdFilial) To 12
                        
                        iMesesMedia = iMesesMedia + 1
                        dConsumoAcum = dConsumoAcum + adConsumoMensal(iMesAux)
                        dVendaAcum = dVendaAcum + adVendaMensal(iMesAux)
                        
                    Next
                    
                    dConsumoMedioQtde = dConsumoAcum / iMesesMedia
                    dVendaMediaQtde = dVendaAcum / iMesesMedia
                    
                Else 'se a implantacao foi ainda anterior
                
                    dConsumoMedioQtde = (adConsumoMensal(1) + adConsumoMensal(2) + adConsumoMensal(3) + adConsumoMensal(4) + adConsumoMensal(5) + adConsumoMensal(6) + adConsumoMensal(7) + adConsumoMensal(8) + adConsumoMensal(9) + adConsumoMensal(10) + adConsumoMensal(11) + adConsumoMensal(12)) / 12
                    dVendaMediaQtde = (adVendaMensal(1) + adVendaMensal(2) + adVendaMensal(3) + adVendaMensal(4) + adVendaMensal(5) + adVendaMensal(6) + adVendaMensal(7) + adVendaMensal(8) + adVendaMensal(9) + adVendaMensal(10) + adVendaMensal(11) + adVendaMensal(12)) / 12
                    
                End If
                
            End If
            
        Else 'se tem que pegar uma parte do ano atual e parte do ano anterior
        
            lErro = Comando_Executar(alComando(1), "SELECT Ano, QuantCons1, QuantCons2, QuantCons3, QuantCons4, QuantCons5, QuantCons6, QuantCons7, QuantCons8, QuantCons9, QuantCons10, QuantCons11, QuantCons12, QuantVend1, QuantVend2, QuantVend3, QuantVend4, QuantVend5, QuantVend6, QuantVend7, QuantVend8, QuantVend9, QuantVend10, QuantVend11, QuantVend12 FROM SldMesEst WHERE FilialEmpresa = ? AND (Ano = ? OR Ano = ?) AND Produto = ? ORDER BY Ano", iAno, adConsumoMensal(1), adConsumoMensal(2), adConsumoMensal(3), adConsumoMensal(4), adConsumoMensal(5), adConsumoMensal(6), adConsumoMensal(7), adConsumoMensal(8), adConsumoMensal(9), adConsumoMensal(10), adConsumoMensal(11), adConsumoMensal(12), adVendaMensal(1), adVendaMensal(2), adVendaMensal(3), adVendaMensal(4), adVendaMensal(5), adVendaMensal(6), adVendaMensal(7), adVendaMensal(8), adVendaMensal(9), adVendaMensal(10), adVendaMensal(11), adVendaMensal(12), iFilialEmpresa, iAnoAtual - 1, iAnoAtual, sProduto)
            If lErro <> AD_SQL_SUCESSO Then Error 41957
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41958
        
            'se nao achou nenhum registro em SldMesEst
            If lErro <> AD_SQL_SUCESSO Then
            
                dConsumoMedioQtde = 0
                dVendaMediaQtde = 0
                
            Else
        
                'se achou registro do ano anterior
                If iAno = iAnoAtual - 1 Then
                                
                    'se o produto foi implantado no ano anterior ao atual
                    If Year(dtDataInicialProdFilial) = iAnoAtual - 1 Then
                    
                        For iMesAux = IIf(iMesAtual > iMesInicial, iMesAtual, iMesInicial) To 12
                            
                            iMesesMedia = iMesesMedia + 1
                            dConsumoAcum = dConsumoAcum + adConsumoMensal(iMesAux)
                            dVendaAcum = dVendaAcum + adVendaMensal(iMesAux)
                            
                        Next
                        
                    Else 'se a implantacao foi ainda anterior
                    
                        For iMesAux = iMesAtual To 12
                            
                            iMesesMedia = iMesesMedia + 1
                            dConsumoAcum = dConsumoAcum + adConsumoMensal(iMesAux)
                            dVendaAcum = dVendaAcum + adVendaMensal(iMesAux)
                            
                        Next
                        
                    End If
                    
                    'buscar registro do ano atual
                    lErro = Comando_BuscarProximo(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41959
                
                    If lErro = AD_SQL_SUCESSO Then
                    
                        For iMesAux = 1 To iMesAtual - 1
                            
                            iMesesMedia = iMesesMedia + 1
                            dConsumoAcum = dConsumoAcum + adConsumoMensal(iMesAux)
                            dVendaAcum = dVendaAcum + adVendaMensal(iMesAux)
                            
                        Next
                                    
                    End If
                    
                Else 'se achou registro apenas do ano atual
                
                    'se o produto foi implantado no ano atual
                    If Year(dtDataInicialProdFilial) = iAnoAtual Then
                    
                        For iMesAux = IIf(1 > iMesInicial, 1, iMesInicial) To iMesAtual - 1
                            
                            iMesesMedia = iMesesMedia + 1
                            dConsumoAcum = dConsumoAcum + adConsumoMensal(iMesAux)
                            dVendaAcum = dVendaAcum + adVendaMensal(iMesAux)
                            
                        Next
                        
                    Else 'se a implantacao foi ainda anterior
                    
                        For iMesAux = 1 To iMesAtual - 1
                            
                            dConsumoAcum = dConsumoAcum + adConsumoMensal(iMesAux)
                            dVendaAcum = dVendaAcum + adVendaMensal(iMesAux)
                            
                        Next
                        
                        iMesesMedia = iMesesMedia + 1
                    
                    End If
                
                End If
            
                If iMesesMedia <> 0 Then
                
                    dConsumoMedioQtde = dConsumoAcum / iMesesMedia
                    dVendaMediaQtde = dVendaAcum / iMesesMedia
                    
                Else
                
                    dConsumoMedioQtde = 0
                    dVendaMediaQtde = 0
                    
                End If
        
            End If
            
        End If
        
    Else
    
        dConsumoMedioQtde = 0
        dVendaMediaQtde = 0
        
    End If
    
    Produto_ObterRelMesMedia12Meses = SUCESSO
     
    Exit Function
    
Erro_Produto_ObterRelMesMedia12Meses:

    Produto_ObterRelMesMedia12Meses = Err
     
    Select Case Err
          
        Case 41952
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 41953, 41954
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", Err)
        
        Case 41955, 41956, 41957, 41958, 41959
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152059)
     
    End Select
     
    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Produto_ObterRelMesEstMedio(ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, dQuantidadeMedia As Double, dValorMedio As Double, alComando() As Long) As Long
'IMPORTANTE: especifica para classe ClassRelMesEst de apoio a execucao de relatorios
'Obtém a quantidade e o valores medios de um Produto num periodo para uma filial
'ATENCAO: se o alComando() vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long, iIndice As Integer, dtPrimDiaMesDataInic As Date
Dim adQtdeEntMes(1 To 12) As Double, adQtdeSaiMes(1 To 12) As Double
Dim adValEntMes(1 To 12) As Double, adValSaiMes(1 To 12) As Double
Dim dQtdeEntDia As Double, dQtdeSaiDia As Double, dValEntDia As Double, dValSaiDia As Double
Dim dtData As Date, dQtdeInicial As Double, dValorInicial As Double

Dim dAcumQtde As Double, dAcumValor As Double 'acumularao qtde e valores multiplicados por numero de dias
Dim dQtde As Double, dValor As Double 'irao mantendo dia a dia o saldo do estoque
Dim dtDataAnterior As Date

On Error GoTo Erro_Produto_ObterRelMesEstMedio

    For iIndice = 0 To 1
        If alComando(iIndice) = 0 Then
            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
            If alComando(iIndice) = 0 Then Error 41943
        End If
    Next
    
    '??? otimizacao possivel: pegar apenas dados dos meses necessarios
    lErro = Comando_Executar(alComando(0), "SELECT QuantInicial, ValorInicial, QuantEnt1, ValorEnt1, QuantSai1, ValorSai1, QuantEnt2, ValorEnt2, QuantSai2, ValorSai2, QuantEnt3, ValorEnt3, QuantSai3, ValorSai3, QuantEnt4, ValorEnt4, QuantSai4, ValorSai4, QuantEnt5, ValorEnt5, QuantSai5, ValorSai5, QuantEnt6, ValorEnt6, QuantSai6, ValorSai6, QuantEnt7, ValorEnt7, QuantSai7, ValorSai7, QuantEnt8, ValorEnt8, QuantSai8, ValorSai8, QuantEnt9, ValorEnt9, QuantSai9, ValorSai9, QuantEnt10, ValorEnt10, QuantSai10, ValorSai10, QuantEnt11, ValorEnt11, QuantSai11, ValorSai11, QuantEnt12, ValorEnt12, QuantSai12, ValorSai12 FROM SldMesEst WHERE FilialEmpresa = ? AND Ano = ? AND Produto = ?", _
        dQtdeInicial, dValorInicial, adQtdeEntMes(1), adValEntMes(1), adQtdeSaiMes(1), adValSaiMes(1), adQtdeEntMes(2), adValEntMes(2), adQtdeSaiMes(2), adValSaiMes(2), adQtdeEntMes(3), adValEntMes(3), adQtdeSaiMes(3), adValSaiMes(3), adQtdeEntMes(4), adValEntMes(4), adQtdeSaiMes(4), adValSaiMes(4), adQtdeEntMes(5), adValEntMes(5), adQtdeSaiMes(5), adValSaiMes(5), adQtdeEntMes(6), adValEntMes(6), adQtdeSaiMes(6), adValSaiMes(6), adQtdeEntMes(7), adValEntMes(7), adQtdeSaiMes(7), adValSaiMes(7), adQtdeEntMes(8), adValEntMes(8), adQtdeSaiMes(8), adValSaiMes(8), adQtdeEntMes(9), adValEntMes(9), adQtdeSaiMes(9), adValSaiMes(9), adQtdeEntMes(10), adValEntMes(10), adQtdeSaiMes(10), adValSaiMes(10), adQtdeEntMes(11), adValEntMes(11), adQtdeSaiMes(11), adValSaiMes(11), adQtdeEntMes(12), adValEntMes(12), adQtdeSaiMes(12), adValSaiMes(12), iFilialEmpresa, Year(dtDataInicial), sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 41944
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41945

    'se achou algum registro
    If lErro <> AD_SQL_SEM_DADOS Then
    
        dQtde = dQtdeInicial
        dValor = dValorInicial
        
        For iIndice = 1 To Month(dtDataInicial) - 1
            
            dQtde = dQtde + adQtdeEntMes(iIndice) - adQtdeSaiMes(iIndice)
            dValor = dValor + adValEntMes(iIndice) - adValSaiMes(iIndice)
            
        Next
        
    End If
    
    dtPrimDiaMesDataInic = CDate("01/" & CStr(Month(dtDataInicial)) & "/" & CStr(Year(dtDataInicial)))
    
    'obtem as movimentacoes diarias do 1o dia do mes da data inicial até a data final
    'obs: as movimentacoes anteriores a data inicial servicao p/obter os saldos nesta data
    lErro = Comando_Executar(alComando(1), "SELECT Data, QuantEntrada, QuantSaida, ValorEntrada, ValorSaida FROM SldDiaEst WHERE FilialEmpresa = ? AND Data >= ? AND Data <= ? AND Produto = ?", dtData, dQtdeEntDia, dQtdeSaiDia, dValEntDia, dValSaiDia, iFilialEmpresa, dtPrimDiaMesDataInic, dtDataFinal, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 41946

    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41947

    dtDataAnterior = dtDataInicial - 1
    
    Do While lErro <> AD_SQL_SEM_DADOS

        'data dentro do periodo vai entrar no calculo do estoque medio
        If dtData >= dtDataInicial Then
        
            dAcumQtde = dAcumQtde + ((dtData - dtDataAnterior) * dQtde) + dQtdeEntDia - dQtdeSaiDia
            dAcumValor = dAcumValor + ((dtData - dtDataAnterior) * dValor) + dValEntDia - dValSaiDia
        
            dtDataAnterior = dtData
    
        End If
        
        dQtde = dQtde + dQtdeEntDia - dQtdeSaiDia
        dValor = dValor + dValEntDia - dValSaiDia
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41948
    
    Loop
        
    If dtDataFinal > dtDataAnterior Then
    
        dAcumQtde = dAcumQtde + ((dtDataFinal - dtDataAnterior) * dQtde)
        dAcumValor = dAcumValor + ((dtDataFinal - dtDataAnterior) * dValor)
    
    End If
    
    dQuantidadeMedia = dAcumQtde / (dtDataFinal - dtDataInicial + 1)
    dValorMedio = dAcumValor / (dtDataFinal - dtDataInicial + 1)
    
    Produto_ObterRelMesEstMedio = SUCESSO

    Exit Function

Erro_Produto_ObterRelMesEstMedio:

    Produto_ObterRelMesEstMedio = Err

    Select Case Err

        Case 41943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41946, 41947, 41948
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST", Err)

        Case 41944, 41945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152060)

    End Select

    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function Produto_ObterRelMesEst(ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, dQuantidadeVendida As Double, dCustoTotal As Double, dQuantidadeEntrada As Double, dQuantidadeSaida As Double, dValorEntrada As Double, dValorSaida As Double, lComando As Long) As Long
'IMPORTANTE: especifica para classe ClassRelMesEst de apoio a execucao de relatorios
'Obtém a quantidade vendida, o custo total (custo do material vendido) e outros dados do Produto correspondentes a um periodo
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long

On Error GoTo Erro_Produto_ObterRelMesEst

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then Error 43328
    End If
    
    If dtDataInicial <> DATA_NULA And dtDataFinal <> DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT SUM(QuantEntrada), SUM(QuantSaida), SUM(QuantVend), SUM(ValorEntrada), SUM(ValorSaida), SUM(ValorVend) FROM SldDiaEst WHERE FilialEmpresa = ? AND Produto = ? AND Data >= ? AND Data <= ?", dQuantidadeEntrada, dQuantidadeSaida, dQuantidadeVendida, dValorEntrada, dValorSaida, dCustoTotal, iFilialEmpresa, sProduto, dtDataInicial, dtDataFinal)

    ElseIf dtDataInicial <> DATA_NULA And dtDataFinal = DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT SUM(QuantEntrada), SUM(QuantSaida), SUM(QuantVend), SUM(ValorEntrada), SUM(ValorSaida), SUM(ValorVend) FROM SldDiaEst WHERE FilialEmpresa = ? AND Produto = ? AND Data >= ?", dQuantidadeEntrada, dQuantidadeSaida, dQuantidadeVendida, dValorEntrada, dValorSaida, dCustoTotal, iFilialEmpresa, sProduto, dtDataInicial)

    ElseIf dtDataInicial = DATA_NULA And dtDataFinal <> DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT SUM(QuantEntrada), SUM(QuantSaida), SUM(QuantVend), SUM(ValorEntrada), SUM(ValorSaida), SUM(ValorVend) FROM SldDiaEst WHERE FilialEmpresa = ? AND Produto = ? AND Data <= ?", dQuantidadeEntrada, dQuantidadeSaida, dQuantidadeVendida, dValorEntrada, dValorSaida, dCustoTotal, iFilialEmpresa, sProduto, dtDataFinal)

    Else
        lErro = Comando_Executar(lComando, "SELECT SUM(QuantEntrada), SUM(QuantSaida), SUM(QuantVend), SUM(ValorEntrada), SUM(ValorSaida), SUM(ValorVend) FROM SldDiaEst WHERE FilialEmpresa = ? AND Produto = ?", dQuantidadeEntrada, dQuantidadeSaida, dQuantidadeVendida, dValorEntrada, dValorSaida, dCustoTotal, iFilialEmpresa, sProduto)

    End If
    If lErro <> AD_SQL_SUCESSO Then Error 43345

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43346

    If lErro = AD_SQL_SEM_DADOS Then

        dQuantidadeVendida = 0
        dQuantidadeEntrada = 0
        dQuantidadeSaida = 0
        dValorEntrada = 0
        dValorSaida = 0
        dCustoTotal = 0

    End If

    Produto_ObterRelMesEst = SUCESSO

    Exit Function

Erro_Produto_ObterRelMesEst:

    Produto_ObterRelMesEst = Err

    Select Case Err

        Case 43328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 43345, 43346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152061)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Function Produto_ObterRelMesSemEst(ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, dQuantidadeVendida As Double, dCustoTotal As Double, lComando1 As Long, lComando2 As Long, lComando3 As Long, lComando4 As Long, Optional lComando5 As Long = 0) As Long

Dim lErro As Long
Dim dtDataEmissao As Date, lFornEntTerc As Long
Dim iFilialFornEntTerc As Integer, sUnidadeMed As String, dQuantidade As Double
Dim lNumIntItemNF As Long, sUMMovEst As String, dQtdeMovEst As Double, dCustoMovEst As Double
Dim sUMEst As String, iClasseUM As Integer, dFatorCusto As Double, dFatorQtde As Double, dCustoUnit As Double
Dim iTipoNF As Integer, iTipoNFFiltro1 As Integer, iTipoNFFiltro2 As Integer, lNumIntNF As Long

On Error GoTo Erro_Produto_ObterRelMesSemEst

    If lComando1 = 0 Then
        lComando1 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando1 = 0 Then gError 211090
    End If
    If lComando2 = 0 Then
        lComando2 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando2 = 0 Then gError 211090
    End If
    If lComando3 = 0 Then
        lComando3 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando3 = 0 Then gError 211090
    End If
    If lComando4 = 0 Then
        lComando4 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando4 = 0 Then gError 211090
    End If
    If lComando5 = 0 Then
        lComando5 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando5 = 0 Then gError 211090
    End If
    
    dQuantidadeVendida = 0
    dCustoTotal = 0
    
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    sUMEst = String(STRING_UM_SIGLA, 0)
    
    'Lê as remessas com entrega por terceiro/entrega futura
    lErro = Comando_Executar(lComando1, "SELECT N.TipoNfiscal, N.NumIntDoc,  N.DataEmissao, N.FornEntTerc, N.FilialFornEntTerc, I.UnidadeMed, I.Quantidade, P.SiglaUmEstoque, P.ClasseUM FROM NFiscal AS N, ItensNFiscal AS I, Produtos AS P WHERE N.NumIntDoc = I.NumIntNF AND N.TipoNfiscal IN (?,?,?) AND N.Status <> 7 AND N.FilialEmpresa = ? AND I.Produto = P.Codigo AND I.Produto = ? AND (N.DataEmissao>= ? OR ?=?) AND (N.DataEmissao<= ? OR ?=?)", _
                        iTipoNF, lNumIntNF, dtDataEmissao, lFornEntTerc, iFilialFornEntTerc, sUnidadeMed, dQuantidade, sUMEst, iClasseUM, DOCINFO_NFIFVEFPV, DOCINFO_NFIFVEF, DOCINFO_NFIFVETPV, iFilialEmpresa, sProduto, dtDataInicial, dtDataInicial, DATA_NULA, dtDataFinal, dtDataFinal, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 211091

    lErro = Comando_BuscarProximo(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211092

    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = CF("UM_Conversao_Trans", iClasseUM, sUnidadeMed, sUMEst, dFatorQtde)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If iTipoNF = DOCINFO_NFIFVETPV Then 'Entrega por terceiros
            
            iTipoNFFiltro1 = DOCINFO_NFEEDSO 'Dev simbólica associada
            iTipoNFFiltro2 = DOCINFO_NFIEDSO
            
            'Lê a primeira devolução simbólica do produto após a remessa para esse terceiro
            lErro = Comando_Executar(lComando2, "SELECT I.NumIntDoc FROM NFiscal AS N, ItensNFiscal AS I WHERE N.NumIntDoc = I.NumIntNF AND N.TipoNfiscal IN (?,?) AND N.Status <> 7 AND N.DataEmissao > ? AND N.FilialEmpresa = ? AND I.Produto = ? AND N.Fornecedor = ? AND N.FilialForn = ? ORDER BY N.DataEmissao ASC", _
                                lNumIntItemNF, iTipoNFFiltro1, iTipoNFFiltro2, dtDataEmissao, iFilialEmpresa, sProduto, lFornEntTerc, iFilialFornEntTerc)
            If lErro <> AD_SQL_SUCESSO Then gError 211093
        
            lErro = Comando_BuscarProximo(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211094
            
            If lErro = AD_SQL_SEM_DADOS Then
            
                'Se não achou devolução posterior pega da mais recente a partir da própria data
                lErro = Comando_Executar(lComando3, "SELECT I.NumIntDoc FROM NFiscal AS N, ItensNFiscal AS I WHERE N.NumIntDoc = I.NumIntNF AND N.TipoNfiscal IN (?, ?) AND N.Status <> 7 AND N.DataEmissao <= ? AND N.FilialEmpresa = ? AND I.Produto = ? AND N.Fornecedor = ? AND N.FilialForn = ? ORDER BY N.DataEmissao ASC", _
                                    lNumIntItemNF, iTipoNFFiltro1, iTipoNFFiltro2, dtDataEmissao, iFilialEmpresa, sProduto, lFornEntTerc, iFilialFornEntTerc)
                If lErro <> AD_SQL_SUCESSO Then gError 211095
            
                lErro = Comando_BuscarProximo(lComando3)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211096
            
            End If
            
            'Se achou a devolução tem que considerar a quantidade como vendida e procurar o custo
            If lErro <> AD_SQL_SEM_DADOS Then
            
                sUMMovEst = String(STRING_UM_SIGLA, 0)
            
                lErro = Comando_Executar(lComando4, "SELECT SiglaUM, Quantidade, Custo FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = 1 AND NumIntDocOrigem = ?", _
                                    sUMMovEst, dQtdeMovEst, dCustoMovEst, lNumIntItemNF)
                If lErro <> AD_SQL_SUCESSO Then gError 211097
            
                lErro = Comando_BuscarProximo(lComando4)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211098
                
                lErro = CF("UM_Conversao_Trans", iClasseUM, sUMMovEst, sUMEst, dFatorCusto)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                dQuantidadeVendida = dQuantidadeVendida + dQuantidade * dFatorQtde
                dCustoTotal = dCustoTotal + (((dCustoMovEst / dQtdeMovEst) / dFatorCusto) * (dQuantidade * dFatorQtde))
                
                lErro = Comando_BuscarProximo(lComando4)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211098
        
            End If
            
        Else 'Entrega Futura
            iTipoNFFiltro1 = DOCINFO_NFISSRPVEF 'Remessa associada
            iTipoNFFiltro2 = DOCINFO_NFISSREF
            
            lErro = Comando_Executar(lComando5, "SELECT I.NumIntDoc FROM NFiscal AS N, ItensNFiscal AS I WHERE N.NumIntDoc = I.NumIntNF AND N.TipoNfiscal IN (?, ?) AND N.Status <> 7 AND N.NumIntNotaOriginal = ? AND I.Produto = ? ", _
                                lNumIntItemNF, iTipoNFFiltro1, iTipoNFFiltro2, lNumIntNF, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 208081
        
            lErro = Comando_BuscarPrimeiro(lComando5)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208082
            
            Do While lErro <> AD_SQL_SEM_DADOS
            
                lErro = Comando_Executar(lComando4, "SELECT SUM(Custo) FROM MovimentoEstoque WHERE TipoNumIntDocOrigem = 1 AND NumIntDocOrigem = ?", _
                                    dCustoMovEst, lNumIntItemNF)
                If lErro <> AD_SQL_SUCESSO Then gError 211097
            
                lErro = Comando_BuscarProximo(lComando4)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211098
                
                dCustoTotal = dCustoTotal + dCustoMovEst
            
                lErro = Comando_BuscarPrimeiro(lComando5)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 208082
            
            Loop
            
            dQuantidadeVendida = dQuantidadeVendida + dQuantidade * dFatorQtde
            
        End If
        
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211099

    Loop

    Produto_ObterRelMesSemEst = SUCESSO

    Exit Function

Erro_Produto_ObterRelMesSemEst:

    Produto_ObterRelMesSemEst = gErr

    Select Case gErr

        Case 211089
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211090 To 211099
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 211100)

    End Select

    Exit Function
    
End Function

Function Produto_ObterRelMesFat(ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, dValorFaturado As Double, dDesconto As Double, dQuantidadePedida As Double, dQuantidadeFaturada As Double, dQuantidadeDevolvida As Double, lComando As Long) As Long
'IMPORTANTE: especifica para classe ClassRelMesFat de apoio a execucao de relatorios
'Obtém as quantidades pedida e faturada, o desconto e valor total faturado por Produto entre datas
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long

On Error GoTo Erro_Produto_ObterRelMesFat

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then Error 43314
    End If
    
    If dtDataInicial <> DATA_NULA And dtDataFinal <> DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT SUM(ValorFaturado-ValorDevolvido), SUM(TotalDescontos), SUM(QuantPedida), SUM(QuantFaturada), SUM(QuantDevolvida) FROM SldDiaFat WHERE FilialEmpresa = ? AND Produto = ? AND Data >= ? AND Data <= ?", dValorFaturado, dDesconto, dQuantidadePedida, dQuantidadeFaturada, dQuantidadeDevolvida, iFilialEmpresa, sProduto, dtDataInicial, dtDataFinal)
                   
    ElseIf dtDataInicial <> DATA_NULA And dtDataFinal = DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT SUM(ValorFaturado-ValorDevolvido), SUM(TotalDescontos), SUM(QuantPedida), SUM(QuantFaturada), SUM(QuantDevolvida) FROM SldDiaFat WHERE FilialEmpresa = ? AND Produto = ? AND Data >= ?", dValorFaturado, dDesconto, dQuantidadePedida, dQuantidadeFaturada, dQuantidadeDevolvida, iFilialEmpresa, sProduto, dtDataInicial)
    
    ElseIf dtDataInicial = DATA_NULA And dtDataFinal <> DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT SUM(ValorFaturado-ValorDevolvido), SUM(TotalDescontos), SUM(QuantPedida), SUM(QuantFaturada), SUM(QuantDevolvida) FROM SldDiaFat WHERE FilialEmpresa = ? AND Produto = ? AND Data <= ?", dValorFaturado, dDesconto, dQuantidadePedida, dQuantidadeFaturada, dQuantidadeDevolvida, iFilialEmpresa, sProduto, dtDataFinal)
    
    Else
        lErro = Comando_Executar(lComando, "SELECT SUM(ValorFaturado-ValorDevolvido), SUM(TotalDescontos), SUM(QuantPedida), SUM(QuantFaturada), SUM(QuantDevolvida) FROM SldDiaFat WHERE FilialEmpresa = ? AND Produto = ?", dValorFaturado, dDesconto, dQuantidadePedida, dQuantidadeFaturada, dQuantidadeDevolvida, iFilialEmpresa, sProduto)
        
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 43315
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43316
    
    If lErro = AD_SQL_SEM_DADOS Then
        
        dValorFaturado = 0
        dDesconto = 0
        dQuantidadePedida = 0
        dQuantidadeFaturada = 0
        dQuantidadeDevolvida = 0

    End If

    Produto_ObterRelMesFat = SUCESSO
    
    Exit Function
    
Erro_Produto_ObterRelMesFat:

    Produto_ObterRelMesFat = Err
    
    Select Case Err

        Case 43314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 43315, 43316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAFAT_PROD", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152062)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Produto_ObterRelQtdSldIni(ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, dQuantInicial As Double, dValorInicial As Double, alComando() As Long, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
'IMPORTANTE: especifica para classe ClassRelMesEstAlm de apoio a execucao de relatorios
'Obtém a quantidade inicial e o valor inicial do Produto/Almoxarifado correspondentes até a data anterior a data inicial
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long
Dim dQuantInicial1 As Double
Dim dValorInicial1 As Double
Dim dtDataInicial1 As Date
Dim iIndice As Integer
Dim iAno As Integer
Dim iMes As Integer
Dim dQuantInicialTerc As Double
Dim dValorInicialTerc As Double

On Error GoTo Erro_Produto_ObterRelQtdSldIni

    For iIndice = 0 To 1
        If alComando(iIndice) = 0 Then
            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
            If alComando(iIndice) = 0 Then Error 55071
        End If
    Next
    
    'obter a quantidade e valor inicial na implantacao do produto no almoxarifado
    lErro = Comando_Executar(alComando(1), "SELECT QuantidadeInicial, SaldoInicial, DataInicial FROM EstoqueProduto WHERE Produto = ? AND Almoxarifado = ?", dQuantInicial, dValorInicial, dtDataInicial1, sProduto, iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then Error 55074
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55075
    
    If lErro = AD_SQL_SEM_DADOS Then Error 55076

    'se a data de inicializacao do produto no estoque for menor que a data final do periodo solicitado
    '==> os valores não estão disponíveis
    If dtDataFinal <> DATA_NULA And dtDataInicial1 > dtDataFinal Then
        dQuantInicial = -1
        dValorInicial = -1
    Else
        
        If dtDataInicial <> DATA_NULA Then
            'obter os totais movimentados até a data anterior a inicial
            lErro = Comando_Executar(alComando(0), "SELECT SUM(QuantEntrada - QuantEntConsig3 - QuantEntDemo3 - QuantEntConserto3 - QuantEntOutros3 - QuantEntBenef3) - SUM(QuantSaida - QuantSaiConsig3 - QuantSaiDemo3 - QuantSaiConserto3 - QuantSaiOutros3 - QuantSaiBenef3) , SUM(ValorEntrada - ValorEntConsig3 - ValorEntDemo3 - ValorEntConserto3 - ValorEntOutros3 - ValorEntBenef3) - SUM(ValorSaida - ValorSaiConsig3 - ValorSaiDemo3 - ValorSaiConserto3 - ValorSaiOutros3 - ValorSaiBenef3) FROM SldDiaEstAlm WHERE Produto = ? AND Almoxarifado = ? AND Data < ?", dQuantInicial1, dValorInicial1, sProduto, iAlmoxarifado, dtDataInicial)
            If lErro <> AD_SQL_SUCESSO Then Error 55072
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55073
        
            dQuantInicial = dQuantInicial + dQuantInicial1
            dValorInicial = dValorInicial + dValorInicial1
        
        End If
        
    End If

    iAno = Year(dtDataInicial)
    iMes = Month(dtDataInicial)
    
    'Se é um produto produzido e a ultima apuracao de custo de producao for menor do que ano/mes inicial do periodo solicitado ==> o saldo em valor não está disponível
    If iApropriacaoCusto = APROPR_CUSTO_REAL And (iAnoEstApurado < iAno Or (iAno = iAnoEstApurado And iMesEstApurado < iMes)) Then dValorInicial = -1
            
    Produto_ObterRelQtdSldIni = SUCESSO

    Exit Function

Erro_Produto_ObterRelQtdSldIni:

    Produto_ObterRelQtdSldIni = gErr

    Select Case gErr

        Case 93724

        Case 55071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 55072, 55073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM1", gErr, iAlmoxarifado, sProduto)

        Case 55074, 55075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", gErr, sProduto, iAlmoxarifado)

        Case 55076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", gErr, sProduto, iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152063)

    End Select

    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function Produto_ObterRelQtdSldDisp(ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, dQuantInicial As Double, dValorInicial As Double, alComando() As Long, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
'IMPORTANTE: especifica para classe ClassRelMesEstAlm de apoio a execucao de relatorios
'Obtém a quantidade inicial e o valor inicial do Produto/Almoxarifado correspondentes até a data anterior a data inicial
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long
Dim dQuantInicial1 As Double
Dim dValorInicial1 As Double
Dim dtDataInicial1 As Date
Dim iIndice As Integer
Dim iAno As Integer
Dim iMes As Integer
Dim dQuantInicialTerc As Double
Dim dValorInicialTerc As Double

On Error GoTo Erro_Produto_ObterRelQtdSldDisp

    For iIndice = 0 To 1
        If alComando(iIndice) = 0 Then
            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
            If alComando(iIndice) = 0 Then Error 55071
        End If
    Next
    
    'obter a quantidade e valor inicial na implantacao do produto no almoxarifado
    lErro = Comando_Executar(alComando(1), "SELECT QuantidadeInicial, SaldoInicial, DataInicial FROM EstoqueProduto WHERE Produto = ? AND Almoxarifado = ?", dQuantInicial, dValorInicial, dtDataInicial1, sProduto, iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then Error 55074
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55075
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        dQuantInicial = 0
        dValorInicial = 0

    Else

        'se a data de inicializacao do produto no estoque for menor que a data final do periodo solicitado
        '==> os valores não estão disponíveis
        If dtDataFinal <> DATA_NULA And dtDataInicial1 > dtDataFinal Then
            dQuantInicial = -1
            dValorInicial = -1
        Else
            
            If dtDataInicial <> DATA_NULA Then
                'obter os totais movimentados até a data anterior a inicial
                lErro = Comando_Executar(alComando(0), "SELECT SUM(QuantEntrada - QuantEntConsig3 - QuantEntDemo3 - QuantEntConserto3 - QuantEntOutros3 - QuantEntBenef3 - QuantEntRecIndl) - SUM(QuantSaida - QuantSaiConsig3 - QuantSaiDemo3 - QuantSaiConserto3 - QuantSaiOutros3 - QuantSaiBenef3 - QuantSaiRecIndl) , SUM(ValorEntrada - ValorEntConsig3 - ValorEntDemo3 - ValorEntConserto3 - ValorEntOutros3 - ValorEntBenef3 - ValorEntRecIndl) - SUM(ValorSaida - ValorSaiConsig3 - ValorSaiDemo3 - ValorSaiConserto3 - ValorSaiOutros3 - ValorSaiBenef3 - ValorSaiRecIndl) FROM SldDiaEstAlm WHERE Produto = ? AND Almoxarifado = ? AND Data < ?", dQuantInicial1, dValorInicial1, sProduto, iAlmoxarifado, dtDataInicial)
                If lErro <> AD_SQL_SUCESSO Then Error 55072
        
                lErro = Comando_BuscarPrimeiro(alComando(0))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55073
            
                dQuantInicial = dQuantInicial + dQuantInicial1
                dValorInicial = dValorInicial + dValorInicial1
            
            End If
            
        End If
    
        iAno = Year(dtDataInicial)
        iMes = Month(dtDataInicial)
        
        'Se é um produto produzido e a ultima apuracao de custo de producao for menor do que ano/mes inicial do periodo solicitado ==> o saldo em valor não está disponível
        If iApropriacaoCusto = APROPR_CUSTO_REAL And (iAnoEstApurado < iAno Or (iAno = iAnoEstApurado And iMesEstApurado < iMes)) Then dValorInicial = -1
            
    End If
            
    Produto_ObterRelQtdSldDisp = SUCESSO

    Exit Function

Erro_Produto_ObterRelQtdSldDisp:

    Produto_ObterRelQtdSldDisp = gErr

    Select Case gErr

        Case 93724

        Case 55071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 55072, 55073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM1", gErr, iAlmoxarifado, sProduto)

        Case 55074, 55075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", gErr, sProduto, iAlmoxarifado)

        Case 55076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", gErr, sProduto, iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152063)

    End Select

    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function Produto_ObterDataUltEmpenho(ByVal sProduto As String, dtDataUltEmpenho As Date, lComando As Long) As Long
'obtem a data do ultimo empenho de um produto, se nao houver nenhum empenho retornar DATA_NULA
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long

On Error GoTo Erro_Produto_ObterDataUltEmpenho

    If lComando = 0 Then
    
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then Error 41555
                
    End If
    
    If giFilialEmpresa = EMPRESA_TODA Then
        lErro = Comando_Executar(lComando, "SELECT MAX(Data) FROM Empenho WHERE Produto = ?", dtDataUltEmpenho, sProduto)
    Else
        lErro = Comando_Executar(lComando, "SELECT MAX(Data) FROM Empenho WHERE Produto = ? AND FilialEmpresa = ?", dtDataUltEmpenho, sProduto, giFilialEmpresa)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 41556
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41557
    
    If lErro = AD_SQL_SEM_DADOS Then dtDataUltEmpenho = DATA_NULA
    
    Produto_ObterDataUltEmpenho = SUCESSO
    
    Exit Function
    
Erro_Produto_ObterDataUltEmpenho:

    Produto_ObterDataUltEmpenho = Err
    
    Select Case Err

        Case 41555
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 41556, 41557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPENHO_PROD", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152064)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TabelaDePrecoItens_Le_DatasVigentes(ByVal objTabelaPrecoItem As ClassTabelaPrecoItem, ByVal colDatasVigentes As Collection) As Long
'Lê todas as Datas De Vigência colocando em ordem na Colecao
'Parametros: Produto, Tabela, FilialEmpresa

Dim lErro As Long
Dim lComando As Long
Dim dtDataVigencia As Date

On Error GoTo Erro_TabelaDePrecoItens_Le_DatasVigentes
    
    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58480

    'Lê todas as datas de Vigentes
    lErro = Comando_Executar(lComando, "SELECT DataVigencia FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodProduto = ? AND CodTabela = ? ORDER BY DataVigencia DESC", dtDataVigencia, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.sCodProduto, objTabelaPrecoItem.iCodTabela)
    If lErro <> AD_SQL_SUCESSO Then Error 58481

    'Busca a Primeira data
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58482

    'Não encontrou a data --- > Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 58483
    
    Do While lErro = AD_SQL_SUCESSO
    
        colDatasVigentes.Add dtDataVigencia
    
        'Buscar Proxima Data
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58484
        
    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    TabelaDePrecoItens_Le_DatasVigentes = SUCESSO
    
    Exit Function
    
Erro_TabelaDePrecoItens_Le_DatasVigentes:
    
    TabelaDePrecoItens_Le_DatasVigentes = Err
    
    Select Case Err
        
        Case 58480
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 58481, 58482, 58484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS", Err)
        
        Case 58483 'Tratado na Rotina Chamadora
                
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152065)
    
    End Select
    
    'Fechamento comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TabelaPrecoItem_Le_DataVigente(ByVal objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
'Lê a tabela de preço a partir da chave fornecida em objTabelaPreco. Retorna os dados em objTabelaPreco

Dim lErro As Long
Dim lComando As Long
Dim tTabelaPrecoItem As typeTabelaPrecoItem

On Error GoTo Erro_TabelaPrecoItem_Le_DataVigente

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58500

    tTabelaPrecoItem.sObservacao = String(STRING_TABELA_OBSERVACAO, 0)
    
    'Pesquisa no BD Ítens de Tabela de Preço
    lErro = Comando_Executar(lComando, "SELECT Preco, Observacao, PercDesconto,PrecoComDesconto FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto = ? AND DataVigencia = ? ", tTabelaPrecoItem.dPreco, tTabelaPrecoItem.sObservacao, tTabelaPrecoItem.dPercDesconto, tTabelaPrecoItem.dPrecoComDesconto, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto, objTabelaPrecoItem.dtDataVigencia)
    If lErro <> AD_SQL_SUCESSO Then Error 58501

    'Verifica se Ítem existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58502

    'Não encontrou o Ítem
    If lErro = AD_SQL_SEM_DADOS Then Error 58503

    'Preenche objTabelaPrecoItem
    objTabelaPrecoItem.dPreco = tTabelaPrecoItem.dPreco
    objTabelaPrecoItem.sObservacao = tTabelaPrecoItem.sObservacao
    objTabelaPrecoItem.dPercDesconto = tTabelaPrecoItem.dPercDesconto
    objTabelaPrecoItem.dPrecoComDesconto = tTabelaPrecoItem.dPrecoComDesconto

    'Fechamento comando
    Call Comando_Fechar(lComando)

    TabelaPrecoItem_Le_DataVigente = SUCESSO

    Exit Function

Erro_TabelaPrecoItem_Le_DataVigente:

    TabelaPrecoItem_Le_DataVigente = Err

    Select Case Err

        Case 58500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 58501, 58502
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS1", Err, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto)

        Case 58503
            'Não encontrou a Ítem da Tabela de Preço. Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152066)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_SituacaoTrib(sSituacaoTrib As String, ByVal sProduto As String, ByVal iTipoTribICMS As Integer, lComando As Long) As Long
'Le a Situacao Tributaria
'Concatena a OrigemMercadoria da Tabela de Produtos com TipoTribCst da Tabela de TiposTribICMS
'IMPORTANTE : O lComando somente é aberto se vier aberto e permanecerá fechado

Dim lErro As Long
Dim iOrigemMercadoria As Integer
Dim iTipoTribCST As Integer

On Error GoTo Erro_Produto_Le_SituacaoTrib

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then Error 58575
    End If
    
    'Obtem dados do cadastro de produtos
    lErro = Comando_Executar(lComando, "SELECT OrigemMercadoria, TipoTribCST FROM Produtos, TiposTribICMS WHERE Produtos.Codigo = ? AND TiposTribICMS.Tipo = ?", iOrigemMercadoria, iTipoTribCST, sProduto, iTipoTribICMS)
    If lErro <> AD_SQL_SUCESSO Then Error 58576

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58577
    
    'Se não encontrar devolve em branco
    If lErro = AD_SQL_SEM_DADOS Then
        sSituacaoTrib = ""
    Else
        'Se não concatena
        sSituacaoTrib = CStr(iOrigemMercadoria) & Format(iTipoTribCST, "00")
    End If
    
    Produto_Le_SituacaoTrib = SUCESSO
    
    Exit Function
    
Erro_Produto_Le_SituacaoTrib:

    Produto_Le_SituacaoTrib = Err
    
    Select Case Err
            
        Case 58575
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 58576, 58577
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152067)
    
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoFiliais_Le(ByVal sProduto As String, ByVal colProdutoFilial As Collection) As Long
'Passa o Produto e obtem uma coleção do Produto para todas as Filiais

Dim lErro As Long
Dim lComando As Long
Dim tProdutoFilial As typeProdutoFilial
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_ProdutoFiliais_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61124
    
    tProdutoFilial.sClasseABC = String(STRING_PRODUTOFILIAL_CLASSEABC, 0)
    tProdutoFilial.sSituacaoTribECF = String(STRING_PRODUTOFILIAL_SITUACAOTRIBECF, 0)
    tProdutoFilial.sICMSAliquota = String(STRING_PRODUTOFILIAL_ICMS, 0)

    'Pesquisa no BD ProdutoFilial
    With tProdutoFilial
        lErro = Comando_Executar(lComando, "SELECT Almoxarifado, Fornecedor, VisibilidadeAlmoxarifados, EstoqueSeguranca, EstoqueMaximo, PontoPedido, ClasseABC, LoteEconomico, IntRessup, TabelaPreco, QuantPedida, SituacaoTribECF, ICMSAliquota FROM ProdutosFilial WHERE Produto = ?", _
            .iAlmoxarifado, .lFornecedor, .iVisibilidadeAlmoxarifados, .dEstoqueSeguranca, .dEstoqueMaximo, .dPontoPedido, .sClasseABC, .dLoteEconomico, .iIntRessup, .iTabelaPreco, .dQuantPedida, .sSituacaoTribECF, .sICMSAliquota, sProduto)
        If lErro <> AD_SQL_SUCESSO Then Error 61125
    End With

    'Tenta selecionar Produto
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61126

    'Não encontrou ProdutoFilial
    If lErro = AD_SQL_SEM_DADOS Then Error 61127
    
    Do While lErro = AD_SQL_SUCESSO
        
        Set objProdutoFilial = New ClassProdutoFilial
        
        'Preenche objProdutoFilial
        With tProdutoFilial
            objProdutoFilial.iAlmoxarifado = .iAlmoxarifado
            objProdutoFilial.lFornecedor = .lFornecedor
            objProdutoFilial.iVisibilidadeAlmoxarifados = .iVisibilidadeAlmoxarifados
            objProdutoFilial.dEstoqueSeguranca = .dEstoqueSeguranca
            objProdutoFilial.dEstoqueMaximo = .dEstoqueMaximo
            objProdutoFilial.dPontoPedido = .dPontoPedido
            objProdutoFilial.sClasseABC = .sClasseABC
            objProdutoFilial.dLoteEconomico = .dLoteEconomico
            objProdutoFilial.iIntRessup = .iIntRessup
            objProdutoFilial.iTabelaPreco = .iTabelaPreco
            objProdutoFilial.dQuantPedida = .dQuantPedida
            objProdutoFilial.sSituacaoTribECF = .sSituacaoTribECF
            objProdutoFilial.sICMSAliquota = .sICMSAliquota
            
        End With
    
        colProdutoFilial.Add objProdutoFilial
        
        'Tenta selecionar ProdutoFilial
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61128
    
    Loop
        
        
    'Fechamento comando
    Call Comando_Fechar(lComando)

    ProdutoFiliais_Le = SUCESSO

    Exit Function

Erro_ProdutoFiliais_Le:

    ProdutoFiliais_Le = Err

    Select Case Err

        Case 61124
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61125, 61126, 61128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err, objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)

        Case 61127
            'Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152068)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Almoxarifado_Le_Total(iTotalAlmox As Integer) As Long
'Le a quantidade de Almoxarifados já gravados na Tabela de Almoxarifados

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Almoxarifado_Le_Total

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61185

    'Seleciona a quantidade de Almoxarifados Gravados
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM Almoxarifado", iTotalAlmox)
    If lErro <> AD_SQL_SUCESSO Then Error 61186
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61187
            
    'Fecha comando
    Call Comando_Fechar(lComando)
        
    Almoxarifado_Le_Total = SUCESSO
    
    Exit Function
    
Erro_Almoxarifado_Le_Total:
    
    Almoxarifado_Le_Total = Err
    
    Select Case Err
    
        Case 61185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61186, 61187
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO2", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152069)
            
    End Select
        
    'Fecha comando
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Function ClasseUM_Le_Total(iTotalClasseUM As Integer) As Long
'Le a quantidade de ClasseUM já gravados na Tabela de ClasseUM

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_ClasseUM_Le_Total

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61188

    'Seleciona a quantidade de ClasseUM gravados
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM ClasseUM", iTotalClasseUM)
    If lErro <> AD_SQL_SUCESSO Then Error 61189
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61190
            
    'Fecha comando
    Call Comando_Fechar(lComando)
        
    ClasseUM_Le_Total = SUCESSO
    
    Exit Function
    
Erro_ClasseUM_Le_Total:
    
    ClasseUM_Le_Total = Err
    
    Select Case Err
    
        Case 61188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61189, 61190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSEUM1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152070)
            
    End Select
        
    'Fecha comando
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

'Function FornecedorProduto_Le_Total(objFornecedorProduto As ClassFornecedorProduto, iTotalFornProduto As Integer) As Long
''Le a quantidade de Fornecedores já gravados na Tabela de FornecedorProduto
'
'Dim lErro As Long
'Dim lComando As Long
'
'On Error GoTo Erro_FornecedorProduto_Le_Total
'
'    'Abre comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then Error 61226
'
'    'Seleciona a quantidade de Fornecedor para o Produto passado
'    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM FornecedorProduto WHERE Produto = ?", iTotalFornProduto, objFornecedorProduto.sProduto)
'    If lErro <> AD_SQL_SUCESSO Then Error 61227
'
'    'Tenta ler
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61228
'
'    'Fecha comando
'    Call Comando_Fechar(lComando)
'
'    FornecedorProduto_Le_Total = SUCESSO
'
'    Exit Function
'
'Erro_FornecedorProduto_Le_Total:
'
'    FornecedorProduto_Le_Total = Err
'
'    Select Case Err
'
'        Case 61226
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 61227, 61228
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTO", Err)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152071)
'
'    End Select
'
'    'Fecha comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Function EstoqueMes_Le_Mes1(ByVal objEstoqueMes As ClassEstoqueMes) As Long
'Le mes e ano em EstoqueMes sendo passado como parametro filial empresa e o fechamento como ABERTO

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueMes As typeEstoqueMes

On Error GoTo Erro_EstoqueMes_Le_Mes1
    
    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60858
    
    'Le Ano e Mes
    'tulio 26/07/02 -> adicionei descs no order by
    lErro = Comando_Executar(lComando, "SELECT Ano, Mes  FROM EstoqueMes WHERE FilialEmpresa = ? AND Fechamento = ? ORDER BY Ano desc, Mes desc", tEstoqueMes.iAno, tEstoqueMes.iMes, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iFechamento)
    If lErro <> AD_SQL_SUCESSO Then Error 60859

    'Busca o primeiro ano e mes que estiverem aberto ( so podera haver um mes aberto
    '??? por tulio 26/07/02 -> so pode haver um mesmo???????????
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60860

    If lErro = AD_SQL_SEM_DADOS Then Error 60861

    'preenche o objeto com o mes e o ano
    objEstoqueMes.iMes = tEstoqueMes.iMes
    objEstoqueMes.iAno = tEstoqueMes.iAno

    'Fechamento comando
    Call Comando_Fechar(lComando)

    EstoqueMes_Le_Mes1 = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_Mes1:

    EstoqueMes_Le_Mes1 = Err

    Select Case Err

        Case 60858
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 60859, 60860
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", Err)

        Case 60861

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152072)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovimentoEstoque_Le_Ccl(ByVal sCcl As String) As Long
'Verifica se existe algum centro de custo com o centro de custo passado como parametro
    
Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long

On Error GoTo Erro_MovimentoEstoque_Le_Ccl

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60865

    'pesquisa um movimento com o centro de custo em questao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM MovimentoEstoque WHERE Ccl = ?", lNumIntDoc, sCcl)
    If lErro <> AD_SQL_SUCESSO Then Error 60866

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60867

    'se o movimento não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then Error 60868

    Call Comando_Fechar(lComando)

    MovimentoEstoque_Le_Ccl = SUCESSO
    
    Exit Function

Erro_MovimentoEstoque_Le_Ccl:

    MovimentoEstoque_Le_Ccl = Err

    Select Case Err

        Case 60865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 60866, 60867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE4", Err, sCcl)

        Case 60868
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152073)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function EstoqueProduto_Le_Todos_Almoxarifados_Filial(ByVal objEstoqueProduto As ClassEstoqueProduto, ByVal iFilialEmpresa As Integer) As Long
'Lê a soma das Quantidades de todos Almoxarifados da Filial que pertence ao Produto passado em objEstoqueProduto

Dim lErro As Long
Dim lComando As Long
Dim tEstoqueProduto As typeEstoqueProduto
Dim dOp As Double

On Error GoTo Erro_EstoqueProduto_Le_Todos_Almoxarifados_Filial

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 64011
    
    lErro = Comando_Executar(lComando, "SELECT SUM(QuantDispNossa), SUM(QuantReservada), MAX(DataInventario), SUM(QuantidadeInicial), SUM(SaldoInicial), SUM(QuantEmpenhada), SUM(QuantPedida), SUM(QuantRecIndl), SUM(QuantIndOutras), SUM(QuantDefeituosa), SUM(QuantConsig3), SUM(QuantConsig), SUM(QuantDemo3), SUM(QuantDemo), SUM(QuantConserto3), SUM(QuantConserto), SUM(QuantOutras3), SUM(QuantOutras), SUM(QuantBenef3), SUM(QuantBenef), SUM(QuantOP), SUM(ValorConsig3), SUM(ValorConsig), SUM(ValorDemo3), SUM(ValorDemo), SUM(ValorConserto3), SUM(ValorConserto), SUM(ValorOutras3), SUM(ValorOutras), SUM(ValorBenef3), SUM(ValorBenef) FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Almoxarifado =  Almoxarifado.Codigo AND EstoqueProduto.Produto = ? AND Almoxarifado.FilialEmpresa = ?", _
        tEstoqueProduto.dQuantDispNossa, tEstoqueProduto.dQuantReservada, tEstoqueProduto.dtDataInventario, tEstoqueProduto.dQuantidadeInicial, tEstoqueProduto.dSaldoInicial, tEstoqueProduto.dQuantEmpenhada, tEstoqueProduto.dQuantPedido, tEstoqueProduto.dQuantRecIndl, tEstoqueProduto.dQuantInd, tEstoqueProduto.dQuantDefeituosa, tEstoqueProduto.dQuantConsig3, tEstoqueProduto.dQuantConsig, tEstoqueProduto.dQuantDemo3, tEstoqueProduto.dQuantDemo, tEstoqueProduto.dQuantConserto3, tEstoqueProduto.dQuantConserto, tEstoqueProduto.dQuantOutras3, tEstoqueProduto.dQuantOutras, tEstoqueProduto.dQuantBenef3, tEstoqueProduto.dQuantBenef, tEstoqueProduto.dQuantOP, _
        tEstoqueProduto.dValorConsig3, tEstoqueProduto.dValorConsig, tEstoqueProduto.dValorDemo3, tEstoqueProduto.dValorDemo, tEstoqueProduto.dValorConserto3, tEstoqueProduto.dValorConserto, tEstoqueProduto.dValorOutras3, tEstoqueProduto.dValorOutras, tEstoqueProduto.dValorBenef3, tEstoqueProduto.dValorBenef, objEstoqueProduto.sProduto, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 64012

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64013

    'Não encontrou ---> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 64014
    
    objEstoqueProduto.dQuantDispNossa = tEstoqueProduto.dQuantDispNossa
    objEstoqueProduto.dQuantReservada = tEstoqueProduto.dQuantReservada
    objEstoqueProduto.dtDataInventario = tEstoqueProduto.dtDataInventario
    objEstoqueProduto.dQuantidadeInicial = tEstoqueProduto.dQuantidadeInicial
    objEstoqueProduto.dSaldoInicial = tEstoqueProduto.dSaldoInicial
    objEstoqueProduto.dQuantEmpenhada = tEstoqueProduto.dQuantEmpenhada
    objEstoqueProduto.dQuantPedido = tEstoqueProduto.dQuantPedido
    objEstoqueProduto.dQuantRecIndl = tEstoqueProduto.dQuantRecIndl
    objEstoqueProduto.dQuantInd = tEstoqueProduto.dQuantInd
    objEstoqueProduto.dQuantDefeituosa = tEstoqueProduto.dQuantDefeituosa
    objEstoqueProduto.dQuantConsig3 = tEstoqueProduto.dQuantConsig3
    objEstoqueProduto.dQuantConsig = tEstoqueProduto.dQuantConsig
    objEstoqueProduto.dQuantDemo3 = tEstoqueProduto.dQuantDemo3
    objEstoqueProduto.dQuantDemo = tEstoqueProduto.dQuantDemo
    objEstoqueProduto.dQuantConserto3 = tEstoqueProduto.dQuantConserto3
    objEstoqueProduto.dQuantConserto = tEstoqueProduto.dQuantConserto
    objEstoqueProduto.dQuantOutras3 = tEstoqueProduto.dQuantOutras3
    objEstoqueProduto.dQuantOutras = tEstoqueProduto.dQuantOutras
    objEstoqueProduto.dQuantBenef3 = tEstoqueProduto.dQuantBenef3
    objEstoqueProduto.dQuantBenef = tEstoqueProduto.dQuantBenef
    objEstoqueProduto.dQuantOP = tEstoqueProduto.dQuantOP
    objEstoqueProduto.dValorConsig3 = tEstoqueProduto.dValorConsig3
    objEstoqueProduto.dValorConsig = tEstoqueProduto.dValorConsig
    objEstoqueProduto.dValorDemo3 = tEstoqueProduto.dValorDemo3
    objEstoqueProduto.dValorDemo = tEstoqueProduto.dValorDemo
    objEstoqueProduto.dValorConserto3 = tEstoqueProduto.dValorConserto3
    objEstoqueProduto.dValorConserto = tEstoqueProduto.dValorConserto
    objEstoqueProduto.dValorOutras3 = tEstoqueProduto.dValorOutras3
    objEstoqueProduto.dValorOutras = tEstoqueProduto.dValorOutras
    objEstoqueProduto.dValorBenef3 = tEstoqueProduto.dValorBenef3
    objEstoqueProduto.dValorBenef = tEstoqueProduto.dValorBenef
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    EstoqueProduto_Le_Todos_Almoxarifados_Filial = SUCESSO

    Exit Function

Erro_EstoqueProduto_Le_Todos_Almoxarifados_Filial:

    EstoqueProduto_Le_Todos_Almoxarifados_Filial = Err

    Select Case Err
        
        Case 64011
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 64012, 64013
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO1", Err, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 64014 'Não encontrou --> Tratar na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152074)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function SldMesEst_Le_QuantidadeCusto(ByVal lComando As Long, ByVal objSldMesEst As ClassSldMesEst) As Long
'le os dados de SldMesEst relativos ao Ano, FilialEmpresa e Produto passados como parametro
    
Dim tSldMesEst As typeSldMesEst
Dim lErro As Long
Dim iIndice As Integer
    
    
On Error GoTo Erro_SldMesEst_Le_QuantidadeCusto
    
    'seleciona o registro de SldMesEst relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialCusto, ValorInicialCusto, SaldoQuantCusto1, SaldoValorCusto1, SaldoQuantCusto2, SaldoValorCusto2, SaldoQuantCusto3, SaldoValorCusto3, SaldoQuantCusto4, SaldoValorCusto4, SaldoQuantCusto5, SaldoValorCusto5, SaldoQuantCusto6, SaldoValorCusto6, SaldoQuantCusto7, SaldoValorCusto7, SaldoQuantCusto8, SaldoValorCusto8, SaldoQuantCusto9, SaldoValorCusto9, SaldoQuantCusto10, SaldoValorCusto10, SaldoQuantCusto11, SaldoValorCusto11, SaldoQuantCusto12, SaldoValorCusto12 FROM SldMesEst WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst.dQuantInicialCusto, tSldMesEst.dValorInicialCusto, tSldMesEst.adSaldoQuantCusto(1), tSldMesEst.adSaldoValorCusto(1), tSldMesEst.adSaldoQuantCusto(2), tSldMesEst.adSaldoValorCusto(2), tSldMesEst.adSaldoQuantCusto(3), tSldMesEst.adSaldoValorCusto(3), tSldMesEst.adSaldoQuantCusto(4), tSldMesEst.adSaldoValorCusto(4), tSldMesEst.adSaldoQuantCusto(5), tSldMesEst.adSaldoValorCusto(5), tSldMesEst.adSaldoQuantCusto(6), tSldMesEst.adSaldoValorCusto(6), tSldMesEst.adSaldoQuantCusto(7), tSldMesEst.adSaldoValorCusto(7), tSldMesEst.adSaldoQuantCusto(8), tSldMesEst.adSaldoValorCusto(8), tSldMesEst.adSaldoQuantCusto(9), tSldMesEst.adSaldoValorCusto(9), tSldMesEst.adSaldoQuantCusto(10), tSldMesEst.adSaldoValorCusto(10), _
                              tSldMesEst.adSaldoQuantCusto(11), tSldMesEst.adSaldoValorCusto(11), tSldMesEst.adSaldoQuantCusto(12), tSldMesEst.adSaldoValorCusto(12), objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 36484
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36485
    
    objSldMesEst.dQuantInicialCusto = tSldMesEst.dQuantInicialCusto
    objSldMesEst.dValorInicialCusto = tSldMesEst.dValorInicialCusto
    
    For iIndice = 1 To 12
    
        objSldMesEst.dSaldoQuantCusto(iIndice) = tSldMesEst.adSaldoQuantCusto(iIndice)
        objSldMesEst.dSaldoValorCusto(iIndice) = tSldMesEst.adSaldoValorCusto(iIndice)
        
    Next
    
    SldMesEst_Le_QuantidadeCusto = SUCESSO

    Exit Function

Erro_SldMesEst_Le_QuantidadeCusto:
    
    SldMesEst_Le_QuantidadeCusto = Err
    
    Select Case Err
    
        Case 36484, 36485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", Err, objSldMesEst.iAno, objSldMesEst.iFilialEmpresa, objSldMesEst.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152075)
        
    End Select
    
    Exit Function

End Function

Function SldMesEstAlm_Le_QuantidadeCusto(ByVal lComando As Long, ByVal objSldMesEstAlm As ClassSldMesEstAlm) As Long
'le os dados de SldMesEst relativos ao Ano, FilialEmpresa e Produto passados como parametro
    
Dim tSldMesEstAlm As typeSldMesEstAlm
Dim lErro As Long
Dim iIndice As Integer
    
On Error GoTo Erro_SldMesEstAlm_Le_QuantidadeCusto
    
    'seleciona o registro de SldMesEstAlm relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialCusto, ValorInicialCusto, SaldoQuantCusto1, SaldoValorCusto1, SaldoQuantCusto2, SaldoValorCusto2, SaldoQuantCusto3, SaldoValorCusto3, SaldoQuantCusto4, SaldoValorCusto4, SaldoQuantCusto5, SaldoValorCusto5, SaldoQuantCusto6, SaldoValorCusto6, SaldoQuantCusto7, SaldoValorCusto7, SaldoQuantCusto8, SaldoValorCusto8, SaldoQuantCusto9, SaldoValorCusto9, SaldoQuantCusto10, SaldoValorCusto10, SaldoQuantCusto11, SaldoValorCusto11, SaldoQuantCusto12, SaldoValorCusto12 FROM SldMesEstAlm WHERE Ano=? AND Almoxarifado=? AND Produto=?", _
                              tSldMesEstAlm.dQuantInicialCusto, tSldMesEstAlm.dValorInicialCusto, tSldMesEstAlm.adSaldoQuantCusto(1), tSldMesEstAlm.adSaldoValorCusto(1), tSldMesEstAlm.adSaldoQuantCusto(2), tSldMesEstAlm.adSaldoValorCusto(2), tSldMesEstAlm.adSaldoQuantCusto(3), tSldMesEstAlm.adSaldoValorCusto(3), tSldMesEstAlm.adSaldoQuantCusto(4), tSldMesEstAlm.adSaldoValorCusto(4), tSldMesEstAlm.adSaldoQuantCusto(5), tSldMesEstAlm.adSaldoValorCusto(5), tSldMesEstAlm.adSaldoQuantCusto(6), tSldMesEstAlm.adSaldoValorCusto(6), tSldMesEstAlm.adSaldoQuantCusto(7), tSldMesEstAlm.adSaldoValorCusto(7), tSldMesEstAlm.adSaldoQuantCusto(8), tSldMesEstAlm.adSaldoValorCusto(8), tSldMesEstAlm.adSaldoQuantCusto(9), tSldMesEstAlm.adSaldoValorCusto(9), tSldMesEstAlm.adSaldoQuantCusto(10), tSldMesEstAlm.adSaldoValorCusto(10), _
                              tSldMesEstAlm.adSaldoQuantCusto(11), tSldMesEstAlm.adSaldoValorCusto(11), tSldMesEstAlm.adSaldoQuantCusto(12), tSldMesEstAlm.adSaldoValorCusto(12), objSldMesEstAlm.iAno, objSldMesEstAlm.iAlmoxarifado, objSldMesEstAlm.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 83587
    
    'le o SldMesEstAlm
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83588
    
    objSldMesEstAlm.dQuantInicialCusto = tSldMesEstAlm.dQuantInicialCusto
    objSldMesEstAlm.dValorInicialCusto = tSldMesEstAlm.dValorInicialCusto
    
    For iIndice = 1 To 12
    
        objSldMesEstAlm.dSaldoQuantCusto(iIndice) = tSldMesEstAlm.adSaldoQuantCusto(iIndice)
        objSldMesEstAlm.dSaldoValorCusto(iIndice) = tSldMesEstAlm.adSaldoValorCusto(iIndice)
        
    Next
    
    SldMesEstAlm_Le_QuantidadeCusto = SUCESSO

    Exit Function

Erro_SldMesEstAlm_Le_QuantidadeCusto:
    
    SldMesEstAlm_Le_QuantidadeCusto = gErr
    
    Select Case gErr
    
        Case 83587, 83588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM", gErr, objSldMesEstAlm.iAno, objSldMesEstAlm.iAlmoxarifado, objSldMesEstAlm.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152076)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantConserto(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dQuantConsertoAtual As Double) As Long
'calcula a quantidade de material em conserto nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double

On Error GoTo Erro_SldMesEst2_Le_QuantConserto

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConserto, SaldoQuantConserto1, SaldoQuantConserto2, SaldoQuantConserto3, SaldoQuantConserto4, SaldoQuantConserto5, SaldoQuantConserto6, SaldoQuantConserto7, SaldoQuantConserto8, SaldoQuantConserto9, SaldoQuantConserto10, SaldoQuantConserto11, SaldoQuantConserto12 FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialConserto, tSldMesEst2.adSaldoQuantConserto(1), tSldMesEst2.adSaldoQuantConserto(2), tSldMesEst2.adSaldoQuantConserto(3), tSldMesEst2.adSaldoQuantConserto(4), tSldMesEst2.adSaldoQuantConserto(5), tSldMesEst2.adSaldoQuantConserto(6), tSldMesEst2.adSaldoQuantConserto(7), tSldMesEst2.adSaldoQuantConserto(8), tSldMesEst2.adSaldoQuantConserto(9), _
                              tSldMesEst2.adSaldoQuantConserto(10), tSldMesEst2.adSaldoQuantConserto(11), tSldMesEst2.adSaldoQuantConserto(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 60967
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 60968
    
    dQuant = tSldMesEst2.dQuantInicialConserto
    
    For iIndice = 1 To iMes
    
        dQuant = dQuant + tSldMesEst2.adSaldoQuantConserto(iIndice)
        
    Next
    
    dQuantConsertoAtual = dQuant
    
    SldMesEst2_Le_QuantConserto = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_QuantConserto:
    
    SldMesEst2_Le_QuantConserto = gErr
    
    Select Case gErr
    
        Case 60967, 60968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152077)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantDemo(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dQuantDemoAtual As Double) As Long
'calcula a quantidade de material em Demonstração nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double

On Error GoTo Erro_SldMesEst2_Le_QuantDemo

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialDemo, SaldoQuantDemo1, SaldoQuantDemo2, SaldoQuantDemo3, SaldoQuantDemo4, SaldoQuantDemo5, SaldoQuantDemo6, SaldoQuantDemo7, SaldoQuantDemo8, SaldoQuantDemo9, SaldoQuantDemo10, SaldoQuantDemo11, SaldoQuantDemo12 FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialDemo, tSldMesEst2.adSaldoQuantDemo(1), tSldMesEst2.adSaldoQuantDemo(2), tSldMesEst2.adSaldoQuantDemo(3), tSldMesEst2.adSaldoQuantDemo(4), tSldMesEst2.adSaldoQuantDemo(5), tSldMesEst2.adSaldoQuantDemo(6), tSldMesEst2.adSaldoQuantDemo(7), tSldMesEst2.adSaldoQuantDemo(8), tSldMesEst2.adSaldoQuantDemo(9), _
                              tSldMesEst2.adSaldoQuantDemo(10), tSldMesEst2.adSaldoQuantDemo(11), tSldMesEst2.adSaldoQuantDemo(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 60967
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 60968
    
    dQuant = tSldMesEst2.dQuantInicialDemo
    
    For iIndice = 1 To iMes
    
        dQuant = dQuant + tSldMesEst2.adSaldoQuantDemo(iIndice)
        
    Next
    
    dQuantDemoAtual = dQuant
    
    SldMesEst2_Le_QuantDemo = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_QuantDemo:
    
    SldMesEst2_Le_QuantDemo = gErr
    
    Select Case gErr
    
        Case 60967, 60968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152078)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantOutros(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dQuantOutrosAtual As Double) As Long
'calcula a quantidade de Outros materiais nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double

On Error GoTo Erro_SldMesEst2_Le_QuantOutros

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialOutros, SaldoQuantOutros1, SaldoQuantOutros2, SaldoQuantOutros3, SaldoQuantOutros4, SaldoQuantOutros5, SaldoQuantOutros6, SaldoQuantOutros7, SaldoQuantOutros8, SaldoQuantOutros9, SaldoQuantOutros10, SaldoQuantOutros11, SaldoQuantOutros12 FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialOutros, tSldMesEst2.adSaldoQuantOutros(1), tSldMesEst2.adSaldoQuantOutros(2), tSldMesEst2.adSaldoQuantOutros(3), tSldMesEst2.adSaldoQuantOutros(4), tSldMesEst2.adSaldoQuantOutros(5), tSldMesEst2.adSaldoQuantOutros(6), tSldMesEst2.adSaldoQuantOutros(7), tSldMesEst2.adSaldoQuantOutros(8), tSldMesEst2.adSaldoQuantOutros(9), _
                              tSldMesEst2.adSaldoQuantOutros(10), tSldMesEst2.adSaldoQuantOutros(11), tSldMesEst2.adSaldoQuantOutros(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 60967
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 60968
    
    dQuant = tSldMesEst2.dQuantInicialOutros
    
    For iIndice = 1 To iMes
    
        dQuant = dQuant + tSldMesEst2.adSaldoQuantOutros(iIndice)
        
    Next
    
    dQuantOutrosAtual = dQuant
    
    SldMesEst2_Le_QuantOutros = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_QuantOutros:
    
    SldMesEst2_Le_QuantOutros = gErr
    
    Select Case gErr
    
        Case 60967, 60968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152079)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantBenef(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dQuantBenefAtual As Double) As Long
'calcula a quantidade de material em Beneficiamento nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double

On Error GoTo Erro_SldMesEst2_Le_QuantBenef

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialBenef, SaldoQuantBenef1, SaldoQuantBenef2, SaldoQuantBenef3, SaldoQuantBenef4, SaldoQuantBenef5, SaldoQuantBenef6, SaldoQuantBenef7, SaldoQuantBenef8, SaldoQuantBenef9, SaldoQuantBenef10, SaldoQuantBenef11, SaldoQuantBenef12 FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialBenef, tSldMesEst2.adSaldoQuantBenef(1), tSldMesEst2.adSaldoQuantBenef(2), tSldMesEst2.adSaldoQuantBenef(3), tSldMesEst2.adSaldoQuantBenef(4), tSldMesEst2.adSaldoQuantBenef(5), tSldMesEst2.adSaldoQuantBenef(6), tSldMesEst2.adSaldoQuantBenef(7), tSldMesEst2.adSaldoQuantBenef(8), tSldMesEst2.adSaldoQuantBenef(9), _
                              tSldMesEst2.adSaldoQuantBenef(10), tSldMesEst2.adSaldoQuantBenef(11), tSldMesEst2.adSaldoQuantBenef(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 60967
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 60968
    
    dQuant = tSldMesEst2.dQuantInicialBenef
    
    For iIndice = 1 To iMes
    
        dQuant = dQuant + tSldMesEst2.adSaldoQuantBenef(iIndice)
        
    Next
    
    dQuantBenefAtual = dQuant
    
    SldMesEst2_Le_QuantBenef = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_QuantBenef:
    
    SldMesEst2_Le_QuantBenef = gErr
    
    Select Case gErr
    
        Case 60967, 60968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152080)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantConsig(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dQuantConsigAtual As Double) As Long
'calcula a quantidade de material em Consignação nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double

On Error GoTo Erro_SldMesEst2_Le_QuantConsig

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConsig, SaldoQuantConsig1, SaldoQuantConsig2, SaldoQuantConsig3, SaldoQuantConsig4, SaldoQuantConsig5, SaldoQuantConsig6, SaldoQuantConsig7, SaldoQuantConsig8, SaldoQuantConsig9, SaldoQuantConsig10, SaldoQuantConsig11, SaldoQuantConsig12 FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialConsig, tSldMesEst2.adSaldoQuantConsig(1), tSldMesEst2.adSaldoQuantConsig(2), tSldMesEst2.adSaldoQuantConsig(3), tSldMesEst2.adSaldoQuantConsig(4), tSldMesEst2.adSaldoQuantConsig(5), tSldMesEst2.adSaldoQuantConsig(6), tSldMesEst2.adSaldoQuantConsig(7), tSldMesEst2.adSaldoQuantConsig(8), tSldMesEst2.adSaldoQuantConsig(9), _
                              tSldMesEst2.adSaldoQuantConsig(10), tSldMesEst2.adSaldoQuantConsig(11), tSldMesEst2.adSaldoQuantConsig(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 60967
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 60968
    
    dQuant = tSldMesEst2.dQuantInicialConsig
    
    For iIndice = 1 To iMes
    
        dQuant = dQuant + tSldMesEst2.adSaldoQuantConsig(iIndice)
        
    Next
    
    dQuantConsigAtual = dQuant
    
    SldMesEst2_Le_QuantConsig = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_QuantConsig:
    
    SldMesEst2_Le_QuantConsig = gErr
    
    Select Case gErr
    
        Case 60967, 60968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152081)
        
    End Select
    
    Exit Function

End Function

Function TipoDeProduto_Le(ByVal objTipoDeProduto As ClassTipoDeProduto) As Long
'Lê a tabela TiposDeProduto, a partir do Tipo De Produto

Dim lErro As Long
Dim tTipoDeProduto As typeTipoDeProduto
Dim lComando As Long

On Error GoTo Erro_TipoDeProduto_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22521

    tTipoDeProduto.sDescricao = String(STRING_TIPODEPRODUTO_DESCRICAO, 0)
    tTipoDeProduto.sSigla = String(STRING_TIPODEPRODUTO_SIGLA, 0)
    tTipoDeProduto.sSiglaUMCompra = String(STRING_UM_SIGLA, 0)
    tTipoDeProduto.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    tTipoDeProduto.sSiglaUMVenda = String(STRING_UM_SIGLA, 0)
    tTipoDeProduto.sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
    tTipoDeProduto.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
    tTipoDeProduto.sISSCodigo = String(STRING_PRODUTO_ISS_CODIGO, 0)
    tTipoDeProduto.sContaContabil = String(STRING_CONTA, 0)
    tTipoDeProduto.sContaProducao = String(STRING_CONTA, 0)
    tTipoDeProduto.sSiglaUMTrib = String(STRING_UM_SIGLA, 0)
    tTipoDeProduto.sGenero = String(STRING_PRODUTOGENERO_CODIGO, 0)
    tTipoDeProduto.sISSQN = String(STRING_ISSQN_CODIGO, 0)
    tTipoDeProduto.sCEST = String(STRING_CEST, 0)
    
    'Lê a tabela TipoDeProduto à partir do Tipo De Produto
    lErro = Comando_Executar(lComando, "SELECT Descricao, Sigla, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, ControleEstoque, Faturamento, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, Apropriacao, MesesConsumoMedio, ConsumoMedioMax, TempoRessupMax, IntRessup, PercentMaisQuantCotAnt, PercentMenosQuantCotAnt, PercentMaisReceb, PercentMenosReceb, Residuo, Natureza, ConsideraQuantCotAnt, TemFaixaReceb, RecebForaFaixa, ContaContabil, ContaContabilProducao, KitVendaComp, Origem, SiglaUMTrib, ExTIPI, ProdutoEspecifico, Genero, ISSQN, CEST FROM TiposDeProduto WHERE TipoDeProduto = ?", _
    tTipoDeProduto.sDescricao, tTipoDeProduto.sSigla, tTipoDeProduto.iClasseUM, tTipoDeProduto.sSiglaUMEstoque, tTipoDeProduto.sSiglaUMCompra, tTipoDeProduto.sSiglaUMVenda, tTipoDeProduto.iControleEstoque, tTipoDeProduto.iFaturamento, tTipoDeProduto.iCompras, tTipoDeProduto.iPCP, tTipoDeProduto.iKitBasico, tTipoDeProduto.iKitInt, tTipoDeProduto.dIPIAliquota, tTipoDeProduto.sIPICodigo, tTipoDeProduto.sIPICodDIPI, tTipoDeProduto.iApropriacaoCusto, _
    tTipoDeProduto.iMesesConsumoMedio, tTipoDeProduto.dConsumoMedioMax, tTipoDeProduto.dTempoRessupMax, tTipoDeProduto.iIntRessup, tTipoDeProduto.dPercentMaisQuantCotAnt, tTipoDeProduto.dPercentMenosQuantCotAnt, tTipoDeProduto.dPercentMaisReceb, tTipoDeProduto.dPercentMenosReceb, tTipoDeProduto.dResiduo, tTipoDeProduto.iNatureza, tTipoDeProduto.iConsideraQuantCotAnt, tTipoDeProduto.iTemFaixaReceb, tTipoDeProduto.iRecebFaixaFora, tTipoDeProduto.sContaContabil, tTipoDeProduto.sContaProducao, tTipoDeProduto.iKitVendaComp, _
    tTipoDeProduto.iOrigem, tTipoDeProduto.sSiglaUMTrib, tTipoDeProduto.iExTIPI, tTipoDeProduto.iProdutoEspecifico, tTipoDeProduto.sGenero, tTipoDeProduto.sISSQN, tTipoDeProduto.sCEST, objTipoDeProduto.iTipo)

    If lErro <> AD_SQL_SUCESSO Then Error 22532

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22533

    'Verifica se o Tipo de Produto foi encontrado
    If lErro = AD_SQL_SEM_DADOS Then Error 22531

    'Transfere os dados para objTipoDeProduto
    objTipoDeProduto.sDescricao = tTipoDeProduto.sDescricao
    objTipoDeProduto.sSigla = tTipoDeProduto.sSigla
    objTipoDeProduto.sSiglaUMCompra = tTipoDeProduto.sSiglaUMCompra
    objTipoDeProduto.sSiglaUMEstoque = tTipoDeProduto.sSiglaUMEstoque
    objTipoDeProduto.sSiglaUMVenda = tTipoDeProduto.sSiglaUMVenda
    objTipoDeProduto.sIPICodDIPI = tTipoDeProduto.sIPICodDIPI
    objTipoDeProduto.sIPICodigo = tTipoDeProduto.sIPICodigo
    objTipoDeProduto.sISSCodigo = tTipoDeProduto.sISSCodigo
    objTipoDeProduto.iClasseUM = tTipoDeProduto.iClasseUM
    objTipoDeProduto.iCompras = tTipoDeProduto.iCompras
    objTipoDeProduto.iControleEstoque = tTipoDeProduto.iControleEstoque
    objTipoDeProduto.iFaturamento = tTipoDeProduto.iFaturamento
    objTipoDeProduto.iKitBasico = tTipoDeProduto.iKitBasico
    objTipoDeProduto.iKitInt = tTipoDeProduto.iKitInt
    objTipoDeProduto.iPCP = tTipoDeProduto.iPCP
    objTipoDeProduto.iApropriacaoCusto = tTipoDeProduto.iApropriacaoCusto
    objTipoDeProduto.iConsideraQuantCotAnt = tTipoDeProduto.iConsideraQuantCotAnt
    objTipoDeProduto.iIntRessup = tTipoDeProduto.iIntRessup
    objTipoDeProduto.iMesesConsumoMedio = tTipoDeProduto.iMesesConsumoMedio
    objTipoDeProduto.dTempoRessupMax = tTipoDeProduto.dTempoRessupMax
    objTipoDeProduto.dIPIAliquota = tTipoDeProduto.dIPIAliquota
    objTipoDeProduto.dConsumoMedioMax = tTipoDeProduto.dConsumoMedioMax
    objTipoDeProduto.dISSAliquota = tTipoDeProduto.dISSAliquota
    objTipoDeProduto.dPercentMaisQuantCotAnt = tTipoDeProduto.dPercentMaisQuantCotAnt
    objTipoDeProduto.dPercentMaisReceb = tTipoDeProduto.dPercentMaisReceb
    objTipoDeProduto.dPercentMenosQuantCotAnt = tTipoDeProduto.dPercentMenosQuantCotAnt
    objTipoDeProduto.dPercentMenosReceb = tTipoDeProduto.dPercentMenosReceb
    objTipoDeProduto.dResiduo = tTipoDeProduto.dResiduo
    objTipoDeProduto.iNatureza = tTipoDeProduto.iNatureza
    objTipoDeProduto.iTemFaixaReceb = tTipoDeProduto.iTemFaixaReceb
    objTipoDeProduto.iRecebForaFaixa = tTipoDeProduto.iRecebFaixaFora
    objTipoDeProduto.sContaContabil = tTipoDeProduto.sContaContabil
    objTipoDeProduto.sContaProducao = tTipoDeProduto.sContaProducao
    objTipoDeProduto.iKitVendaComp = tTipoDeProduto.iKitVendaComp

    objTipoDeProduto.iOrigem = tTipoDeProduto.iOrigem
    objTipoDeProduto.sSiglaUMTrib = tTipoDeProduto.sSiglaUMTrib
    objTipoDeProduto.iExTIPI = tTipoDeProduto.iExTIPI
    objTipoDeProduto.iProdutoEspecifico = tTipoDeProduto.iProdutoEspecifico
    objTipoDeProduto.sGenero = tTipoDeProduto.sGenero
    objTipoDeProduto.sISSQN = tTipoDeProduto.sISSQN
    objTipoDeProduto.sCEST = tTipoDeProduto.sCEST
    
    Call Comando_Fechar(lComando)

    TipoDeProduto_Le = SUCESSO

    Exit Function

Erro_TipoDeProduto_Le:

    TipoDeProduto_Le = Err

    Select Case Err

        Case 22521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22531
            'Erro já tratado na rotina chamadora

        Case 22532, 22533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEPRODUTO", Err, objTipoDeProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152082)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TipoDeProduto_Le_Todos(ByVal colTipos As Collection) As Long
'Lê a tabela TiposDeProduto, a partir do Tipo De Produto

Dim lErro As Long
Dim tTipoDeProduto As typeTipoDeProduto
Dim objTipoDeProduto As ClassTipoDeProduto
Dim lComando As Long

On Error GoTo Erro_TipoDeProduto_Le_Todos

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22521

    tTipoDeProduto.sDescricao = String(STRING_TIPODEPRODUTO_DESCRICAO, 0)
    tTipoDeProduto.sSigla = String(STRING_TIPODEPRODUTO_SIGLA, 0)
    tTipoDeProduto.sSiglaUMCompra = String(STRING_UM_SIGLA, 0)
    tTipoDeProduto.sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    tTipoDeProduto.sSiglaUMVenda = String(STRING_UM_SIGLA, 0)
    tTipoDeProduto.sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
    tTipoDeProduto.sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
    tTipoDeProduto.sISSCodigo = String(STRING_PRODUTO_ISS_CODIGO, 0)
    tTipoDeProduto.sContaContabil = String(STRING_CONTA, 0)
    tTipoDeProduto.sContaProducao = String(STRING_CONTA, 0)
    tTipoDeProduto.sSiglaUMTrib = String(STRING_UM_SIGLA, 0)
    tTipoDeProduto.sGenero = String(STRING_PRODUTOGENERO_CODIGO, 0)
    tTipoDeProduto.sISSQN = String(STRING_ISSQN_CODIGO, 0)
    tTipoDeProduto.sCEST = String(STRING_CEST, 0)
    
    'Lê a tabela TipoDeProduto à partir do Tipo De Produto
    lErro = Comando_Executar(lComando, "SELECT TipoDeProduto, Descricao, Sigla, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, ControleEstoque, Faturamento, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, Apropriacao, MesesConsumoMedio, ConsumoMedioMax, TempoRessupMax, IntRessup, PercentMaisQuantCotAnt, PercentMenosQuantCotAnt, PercentMaisReceb, PercentMenosReceb, Residuo, Natureza, ConsideraQuantCotAnt, TemFaixaReceb, RecebForaFaixa, ContaContabil, ContaContabilProducao, KitVendaComp, Origem, SiglaUMTrib, ExTIPI, ProdutoEspecifico, Genero, ISSQN, CEST FROM TiposDeProduto ORDER BY TipoDeProduto ", _
    tTipoDeProduto.iTipo, tTipoDeProduto.sDescricao, tTipoDeProduto.sSigla, tTipoDeProduto.iClasseUM, tTipoDeProduto.sSiglaUMEstoque, tTipoDeProduto.sSiglaUMCompra, tTipoDeProduto.sSiglaUMVenda, tTipoDeProduto.iControleEstoque, tTipoDeProduto.iFaturamento, tTipoDeProduto.iCompras, tTipoDeProduto.iPCP, tTipoDeProduto.iKitBasico, tTipoDeProduto.iKitInt, tTipoDeProduto.dIPIAliquota, tTipoDeProduto.sIPICodigo, tTipoDeProduto.sIPICodDIPI, tTipoDeProduto.iApropriacaoCusto, _
    tTipoDeProduto.iMesesConsumoMedio, tTipoDeProduto.dConsumoMedioMax, tTipoDeProduto.dTempoRessupMax, tTipoDeProduto.iIntRessup, tTipoDeProduto.dPercentMaisQuantCotAnt, tTipoDeProduto.dPercentMenosQuantCotAnt, tTipoDeProduto.dPercentMaisReceb, tTipoDeProduto.dPercentMenosReceb, tTipoDeProduto.dResiduo, tTipoDeProduto.iNatureza, tTipoDeProduto.iConsideraQuantCotAnt, tTipoDeProduto.iTemFaixaReceb, tTipoDeProduto.iRecebFaixaFora, tTipoDeProduto.sContaContabil, tTipoDeProduto.sContaProducao, tTipoDeProduto.iKitVendaComp, _
    tTipoDeProduto.iOrigem, tTipoDeProduto.sSiglaUMTrib, tTipoDeProduto.iExTIPI, tTipoDeProduto.iProdutoEspecifico, tTipoDeProduto.sGenero, tTipoDeProduto.sISSQN, tTipoDeProduto.sCEST)

    If lErro <> AD_SQL_SUCESSO Then Error 22532

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22533

    'Verifica se o Tipo de Produto foi encontrado
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTipoDeProduto = New ClassTipoDeProduto

        'Transfere os dados para objTipoDeProduto
        objTipoDeProduto.iTipo = tTipoDeProduto.iTipo
        objTipoDeProduto.sDescricao = tTipoDeProduto.sDescricao
        objTipoDeProduto.sSigla = tTipoDeProduto.sSigla
        objTipoDeProduto.sSiglaUMCompra = tTipoDeProduto.sSiglaUMCompra
        objTipoDeProduto.sSiglaUMEstoque = tTipoDeProduto.sSiglaUMEstoque
        objTipoDeProduto.sSiglaUMVenda = tTipoDeProduto.sSiglaUMVenda
        objTipoDeProduto.sIPICodDIPI = tTipoDeProduto.sIPICodDIPI
        objTipoDeProduto.sIPICodigo = tTipoDeProduto.sIPICodigo
        objTipoDeProduto.sISSCodigo = tTipoDeProduto.sISSCodigo
        objTipoDeProduto.iClasseUM = tTipoDeProduto.iClasseUM
        objTipoDeProduto.iCompras = tTipoDeProduto.iCompras
        objTipoDeProduto.iControleEstoque = tTipoDeProduto.iControleEstoque
        objTipoDeProduto.iFaturamento = tTipoDeProduto.iFaturamento
        objTipoDeProduto.iKitBasico = tTipoDeProduto.iKitBasico
        objTipoDeProduto.iKitInt = tTipoDeProduto.iKitInt
        objTipoDeProduto.iPCP = tTipoDeProduto.iPCP
        objTipoDeProduto.iApropriacaoCusto = tTipoDeProduto.iApropriacaoCusto
        objTipoDeProduto.iConsideraQuantCotAnt = tTipoDeProduto.iConsideraQuantCotAnt
        objTipoDeProduto.iIntRessup = tTipoDeProduto.iIntRessup
        objTipoDeProduto.iMesesConsumoMedio = tTipoDeProduto.iMesesConsumoMedio
        objTipoDeProduto.dTempoRessupMax = tTipoDeProduto.dTempoRessupMax
        objTipoDeProduto.dIPIAliquota = tTipoDeProduto.dIPIAliquota
        objTipoDeProduto.dConsumoMedioMax = tTipoDeProduto.dConsumoMedioMax
        objTipoDeProduto.dISSAliquota = tTipoDeProduto.dISSAliquota
        objTipoDeProduto.dPercentMaisQuantCotAnt = tTipoDeProduto.dPercentMaisQuantCotAnt
        objTipoDeProduto.dPercentMaisReceb = tTipoDeProduto.dPercentMaisReceb
        objTipoDeProduto.dPercentMenosQuantCotAnt = tTipoDeProduto.dPercentMenosQuantCotAnt
        objTipoDeProduto.dPercentMenosReceb = tTipoDeProduto.dPercentMenosReceb
        objTipoDeProduto.dResiduo = tTipoDeProduto.dResiduo
        objTipoDeProduto.iNatureza = tTipoDeProduto.iNatureza
        objTipoDeProduto.iTemFaixaReceb = tTipoDeProduto.iTemFaixaReceb
        objTipoDeProduto.iRecebForaFaixa = tTipoDeProduto.iRecebFaixaFora
        objTipoDeProduto.sContaContabil = tTipoDeProduto.sContaContabil
        objTipoDeProduto.sContaProducao = tTipoDeProduto.sContaProducao
        objTipoDeProduto.iKitVendaComp = tTipoDeProduto.iKitVendaComp
    
        objTipoDeProduto.iOrigem = tTipoDeProduto.iOrigem
        objTipoDeProduto.sSiglaUMTrib = tTipoDeProduto.sSiglaUMTrib
        objTipoDeProduto.iExTIPI = tTipoDeProduto.iExTIPI
        objTipoDeProduto.iProdutoEspecifico = tTipoDeProduto.iProdutoEspecifico
        objTipoDeProduto.sGenero = tTipoDeProduto.sGenero
        objTipoDeProduto.sISSQN = tTipoDeProduto.sISSQN
        objTipoDeProduto.sCEST = tTipoDeProduto.sCEST
    
        colTipos.Add objTipoDeProduto
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22533
        
    Loop
    
    Call Comando_Fechar(lComando)

    TipoDeProduto_Le_Todos = SUCESSO

    Exit Function

Erro_TipoDeProduto_Le_Todos:

    TipoDeProduto_Le_Todos = Err

    Select Case Err

        Case 22521
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22531
            'Erro já tratado na rotina chamadora

        Case 22532, 22533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEPRODUTO", Err, objTipoDeProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152082)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoFilial_Le(ByVal objProdutoFilial As ClassProdutoFilial) As Long
'typeProdutoFilial 'Lê os dados do Produto da Filial à partir da chave fornecida em objProdutoFilial. Retorna os dados em objProdutoFilial

Dim lErro As Long
Dim lComando As Long
Dim tProdutoFilial As typeProdutoFilial

On Error GoTo Erro_ProdutoFilial_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28258

    tProdutoFilial.sClasseABC = String(STRING_PRODUTOFILIAL_CLASSEABC, 0)
    tProdutoFilial.sSituacaoTribECF = String(STRING_PRODUTOFILIAL_SITUACAOTRIBECF, 0)
    tProdutoFilial.sICMSAliquota = String(STRING_PRODUTOFILIAL_ICMS, 0)
    
    'Pesquisa no BD ProdutoFilial
    With tProdutoFilial
        lErro = Comando_Executar(lComando, "SELECT  Fornecedor, FilialForn, Almoxarifado, EstoqueSeguranca,ESAuto,EstoqueMaximo, PontoPedido,PPAuto, ClasseABC, LoteEconomico, IntRessup,TempoRessup, TRAuto,TempoRessupMax,ConsumoMedio,CMAuto,ConsumoMedioMax,MesesConsumoMedio, QuantPedida, TabelaPreco,SituacaoTribECF,ICMSAliquota , LoteMinimo, ProdNaFilial FROM ProdutosFilial WHERE FilialEmpresa= ? AND Produto = ?", _
             .lFornecedor, .iFilialForn, .iAlmoxarifado, .dEstoqueSeguranca, .iESAuto, .dEstoqueMaximo, .dPontoPedido, .iPPAuto, .sClasseABC, .dLoteEconomico, .iIntRessup, .iTempoRessup, .iTRAuto, .dTempoRessupMax, .dConsumoMedio, .iCMAuto, .dConsumoMedioMax, .iMesesConsumoMedio, .dQuantPedida, .iTabelaPreco, .sSituacaoTribECF, .sICMSAliquota, .dLoteMinimo, .iProdNaFilial, objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)
        If lErro <> AD_SQL_SUCESSO Then Error 28259
    End With

    'Tenta selecionar ProdutoFilial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28260

    'Não encontrou ProdutoFilial
    If lErro = AD_SQL_SEM_DADOS Then Error 28261

    'Preenche objProdutoFilial
    With tProdutoFilial

        objProdutoFilial.lFornecedor = .lFornecedor
        objProdutoFilial.iFilialForn = .iFilialForn
        objProdutoFilial.iAlmoxarifado = .iAlmoxarifado
        objProdutoFilial.dEstoqueSeguranca = .dEstoqueSeguranca
        objProdutoFilial.dEstoqueMaximo = .dEstoqueMaximo
        objProdutoFilial.dPontoPedido = .dPontoPedido
        objProdutoFilial.sClasseABC = .sClasseABC
        objProdutoFilial.dLoteEconomico = .dLoteEconomico
        objProdutoFilial.iIntRessup = .iIntRessup
        objProdutoFilial.dConsumoMedio = .dConsumoMedio
        objProdutoFilial.dConsumoMedioMax = .dConsumoMedioMax
        objProdutoFilial.iCMCalculado = .iCMAuto
        objProdutoFilial.iMesesConsumoMedio = .iMesesConsumoMedio
        objProdutoFilial.iTempoRessup = .iTempoRessup
        objProdutoFilial.dTempoRessupMax = .dTempoRessupMax
        objProdutoFilial.iTRCalculado = .iTRAuto
        objProdutoFilial.iPPCalculado = .iPPAuto
        objProdutoFilial.iESCalculado = .iESAuto
        objProdutoFilial.dQuantPedida = .dQuantPedida
        objProdutoFilial.iTabelaPreco = .iTabelaPreco
        objProdutoFilial.sSituacaoTribECF = .sSituacaoTribECF
        objProdutoFilial.sICMSAliquota = .sICMSAliquota
        objProdutoFilial.dLoteMinimo = .dLoteMinimo
        objProdutoFilial.iProdNaFilial = .iProdNaFilial
        
    End With

    'Fechamento comando
    Call Comando_Fechar(lComando)

    ProdutoFilial_Le = SUCESSO

    Exit Function

Erro_ProdutoFilial_Le:

    ProdutoFilial_Le = Err

    Select Case Err

        Case 28258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28259, 28260
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", Err, objProdutoFilial.iFilialEmpresa, objProdutoFilial.sProduto)

        Case 28261
            'Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152083)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FornecedorProdutoFF_Le(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Lê dados da associacao Fornecedor-Produto à partir da chave fornecida em objFornecedorProdutoFF. Retorna os dados em objFornecedorProdutoFF

Dim lErro As Long
Dim lComando As Long
Dim tFornecedorProdutoFF As typeFornecedorProdutoFF
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_FornecedorProdutoFF_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 54214
    
    tFornecedorProdutoFF.sProdutoFornecedor = String(STRING_PRODUTO_FORNECEDOR, 0)
    tFornecedorProdutoFF.sDescricao = String(STRING_FORNECEDOR_PRODUTO_DESCRICAO, 0)
    
    'Pesquisa no BD o Fornecedor do Produto
    lErro = Comando_Executar(lComando, "SELECT ProdutoFornecedor, LoteMinimo, Nota, Descricao FROM FornecedorProdutoFF WHERE Fornecedor = ? AND Produto = ? AND FilialEmpresa = ? AND FilialForn = ?", _
    tFornecedorProdutoFF.sProdutoFornecedor, tFornecedorProdutoFF.dLoteMinimo, tFornecedorProdutoFF.iNota, tFornecedorProdutoFF.sDescricao, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.iFilialForn)
    If lErro <> AD_SQL_SUCESSO Then Error 54215

    'Verifica se o Fornecedor do Produto existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54216

    'Não encontrou o FornecedorProduto
    If lErro = AD_SQL_SEM_DADOS Then Error 54217

    'Preenche objFornecedorProdutoFF
    With objFornecedorProdutoFF
        'Frame Dados Principais
        .sProdutoFornecedor = tFornecedorProdutoFF.sProdutoFornecedor
        .iNota = tFornecedorProdutoFF.iNota
        .dLoteMinimo = tFornecedorProdutoFF.dLoteMinimo
        .sDescricao = tFornecedorProdutoFF.sDescricao
        
    End With

    objCondicaoPagto.iCodigo = objFornecedorProdutoFF.iCondPagto
    
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then Error 54727
    
    objFornecedorProdutoFF.sCondPagto = objCondicaoPagto.sDescReduzida
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    FornecedorProdutoFF_Le = SUCESSO

    Exit Function

Erro_FornecedorProdutoFF_Le:

    FornecedorProdutoFF_Le = Err

    Select Case Err

        Case 54214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54215, 54216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", Err, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.sProduto)

        Case 54217, 54727
            'Não encontrou o Fornecedor do Produto. Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152084)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FornecedorProdutoFF_Le_Fornecedores(ByVal colFornecedores As AdmCollCodigoNome, ByVal objProduto As ClassProduto) As Long
'Lê todos nomes reduzidos de Fornecedores associados a sProduto e coloca-os em colFornecedores

Dim lErro As Long
Dim lComando As Long
Dim sNomeReduzido As String
Dim lCodigo As Long

On Error GoTo Erro_FornecedorProdutoFF_Le_Fornecedores

    'Abertura comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 54218

    sNomeReduzido = String(STRING_FORNECEDOR_NOME_REDUZIDO, 0)

    'Lê a tabela de FornecedorProduto e Fornecedores
    lErro = Comando_Executar(lComando, "SELECT DISTINCT Fornecedores.Codigo, Fornecedores.NomeReduzido FROM FornecedorProdutoFF, Fornecedores WHERE FornecedorProdutoFF.Produto = ? AND FornecedorProdutoFF.Fornecedor = Fornecedores.Codigo AND FornecedorProdutoFF.FilialEmpresa = ?", lCodigo, sNomeReduzido, objProduto.sCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 54219

    'Lê o primeiro Fornecedor de Produto
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54220

    Do While lErro <> AD_SQL_SEM_DADOS

        'Coloca o código lido na coleção
        colFornecedores.Add lCodigo, sNomeReduzido

        'Lê o próximo código
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54221

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    FornecedorProdutoFF_Le_Fornecedores = SUCESSO

    Exit Function

Erro_FornecedorProdutoFF_Le_Fornecedores:

    FornecedorProdutoFF_Le_Fornecedores = Err

    Select Case Err

        Case 54218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 54219, 54220, 54221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTO_FORNECEDOR", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152085)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FornecedorProdutoFF_Existe_Relacionamento(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Verifica se existe relacionamento com as tabelas: PedidoCompra, RequisicaoCompra, Concorrencia

Dim lErro As Long
Dim lComando As Long
Dim lTransacao As Long
Dim lFornecedor As Long
Dim lCodCompra As Long

On Error GoTo Erro_FornecedorProdutoFF_Existe_Relacionamento

    'Abertura do comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 54253

    'PedidoCompra
    lErro = Comando_Executar(lComando, "SELECT PedidoCompra.Codigo FROM PedidoCompra, ItensPedCompra WHERE PedidoCompra.NumIntDoc = ItensPedCompra.PedCompra AND PedidoCompra.Fornecedor = ? AND PedidoCompra.Filial = ? AND PedidoCompra.FilialEmpresa = ? AND ItensPedCompra.Produto = ?", _
    lCodCompra, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 54255

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54256

    'Existe Relacionamento
    If lErro = AD_SQL_SUCESSO Then Error 54257

    'RequisicaoCompra
    lErro = Comando_Executar(lComando, "SELECT RequisicaoCompra.Codigo FROM RequisicaoCompra, ItensReqCompra WHERE RequisicaoCompra.NumIntDoc = ItensReqCompra.ReqCompra AND ItensReqCompra.Fornecedor = ? AND ItensReqCompra.Filial = ? AND RequisicaoCompra.FilialEmpresa = ? AND ItensReqCompra.Produto = ?", _
    lCodCompra, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 54260

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54261

    'Existe Relacionamento
    If lErro = AD_SQL_SUCESSO Then Error 54262

    'Concorrencia
    lErro = Comando_Executar(lComando, "SELECT Concorrencia.Codigo FROM Concorrencia, ItensConcorrencia WHERE Concorrencia.NumIntDoc = ItensConcorrencia.Concorrencia AND Concorrencia.FilialEmpresa = ? AND ItensConcorrencia.Fornecedor = ? AND ItensConcorrencia.Filial = ? AND ItensConcorrencia.Produto = ?", _
    lCodCompra, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 54264

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54265

    'Existe Relacionamento
    If lErro = AD_SQL_SUCESSO Then Error 54266

    'Fechamento do comando
    Call Comando_Fechar(lComando)

    FornecedorProdutoFF_Existe_Relacionamento = SUCESSO

    Exit Function

Erro_FornecedorProdutoFF_Existe_Relacionamento:

    FornecedorProdutoFF_Existe_Relacionamento = Err

    Select Case Err

        Case 54253
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 54255, 54256
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRA", Err)

        Case 54257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_UTILIZADO_PEDCOMPRA", Err, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, lCodCompra)

        Case 54260, 54261
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REQUISICAOCOMPRA", Err)

        Case 54262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_UTILIZADO_REQUISICAOCOMPRA", Err, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, lCodCompra)

        Case 54264, 54265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONCORRENCIA", Err)

        Case 54266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDORPRODUTOFF_UTILIZADO_CONCORRENCIA", Err, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, lCodCompra)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152086)

    End Select

    'Fechamento do comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FornecedoresProdutoFF_Le(colFornecedorProdutoFF As Collection, ByVal objProduto As ClassProduto) As Long
'Le os registros da tabela FornecedorProdutoFF de acordo com giFilial e código do Produto e armazena na colecao

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim lTransacao As Long
Dim tFornecedorProdutoFF As typeFornecedorProdutoFF
Dim iIndice As Integer
Dim objFilFornFilEmp As New ClassFilFornFilEmp
Dim iTipoFreteUltimaCotacao As Integer
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF

On Error GoTo Erro_FornecedoresProdutoFF_Le

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then Error 63110
    Next

    tFornecedorProdutoFF.sProduto = String(STRING_PRODUTO, 0)
    tFornecedorProdutoFF.sProdutoFornecedor = String(STRING_PRODUTO, 0)

    'Lê a tabela FornecedorProdutoFF de acordo com o Produto e FilialEmpresa, ordenados por Fornecedor e FilialForn
    lErro = Comando_Executar(alComando(1), "SELECT FilialEmpresa,Produto,Fornecedor,FilialForn,ProdutoFornecedor,LoteMinimo,Nota FROM FornecedorProdutoFF WHERE FilialEmpresa=? AND Produto=? ORDER BY Fornecedor, FilialForn", tFornecedorProdutoFF.iFilialEmpresa, tFornecedorProdutoFF.sProduto, tFornecedorProdutoFF.lFornecedor, tFornecedorProdutoFF.iFilialForn, tFornecedorProdutoFF.sProdutoFornecedor, tFornecedorProdutoFF.dLoteMinimo, tFornecedorProdutoFF.iNota, giFilialEmpresa, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 63112

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 63113

    'Se nao encontrou registro em FornecedorProdutoFF
    If lErro = AD_SQL_SEM_DADOS Then Error 63156

    'Set colFilFornFilEmp = New Collection
    Set colFornecedorProdutoFF = New Collection
    
    Do While lErro = AD_SQL_SUCESSO

        'Guarda na coleção os regitros lidos de FornecedorProdutoFF
        'Set objFilFornFilEmp = New ClassFilFornFilEmp
        Set objFornecedorProdutoFF = New ClassFornecedorProdutoFF
        
        'Preenche objFilFornFilEmp com os dados lidos do BD
        objFornecedorProdutoFF.dLoteMinimo = tFornecedorProdutoFF.dLoteMinimo
        objFornecedorProdutoFF.iFilialEmpresa = tFornecedorProdutoFF.iFilialEmpresa
        objFornecedorProdutoFF.iFilialForn = tFornecedorProdutoFF.iFilialForn
        objFornecedorProdutoFF.iNota = tFornecedorProdutoFF.iNota
        objFornecedorProdutoFF.iPadrao = tFornecedorProdutoFF.iPadrao
        objFornecedorProdutoFF.lFornecedor = tFornecedorProdutoFF.lFornecedor
        objFornecedorProdutoFF.sProduto = tFornecedorProdutoFF.sProduto
        objFornecedorProdutoFF.sProdutoFornecedor = tFornecedorProdutoFF.sProdutoFornecedor
        
        'Adiciona na colecao
        colFornecedorProdutoFF.Add objFornecedorProdutoFF
        
        'Le o próximo registro existente no banco de dados
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 63114

    Loop

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FornecedoresProdutoFF_Le = SUCESSO

    Exit Function

Erro_FornecedoresProdutoFF_Le:

    FornecedoresProdutoFF_Le = Err

    Select Case Err

        Case 63110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 63112, 63113, 63114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_FORNECEDOR", objProduto.sCodigo)

        Case 63156
            'Erro tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152087)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'??? jones 2002: deve ter erro pois o parametro lFornecedor nao é utilizado
Function FornecedorProdutoFF_Le_FilForn(ByVal sProduto As String, lFornecedor As Long, ByVal colCodigoNome As AdmColCodigoNome) As Long
'Lê código e nome das Filiais do Fornecedor das tabelas FornedorProdutoFF e FiliaisFornecedores

Dim lErro As Long
Dim tFilialForn As typeFilialFornecedor
Dim lComando As Long

On Error GoTo Erro_FornecedorProdutoFF_Le_FilForn

    'Abre o comando
    lErro = Comando_Abrir()
    If lErro <> SUCESSO Then Error 63221

    tFilialForn.sNome = String(STRING_NOME_FILIAL_FORN, 0)

    'Lê códigos e nomes de Filiais do Fornecedor
    lErro = Comando_Executar(lComando, "SELECT FornecedorProdutoFF.FilialForn,FiliaisFornecedores.Nome FROM FornecedorProdutoFF,FiliaisFornecedores as FiliaisFornecedor WHERE FornecedorProdutoFF.Fornecedor=FiliaisFornecedores.CodFornecedor AND FornecedorProdutoFF.FilialForn= FiliaisFornecedores.CodFilial AND FornecedorProdutoFF.Produto= sProduto AND FornecedorProdutoFF.FilialEmpresa= giFilialEmpresa AND FornecedorProdutoFF.Fornecedor=lFornecedor", tFilialForn.iCodFilial, tFilialForn.sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 63217

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 63218

    'Se nao encontrou nenhuma filial==>erro
    If lErro = AD_SQL_SEM_DADOS Then Error 63219

    Do While lErro <> AD_SQL_SEM_DADOS

        'Coloca o Codigo e Nome lidos na coleção
        colCodigoNome.Add tFilialForn.iCodFilial, tFilialForn.sNome

        'Lê o próximo Codigo e DescReduzida
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 63220

    Loop

    'Fecha comando
    lErro = Comando_Fechar(lComando)

    FornecedorProdutoFF_Le_FilForn = SUCESSO

    Exit Function

Erro_FornecedorProdutoFF_Le_FilForn:

    FornecedorProdutoFF_Le_FilForn = Err

    Select Case Err

        Case 63219
            'Tratado na rotina chamadora

        Case 63218, 63220

        Case 63221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152088)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FornecedorProdutoFF_Le_CotacaoProdutoForn(ByVal objCotacaoProduto As ClassCotacaoProduto, ByVal colFornecedorProdutoFF As Collection) As Long
'Lê a tabela FornecedorProdutoFF em busca dos fornecedores associados ao produto passado como parametro
'Devolve-os em colCotacaoProdutoFornecedor

Dim lErro As Long
Dim lComando As Long
Dim tFornecedorProdutoFF As typeFornecedorProdutoFF
Dim objFornecedorProdutoFF As ClassFornecedorProdutoFF

On Error GoTo Erro_FornecedorProdutoFF_Le_CotacaoProdutoForn

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 63590

    tFornecedorProdutoFF.sProdutoFornecedor = String(STRING_PRODUTO_FORNECEDOR, 0)

    'Pesquisa no BD o Fornecedor do Produto
    lErro = Comando_Executar(lComando, "SELECT FP.ProdutoFornecedor, FP.LoteMinimo, FP.Nota, FP.Fornecedor,FP.FilialForn FROM FornecedorProdutoFF AS FP, Fornecedores AS F WHERE FP.Fornecedor = F.Codigo AND FP.Produto = ? AND FP.FilialEmpresa = ? AND F.Ativo = ?", _
    tFornecedorProdutoFF.sProdutoFornecedor, tFornecedorProdutoFF.dLoteMinimo, tFornecedorProdutoFF.iNota, tFornecedorProdutoFF.lFornecedor, tFornecedorProdutoFF.iFilialForn, objCotacaoProduto.sProduto, giFilialEmpresa, MARCADO)
    If lErro <> AD_SQL_SUCESSO Then Error 63591

    'Verifica se o Fornecedor do Produto existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 63592

    'Não encontrou o FornecedorProduto
    If lErro = AD_SQL_SEM_DADOS Then Error 63593

    Do While lErro = AD_SQL_SUCESSO
        
        Set objFornecedorProdutoFF = New ClassFornecedorProdutoFF

        'Preenche objCotacaoProdutoFornecedor com os dados lidos do BD
        With objFornecedorProdutoFF
            .iSelecionado = objCotacaoProduto.iSelecionado
            .iEscolhido = Selecionado
            .sProdutoFornecedor = tFornecedorProdutoFF.sProdutoFornecedor
            .dLoteMinimo = tFornecedorProdutoFF.dLoteMinimo
            .iNota = tFornecedorProdutoFF.iNota
            .sProduto = objCotacaoProduto.sProduto
            .lFornecedor = tFornecedorProdutoFF.lFornecedor
            .iFilialForn = tFornecedorProdutoFF.iFilialForn
            .iFilialEmpresa = giFilialEmpresa
            'Referencia CotacaoProduto em objFornecedorProdutoFF
            .lNumIntCotacaoProduto = objCotacaoProduto.lNumIntDoc
            .lNumIntDoc = colFornecedorProdutoFF.Count + 1
        End With

        'Adiciona na colecao
        colFornecedorProdutoFF.Add objFornecedorProdutoFF

        'Busca o próximo Fornecedor associado ao Produto
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 63589

    Loop

    'Fechamento comando
    Call Comando_Fechar(lComando)

    FornecedorProdutoFF_Le_CotacaoProdutoForn = SUCESSO

    Exit Function

Erro_FornecedorProdutoFF_Le_CotacaoProdutoForn:

    FornecedorProdutoFF_Le_CotacaoProdutoForn = Err

    Select Case Err

        Case 63590
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 63589, 63591, 63592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", Err, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.sProduto)

        Case 63593
            'Não encontrou o Fornecedor do Produto. Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152089)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FornecedorProdutoFF_Le_NomeRed(ByVal sFornNomeRed As String, sNomeFilial As String, ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Verifica se existe associação entre a Filial de um Fornecedor e um Produto

Dim lErro As Long
Dim lComando As Long
Dim objFornecedor As New ClassFornecedor
Dim iFilialForn As Integer

On Error GoTo Erro_FornecedorProdutoFF_Le_NomeRed

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61775

    'Lê dados do Fornecedor a partir do seu nome Reduzido
    objFornecedor.sNomeReduzido = sFornNomeRed
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then Error 61776

    'Se não encontrou o Fornecedor, erro
    If lErro = 6681 Then Error 61777

    objFornecedorProdutoFF.lFornecedor = objFornecedor.lCodigo
        
    sNomeFilial = String(STRING_NOME_FILIAL_FORN, 0)
    
    'Pesquisa no BD se existe a associação da Filial de um Fornecedor com um Produto
    lErro = Comando_Executar(lComando, "SELECT FilialForn, Nome FROM FornecedorProdutoFF, FiliaisFornecedores WHERE FiliaisFornecedores.CodFornecedor = FornecedorProdutoFF.Fornecedor AND FiliaisFornecedores.CodFilial = FornecedorProdutoFF.FilialForn AND FilialEmpresa = ? AND Produto = ? AND Fornecedor = ? AND FilialForn = ?", _
        iFilialForn, sNomeFilial, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn)
    If lErro <> AD_SQL_SUCESSO Then Error 61778

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61779

    'Se não encontrar, erro
    If lErro = AD_SQL_SEM_DADOS Then Error 61780

    Call Comando_Fechar(lComando)

    FornecedorProdutoFF_Le_NomeRed = SUCESSO

    Exit Function

Erro_FornecedorProdutoFF_Le_NomeRed:

    FornecedorProdutoFF_Le_NomeRed = Err

    Select Case Err

        Case 61775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61776

        Case 61777
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case 61778, 61779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF1", Err, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn, objFornecedorProdutoFF.sProduto)

        Case 61780 'Associação Filial do Fornecedor X Produto não encontrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152090)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FornecedorProdutoFF_Le_FilialForn(ByVal sProduto As String, ByVal lFornecedor As Long, ByVal iFilialEmpresa As Integer, ByVal colCodigoNome As AdmColCodigoNome) As Long
'Carrega em colCodigoNome o código e o nome das filiais que possuem associação com o produto passado

Dim lErro As Long
Dim lComando As Long
Dim iFilialForn As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_FornecedorProdutoFF_Le_FilialForn

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 61781

    'Pesquisa no BD as Filiais que possui associação com o Produto Passado
    lErro = Comando_Executar(lComando, "SELECT FilialForn FROM FornecedorProdutoFF WHERE FilialEmpresa = ? AND Produto = ? AND Fornecedor = ?", iFilialForn, iFilialEmpresa, sProduto, lFornecedor)
    If lErro <> AD_SQL_SUCESSO Then Error 61782

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61783

    'Enquanto encontrar
    Do While lErro = AD_SQL_SUCESSO

        objFilialFornecedor.iCodFilial = iFilialForn
        objFilialFornecedor.lCodFornecedor = lFornecedor

        'Lê o Nome da Filial
        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 12929 Then Error 61785
        If lErro = 12929 Then Error 61786

        'Adiciona a Filial em colCodigoNome
        colCodigoNome.Add iFilialForn, objFilialFornecedor.sNome

        'Busca Próxima Filial
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 61784

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComando)

    FornecedorProdutoFF_Le_FilialForn = SUCESSO

    Exit Function

Erro_FornecedorProdutoFF_Le_FilialForn:

    FornecedorProdutoFF_Le_FilialForn = Err

    Select Case Err

        Case 61781
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61782, 61783, 61784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", Err, lFornecedor, sProduto)

        Case 61785

        Case 61786
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_CADASTRADA", Err, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152091)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FornecedorProdutoFF_Lock(ByVal objFornecedorProdutoFF As ClassFornecedorProdutoFF) As Long
'Lê e Faz "LockShared" do FornecedorProdutoFF passado

Dim lErro As Long
Dim lComando As Long
Dim lFornecedor As Long

On Error GoTo Erro_FornecedorProdutoFF_Lock

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67869
        
    'Lê o FornecedorProdutoFF passado
    lErro = Comando_ExecutarPos(lComando, "SELECT Fornecedor FROM FornecedorProdutoFF WHERE FilialEmpresa = ? AND Produto = ? AND Fornecedor = ? AND FilialForn = ?", 0, lFornecedor, objFornecedorProdutoFF.iFilialEmpresa, objFornecedorProdutoFF.sProduto, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.iFilialForn)
    If lErro <> AD_SQL_SUCESSO Then gError 67870

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67871
    
    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 67872

    'Faz LockShared no FornecedorProdutoFF
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 67873

    Call Comando_Fechar(lComando)
    
    FornecedorProdutoFF_Lock = SUCESSO
    
    Exit Function
    
    FornecedorProdutoFF_Lock = gErr
    
Erro_FornecedorProdutoFF_Lock:

    FornecedorProdutoFF_Lock = gErr
    
    Select Case gErr
        
        Case 67869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 67870, 67871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.sProduto)
        
        Case 67872 'FornecedorProdutoFF não cadastrado
        
        Case 67873
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FORNECEDORPRODUTOFF", gErr, objFornecedorProdutoFF.lFornecedor, objFornecedorProdutoFF.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152092)
            
    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Produto_Le_Descricao(ByVal objProduto As ClassProduto) As Long
'Lê o Produto com a Descricao Passada

Dim lErro As Long
Dim lComando As Long
Dim tProduto As typeProduto

On Error GoTo Erro_Produto_Le_Descricao

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 67015

    'Inicializa as variáveis do tipo String
    With tProduto

        .sCodigo = String(STRING_PRODUTO, 0)
        .sCodigoBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)
        .sCor = String(STRING_PRODUTO_COR, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        .sNomeReduzido = String(STRING_PRODUTO_NOME_REDUZIDO, 0)
        .sIPICodDIPI = String(STRING_PRODUTO_IPI_COD_DIPI, 0)
        .sIPICodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)
        .sModelo = String(STRING_PRODUTO_MODELO, 0)
        .sObsFisica = String(STRING_PRODUTO_OBS_FISICA, 0)
        .sSiglaUMCompra = String(STRING_PRODUTO_SIGLAUMCOMPRA, 0)
        .sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
        .sSiglaUMVenda = String(STRING_PRODUTO_SIGLAUMVENDA, 0)
        .sSubstituto1 = String(STRING_PRODUTO_SUBSTITUTO1, 0)
        .sSubstituto2 = String(STRING_PRODUTO_SUBSTITUTO2, 0)
        .sReferencia = String(STRING_PRODUTO_REFERENCIA, 0)
        .sFigura = String(STRING_NOME_ARQ_COMPLETO, 0)


        'Tenta ler
        lErro = Comando_Executar(lComando, "SELECT  Descricao, NomeReduzido, Gerencial, Ativo, Faturamento, ClasseUM, SiglaUMEstoque, SiglaUMCompra, SiglaUMVenda, Substituto1, Substituto2, Nivel, Tipo, Codigo, Modelo, PrazoValidade, CodigoBarras, EtiquetasCodBarras, PesoLiq, PesoBruto, Comprimento, Espessura, Largura, Cor, ObsFisica, Compras, PCP, KitBasico, KitInt, IPIAliquota, IPICodigo, IPICodDIPI, ControleEstoque, Apropriacao,TempoProducao,Embalagem, Referencia, Figura FROM Produtos WHERE Descricao = ? ", .sDescricao, .sNomeReduzido, .iGerencial, .iAtivo, .iFaturamento, .iClasseUM, .sSiglaUMEstoque, .sSiglaUMCompra, .sSiglaUMVenda, .sSubstituto1, .sSubstituto2, .iNivel, .iTipo, .sCodigo, .sModelo, .iPrazoValidade, .sCodigoBarras, .iEtiquetasCodBarras, .dPesoLiq, .dPesoBruto, .dComprimento, .dEspessura, .dLargura, .sCor, _
            .sObsFisica, .iCompras, .iPCP, .iKitBasico, .iKitInt, .dIPIAliquota, .sIPICodigo, .sIPICodDIPI, .iControleEstoque, .iApropriacaoCusto, .iTempoProducao, .iEmbalagem, .sReferencia, .sFigura, objProduto.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 67016

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 67017
        
        'Não encontra
        If lErro = AD_SQL_SEM_DADOS Then gError 67018

        'Guarda os dados no objProduto
        objProduto.dComprimento = .dComprimento
        objProduto.dEspessura = .dEspessura
        objProduto.dIPIAliquota = .dIPIAliquota
        objProduto.dLargura = .dLargura
        objProduto.dPesoBruto = .dPesoBruto
        objProduto.dPesoLiq = .dPesoLiq
        objProduto.iAtivo = .iAtivo
        objProduto.iClasseUM = .iClasseUM
        objProduto.iCompras = .iCompras
        objProduto.iEtiquetasCodBarras = .iEtiquetasCodBarras
        objProduto.iFaturamento = .iFaturamento
        objProduto.iGerencial = .iGerencial
        objProduto.iKitBasico = .iKitBasico
        objProduto.iKitInt = .iKitInt
        objProduto.iNivel = .iNivel
        objProduto.iPCP = .iPCP
        objProduto.iPrazoValidade = .iPrazoValidade
        objProduto.iTipo = .iTipo
        objProduto.sCodigo = .sCodigo
        objProduto.sCodigoBarras = .sCodigoBarras
        objProduto.sCor = .sCor
        objProduto.sDescricao = .sDescricao
        objProduto.sNomeReduzido = .sNomeReduzido
        objProduto.sIPICodDIPI = .sIPICodDIPI
        objProduto.sIPICodigo = .sIPICodigo
        objProduto.sModelo = .sModelo
        objProduto.sObsFisica = .sObsFisica
        objProduto.sSiglaUMCompra = .sSiglaUMCompra
        objProduto.sSiglaUMEstoque = .sSiglaUMEstoque
        objProduto.sSiglaUMVenda = .sSiglaUMVenda
        objProduto.sSubstituto1 = .sSubstituto1
        objProduto.sSubstituto2 = .sSubstituto2
        objProduto.iControleEstoque = .iControleEstoque
        objProduto.iApropriacaoCusto = .iApropriacaoCusto
        objProduto.iTempoProducao = .iTempoProducao
        objProduto.iEmbalagem = .iEmbalagem
        objProduto.sReferencia = .sReferencia
        objProduto.sFigura = .sFigura
        
    End With

    'fecha o comando
    Call Comando_Fechar(lComando)

    Produto_Le_Descricao = SUCESSO

    Exit Function

Erro_Produto_Le_Descricao:

    Produto_Le_Descricao = gErr

    Select Case gErr

        Case 67015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 67016, 67017
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS2", gErr, objProduto.sNomeReduzido)

        Case 67018  'Não encontrou o Produto
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152093)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function SldMesEstAlm2_Le_QuantConserto(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, ByVal iAlmoxarifado As Integer, dQuantConsertoAtual As Double) As Long
'calcula a quantidade de material em conserto nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double

On Error GoTo Erro_SldMesEstAlm2_Le_QuantConserto

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConserto, SaldoQuantConserto1, SaldoQuantConserto2, SaldoQuantConserto3, SaldoQuantConserto4, SaldoQuantConserto5, SaldoQuantConserto6, SaldoQuantConserto7, SaldoQuantConserto8, SaldoQuantConserto9, SaldoQuantConserto10, SaldoQuantConserto11, SaldoQuantConserto12 FROM SldMesEstAlm2 WHERE Ano=? AND Produto=? AND Almoxarifado = ?", _
                              tSldMesEst2.dQuantInicialConserto, tSldMesEst2.adSaldoQuantConserto(1), tSldMesEst2.adSaldoQuantConserto(2), tSldMesEst2.adSaldoQuantConserto(3), tSldMesEst2.adSaldoQuantConserto(4), tSldMesEst2.adSaldoQuantConserto(5), tSldMesEst2.adSaldoQuantConserto(6), tSldMesEst2.adSaldoQuantConserto(7), tSldMesEst2.adSaldoQuantConserto(8), tSldMesEst2.adSaldoQuantConserto(9), _
                              tSldMesEst2.adSaldoQuantConserto(10), tSldMesEst2.adSaldoQuantConserto(11), tSldMesEst2.adSaldoQuantConserto(12), iAno, sProduto, iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 71534
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71535
    
    dQuant = tSldMesEst2.dQuantInicialConserto
    
    For iIndice = 1 To iMes
    
        dQuant = dQuant + tSldMesEst2.adSaldoQuantConserto(iIndice)
        
    Next
    
    dQuantConsertoAtual = dQuant
    
    SldMesEstAlm2_Le_QuantConserto = SUCESSO

    Exit Function

Erro_SldMesEstAlm2_Le_QuantConserto:
    
    SldMesEstAlm2_Le_QuantConserto = gErr
    
    Select Case gErr
    
        Case 71534, 71535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM2", gErr, iAno, iFilialEmpresa, sProduto, iAlmoxarifado)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152094)
        
    End Select
    
    Exit Function

End Function

Function SldMesEstAlm2_Le_QuantDemo(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, ByVal iAlmoxarifado As Integer, dQuantDemoAtual As Double) As Long
'calcula a quantidade de material em Demonstração nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double

On Error GoTo Erro_SldMesEstAlm2_Le_QuantDemo

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialDemo, SaldoQuantDemo1, SaldoQuantDemo2, SaldoQuantDemo3, SaldoQuantDemo4, SaldoQuantDemo5, SaldoQuantDemo6, SaldoQuantDemo7, SaldoQuantDemo8, SaldoQuantDemo9, SaldoQuantDemo10, SaldoQuantDemo11, SaldoQuantDemo12 FROM SldMesEstAlm2 WHERE Ano=? AND Produto=? AND Almoxarifado=?", _
                              tSldMesEst2.dQuantInicialDemo, tSldMesEst2.adSaldoQuantDemo(1), tSldMesEst2.adSaldoQuantDemo(2), tSldMesEst2.adSaldoQuantDemo(3), tSldMesEst2.adSaldoQuantDemo(4), tSldMesEst2.adSaldoQuantDemo(5), tSldMesEst2.adSaldoQuantDemo(6), tSldMesEst2.adSaldoQuantDemo(7), tSldMesEst2.adSaldoQuantDemo(8), tSldMesEst2.adSaldoQuantDemo(9), _
                              tSldMesEst2.adSaldoQuantDemo(10), tSldMesEst2.adSaldoQuantDemo(11), tSldMesEst2.adSaldoQuantDemo(12), iAno, sProduto, iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 71536
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71537
    
    dQuant = tSldMesEst2.dQuantInicialDemo
    
    For iIndice = 1 To iMes
    
        dQuant = dQuant + tSldMesEst2.adSaldoQuantDemo(iIndice)
        
    Next
    
    dQuantDemoAtual = dQuant
    
    SldMesEstAlm2_Le_QuantDemo = SUCESSO

    Exit Function

Erro_SldMesEstAlm2_Le_QuantDemo:
    
    SldMesEstAlm2_Le_QuantDemo = gErr
    
    Select Case gErr
    
        Case 71536, 71537
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM2", gErr, iAno, iFilialEmpresa, sProduto, iAlmoxarifado)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152095)
        
    End Select
    
    Exit Function

End Function

Function SldMesEstAlm2_Le_QuantOutros(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, ByVal iAlmoxarifado As Integer, dQuantOutrosAtual As Double) As Long
'calcula a quantidade de Outros materiais nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double

On Error GoTo Erro_SldMesEstAlm2_Le_QuantOutros

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialOutros, SaldoQuantOutros1, SaldoQuantOutros2, SaldoQuantOutros3, SaldoQuantOutros4, SaldoQuantOutros5, SaldoQuantOutros6, SaldoQuantOutros7, SaldoQuantOutros8, SaldoQuantOutros9, SaldoQuantOutros10, SaldoQuantOutros11, SaldoQuantOutros12 FROM SldMesEstAlm2 WHERE Ano=? AND Produto=? AND Almoxarifado=?", _
                              tSldMesEst2.dQuantInicialOutros, tSldMesEst2.adSaldoQuantOutros(1), tSldMesEst2.adSaldoQuantOutros(2), tSldMesEst2.adSaldoQuantOutros(3), tSldMesEst2.adSaldoQuantOutros(4), tSldMesEst2.adSaldoQuantOutros(5), tSldMesEst2.adSaldoQuantOutros(6), tSldMesEst2.adSaldoQuantOutros(7), tSldMesEst2.adSaldoQuantOutros(8), tSldMesEst2.adSaldoQuantOutros(9), _
                              tSldMesEst2.adSaldoQuantOutros(10), tSldMesEst2.adSaldoQuantOutros(11), tSldMesEst2.adSaldoQuantOutros(12), iAno, sProduto, iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 71538
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71539
    
    dQuant = tSldMesEst2.dQuantInicialOutros
    
    For iIndice = 1 To iMes
    
        dQuant = dQuant + tSldMesEst2.adSaldoQuantOutros(iIndice)
        
    Next
    
    dQuantOutrosAtual = dQuant
    
    SldMesEstAlm2_Le_QuantOutros = SUCESSO

    Exit Function

Erro_SldMesEstAlm2_Le_QuantOutros:
    
    SldMesEstAlm2_Le_QuantOutros = gErr
    
    Select Case gErr
    
        Case 71538, 71539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM2", gErr, iAno, iFilialEmpresa, sProduto, iAlmoxarifado)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152096)
        
    End Select
    
    Exit Function

End Function

Function SldMesEstAlm2_Le_QuantBenef(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, ByVal iAlmoxarifado As Integer, dQuantBenefAtual As Double) As Long
'calcula a quantidade de material em Beneficiamento nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double

On Error GoTo Erro_SldMesEstAlm2_Le_QuantBenef

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialBenef, SaldoQuantBenef1, SaldoQuantBenef2, SaldoQuantBenef3, SaldoQuantBenef4, SaldoQuantBenef5, SaldoQuantBenef6, SaldoQuantBenef7, SaldoQuantBenef8, SaldoQuantBenef9, SaldoQuantBenef10, SaldoQuantBenef11, SaldoQuantBenef12 FROM SldMesEstAlm2 WHERE Ano=? AND Produto=? AND Almoxarifado=?", _
                              tSldMesEst2.dQuantInicialBenef, tSldMesEst2.adSaldoQuantBenef(1), tSldMesEst2.adSaldoQuantBenef(2), tSldMesEst2.adSaldoQuantBenef(3), tSldMesEst2.adSaldoQuantBenef(4), tSldMesEst2.adSaldoQuantBenef(5), tSldMesEst2.adSaldoQuantBenef(6), tSldMesEst2.adSaldoQuantBenef(7), tSldMesEst2.adSaldoQuantBenef(8), tSldMesEst2.adSaldoQuantBenef(9), _
                              tSldMesEst2.adSaldoQuantBenef(10), tSldMesEst2.adSaldoQuantBenef(11), tSldMesEst2.adSaldoQuantBenef(12), iAno, sProduto, iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 71540
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71541
    
    dQuant = tSldMesEst2.dQuantInicialBenef
    
    For iIndice = 1 To iMes
    
        dQuant = dQuant + tSldMesEst2.adSaldoQuantBenef(iIndice)
        
    Next
    
    dQuantBenefAtual = dQuant
    
    SldMesEstAlm2_Le_QuantBenef = SUCESSO

    Exit Function

Erro_SldMesEstAlm2_Le_QuantBenef:
    
    SldMesEstAlm2_Le_QuantBenef = gErr
    
    Select Case gErr
    
        Case 71540, 71541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM2", gErr, iAno, iFilialEmpresa, sProduto, iAlmoxarifado)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152097)
        
    End Select
    
    Exit Function

End Function

Function SldMesEstAlm2_Le_QuantConsig(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, ByVal iAlmoxarifado As Integer, dQuantConsigAtual As Double) As Long
'calcula a quantidade de material em Consignação nosso em poder de terceiros.

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double

On Error GoTo Erro_SldMesEstAlm2_Le_QuantConsig

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConsig, SaldoQuantConsig1, SaldoQuantConsig2, SaldoQuantConsig3, SaldoQuantConsig4, SaldoQuantConsig5, SaldoQuantConsig6, SaldoQuantConsig7, SaldoQuantConsig8, SaldoQuantConsig9, SaldoQuantConsig10, SaldoQuantConsig11, SaldoQuantConsig12 FROM SldMesEstAlm2 WHERE Ano=? AND Produto=? AND Almoxarifado=?", _
                              tSldMesEst2.dQuantInicialConsig, tSldMesEst2.adSaldoQuantConsig(1), tSldMesEst2.adSaldoQuantConsig(2), tSldMesEst2.adSaldoQuantConsig(3), tSldMesEst2.adSaldoQuantConsig(4), tSldMesEst2.adSaldoQuantConsig(5), tSldMesEst2.adSaldoQuantConsig(6), tSldMesEst2.adSaldoQuantConsig(7), tSldMesEst2.adSaldoQuantConsig(8), tSldMesEst2.adSaldoQuantConsig(9), _
                              tSldMesEst2.adSaldoQuantConsig(10), tSldMesEst2.adSaldoQuantConsig(11), tSldMesEst2.adSaldoQuantConsig(12), iAno, sProduto, iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 71542
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71543
    
    dQuant = tSldMesEst2.dQuantInicialConsig
    
    For iIndice = 1 To iMes
    
        dQuant = dQuant + tSldMesEst2.adSaldoQuantConsig(iIndice)
        
    Next
    
    dQuantConsigAtual = dQuant
    
    SldMesEstAlm2_Le_QuantConsig = SUCESSO

    Exit Function

Erro_SldMesEstAlm2_Le_QuantConsig:
    
    SldMesEstAlm2_Le_QuantConsig = gErr
    
    Select Case gErr
    
        Case 71542, 71543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM2", gErr, iAno, iFilialEmpresa, sProduto, iAlmoxarifado)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152098)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_Le_QuantVendConsig(ByVal objItemMovEstoque As ClassItemMovEstoque) As Long
'Verifica se a NF movimentou produto consig e retorna a quantidade

Dim lComando As Long
Dim lErro As Long
Dim dQuantidade As Double
Dim sComando_SQL As String

On Error GoTo Erro_MovEstoque_Le_QuantVendConsig

    sComando_SQL = "SELECT Quantidade FROM MovimentoEstoque WHERE Produto=? AND Almoxarifado=? AND TipoMov=? AND DocOrigem=? AND TipoNumIntDocOrigem=?"
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 79000

    'procura o movimento que está sendo gravado a partir do produto em GridAlocacao
    lErro = Comando_Executar(lComando, sComando_SQL, dQuantidade, objItemMovEstoque.sProduto, objItemMovEstoque.iAlmoxarifado, objItemMovEstoque.iTipoMov, objItemMovEstoque.sDocOrigem, objItemMovEstoque.iTipoNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 79001
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79002
    
    'se o movimento não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 79003
    
    'informa a quantidade encontrada
    objItemMovEstoque.dQuantidade = dQuantidade
    
    Call Comando_Fechar(lComando)
    
    MovEstoque_Le_QuantVendConsig = SUCESSO
    
    Exit Function

Erro_MovEstoque_Le_QuantVendConsig:

    MovEstoque_Le_QuantVendConsig = gErr
    
    Select Case gErr
    
        Case 79000
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 79001, 79002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
        
        Case 79003
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152099)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function SldMesEst1_Le_QuantValorConsig3(ByVal lComando As Long, ByVal objSldMesEst1 As ClassSldMesEst1) As Long

Dim tSldMesEst1 As typeSldMesEst1
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double
Dim dValor As Double

On Error GoTo Erro_SldMesEst1_Le_QuantValorConsig3

    'seleciona o registro de SldMesEst1 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConsig3, ValorInicialConsig3, SaldoQuantConsig31, SaldoValorConsig31, SaldoQuantConsig32, SaldoValorConsig32, SaldoQuantConsig33, SaldoValorConsig33, SaldoQuantConsig34, SaldoValorConsig34, SaldoQuantConsig35, SaldoValorConsig35, SaldoQuantConsig36, SaldoValorConsig36, SaldoQuantConsig37, SaldoValorConsig37, SaldoQuantConsig38, SaldoValorConsig38, SaldoQuantConsig39, SaldoValorConsig39, SaldoQuantConsig310, SaldoValorConsig310, SaldoQuantConsig311, SaldoValorConsig311, SaldoQuantConsig312, SaldoValorConsig312  FROM SldMesEst1 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst1.dQuantInicialConsig3, tSldMesEst1.dValorInicialConsig3, tSldMesEst1.adSaldoQuantConsig3(1), tSldMesEst1.adSaldoValorConsig3(1), tSldMesEst1.adSaldoQuantConsig3(2), tSldMesEst1.adSaldoValorConsig3(2), tSldMesEst1.adSaldoQuantConsig3(3), tSldMesEst1.adSaldoValorConsig3(3), tSldMesEst1.adSaldoQuantConsig3(4), tSldMesEst1.adSaldoValorConsig3(4), tSldMesEst1.adSaldoQuantConsig3(5), tSldMesEst1.adSaldoValorConsig3(5), tSldMesEst1.adSaldoQuantConsig3(6), tSldMesEst1.adSaldoValorConsig3(6), tSldMesEst1.adSaldoQuantConsig3(7), tSldMesEst1.adSaldoValorConsig3(7), tSldMesEst1.adSaldoQuantConsig3(8), tSldMesEst1.adSaldoValorConsig3(8), tSldMesEst1.adSaldoQuantConsig3(9), tSldMesEst1.adSaldoValorConsig3(9), _
                              tSldMesEst1.adSaldoQuantConsig3(10), tSldMesEst1.adSaldoValorConsig3(10), tSldMesEst1.adSaldoQuantConsig3(11), tSldMesEst1.adSaldoValorConsig3(11), tSldMesEst1.adSaldoQuantConsig3(12), tSldMesEst1.adSaldoValorConsig3(12), objSldMesEst1.iAno, objSldMesEst1.iFilialEmpresa, objSldMesEst1.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71608
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71609
    
    objSldMesEst1.dQuantInicialConsig3 = tSldMesEst1.dQuantInicialConsig3
    objSldMesEst1.dValorInicialConsig3 = tSldMesEst1.dValorInicialConsig3
    
    For iIndice = 1 To 12
    
        objSldMesEst1.dSaldoQuantConsig3(iIndice) = tSldMesEst1.adSaldoQuantConsig3(iIndice)
        objSldMesEst1.dSaldoValorConsig3(iIndice) = tSldMesEst1.adSaldoValorConsig3(iIndice)
        
    Next
    
    SldMesEst1_Le_QuantValorConsig3 = SUCESSO

    Exit Function

Erro_SldMesEst1_Le_QuantValorConsig3:
    
    SldMesEst1_Le_QuantValorConsig3 = gErr
    
    Select Case gErr
    
        Case 71608, 71609
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, objSldMesEst1.iAno, objSldMesEst1.iFilialEmpresa, objSldMesEst1.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152100)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_QuantValorDemo3(ByVal lComando As Long, ByVal objSldMesEst1 As ClassSldMesEst1) As Long

Dim tSldMesEst1 As typeSldMesEst1
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double
Dim dValor As Double

On Error GoTo Erro_SldMesEst1_Le_QuantValorDemo3

    'seleciona o registro de SldMesEst1 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialDemo3, ValorInicialDemo3, SaldoQuantDemo31, SaldoValorDemo31, SaldoQuantDemo32, SaldoValorDemo32, SaldoQuantDemo33, SaldoValorDemo33, SaldoQuantDemo34, SaldoValorDemo34, SaldoQuantDemo35, SaldoValorDemo35, SaldoQuantDemo36, SaldoValorDemo36, SaldoQuantDemo37, SaldoValorDemo37, SaldoQuantDemo38, SaldoValorDemo38, SaldoQuantDemo39, SaldoValorDemo39, SaldoQuantDemo310, SaldoValorDemo310, SaldoQuantDemo311, SaldoValorDemo311, SaldoQuantDemo312, SaldoValorDemo312  FROM SldMesEst1 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst1.dQuantInicialDemo3, tSldMesEst1.dValorInicialDemo3, tSldMesEst1.adSaldoQuantDemo3(1), tSldMesEst1.adSaldoValorDemo3(1), tSldMesEst1.adSaldoQuantDemo3(2), tSldMesEst1.adSaldoValorDemo3(2), tSldMesEst1.adSaldoQuantDemo3(3), tSldMesEst1.adSaldoValorDemo3(3), tSldMesEst1.adSaldoQuantDemo3(4), tSldMesEst1.adSaldoValorDemo3(4), tSldMesEst1.adSaldoQuantDemo3(5), tSldMesEst1.adSaldoValorDemo3(5), tSldMesEst1.adSaldoQuantDemo3(6), tSldMesEst1.adSaldoValorDemo3(6), tSldMesEst1.adSaldoQuantDemo3(7), tSldMesEst1.adSaldoValorDemo3(7), tSldMesEst1.adSaldoQuantDemo3(8), tSldMesEst1.adSaldoValorDemo3(8), tSldMesEst1.adSaldoQuantDemo3(9), tSldMesEst1.adSaldoValorDemo3(9), _
                              tSldMesEst1.adSaldoQuantDemo3(10), tSldMesEst1.adSaldoValorDemo3(10), tSldMesEst1.adSaldoQuantDemo3(11), tSldMesEst1.adSaldoValorDemo3(11), tSldMesEst1.adSaldoQuantDemo3(12), tSldMesEst1.adSaldoValorDemo3(12), objSldMesEst1.iAno, objSldMesEst1.iFilialEmpresa, objSldMesEst1.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71610
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71611
    
    objSldMesEst1.dQuantInicialDemo3 = tSldMesEst1.dQuantInicialDemo3
    objSldMesEst1.dValorInicialDemo3 = tSldMesEst1.dValorInicialDemo3
    
    For iIndice = 1 To 12
    
        objSldMesEst1.dSaldoQuantDemo3(iIndice) = tSldMesEst1.adSaldoQuantDemo3(iIndice)
        objSldMesEst1.dSaldoValorDemo3(iIndice) = tSldMesEst1.adSaldoValorDemo3(iIndice)
        
    Next
    
    SldMesEst1_Le_QuantValorDemo3 = SUCESSO

    Exit Function

Erro_SldMesEst1_Le_QuantValorDemo3:
    
    SldMesEst1_Le_QuantValorDemo3 = gErr
    
    Select Case gErr
    
        Case 71610, 71611
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, objSldMesEst1.iAno, objSldMesEst1.iFilialEmpresa, objSldMesEst1.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152101)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_QuantValorConserto3(ByVal lComando As Long, ByVal objSldMesEst1 As ClassSldMesEst1) As Long

Dim tSldMesEst1 As typeSldMesEst1
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double
Dim dValor As Double

On Error GoTo Erro_SldMesEst1_Le_QuantValorConserto3

    'seleciona o registro de SldMesEst1 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConserto3, ValorInicialConserto3, SaldoQuantConserto31, SaldoValorConserto31, SaldoQuantConserto32, SaldoValorConserto32, SaldoQuantConserto33, SaldoValorConserto33, SaldoQuantConserto34, SaldoValorConserto34, SaldoQuantConserto35, SaldoValorConserto35, SaldoQuantConserto36, SaldoValorConserto36, SaldoQuantConserto37, SaldoValorConserto37, SaldoQuantConserto38, SaldoValorConserto38, SaldoQuantConserto39, SaldoValorConserto39, SaldoQuantConserto310, SaldoValorConserto310, SaldoQuantConserto311, SaldoValorConserto311, SaldoQuantConserto312, SaldoValorConserto312  FROM SldMesEst1 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst1.dQuantInicialConserto3, tSldMesEst1.dValorInicialConserto3, tSldMesEst1.adSaldoQuantConserto3(1), tSldMesEst1.adSaldoValorConserto3(1), tSldMesEst1.adSaldoQuantConserto3(2), tSldMesEst1.adSaldoValorConserto3(2), tSldMesEst1.adSaldoQuantConserto3(3), tSldMesEst1.adSaldoValorConserto3(3), tSldMesEst1.adSaldoQuantConserto3(4), tSldMesEst1.adSaldoValorConserto3(4), tSldMesEst1.adSaldoQuantConserto3(5), tSldMesEst1.adSaldoValorConserto3(5), tSldMesEst1.adSaldoQuantConserto3(6), tSldMesEst1.adSaldoValorConserto3(6), tSldMesEst1.adSaldoQuantConserto3(7), tSldMesEst1.adSaldoValorConserto3(7), tSldMesEst1.adSaldoQuantConserto3(8), tSldMesEst1.adSaldoValorConserto3(8), tSldMesEst1.adSaldoQuantConserto3(9), tSldMesEst1.adSaldoValorConserto3(9), _
                              tSldMesEst1.adSaldoQuantConserto3(10), tSldMesEst1.adSaldoValorConserto3(10), tSldMesEst1.adSaldoQuantConserto3(11), tSldMesEst1.adSaldoValorConserto3(11), tSldMesEst1.adSaldoQuantConserto3(12), tSldMesEst1.adSaldoValorConserto3(12), objSldMesEst1.iAno, objSldMesEst1.iFilialEmpresa, objSldMesEst1.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71606
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71607
    
    objSldMesEst1.dQuantInicialConserto3 = tSldMesEst1.dQuantInicialConserto3
    objSldMesEst1.dValorInicialConserto3 = tSldMesEst1.dValorInicialConserto3
    
    For iIndice = 1 To 12
    
        objSldMesEst1.dSaldoQuantConserto3(iIndice) = tSldMesEst1.adSaldoQuantConserto3(iIndice)
        objSldMesEst1.dSaldoValorConserto3(iIndice) = tSldMesEst1.adSaldoValorConserto3(iIndice)
        
    Next
    
    SldMesEst1_Le_QuantValorConserto3 = SUCESSO

    Exit Function

Erro_SldMesEst1_Le_QuantValorConserto3:
    
    SldMesEst1_Le_QuantValorConserto3 = gErr
    
    Select Case gErr
    
        Case 71606, 71607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, objSldMesEst1.iAno, objSldMesEst1.iFilialEmpresa, objSldMesEst1.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152102)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_QuantValorOutros3(ByVal lComando As Long, ByVal objSldMesEst1 As ClassSldMesEst1) As Long

Dim tSldMesEst1 As typeSldMesEst1
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double
Dim dValor As Double

On Error GoTo Erro_SldMesEst1_Le_QuantValorOutros3

    'seleciona o registro de SldMesEst1 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialOutros3, ValorInicialOutros3, SaldoQuantOutros31, SaldoValorOutros31, SaldoQuantOutros32, SaldoValorOutros32, SaldoQuantOutros33, SaldoValorOutros33, SaldoQuantOutros34, SaldoValorOutros34, SaldoQuantOutros35, SaldoValorOutros35, SaldoQuantOutros36, SaldoValorOutros36, SaldoQuantOutros37, SaldoValorOutros37, SaldoQuantOutros38, SaldoValorOutros38, SaldoQuantOutros39, SaldoValorOutros39, SaldoQuantOutros310, SaldoValorOutros310, SaldoQuantOutros311, SaldoValorOutros311, SaldoQuantOutros312, SaldoValorOutros312  FROM SldMesEst1 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst1.dQuantInicialOutros3, tSldMesEst1.dValorInicialOutros3, tSldMesEst1.adSaldoQuantOutros3(1), tSldMesEst1.adSaldoValorOutros3(1), tSldMesEst1.adSaldoQuantOutros3(2), tSldMesEst1.adSaldoValorOutros3(2), tSldMesEst1.adSaldoQuantOutros3(3), tSldMesEst1.adSaldoValorOutros3(3), tSldMesEst1.adSaldoQuantOutros3(4), tSldMesEst1.adSaldoValorOutros3(4), tSldMesEst1.adSaldoQuantOutros3(5), tSldMesEst1.adSaldoValorOutros3(5), tSldMesEst1.adSaldoQuantOutros3(6), tSldMesEst1.adSaldoValorOutros3(6), tSldMesEst1.adSaldoQuantOutros3(7), tSldMesEst1.adSaldoValorOutros3(7), tSldMesEst1.adSaldoQuantOutros3(8), tSldMesEst1.adSaldoValorOutros3(8), tSldMesEst1.adSaldoQuantOutros3(9), tSldMesEst1.adSaldoValorOutros3(9), _
                              tSldMesEst1.adSaldoQuantOutros3(10), tSldMesEst1.adSaldoValorOutros3(10), tSldMesEst1.adSaldoQuantOutros3(11), tSldMesEst1.adSaldoValorOutros3(11), tSldMesEst1.adSaldoQuantOutros3(12), tSldMesEst1.adSaldoValorOutros3(12), objSldMesEst1.iAno, objSldMesEst1.iFilialEmpresa, objSldMesEst1.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71612
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71613
    
    objSldMesEst1.dQuantInicialOutros3 = tSldMesEst1.dQuantInicialOutros3
    objSldMesEst1.dValorInicialOutros3 = tSldMesEst1.dValorInicialOutros3
    
    For iIndice = 1 To 12
    
        objSldMesEst1.dSaldoQuantOutros3(iIndice) = tSldMesEst1.adSaldoQuantOutros3(iIndice)
        objSldMesEst1.dSaldoValorOutros3(iIndice) = tSldMesEst1.adSaldoValorOutros3(iIndice)
        
    Next
    
    SldMesEst1_Le_QuantValorOutros3 = SUCESSO

    Exit Function

Erro_SldMesEst1_Le_QuantValorOutros3:
    
    SldMesEst1_Le_QuantValorOutros3 = gErr
    
    Select Case gErr
    
        Case 71612, 71613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, objSldMesEst1.iAno, objSldMesEst1.iFilialEmpresa, objSldMesEst1.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152103)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_QuantValorBenef3(ByVal lComando As Long, ByVal objSldMesEst1 As ClassSldMesEst1) As Long

Dim tSldMesEst1 As typeSldMesEst1
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double
Dim dValor As Double

On Error GoTo Erro_SldMesEst1_Le_QuantValorBenef3

    'seleciona o registro de SldMesEst1 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialBenef3, ValorInicialBenef3, SaldoQuantBenef31, SaldoValorBenef31, SaldoQuantBenef32, SaldoValorBenef32, SaldoQuantBenef33, SaldoValorBenef33, SaldoQuantBenef34, SaldoValorBenef34, SaldoQuantBenef35, SaldoValorBenef35, SaldoQuantBenef36, SaldoValorBenef36, SaldoQuantBenef37, SaldoValorBenef37, SaldoQuantBenef38, SaldoValorBenef38, SaldoQuantBenef39, SaldoValorBenef39, SaldoQuantBenef310, SaldoValorBenef310, SaldoQuantBenef311, SaldoValorBenef311, SaldoQuantBenef312, SaldoValorBenef312  FROM SldMesEst1 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst1.dQuantInicialBenef3, tSldMesEst1.dValorInicialBenef3, tSldMesEst1.adSaldoQuantBenef3(1), tSldMesEst1.adSaldoValorBenef3(1), tSldMesEst1.adSaldoQuantBenef3(2), tSldMesEst1.adSaldoValorBenef3(2), tSldMesEst1.adSaldoQuantBenef3(3), tSldMesEst1.adSaldoValorBenef3(3), tSldMesEst1.adSaldoQuantBenef3(4), tSldMesEst1.adSaldoValorBenef3(4), tSldMesEst1.adSaldoQuantBenef3(5), tSldMesEst1.adSaldoValorBenef3(5), tSldMesEst1.adSaldoQuantBenef3(6), tSldMesEst1.adSaldoValorBenef3(6), tSldMesEst1.adSaldoQuantBenef3(7), tSldMesEst1.adSaldoValorBenef3(7), tSldMesEst1.adSaldoQuantBenef3(8), tSldMesEst1.adSaldoValorBenef3(8), tSldMesEst1.adSaldoQuantBenef3(9), tSldMesEst1.adSaldoValorBenef3(9), _
                              tSldMesEst1.adSaldoQuantBenef3(10), tSldMesEst1.adSaldoValorBenef3(10), tSldMesEst1.adSaldoQuantBenef3(11), tSldMesEst1.adSaldoValorBenef3(11), tSldMesEst1.adSaldoQuantBenef3(12), tSldMesEst1.adSaldoValorBenef3(12), objSldMesEst1.iAno, objSldMesEst1.iFilialEmpresa, objSldMesEst1.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71612
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71613
    
    objSldMesEst1.dQuantInicialBenef3 = tSldMesEst1.dQuantInicialBenef3
    objSldMesEst1.dValorInicialBenef3 = tSldMesEst1.dValorInicialBenef3
    
    For iIndice = 1 To 12
    
        objSldMesEst1.dSaldoQuantBenef3(iIndice) = tSldMesEst1.adSaldoQuantBenef3(iIndice)
        objSldMesEst1.dSaldoValorBenef3(iIndice) = tSldMesEst1.adSaldoValorBenef3(iIndice)
        
    Next
    
    SldMesEst1_Le_QuantValorBenef3 = SUCESSO

    Exit Function

Erro_SldMesEst1_Le_QuantValorBenef3:
    
    SldMesEst1_Le_QuantValorBenef3 = gErr
    
    Select Case gErr
    
        Case 71612, 71613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST11", gErr, objSldMesEst1.iAno, objSldMesEst1.iFilialEmpresa, objSldMesEst1.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152104)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantValorConsig(ByVal lComando As Long, ByVal objSldMesEst2 As ClassSldMesEst2) As Long

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double
Dim dValor As Double

On Error GoTo Erro_SldMesEst2_Le_QuantValorConsig

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConsig, ValorInicialConsig, SaldoQuantConsig1, SaldoValorConsig1, SaldoQuantConsig2, SaldoValorConsig2, SaldoQuantConsig3, SaldoValorConsig3, SaldoQuantConsig4, SaldoValorConsig4, SaldoQuantConsig5, SaldoValorConsig5, SaldoQuantConsig6, SaldoValorConsig6, SaldoQuantConsig7, SaldoValorConsig7, SaldoQuantConsig8, SaldoValorConsig8, SaldoQuantConsig9, SaldoValorConsig9, SaldoQuantConsig10, SaldoValorConsig10, SaldoQuantConsig11, SaldoValorConsig11, SaldoQuantConsig12, SaldoValorConsig12  FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialConsig, tSldMesEst2.dValorInicialConsig, tSldMesEst2.adSaldoQuantConsig(1), tSldMesEst2.adSaldoValorConsig(1), tSldMesEst2.adSaldoQuantConsig(2), tSldMesEst2.adSaldoValorConsig(2), tSldMesEst2.adSaldoQuantConsig(3), tSldMesEst2.adSaldoValorConsig(3), tSldMesEst2.adSaldoQuantConsig(4), tSldMesEst2.adSaldoValorConsig(4), tSldMesEst2.adSaldoQuantConsig(5), tSldMesEst2.adSaldoValorConsig(5), tSldMesEst2.adSaldoQuantConsig(6), tSldMesEst2.adSaldoValorConsig(6), tSldMesEst2.adSaldoQuantConsig(7), tSldMesEst2.adSaldoValorConsig(7), tSldMesEst2.adSaldoQuantConsig(8), tSldMesEst2.adSaldoValorConsig(8), tSldMesEst2.adSaldoQuantConsig(9), tSldMesEst2.adSaldoValorConsig(9), _
                              tSldMesEst2.adSaldoQuantConsig(10), tSldMesEst2.adSaldoValorConsig(10), tSldMesEst2.adSaldoQuantConsig(11), tSldMesEst2.adSaldoValorConsig(11), tSldMesEst2.adSaldoQuantConsig(12), tSldMesEst2.adSaldoValorConsig(12), objSldMesEst2.iAno, objSldMesEst2.iFilialEmpresa, objSldMesEst2.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71619
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71620
    
    objSldMesEst2.dQuantInicialConsig = tSldMesEst2.dQuantInicialConsig
    objSldMesEst2.dValorInicialConsig = tSldMesEst2.dValorInicialConsig
    
    For iIndice = 1 To 12
    
        objSldMesEst2.dSaldoQuantConsig(iIndice) = tSldMesEst2.adSaldoQuantConsig(iIndice)
        objSldMesEst2.dSaldoValorConsig(iIndice) = tSldMesEst2.adSaldoValorConsig(iIndice)
        
    Next
    
    SldMesEst2_Le_QuantValorConsig = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_QuantValorConsig:
    
    SldMesEst2_Le_QuantValorConsig = gErr
    
    Select Case gErr
    
        Case 71619, 71620
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, objSldMesEst2.iAno, objSldMesEst2.iFilialEmpresa, objSldMesEst2.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152105)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantValorDemo(ByVal lComando As Long, ByVal objSldMesEst2 As ClassSldMesEst2) As Long

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double
Dim dValor As Double

On Error GoTo Erro_SldMesEst2_Le_QuantValorDemo

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialDemo, ValorInicialDemo, SaldoQuantDemo1, SaldoValorDemo1, SaldoQuantDemo2, SaldoValorDemo2, SaldoQuantDemo3, SaldoValorDemo3, SaldoQuantDemo4, SaldoValorDemo4, SaldoQuantDemo5, SaldoValorDemo5, SaldoQuantDemo6, SaldoValorDemo6, SaldoQuantDemo7, SaldoValorDemo7, SaldoQuantDemo8, SaldoValorDemo8, SaldoQuantDemo9, SaldoValorDemo9, SaldoQuantDemo10, SaldoValorDemo10, SaldoQuantDemo11, SaldoValorDemo11, SaldoQuantDemo12, SaldoValorDemo12  FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialDemo, tSldMesEst2.dValorInicialDemo, tSldMesEst2.adSaldoQuantDemo(1), tSldMesEst2.adSaldoValorDemo(1), tSldMesEst2.adSaldoQuantDemo(2), tSldMesEst2.adSaldoValorDemo(2), tSldMesEst2.adSaldoQuantDemo(3), tSldMesEst2.adSaldoValorDemo(3), tSldMesEst2.adSaldoQuantDemo(4), tSldMesEst2.adSaldoValorDemo(4), tSldMesEst2.adSaldoQuantDemo(5), tSldMesEst2.adSaldoValorDemo(5), tSldMesEst2.adSaldoQuantDemo(6), tSldMesEst2.adSaldoValorDemo(6), tSldMesEst2.adSaldoQuantDemo(7), tSldMesEst2.adSaldoValorDemo(7), tSldMesEst2.adSaldoQuantDemo(8), tSldMesEst2.adSaldoValorDemo(8), tSldMesEst2.adSaldoQuantDemo(9), tSldMesEst2.adSaldoValorDemo(9), _
                              tSldMesEst2.adSaldoQuantDemo(10), tSldMesEst2.adSaldoValorDemo(10), tSldMesEst2.adSaldoQuantDemo(11), tSldMesEst2.adSaldoValorDemo(11), tSldMesEst2.adSaldoQuantDemo(12), tSldMesEst2.adSaldoValorDemo(12), objSldMesEst2.iAno, objSldMesEst2.iFilialEmpresa, objSldMesEst2.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71621
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71622
    
    objSldMesEst2.dQuantInicialDemo = tSldMesEst2.dQuantInicialDemo
    objSldMesEst2.dValorInicialDemo = tSldMesEst2.dValorInicialDemo
    
    For iIndice = 1 To 12
    
        objSldMesEst2.dSaldoQuantDemo(iIndice) = tSldMesEst2.adSaldoQuantDemo(iIndice)
        objSldMesEst2.dSaldoValorDemo(iIndice) = tSldMesEst2.adSaldoValorDemo(iIndice)
        
    Next
    
    SldMesEst2_Le_QuantValorDemo = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_QuantValorDemo:
    
    SldMesEst2_Le_QuantValorDemo = gErr
    
    Select Case gErr
    
        Case 71621, 71622
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, objSldMesEst2.iAno, objSldMesEst2.iFilialEmpresa, objSldMesEst2.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152106)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantValorConserto(ByVal lComando As Long, ByVal objSldMesEst2 As ClassSldMesEst2) As Long

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double
Dim dValor As Double

On Error GoTo Erro_SldMesEst2_Le_QuantValorConserto

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialConserto, ValorInicialConserto, SaldoQuantConserto1, SaldoValorConserto1, SaldoQuantConserto2, SaldoValorConserto2, SaldoQuantConserto3, SaldoValorConserto3, SaldoQuantConserto4, SaldoValorConserto4, SaldoQuantConserto5, SaldoValorConserto5, SaldoQuantConserto6, SaldoValorConserto6, SaldoQuantConserto7, SaldoValorConserto7, SaldoQuantConserto8, SaldoValorConserto8, SaldoQuantConserto9, SaldoValorConserto9, SaldoQuantConserto10, SaldoValorConserto10, SaldoQuantConserto11, SaldoValorConserto11, SaldoQuantConserto12, SaldoValorConserto12  FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialConserto, tSldMesEst2.dValorInicialConserto, tSldMesEst2.adSaldoQuantConserto(1), tSldMesEst2.adSaldoValorConserto(1), tSldMesEst2.adSaldoQuantConserto(2), tSldMesEst2.adSaldoValorConserto(2), tSldMesEst2.adSaldoQuantConserto(3), tSldMesEst2.adSaldoValorConserto(3), tSldMesEst2.adSaldoQuantConserto(4), tSldMesEst2.adSaldoValorConserto(4), tSldMesEst2.adSaldoQuantConserto(5), tSldMesEst2.adSaldoValorConserto(5), tSldMesEst2.adSaldoQuantConserto(6), tSldMesEst2.adSaldoValorConserto(6), tSldMesEst2.adSaldoQuantConserto(7), tSldMesEst2.adSaldoValorConserto(7), tSldMesEst2.adSaldoQuantConserto(8), tSldMesEst2.adSaldoValorConserto(8), tSldMesEst2.adSaldoQuantConserto(9), tSldMesEst2.adSaldoValorConserto(9), _
                              tSldMesEst2.adSaldoQuantConserto(10), tSldMesEst2.adSaldoValorConserto(10), tSldMesEst2.adSaldoQuantConserto(11), tSldMesEst2.adSaldoValorConserto(11), tSldMesEst2.adSaldoQuantConserto(12), tSldMesEst2.adSaldoValorConserto(12), objSldMesEst2.iAno, objSldMesEst2.iFilialEmpresa, objSldMesEst2.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71623
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71624
    
    objSldMesEst2.dQuantInicialConserto = tSldMesEst2.dQuantInicialConserto
    objSldMesEst2.dValorInicialConserto = tSldMesEst2.dValorInicialConserto
    
    For iIndice = 1 To 12
    
        objSldMesEst2.dSaldoQuantConserto(iIndice) = tSldMesEst2.adSaldoQuantConserto(iIndice)
        objSldMesEst2.dSaldoValorConserto(iIndice) = tSldMesEst2.adSaldoValorConserto(iIndice)
        
    Next
    
    SldMesEst2_Le_QuantValorConserto = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_QuantValorConserto:
    
    SldMesEst2_Le_QuantValorConserto = gErr
    
    Select Case gErr
    
        Case 71623, 71624
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, objSldMesEst2.iAno, objSldMesEst2.iFilialEmpresa, objSldMesEst2.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152107)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantValorOutros(ByVal lComando As Long, ByVal objSldMesEst2 As ClassSldMesEst2) As Long

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double
Dim dValor As Double

On Error GoTo Erro_SldMesEst2_Le_QuantValorOutros

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialOutros, ValorInicialOutros, SaldoQuantOutros1, SaldoValorOutros1, SaldoQuantOutros2, SaldoValorOutros2, SaldoQuantOutros3, SaldoValorOutros3, SaldoQuantOutros4, SaldoValorOutros4, SaldoQuantOutros5, SaldoValorOutros5, SaldoQuantOutros6, SaldoValorOutros6, SaldoQuantOutros7, SaldoValorOutros7, SaldoQuantOutros8, SaldoValorOutros8, SaldoQuantOutros9, SaldoValorOutros9, SaldoQuantOutros10, SaldoValorOutros10, SaldoQuantOutros11, SaldoValorOutros11, SaldoQuantOutros12, SaldoValorOutros12  FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialOutros, tSldMesEst2.dValorInicialOutros, tSldMesEst2.adSaldoQuantOutros(1), tSldMesEst2.adSaldoValorOutros(1), tSldMesEst2.adSaldoQuantOutros(2), tSldMesEst2.adSaldoValorOutros(2), tSldMesEst2.adSaldoQuantOutros(3), tSldMesEst2.adSaldoValorOutros(3), tSldMesEst2.adSaldoQuantOutros(4), tSldMesEst2.adSaldoValorOutros(4), tSldMesEst2.adSaldoQuantOutros(5), tSldMesEst2.adSaldoValorOutros(5), tSldMesEst2.adSaldoQuantOutros(6), tSldMesEst2.adSaldoValorOutros(6), tSldMesEst2.adSaldoQuantOutros(7), tSldMesEst2.adSaldoValorOutros(7), tSldMesEst2.adSaldoQuantOutros(8), tSldMesEst2.adSaldoValorOutros(8), tSldMesEst2.adSaldoQuantOutros(9), tSldMesEst2.adSaldoValorOutros(9), _
                              tSldMesEst2.adSaldoQuantOutros(10), tSldMesEst2.adSaldoValorOutros(10), tSldMesEst2.adSaldoQuantOutros(11), tSldMesEst2.adSaldoValorOutros(11), tSldMesEst2.adSaldoQuantOutros(12), tSldMesEst2.adSaldoValorOutros(12), objSldMesEst2.iAno, objSldMesEst2.iFilialEmpresa, objSldMesEst2.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71625
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71626
    
    objSldMesEst2.dQuantInicialOutros = tSldMesEst2.dQuantInicialOutros
    objSldMesEst2.dValorInicialOutros = tSldMesEst2.dValorInicialOutros
    
    For iIndice = 1 To 12
    
        objSldMesEst2.dSaldoQuantOutros(iIndice) = tSldMesEst2.adSaldoQuantOutros(iIndice)
        objSldMesEst2.dSaldoValorOutros(iIndice) = tSldMesEst2.adSaldoValorOutros(iIndice)
        
    Next
    
    SldMesEst2_Le_QuantValorOutros = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_QuantValorOutros:
    
    SldMesEst2_Le_QuantValorOutros = gErr
    
    Select Case gErr
    
        Case 71625, 71626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, objSldMesEst2.iAno, objSldMesEst2.iFilialEmpresa, objSldMesEst2.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152108)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantValorBenef(ByVal lComando As Long, ByVal objSldMesEst2 As ClassSldMesEst2) As Long

Dim tSldMesEst2 As typeSldMesEst2
Dim lErro As Long
Dim iIndice As Integer
Dim dQuant As Double
Dim dValor As Double

On Error GoTo Erro_SldMesEst2_Le_QuantValorBenef

    'seleciona o registro de SldMesEst2 relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialBenef, ValorInicialBenef, SaldoQuantBenef1, SaldoValorBenef1, SaldoQuantBenef2, SaldoValorBenef2, SaldoQuantBenef3, SaldoValorBenef3, SaldoQuantBenef4, SaldoValorBenef4, SaldoQuantBenef5, SaldoValorBenef5, SaldoQuantBenef6, SaldoValorBenef6, SaldoQuantBenef7, SaldoValorBenef7, SaldoQuantBenef8, SaldoValorBenef8, SaldoQuantBenef9, SaldoValorBenef9, SaldoQuantBenef10, SaldoValorBenef10, SaldoQuantBenef11, SaldoValorBenef11, SaldoQuantBenef12, SaldoValorBenef12  FROM SldMesEst2 WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst2.dQuantInicialBenef, tSldMesEst2.dValorInicialBenef, tSldMesEst2.adSaldoQuantBenef(1), tSldMesEst2.adSaldoValorBenef(1), tSldMesEst2.adSaldoQuantBenef(2), tSldMesEst2.adSaldoValorBenef(2), tSldMesEst2.adSaldoQuantBenef(3), tSldMesEst2.adSaldoValorBenef(3), tSldMesEst2.adSaldoQuantBenef(4), tSldMesEst2.adSaldoValorBenef(4), tSldMesEst2.adSaldoQuantBenef(5), tSldMesEst2.adSaldoValorBenef(5), tSldMesEst2.adSaldoQuantBenef(6), tSldMesEst2.adSaldoValorBenef(6), tSldMesEst2.adSaldoQuantBenef(7), tSldMesEst2.adSaldoValorBenef(7), tSldMesEst2.adSaldoQuantBenef(8), tSldMesEst2.adSaldoValorBenef(8), tSldMesEst2.adSaldoQuantBenef(9), tSldMesEst2.adSaldoValorBenef(9), _
                              tSldMesEst2.adSaldoQuantBenef(10), tSldMesEst2.adSaldoValorBenef(10), tSldMesEst2.adSaldoQuantBenef(11), tSldMesEst2.adSaldoValorBenef(11), tSldMesEst2.adSaldoQuantBenef(12), tSldMesEst2.adSaldoValorBenef(12), objSldMesEst2.iAno, objSldMesEst2.iFilialEmpresa, objSldMesEst2.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 71627
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71628
    
    objSldMesEst2.dQuantInicialBenef = tSldMesEst2.dQuantInicialBenef
    objSldMesEst2.dValorInicialBenef = tSldMesEst2.dValorInicialBenef
    
    For iIndice = 1 To 12
    
        objSldMesEst2.dSaldoQuantBenef(iIndice) = tSldMesEst2.adSaldoQuantBenef(iIndice)
        objSldMesEst2.dSaldoValorBenef(iIndice) = tSldMesEst2.adSaldoValorBenef(iIndice)
        
    Next
    
    SldMesEst2_Le_QuantValorBenef = SUCESSO

    Exit Function

Erro_SldMesEst2_Le_QuantValorBenef:
    
    SldMesEst2_Le_QuantValorBenef = gErr
    
    Select Case gErr
    
        Case 71627, 71628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST21", gErr, objSldMesEst2.iAno, objSldMesEst2.iFilialEmpresa, objSldMesEst2.sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152109)
        
    End Select
    
    Exit Function

End Function

Function MovEstoqueItens_Le_Desc(ByVal objMovEstoque As ClassMovEstoque) As Long
'Lê todos os Itens em Ordem Descendente de NumIntDoc correspondentes a partir da FilialEmpresa, Código e do NumIntDocEst=0 e retorna no colItens do objeto
'Ordena pelo Item.

Dim lErro As Long
Dim lComando As Long
Dim dtData As Date
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovEstoqueItens_Le_Desc

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 71815

    With tItemMovEstoque

        .sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
        .sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sProdutoOP = String(STRING_PRODUTO, 0)
        .sOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        .sContaContabilEst = String(STRING_CONTA, 0)
        .sContaContabilAplic = String(STRING_CONTA, 0)
        .sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)

        lErro = Comando_Executar(lComando, "SELECT MovimentoEstoque.ContaContabilEst, MovimentoEstoque.ContaContabilAplic, MovimentoEstoque.Data, MovimentoEstoque.NumIntDoc, MovimentoEstoque.Custo, MovimentoEstoque.Apropriacao, MovimentoEstoque.Produto, Produtos.Descricao, MovimentoEstoque.SiglaUM, MovimentoEstoque.Quantidade, MovimentoEstoque.Almoxarifado, Almoxarifado.NomeReduzido, MovimentoEstoque.NumIntDocOrigem, MovimentoEstoque.Ccl, MovimentoEstoque.NumIntDocEst, MovimentoEstoque.TipoMov,MovimentoEstoque.DocOrigem,MovimentoEstoque.Cliente,MovimentoEstoque.Fornecedor FROM MovimentoEstoque, Produtos, Almoxarifado WHERE Almoxarifado.Codigo = MovimentoEstoque.Almoxarifado AND Produtos.Codigo = MovimentoEstoque.Produto AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.Codigo = ? AND NumIntDocEst = 0 ORDER BY NumIntDoc DESC", _
        .sContaContabilEst, .sContaContabilAplic, dtData, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .iTipoMov, .sDocOrigem, .lCliente, .lFornecedor, objMovEstoque.iFilialEmpresa, objMovEstoque.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 71816

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71817

        Do While lErro <> AD_SQL_SEM_DADOS

            objMovEstoque.dtData = dtData

            objMovEstoque.colItens.Add .lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, 0, Nothing, Nothing, DATA_NULA
            objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).sDocOrigem = .sDocOrigem
            objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lCliente = .lCliente
            objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lFornecedor = .lFornecedor

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71818

        Loop

    End With

    If objMovEstoque.colItens.Count = 0 Then gError 71819

    Call Comando_Fechar(lComando)

    MovEstoqueItens_Le_Desc = SUCESSO

    Exit Function

Erro_MovEstoqueItens_Le_Desc:

    MovEstoqueItens_Le_Desc = gErr

    Select Case gErr

        Case 71815
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 71816, 71817, 71818
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 71819 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152110)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_ObterRelMesFatVendCli(ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, ByVal iVendedor As Integer, ByVal lCliente As Long, dValorFaturado As Double, dQuantidadeFaturada As Double, dQuantidadeDevolvida As Double, lComando As Long, lComando2 As Long) As Long
'IMPORTANTE: especifica para classe ClassRelMesFat de apoio a execucao de relatorios
'Obtém as quantidades  faturada e devolvida e valor total faturado por Produto entre datas passadas, relacionados ao Vendedor e Cliente passados
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long

On Error GoTo Erro_Produto_ObterRelMesFatVendCli

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then gError 500218
    End If
    
    If lComando2 = 0 Then
        lComando2 = Comando_AbrirExt(GL_lConexaoRel)
        If lComando2 = 0 Then gError 500226
    End If
                
    'Acumula os valores e quantidades das notas ficais que satisfazem a seleção passada, fazem parte do faturamento e NÃO são de devolução
    lErro = Comando_Executar(lComando, "SELECT SUM(ItensNFiscal.Quantidade * ItensNFiscal.PrecoUnitario), SUM(ItensNFiscal.Quantidade) FROM NFiscal, ItensNFiscal, ComissoesNF, TiposDocInfo WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.NumIntDoc = ComissoesNF.NumIntDoc AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.FilialEmpresa = ? AND ItensNFiscal.Produto = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ? AND ComissoesNF.CodVendedor = ? AND NFiscal.Cliente = ? AND TiposDocInfo.Faturamento = ?", dValorFaturado, dQuantidadeFaturada, iFilialEmpresa, sProduto, dtDataInicial, dtDataFinal, iVendedor, lCliente, TIPO_FATURAMENTO)
    If lErro <> AD_SQL_SUCESSO Then gError 500219
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 500220
        
    If lErro = AD_SQL_SEM_DADOS Then
        
        dValorFaturado = 0
        dQuantidadeFaturada = 0

    End If
            
    'Acumula quantidade devolvida das notas ficais que satisfazem a seleção passada, fazem parte do faturamento e são de devolução
    lErro = Comando_Executar(lComando2, "SELECT SUM(ItensNFiscal.Quantidade) FROM NFiscal, ItensNFiscal, ComissoesNF, TiposDocInfo WHERE NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.NumIntDoc = ComissoesNF.NumIntDoc AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.FilialEmpresa = ? AND ItensNFiscal.Produto = ? AND NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ? AND ComissoesNF.CodVendedor = ? AND NFiscal.Cliente = ? AND TiposDocInfo.Faturamento = ?", dQuantidadeDevolvida, iFilialEmpresa, sProduto, dtDataInicial, dtDataFinal, iVendedor, lCliente, TIPO_FATURAMENTO_DEVOLUCAO)
    If lErro <> AD_SQL_SUCESSO Then gError 500221
        
    lErro = Comando_BuscarProximo(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 500222
    
    If lErro = AD_SQL_SEM_DADOS Then
        
        dQuantidadeDevolvida = 0

    End If
    
    Produto_ObterRelMesFatVendCli = SUCESSO
    
    Exit Function
    
Erro_Produto_ObterRelMesFatVendCli:

    Produto_ObterRelMesFatVendCli = gErr
    
    Select Case gErr

        Case 500218, 500226
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 500219, 500220, 500221, 500222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152111)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function
    
End Function

Function MATConfig_Le_Lock(ByVal objMATConfig As ClassMATConfig, ByVal lComando As Long) As Long
'Lê registro em MATConfig.

Dim sConteudo As String
Dim lErro As Long

On Error GoTo Erro_MATConfig_Le_Lock

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM MATConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 71738
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 71739
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 71740

    objMATConfig.sConteudo = sConteudo

    MATConfig_Le_Lock = SUCESSO
    
    Exit Function

Erro_MATConfig_Le_Lock:

    MATConfig_Le_Lock = gErr
    
    Select Case gErr
    
        Case 71738, 71739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG1", gErr, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
        
        Case 71740
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MATCONFIG1", gErr, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152112)
        
    End Select
    
    Exit Function

End Function

Function MovimentoEstoque_Le_ProdutoData(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal dtData As Date) As Long
'verifica se há movimento para o produto a partir da data especificada

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_MovimentoEstoque_Le_ProdutoData

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 83593

    'le os movimentos a partir da data especificada
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM MovimentoEstoque WHERE FilialEmpresa = ? AND Produto = ? AND Data >= ?", lNumIntDoc, iFilialEmpresa, sProduto, dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 83594

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83595

    'se não há movimento para o produto a partir da data especificada ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 83596
    
    Call Comando_Fechar(lComando)

    MovimentoEstoque_Le_ProdutoData = SUCESSO
    
    Exit Function

Erro_MovimentoEstoque_Le_ProdutoData:

    MovimentoEstoque_Le_ProdutoData = gErr

    Select Case gErr

        Case 83593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 83594, 83595
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE5", gErr, iFilialEmpresa, sProduto, dtData)

        Case 83596
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152113)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovimentoEstoque_Le_NumIntDoc(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque) As Long
'Carrega os dados do movimento com o NumIntDoc passado como parametro
    
Dim lErro As Long
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovimentoEstoque_Le_NumIntDoc

    tItemMovEstoque.sProduto = String(STRING_PRODUTO, 0)
    tItemMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemMovEstoque.sCcl = String(STRING_CCL, 0)
    tItemMovEstoque.sOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
    tItemMovEstoque.sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)
    tItemMovEstoque.sContaContabilAplic = String(STRING_CONTA, 0)
    tItemMovEstoque.sContaContabilEst = String(STRING_CONTA, 0)
    
    'le os movimento de estoque referente ao numintdoc  passado como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa, Codigo, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao, Hora FROM MovimentoEstoque WHERE NumIntDoc = ?", 0, _
    tItemMovEstoque.iFilialEmpresa, tItemMovEstoque.lCodigo, tItemMovEstoque.dCusto, tItemMovEstoque.iApropriacao, tItemMovEstoque.sProduto, tItemMovEstoque.sSiglaUM, tItemMovEstoque.dQuantidade, tItemMovEstoque.iAlmoxarifado, tItemMovEstoque.iTipoMov, tItemMovEstoque.lNumIntDocOrigem, tItemMovEstoque.iTipoNumIntDocOrigem, tItemMovEstoque.dtData, tItemMovEstoque.sCcl, tItemMovEstoque.lNumIntDocEst, tItemMovEstoque.lCliente, tItemMovEstoque.lFornecedor, tItemMovEstoque.sOPCodigo, tItemMovEstoque.sDocOrigem, tItemMovEstoque.sContaContabilEst, tItemMovEstoque.sContaContabilAplic, tItemMovEstoque.lHorasMaquina, tItemMovEstoque.dtDataInicioProducao, tItemMovEstoque.dHora, objItemMovEst.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 83746

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83747

    'se o movimento não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 83748

    objItemMovEst.iFilialEmpresa = tItemMovEstoque.iFilialEmpresa
    objItemMovEst.lCodigo = tItemMovEstoque.lCodigo
    objItemMovEst.dCusto = tItemMovEstoque.dCusto
    objItemMovEst.iApropriacao = tItemMovEstoque.iApropriacao
    objItemMovEst.sProduto = tItemMovEstoque.sProduto
    objItemMovEst.sSiglaUM = tItemMovEstoque.sSiglaUM
    objItemMovEst.dQuantidade = tItemMovEstoque.dQuantidade
    objItemMovEst.iAlmoxarifado = tItemMovEstoque.iAlmoxarifado
    objItemMovEst.iTipoMov = tItemMovEstoque.iTipoMov
    objItemMovEst.lNumIntDocOrigem = tItemMovEstoque.lNumIntDocOrigem
    objItemMovEst.iTipoNumIntDocOrigem = tItemMovEstoque.iTipoNumIntDocOrigem
    objItemMovEst.dtData = tItemMovEstoque.dtData
    objItemMovEst.sCcl = tItemMovEstoque.sCcl
    objItemMovEst.lNumIntDocEst = tItemMovEstoque.lNumIntDocEst
    objItemMovEst.lCliente = tItemMovEstoque.lCliente
    objItemMovEst.lFornecedor = tItemMovEstoque.lFornecedor
    objItemMovEst.sOPCodigo = tItemMovEstoque.sOPCodigo
    objItemMovEst.sDocOrigem = tItemMovEstoque.sDocOrigem
    objItemMovEst.sContaContabilEst = tItemMovEstoque.sContaContabilEst
    objItemMovEst.sContaContabilAplic = tItemMovEstoque.sContaContabilAplic
    objItemMovEst.lHorasMaquina = tItemMovEstoque.lHorasMaquina
    objItemMovEst.dtDataInicioProducao = tItemMovEstoque.dtDataInicioProducao
    objItemMovEst.dtHora = tItemMovEstoque.dHora
        
    MovimentoEstoque_Le_NumIntDoc = SUCESSO
    
    Exit Function

Erro_MovimentoEstoque_Le_NumIntDoc:

    MovimentoEstoque_Le_NumIntDoc = gErr

    Select Case gErr

        Case 83746, 83747
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)

        Case 83748
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152114)

    End Select

    Exit Function

End Function

Function MATConfig_Le_Lock1(ByVal objMATConfig As ClassMATConfig, ByVal lComando As Long) As Long
'Lê registro em MATConfig.

Dim sConteudo As String
Dim lErro As Long

On Error GoTo Erro_MATConfig_Le_Lock1

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)

    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM MATConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 83774
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83775
    
    If lErro = AD_SQL_SEM_DADOS Then gError 83776
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 83779

    objMATConfig.sConteudo = sConteudo
    
    MATConfig_Le_Lock1 = SUCESSO
    
    Exit Function

Erro_MATConfig_Le_Lock1:

    MATConfig_Le_Lock1 = gErr
    
    Select Case gErr
    
        Case 83773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 83774, 83775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG1", gErr, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
            
        Case 83776
        
        Case 83779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MATCONFIG1", gErr, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152115)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_Le_ItemNF(ByVal objItemMovEstoque As ClassItemMovEstoque, ByVal colMovEstoque As Collection) As Long
'Lê o NumIntDoc e Código do MovimentoEstoque a partir do NumIntDoc do ItemNF passado

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim lCodigo As Long
Dim dtData As Date
Dim tItemMovEstoque As typeItemMovEstoque
Dim objItemMovEstoque1 As ClassItemMovEstoque
Dim dHora As Double

On Error GoTo Erro_MovEstoque_Le_ItemNF
    
    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 75793
    
    With tItemMovEstoque

        .sProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUM = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCcl = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sOPCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sDocOrigem = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilEst = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilAplic = String(STRING_BUFFER_MAX_TEXTO, 0)
    
        'Lê NumIntDoc de MovimentoEstoque
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo, NumIntDoc, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao, Hora FROM MovimentoEstoque WHERE NumIntDocOrigem = ? AND TipoNumIntDocOrigem = ? AND FilialEmpresa = ? ORDER BY NumIntDoc", .iFilialEmpresa, .lCodigo, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .iTipoMov, .lNumIntDocOrigem, .iTipoNumIntDocOrigem, .dtData, .sCcl, .lNumIntDocEst, .lCliente, .lFornecedor, .sOPCodigo, .sDocOrigem, .sContaContabilEst, .sContaContabilAplic, .lHorasMaquina, .dtDataInicioProducao, dHora, objItemMovEstoque.lNumIntDocOrigem, objItemMovEstoque.iTipoNumIntDocOrigem, objItemMovEstoque.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 75794

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75795
    
        Do While lErro = AD_SQL_SUCESSO
    
            Set objItemMovEstoque1 = New ClassItemMovEstoque
            
            objItemMovEstoque1.iFilialEmpresa = .iFilialEmpresa
            objItemMovEstoque1.lCodigo = .lCodigo
            objItemMovEstoque1.lNumIntDoc = .lNumIntDoc
            objItemMovEstoque1.dCusto = .dCusto
            objItemMovEstoque1.iApropriacao = .iApropriacao
            objItemMovEstoque1.sProduto = .sProduto
            objItemMovEstoque1.sSiglaUM = .sSiglaUM
            objItemMovEstoque1.dQuantidade = .dQuantidade
            objItemMovEstoque1.iAlmoxarifado = .iAlmoxarifado
            objItemMovEstoque1.iTipoMov = .iTipoMov
            objItemMovEstoque1.lNumIntDocOrigem = .lNumIntDocOrigem
            objItemMovEstoque1.iTipoNumIntDocOrigem = .iTipoNumIntDocOrigem
            objItemMovEstoque1.dtData = .dtData
            objItemMovEstoque1.sCcl = .sCcl
            objItemMovEstoque1.lNumIntDocEst = .lNumIntDocEst
            objItemMovEstoque1.lCliente = .lCliente
            objItemMovEstoque1.lFornecedor = .lFornecedor
            objItemMovEstoque1.sOPCodigo = .sOPCodigo
            objItemMovEstoque1.sDocOrigem = .sDocOrigem
            objItemMovEstoque1.sContaContabilEst = .sContaContabilEst
            objItemMovEstoque1.sContaContabilAplic = .sContaContabilAplic
            objItemMovEstoque1.lHorasMaquina = .lHorasMaquina
            objItemMovEstoque1.dtDataInicioProducao = .dtDataInicioProducao
            objItemMovEstoque1.dtHora = CDate(dHora)
            
            colMovEstoque.Add objItemMovEstoque1
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83227
            
        Loop
        
    End With
    'Fecha comandos
    Call Comando_Fechar(lComando)
    
    MovEstoque_Le_ItemNF = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Le_ItemNF:

    MovEstoque_Le_ItemNF = gErr
    
    Select Case gErr
    
        Case 75793
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 75794, 75795, 83227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152116)

    End Select
    
    'Fecha comandos
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function RastreamentoMovto_Le_DocOrigem(ByVal lNumIntDocOrigem As Long, ByVal iTipoDocOrigem As Integer, ByVal colRastreamentoMovto As Collection) As Long
'Lê a tabela de RastreamentoMovto através do Movimento de Estoque

Dim lErro As Long
Dim tRastreamentoMovto As typeRastreamentoMovto
Dim objRastreamentoMovto As New ClassRastreamentoMovto
Dim lComando As Long
Dim sSiglaUM As String

'###########################################
'Inserido por Wagner 13/03/2006
Dim objProduto As ClassProduto
Dim colRastreamentoSerie As Collection
Dim lNumIntDocLoteIni As Long
Dim lNumIntDocLoteIniAnt As Long
Dim lNumIntDocLote As Long
Dim dQuantidadeFaixaSerie As Double
'###########################################

On Error GoTo Erro_RastreamentoMovto_Le_DocOrigem
    
    'Abertura de comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 78411
    
    tRastreamentoMovto.sProduto = String(STRING_PRODUTO, 0)
    tRastreamentoMovto.sLote = String(STRING_LOTE_RASTREAMENTO, 0)
    sSiglaUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
    
    '#################################################################
    'Alterado por Wagner 13/03/2006
    'Lê o Rastreamento Movto
    lErro = Comando_Executar(lComando, "SELECT RastreamentoMovto.NumIntDoc, RastreamentoMovto.TipoDocOrigem, RastreamentoMovto.NumIntDocOrigem, RastreamentoMovto.Produto, RastreamentoMovto.Quantidade, RastreamentoMovto.SiglaUM, RastreamentoLote.Lote, RastreamentoLote.FilialOP, RastreamentoMovto.QuantidadeFaixaSerie, RastreamentoMovto.NumIntDocLoteSerieIni,RastreamentoMovto.NumIntDocLote FROM RastreamentoMovto, RastreamentoLote WHERE RastreamentoMovto.NumIntDocLote = RastreamentoLote.NumIntDoc AND TipoDocOrigem = ? AND NumIntDocOrigem = ? ORDER BY RastreamentoMovto.NumIntDocLoteSerieIni" _
    , tRastreamentoMovto.lNumIntDoc, tRastreamentoMovto.iTipoDocOrigem, tRastreamentoMovto.lNumIntDocOrigem, tRastreamentoMovto.sProduto, tRastreamentoMovto.dQuantidade, sSiglaUM, tRastreamentoMovto.sLote, tRastreamentoMovto.iFilialOP, dQuantidadeFaixaSerie, lNumIntDocLoteIni, lNumIntDocLote, iTipoDocOrigem, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 78412
    '#################################################################
           
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78413

    Do While lErro = AD_SQL_SUCESSO
        
        Set objRastreamentoMovto = New ClassRastreamentoMovto
        
        '##############################################################
        'Inserido por Wagner 13/03/2006
        Set objProduto = New ClassProduto
        
        'Se inda não mudou o lote inicial tem que continuar preenchendo a coleção
        'que agrupa as séries por lote inicial
        If lNumIntDocLoteIni <> lNumIntDocLoteIniAnt Then
            lNumIntDocLoteIniAnt = lNumIntDocLoteIni
            Set colRastreamentoSerie = New Collection
        End If
        
        objProduto.sCodigo = tRastreamentoMovto.sProduto

        'le o produto para descobrir o tipo de rastreamento
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 141802
        '##############################################################
        
        'passa para o objeto
        objRastreamentoMovto.lNumIntDoc = tRastreamentoMovto.lNumIntDoc
        objRastreamentoMovto.iTipoDocOrigem = tRastreamentoMovto.iTipoDocOrigem
        objRastreamentoMovto.lNumIntDocOrigem = tRastreamentoMovto.lNumIntDocOrigem
        objRastreamentoMovto.sProduto = tRastreamentoMovto.sProduto
        objRastreamentoMovto.dQuantidade = tRastreamentoMovto.dQuantidade
        objRastreamentoMovto.sSiglaUM = sSiglaUM
        objRastreamentoMovto.sLote = tRastreamentoMovto.sLote
        objRastreamentoMovto.iFilialOP = tRastreamentoMovto.iFilialOP
        
        '##############################################################
        'Alterado por Wagner 13/03/2006
        'Se for rastreamento por número de série tem que agrupar as séries só
        'exibindo a série inicial
        If objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE And lNumIntDocLoteIni <> 0 Then
        
            objRastreamentoMovto.dQuantidadeSerie = dQuantidadeFaixaSerie
            objRastreamentoMovto.lNumIntDocLote = lNumIntDocLote
            objRastreamentoMovto.lNumIntDocLoteSerieIni = lNumIntDocLoteIni
            objRastreamentoMovto.iFilialOP = 0
            
            colRastreamentoSerie.Add objRastreamentoMovto
            
            If objRastreamentoMovto.lNumIntDocLote = objRastreamentoMovto.lNumIntDocLoteSerieIni Then
            
                Set objRastreamentoMovto.colRastreamentoSerie = colRastreamentoSerie
            
                colRastreamentoMovto.Add objRastreamentoMovto
            
            End If
            
        Else
            colRastreamentoMovto.Add objRastreamentoMovto
        End If
        '##############################################################
            
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78421
    
    Loop
            
    Call Comando_Fechar(lComando)
    
    RastreamentoMovto_Le_DocOrigem = SUCESSO
        
    Exit Function
    
Erro_RastreamentoMovto_Le_DocOrigem:

    RastreamentoMovto_Le_DocOrigem = gErr
    
    Select Case gErr
            
        Case 78411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 78412, 78413, 78421
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_RASTREAMENTOMOVTO", gErr)
            
        Case 141802
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152117)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RastreamentoLote_Le(ByVal objRastroLote As ClassRastreamentoLote) As Long
'Lê rastreamento do lote a partir do produto, filialOP e código do lote passados

Dim lErro As Long
Dim lComando As Long
Dim tRastroLote As typeRastreamentoLote

On Error GoTo Erro_RastreamentoLote_Le

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 75707

    tRastroLote.sObservacao = String(STRING_RASTRO_OBSERVACAO, 0)
    tRastroLote.sLocalizacao = String(STRING_RASTRO_LOCALIZACAO, 0)

    'Lê dados de RastrementoLote a partir de Produto, FilialOP e Lote
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Status, DataValidade, DataEntrada, DataFabricacao, Observacao, Localizacao, Cliente, FilialCli FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", tRastroLote.lNumIntDoc, tRastroLote.iStatus, tRastroLote.dtDataValidade, tRastroLote.dtDataEntrada, tRastroLote.dtDataFabricacao, tRastroLote.sObservacao, tRastroLote.sLocalizacao, tRastroLote.lCliente, tRastroLote.iFilialCli, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)
    If lErro <> AD_SQL_SUCESSO Then gError 75708

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75709

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 75710
    
    With objRastroLote
        .lNumIntDoc = tRastroLote.lNumIntDoc
        .iStatus = tRastroLote.iStatus
        .dtDataEntrada = tRastroLote.dtDataEntrada
        .dtDataFabricacao = tRastroLote.dtDataFabricacao
        .dtDataValidade = tRastroLote.dtDataValidade
        .sObservacao = tRastroLote.sObservacao
        .sLocalizacao = tRastroLote.sLocalizacao
        .lCliente = tRastroLote.lCliente
        .iFilialCli = tRastroLote.iFilialCli
    End With
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    RastreamentoLote_Le = SUCESSO

    Exit Function

Erro_RastreamentoLote_Le:

    RastreamentoLote_Le = gErr

    Select Case gErr

        Case 75707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 75708, 75709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)

        Case 75710 'RastreamentoLote não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152118)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RastreamentoLote_Le1(ByVal lComando As Long, ByVal objRastroLote As ClassRastreamentoLote) As Long
'Lê rastreamento do lote a partir do produto, filialOP e código do lote passados

Dim lErro As Long
Dim tRastroLote As typeRastreamentoLote

On Error GoTo Erro_RastreamentoLote_Le1

    tRastroLote.sObservacao = String(STRING_RASTRO_OBSERVACAO, 0)
    tRastroLote.sLocalizacao = String(STRING_RASTRO_LOCALIZACAO, 0)

    'Lê dados de RastrementoLote a partir de Produto, FilialOP e Lote
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, Status, DataValidade, DataEntrada, DataFabricacao, Observacao, Localizacao, Cliente, FilialCli FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", 0, tRastroLote.lNumIntDoc, tRastroLote.iStatus, tRastroLote.dtDataValidade, tRastroLote.dtDataEntrada, tRastroLote.dtDataFabricacao, tRastroLote.sObservacao, tRastroLote.sLocalizacao, tRastroLote.lCliente, tRastroLote.iFilialCli, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)
    If lErro <> AD_SQL_SUCESSO Then gError 71887

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71888
    
    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 71889

    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 71901
    
    objRastroLote.dtDataEntrada = tRastroLote.dtDataEntrada
    objRastroLote.dtDataFabricacao = tRastroLote.dtDataFabricacao
    objRastroLote.dtDataValidade = tRastroLote.dtDataValidade
    objRastroLote.sObservacao = tRastroLote.sObservacao
    objRastroLote.iStatus = tRastroLote.iStatus
    objRastroLote.lNumIntDoc = tRastroLote.lNumIntDoc
    objRastroLote.sLocalizacao = tRastroLote.sLocalizacao
    objRastroLote.lCliente = tRastroLote.lCliente
    objRastroLote.iFilialCli = tRastroLote.iFilialCli
    
    RastreamentoLote_Le1 = SUCESSO

    Exit Function

Erro_RastreamentoLote_Le1:

    RastreamentoLote_Le1 = gErr

    Select Case gErr

        Case 71887, 71888
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE2", gErr, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)

        Case 71889

        Case 71901
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_RASTREAMENTOLOTE", gErr, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152119)

    End Select

    Exit Function

End Function

Function MovEstoqueItens_Le_Verifica_Estorno(ByVal objMovEstoque As ClassMovEstoque, Optional ByVal iTipoMovPassado As Integer = 0) As Long
'Lê todos os Itens correspondentes a partir da FilialEmpresa, Código e Opcional(TipoMov) retorna no colItens do objeto
'Ordena pelo Item.
'Verifica se todos os Movimentos estão estornados

Dim lErro As Long
Dim lComando As Long
Dim dtData As Date
Dim tItemMovEstoque As typeItemMovEstoque
Dim iEstorno As Integer

On Error GoTo Erro_MovEstoqueItens_Le_Verifica_Estorno

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 78880

    With tItemMovEstoque

        .sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
        .sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sProdutoOP = String(STRING_PRODUTO, 0)
        .sOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        .sContaContabilEst = String(STRING_CONTA, 0)
        .sContaContabilAplic = String(STRING_CONTA, 0)
        .sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)
        
        '''11/09/01 Marcelo inclusao do campo Hora no Select
        lErro = Comando_Executar(lComando, "SELECT MovimentoEstoque.ContaContabilEst, MovimentoEstoque.ContaContabilAplic, MovimentoEstoque.Data, MovimentoEstoque.NumIntDoc, MovimentoEstoque.Custo, MovimentoEstoque.Apropriacao, MovimentoEstoque.Produto, Produtos.Descricao, MovimentoEstoque.SiglaUM, MovimentoEstoque.Quantidade, MovimentoEstoque.Almoxarifado, Almoxarifado.NomeReduzido, MovimentoEstoque.NumIntDocOrigem, MovimentoEstoque.Ccl, MovimentoEstoque.NumIntDocEst, MovimentoEstoque.TipoMov,MovimentoEstoque.DocOrigem,MovimentoEstoque.Cliente,MovimentoEstoque.Fornecedor, MovimentoEstoque.Hora FROM MovimentoEstoque, Produtos, Almoxarifado WHERE Almoxarifado.Codigo = MovimentoEstoque.Almoxarifado AND Produtos.Codigo = MovimentoEstoque.Produto AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.Codigo = ? ORDER BY NumIntDoc", _
        .sContaContabilEst, .sContaContabilAplic, dtData, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .iTipoMov, .sDocOrigem, .lCliente, .lFornecedor, .dHora, giFilialEmpresa, objMovEstoque.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 78881
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78882

        If lErro = AD_SQL_SEM_DADOS Then gError 78883

        Do While lErro <> AD_SQL_SEM_DADOS
            
            If .lNumIntDocEst = 0 Then
                
                
                objMovEstoque.dtData = dtData
                objMovEstoque.dtHora = tItemMovEstoque.dHora ''' 11/09/01 inclusao da atribuição
                objMovEstoque.colItens.Add .lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, 0, Nothing, Nothing, DATA_NULA
                objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).sDocOrigem = .sDocOrigem
                objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lCliente = .lCliente
                objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lFornecedor = .lFornecedor
                
                'Se passou um Tipo como parametro verifica se o movimento tem o Tipo, se tiver seta como o tipo do Movimento
                If iTipoMovPassado > 0 Then
                    If .iTipoMov = iTipoMovPassado Then
                        objMovEstoque.iTipoMov = iTipoMovPassado
                    End If
                Else
                    objMovEstoque.iTipoMov = .iTipoMov
                End If
            Else
                iEstorno = 1
            End If
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78884
            
        Loop
        
        'Se Só houve estorno --> Erro
        If iEstorno = 1 And objMovEstoque.colItens.Count = 0 Then gError 78885
        
        
    End With

    Call Comando_Fechar(lComando)

    MovEstoqueItens_Le_Verifica_Estorno = SUCESSO

    Exit Function

Erro_MovEstoqueItens_Le_Verifica_Estorno:

    MovEstoqueItens_Le_Verifica_Estorno = gErr

    Select Case gErr

        Case 78880
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 78881, 78882, 78884
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 78883 'Não encontrou o Movimento de Estoque
        
        Case 78885 'Todos os Itens do Movimento estão estornados
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152120)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RastreamentoLoteSaldo_Le(ByVal objRastroLoteSaldo As ClassRastreamentoLoteSaldo) As Long
'Lê a tabela de Rastreamento Lote Saldo

Dim lErro As Long
Dim lComando As Long
Dim tRastroLoteSaldo As typeRastreamentoLoteSaldo

On Error GoTo Erro_RastreamentoLoteSaldo_Le

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 78630

    tRastroLoteSaldo.sProduto = String(STRING_PRODUTO, 0)

    'Lê dados de RastrementoLote a partir de Produto, FilialOP e Lote
    lErro = Comando_Executar(lComando, "SELECT RastreamentoLoteSaldo.Produto, Almoxarifado, NumIntDocLote, QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantEmpenhada, QuantPedida, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantOP, QuantBenef, QuantBenef3 FROM RastreamentoLoteSaldo, RastreamentoLote WHERE RastreamentoLote.NumIntDoc = RastreamentoLoteSaldo.NumIntDocLote AND RastreamentoLoteSaldo.Produto = ? AND Almoxarifado = ? AND RastreamentoLote.Lote = ? AND RastreamentoLote.FilialOP = ?", _
    tRastroLoteSaldo.sProduto, tRastroLoteSaldo.iAlmoxarifado, tRastroLoteSaldo.lNumIntDocLote, tRastroLoteSaldo.dQuantDispNossa, tRastroLoteSaldo.dQuantReservada, tRastroLoteSaldo.dQuantReservadaConsig, tRastroLoteSaldo.dQuantEmpenhada, tRastroLoteSaldo.dQuantPedida, tRastroLoteSaldo.dQuantRecIndl, tRastroLoteSaldo.dQuantIndOutras, tRastroLoteSaldo.dQuantDefeituosa, tRastroLoteSaldo.dQuantConsig3, tRastroLoteSaldo.dQuantConsig, tRastroLoteSaldo.dQuantDemo3, tRastroLoteSaldo.dQuantDemo, tRastroLoteSaldo.dQuantConserto3, tRastroLoteSaldo.dQuantConserto, tRastroLoteSaldo.dQuantOutras3, tRastroLoteSaldo.dQuantOutras, tRastroLoteSaldo.dQuantOP, tRastroLoteSaldo.dQuantBenef, tRastroLoteSaldo.dQuantBenef3, objRastroLoteSaldo.sProduto, objRastroLoteSaldo.iAlmoxarifado, objRastroLoteSaldo.sLote, objRastroLoteSaldo.iFilialOP)
    If lErro <> AD_SQL_SUCESSO Then gError 78631

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78632

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 78633

    objRastroLoteSaldo.sProduto = tRastroLoteSaldo.sProduto
    objRastroLoteSaldo.iAlmoxarifado = tRastroLoteSaldo.iAlmoxarifado
    objRastroLoteSaldo.lNumIntDocLote = tRastroLoteSaldo.lNumIntDocLote
    objRastroLoteSaldo.dQuantDispNossa = tRastroLoteSaldo.dQuantDispNossa
    objRastroLoteSaldo.dQuantReservada = tRastroLoteSaldo.dQuantReservada
    objRastroLoteSaldo.dQuantReservadaConsig = tRastroLoteSaldo.dQuantReservadaConsig
    objRastroLoteSaldo.dQuantEmpenhada = tRastroLoteSaldo.dQuantEmpenhada
    objRastroLoteSaldo.dQuantPedida = tRastroLoteSaldo.dQuantPedida
    objRastroLoteSaldo.dQuantRecIndl = tRastroLoteSaldo.dQuantRecIndl
    objRastroLoteSaldo.dQuantIndOutras = tRastroLoteSaldo.dQuantIndOutras
    objRastroLoteSaldo.dQuantDefeituosa = tRastroLoteSaldo.dQuantDefeituosa
    objRastroLoteSaldo.dQuantConsig3 = tRastroLoteSaldo.dQuantConsig3
    objRastroLoteSaldo.dQuantConsig = tRastroLoteSaldo.dQuantConsig
    objRastroLoteSaldo.dQuantDemo3 = tRastroLoteSaldo.dQuantDemo3
    objRastroLoteSaldo.dQuantDemo = tRastroLoteSaldo.dQuantDemo
    objRastroLoteSaldo.dQuantConserto3 = tRastroLoteSaldo.dQuantConserto3
    objRastroLoteSaldo.dQuantConserto = tRastroLoteSaldo.dQuantConserto
    objRastroLoteSaldo.dQuantOutras3 = tRastroLoteSaldo.dQuantOutras3
    objRastroLoteSaldo.dQuantOutras = tRastroLoteSaldo.dQuantOutras
    objRastroLoteSaldo.dQuantOP = tRastroLoteSaldo.dQuantOP
    objRastroLoteSaldo.dQuantBenef = tRastroLoteSaldo.dQuantBenef
    objRastroLoteSaldo.dQuantBenef3 = tRastroLoteSaldo.dQuantBenef3

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    RastreamentoLoteSaldo_Le = SUCESSO

    Exit Function

Erro_RastreamentoLoteSaldo_Le:

    RastreamentoLoteSaldo_Le = gErr

    Select Case gErr

        Case 78630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 78631, 78632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTESALDO", gErr, objRastroLoteSaldo.sProduto, objRastroLoteSaldo.iAlmoxarifado, objRastroLoteSaldo.sLote)

        Case 78633 'RastreamentoLoteSaldo não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152121)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CustoMedioProducaoApurado_Le_Escaninhos(ByVal sProduto As String, dCustoConsig As Double, dCustoDemo As Double, dCustoConserto As Double, dCustoOutras As Double, dCustoBenef As Double, dCustoConsig3 As Double, dCustoDemo3 As Double, dCustoConserto3 As Double, dCustoOutras3 As Double, dCustoBenef3 As Double, ByVal iFilialEmpresa As Integer) As Long
'Calcula último Custo Médio de Produção Apurado do Produto passado

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes
Dim alComando(1 To 10) As Long
Dim iIndice As Integer
Dim objSldMesEst As New ClassSldMesEst

On Error GoTo Erro_CustoMedioProducaoApurado_Le_Escaninhos

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 69791
    Next

    objEstoqueMes.iFilialEmpresa = iFilialEmpresa
    
    lErro = CF("EstoqueMes_Le_Apurado", objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 46225 Then gError 69792
    
    'Se não foi apurado nenhum mês ==> descobre o ano em que o estoque foi criado
    If lErro = 46225 Then
        
        objSldMesEst.sProduto = sProduto
        objSldMesEst.iFilialEmpresa = objEstoqueMes.iFilialEmpresa
        
        'Se não  encontrou mês apurado então lê o SldMesEst para pegar o primeiro Ano
        lErro = CF("SldMesEst_Le_PrimeiroAno", objSldMesEst)
        If lErro <> SUCESSO Then gError 69831
        
        objEstoqueMes.iMes = 0
        objEstoqueMes.iAno = objSldMesEst.iAno
    End If
    
    If objEstoqueMes.iMes = 0 Then
        
        dCustoConserto = 0
        dCustoConserto3 = 0
        dCustoConsig = 0
        dCustoConsig3 = 0
        dCustoDemo = 0
        dCustoDemo3 = 0
        dCustoOutras = 0
        dCustoOutras3 = 0
        dCustoBenef = 0
        dCustoBenef3 = 0
        
    Else
        
        'calcula o custo medio de produção Benef3 do produto em questão
        lErro = CF("SldMesEst1_Le_CustoMedioBenef3", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoBenef3)
        If lErro <> SUCESSO Then gError 69793
        
        'calcula o custo medio de produção Conserto3 do produto em questão
        lErro = CF("SldMesEst1_Le_CustoMedioConserto3", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoConserto3)
        If lErro <> SUCESSO Then gError 69794
        
        'calcula o custo medio de produção Consig3 do produto em questão
        lErro = CF("SldMesEst1_Le_CustoMedioConsig3", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoConsig3)
        If lErro <> SUCESSO Then gError 69795
        
        'calcula o custo medio de produção Demo3 do produto em questão
        lErro = CF("SldMesEst1_Le_CustoMedioDemo3", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoDemo3)
        If lErro <> SUCESSO Then gError 69796
        
        'calcula o custo medio de produção Outros3  do produto em questão
        lErro = CF("SldMesEst1_Le_CustoMedioOutros3", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoOutras3)
        If lErro <> SUCESSO Then gError 69797
            
        'calcula o custo medio de produção Benef do produto em questão
        lErro = CF("SldMesEst2_Le_CustoMedioBenef", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoBenef)
        If lErro <> SUCESSO Then gError 69798
        
        'calcula o custo medio de produção Conserto  do produto em questão
        lErro = CF("SldMesEst2_Le_CustoMedioConserto", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoConserto)
        If lErro <> SUCESSO Then gError 69799
        
        'calcula o custo medio de produção Consug  do produto em questão
        lErro = CF("SldMesEst2_Le_CustoMedioConsig", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoConsig)
        If lErro <> SUCESSO Then gError 69800
        
        'calcula o custo medio de produção Demo  do produto em questão
        lErro = CF("SldMesEst2_Le_CustoMedioDemo", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoDemo)
        If lErro <> SUCESSO Then gError 69801
        
        'calcula o custo medio de produção Outros do produto em questão
        lErro = CF("SldMesEst2_Le_CustoMedioOutros", alComando(1), iFilialEmpresa, objEstoqueMes.iAno, sProduto, objEstoqueMes.iMes, dCustoOutras)
        If lErro <> SUCESSO Then gError 69802

    End If

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CustoMedioProducaoApurado_Le_Escaninhos = SUCESSO

    Exit Function

Erro_CustoMedioProducaoApurado_Le_Escaninhos:

    CustoMedioProducaoApurado_Le_Escaninhos = gErr

    Select Case gErr
        
        Case 69791
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 69792 To 69802 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152122)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RastroEstIni_Le(ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal colRastreamento As Collection) As Long

Dim lErro As Long
Dim objRastroEstIni As ClassRastroEstIni
Dim lComando As Long
Dim tRastroEstIni As typeRastroEstIni

On Error GoTo Erro_RastroEstIni_Le

    With tRastroEstIni
    
        .sLote = String(STRING_LOTE_RASTREAMENTO, 0)

        lComando = Comando_Abrir()
        If lComando = 0 Then gError 71831
    
        'Selecionar registros na tabela de RastroEstIni associados ao produto, almoxarifado  em questão
        lErro = Comando_Executar(lComando, "SELECT RastroEstIni.Escaninho, RastroEstIni.Quantidade, RastroEstIni.NumIntDocLote, RastreamentoLote.Lote, RastreamentoLote.FilialOP, RastreamentoLote.DataEntrada FROM RastroEstIni, RastreamentoLote WHERE RastroEstIni.NumIntDocLote = RastreamentoLote.NumIntDoc AND RastroEstIni.Produto=? AND RastroEstIni.Almoxarifado=?", .iEscaninho, .dQuantidade, .lNumIntDocLote, .sLote, .iFilialOP, .dtDataEntrada, sProduto, iAlmoxarifado)
        If lErro <> AD_SQL_SUCESSO Then gError 71832
    
        'leitura do registro referente ao RastreamentoMovto
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71833
    
        Do While lErro = AD_SQL_SUCESSO
            
            Set objRastroEstIni = New ClassRastroEstIni
            
            objRastroEstIni.sProduto = sProduto
            objRastroEstIni.iAlmoxarifado = iAlmoxarifado
            objRastroEstIni.iEscaninho = .iEscaninho
            objRastroEstIni.lNumIntDocLote = .lNumIntDocLote
            objRastroEstIni.dQuantidade = .dQuantidade
            objRastroEstIni.dtDataEntrada = .dtDataEntrada
            objRastroEstIni.iFilialOP = .iFilialOP
            objRastroEstIni.sLote = .sLote
            
            colRastreamento.Add objRastroEstIni
            
            'leitura do proximo registro referente ao RastreamentoMovto
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71834
    
        Loop

    End With

    Call Comando_Fechar(lComando)

    RastroEstIni_Le = SUCESSO
    
    Exit Function
    
Erro_RastroEstIni_Le:

    Select Case gErr
    
        Case 71831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 71832, 71833, 71834
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTROESTINI", gErr, sProduto, iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152123)
    
    End Select

    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Function MovEstoqueItens_Le_Desc1(ByVal objMovEstoque As ClassMovEstoque) As Long
'Lê todos os Itens (inclusive os estornados) em Ordem Descendente de NumIntDoc correspondentes a partir da FilialEmpresa, Código e retorna no colItens do objeto
'Ordena pelo Item.

Dim lErro As Long
Dim lComando As Long
Dim dtData As Date
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovEstoqueItens_Le_Desc1

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89046

    With tItemMovEstoque

        .sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
        .sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
        .sProdutoOP = String(STRING_PRODUTO, 0)
        .sOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        .sContaContabilEst = String(STRING_CONTA, 0)
        .sContaContabilAplic = String(STRING_CONTA, 0)
        .sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)

        lErro = Comando_Executar(lComando, "SELECT MovimentoEstoque.ContaContabilEst, MovimentoEstoque.ContaContabilAplic, MovimentoEstoque.Data, MovimentoEstoque.NumIntDoc, MovimentoEstoque.Custo, MovimentoEstoque.Apropriacao, MovimentoEstoque.Produto, Produtos.Descricao, MovimentoEstoque.SiglaUM, MovimentoEstoque.Quantidade, MovimentoEstoque.Almoxarifado, Almoxarifado.NomeReduzido, MovimentoEstoque.NumIntDocOrigem, MovimentoEstoque.Ccl, MovimentoEstoque.NumIntDocEst, MovimentoEstoque.TipoMov,MovimentoEstoque.DocOrigem,MovimentoEstoque.Cliente,MovimentoEstoque.Fornecedor FROM MovimentoEstoque, Produtos, Almoxarifado WHERE Almoxarifado.Codigo = MovimentoEstoque.Almoxarifado AND Produtos.Codigo = MovimentoEstoque.Produto AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.Codigo = ? ORDER BY NumIntDoc DESC", _
        .sContaContabilEst, .sContaContabilAplic, dtData, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .iTipoMov, .sDocOrigem, .lCliente, .lFornecedor, objMovEstoque.iFilialEmpresa, objMovEstoque.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 89047

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89048

        Do While lErro <> AD_SQL_SEM_DADOS

            objMovEstoque.dtData = dtData

            objMovEstoque.colItens.Add .lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, 0, Nothing, Nothing, DATA_NULA
            objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).sDocOrigem = .sDocOrigem
            objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lCliente = .lCliente
            objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lFornecedor = .lFornecedor
            objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).dtData = dtData
            objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).iFilialEmpresa = objMovEstoque.iFilialEmpresa

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89049

        Loop

    End With

    If objMovEstoque.colItens.Count = 0 Then gError 89050

    Call Comando_Fechar(lComando)

    MovEstoqueItens_Le_Desc1 = SUCESSO

    Exit Function

Erro_MovEstoqueItens_Le_Desc1:

    MovEstoqueItens_Le_Desc1 = gErr

    Select Case gErr

        Case 89046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89047, 89048, 89049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 89050 'Tratado na Rotina Chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152124)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MATConfig_Le(ByVal objMATConfig As ClassMATConfig) As Long
'Lê registro em MATConfig.

Dim sConteudo As String
Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_MATConfig_Le

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89115

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)

    lErro = Comando_Executar(lComando, "SELECT Conteudo FROM MATConfig WHERE Codigo = ? AND FilialEmpresa = ?", sConteudo, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 89116
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89117
    
    If lErro = AD_SQL_SEM_DADOS Then gError 89118
    
    objMATConfig.sConteudo = sConteudo
    
    Call Comando_Fechar(lComando)
    
    MATConfig_Le = SUCESSO
    
    Exit Function

Erro_MATConfig_Le:

    MATConfig_Le = gErr
    
    Select Case gErr
    
        Case 89115
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89116, 89117
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG1", gErr, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
            
        Case 89118
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152125)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function MovEstoque_NFiscal_Le(ByVal objNFiscal As Object, ByVal objMovEstoque As ClassMovEstoque) As Long
'Lê o codigo do movimento de estoque associado aos itens da Nota Fiscal passada

Dim lErro As Long
Dim lComando As Long
Dim objItemNF As Object
Dim iAchou As Integer
Dim objItemMovEstoque1 As ClassItemMovEstoque
Dim tItemMovEstoque As typeItemMovEstoque
Dim dtDataRegistro As Date

On Error GoTo Erro_MovEstoque_NFiscal_Le

    Set objMovEstoque.colItens = New colItensMovEstoque

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89695

    With tItemMovEstoque

        For Each objItemNF In objNFiscal.ColItensNF
    
            .sProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
            .sSiglaUM = String(STRING_BUFFER_MAX_TEXTO, 0)
            .sCcl = String(STRING_BUFFER_MAX_TEXTO, 0)
            .sOPCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
            .sDocOrigem = String(STRING_BUFFER_MAX_TEXTO, 0)
            .sContaContabilEst = String(STRING_BUFFER_MAX_TEXTO, 0)
            .sContaContabilAplic = String(STRING_BUFFER_MAX_TEXTO, 0)
    
            'Lê NumIntDoc de MovimentoEstoque
            lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo, NumIntDoc, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao, DataRegistro FROM MovimentoEstoque WHERE NumIntDocOrigem = ? AND TipoNumIntDocOrigem = ? AND FilialEmpresa = ? ORDER BY NumIntDoc", .iFilialEmpresa, .lCodigo, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .iTipoMov, .lNumIntDocOrigem, .iTipoNumIntDocOrigem, .dtData, .sCcl, .lNumIntDocEst, .lCliente, .lFornecedor, .sOPCodigo, .sDocOrigem, .sContaContabilEst, .sContaContabilAplic, .lHorasMaquina, .dtDataInicioProducao, dtDataRegistro, objItemNF.lNumIntDoc, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL, objNFiscal.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 89696
    
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89697
        
            Do While lErro = AD_SQL_SUCESSO
        
                Set objItemMovEstoque1 = New ClassItemMovEstoque
                
                objItemMovEstoque1.iFilialEmpresa = .iFilialEmpresa
                objItemMovEstoque1.lCodigo = .lCodigo
                objItemMovEstoque1.lNumIntDoc = .lNumIntDoc
                objItemMovEstoque1.dCusto = .dCusto
                objItemMovEstoque1.iApropriacao = .iApropriacao
                objItemMovEstoque1.sProduto = .sProduto
                objItemMovEstoque1.sSiglaUM = .sSiglaUM
                objItemMovEstoque1.dQuantidade = .dQuantidade
                objItemMovEstoque1.iAlmoxarifado = .iAlmoxarifado
                objItemMovEstoque1.iTipoMov = .iTipoMov
                objItemMovEstoque1.lNumIntDocOrigem = .lNumIntDocOrigem
                objItemMovEstoque1.iTipoNumIntDocOrigem = .iTipoNumIntDocOrigem
                objItemMovEstoque1.dtData = .dtData
                objItemMovEstoque1.sCcl = .sCcl
                objItemMovEstoque1.lNumIntDocEst = .lNumIntDocEst
                objItemMovEstoque1.lCliente = .lCliente
                objItemMovEstoque1.lFornecedor = .lFornecedor
                objItemMovEstoque1.sOPCodigo = .sOPCodigo
                objItemMovEstoque1.sDocOrigem = .sDocOrigem
                objItemMovEstoque1.sContaContabilEst = .sContaContabilEst
                objItemMovEstoque1.sContaContabilAplic = .sContaContabilAplic
                objItemMovEstoque1.lHorasMaquina = .lHorasMaquina
                objItemMovEstoque1.dtDataInicioProducao = .dtDataInicioProducao
                objItemMovEstoque1.dtDataRegistro = dtDataRegistro
                
                objMovEstoque.colItens.Add1 objItemMovEstoque1
                
                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89714
                
            Loop

        Next

    End With
    
    'se não tiver movimento de estoque associado a nenhum dos itens da nota fiscal
    If objMovEstoque.colItens.Count = 0 Then gError 89699

    'Fecha o Comando
    Call Comando_Fechar(lComando)

    MovEstoque_NFiscal_Le = SUCESSO

    Exit Function

Erro_MovEstoque_NFiscal_Le:

    MovEstoque_NFiscal_Le = gErr

    Select Case gErr

        Case 89696
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89697, 89698
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 89699

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152126)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovEstoque_Le_NumIntDocOrigem(ByVal objMovEstoque As ClassMovEstoque, ByVal iTipoNumIntDocOrigem As Integer, ByVal lNumIntDocOrigem As Long) As Long
'Lê os movimentos de estoque associados ao NumIntDocOrigem, TipoNumIntDocOrigem em questão

Dim lErro As Long
Dim lComando As Long
Dim objItemNF As Object
Dim iAchou As Integer
Dim objItemMovEstoque1 As ClassItemMovEstoque
Dim tItemMovEstoque As typeItemMovEstoque
Dim dtDataRegistro As Date

On Error GoTo Erro_MovEstoque_Le_NumIntDocOrigem

    Set objMovEstoque.colItens = New colItensMovEstoque

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 89846

    With tItemMovEstoque

        .sProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUM = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sCcl = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sOPCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sDocOrigem = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilEst = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilAplic = String(STRING_BUFFER_MAX_TEXTO, 0)
            
        'Lê NumIntDoc de MovimentoEstoque
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Codigo, NumIntDoc, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao, DataRegistro, Hora FROM MovimentoEstoque WHERE  FilialEmpresa = ? AND TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ?", .iFilialEmpresa, .lCodigo, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .iTipoMov, .lNumIntDocOrigem, .iTipoNumIntDocOrigem, .dtData, .sCcl, .lNumIntDocEst, .lCliente, .lFornecedor, .sOPCodigo, .sDocOrigem, .sContaContabilEst, .sContaContabilAplic, .lHorasMaquina, .dtDataInicioProducao, .dtDataRegistro, .dHora, objMovEstoque.iFilialEmpresa, iTipoNumIntDocOrigem, lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 89847

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89848
    
        Do While lErro = AD_SQL_SUCESSO
    
            Set objItemMovEstoque1 = New ClassItemMovEstoque
            
            objItemMovEstoque1.iFilialEmpresa = .iFilialEmpresa
            objItemMovEstoque1.lCodigo = .lCodigo
            objItemMovEstoque1.lNumIntDoc = .lNumIntDoc
            objItemMovEstoque1.dCusto = .dCusto
            objItemMovEstoque1.iApropriacao = .iApropriacao
            objItemMovEstoque1.sProduto = .sProduto
            objItemMovEstoque1.sSiglaUM = .sSiglaUM
            objItemMovEstoque1.dQuantidade = .dQuantidade
            objItemMovEstoque1.iAlmoxarifado = .iAlmoxarifado
            objItemMovEstoque1.iTipoMov = .iTipoMov
            objItemMovEstoque1.lNumIntDocOrigem = .lNumIntDocOrigem
            objItemMovEstoque1.iTipoNumIntDocOrigem = .iTipoNumIntDocOrigem
            objItemMovEstoque1.dtData = .dtData
            objItemMovEstoque1.sCcl = .sCcl
            objItemMovEstoque1.lNumIntDocEst = .lNumIntDocEst
            objItemMovEstoque1.lCliente = .lCliente
            objItemMovEstoque1.lFornecedor = .lFornecedor
            objItemMovEstoque1.sOPCodigo = .sOPCodigo
            objItemMovEstoque1.sDocOrigem = .sDocOrigem
            objItemMovEstoque1.sContaContabilEst = .sContaContabilEst
            objItemMovEstoque1.sContaContabilAplic = .sContaContabilAplic
            objItemMovEstoque1.lHorasMaquina = .lHorasMaquina
            objItemMovEstoque1.dtDataInicioProducao = .dtDataInicioProducao
            objItemMovEstoque1.dtDataRegistro = .dtDataRegistro
            objItemMovEstoque1.dtHora = CDate(.dHora)
            
            objMovEstoque.colItens.Add1 objItemMovEstoque1
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89849
            
        Loop

    End With
    
    'se não tiver movimento de estoque associado a nenhum dos itens da nota fiscal
    If objMovEstoque.colItens.Count = 0 Then gError 89850

    'Fecha o Comando
    Call Comando_Fechar(lComando)

    MovEstoque_Le_NumIntDocOrigem = SUCESSO

    Exit Function

Erro_MovEstoque_Le_NumIntDocOrigem:

    MovEstoque_Le_NumIntDocOrigem = gErr

    Select Case gErr

        Case 89846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89847, 89848, 89849
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE6", gErr, objMovEstoque.iFilialEmpresa, iTipoNumIntDocOrigem, lNumIntDocOrigem)

        Case 89850

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152127)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdProduzidos_Custos_MesAnt_Le(ByVal iMes As Integer, ByVal iAno As Integer, ByVal colProdutosProduzidos As Collection) As Long
'Função criada em 13/06/2001 por Luiz Gustavo de Freitas Nogueira
'Usada para preenchimento do grid na tela 'Custo de Produção'
'Essa função recebe uma coleção os produtos para os quais se deseja obter o custo de produção
'e lê no BD os custos (em iMes / iAno) de todos os produtos produzíveis
'e então, para os produtos que foram passados na coleção, guarda os custos lidos

Dim lComando As Long
Dim lErro As Long
Dim sMes As String
Dim objProdutoCusto As New ClassProdutoCusto
Dim tProdutoCusto As typeProdutoCusto

On Error GoTo Erro_ProdProduzidos_Custos_MesAnt_Le

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90544
    
    'Converte iMes em string para ser utilizada no Comando SELECT
    sMes = CStr(iMes)
    
    'Inicializa as variáveis que receberão strings
    tProdutoCusto.sCodProduto = String(STRING_PRODUTO, 0)
    tProdutoCusto.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    
    'Lê no BD os produtos produzíveis ativos para iMes / Ano
    lErro = Comando_Executar(lComando, "SELECT Produtos.Codigo, Produtos.Descricao, SldMesEst.CustoProducao" + sMes + " FROM Produtos, SldMesEst WHERE SldMesEst.Produto =  Produtos.Codigo AND SldMesEst.FilialEmpresa =? And SldMesEst.Ano =? And Produtos.Compras =? And Produtos.Gerencial =?  AND Produtos.Ativo =?  ORDER BY Produtos.Codigo", tProdutoCusto.sCodProduto, tProdutoCusto.sDescProduto, tProdutoCusto.dCustoMesAnterior, giFilialEmpresa, iAno, PRODUTO_PRODUZIVEL, NAO_GERENCIAL, PRODUTO_ATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 90545
    
    'Busca o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90546
    
    'Enquanto houverem produtos
    Do While lErro = AD_SQL_SUCESSO

        'Instancia o obj como o obj que armazena os dados de tProdutoCusto.sCodProduto
        'Se tProdutoCusto.sCodProduto não estiver na coleção, ocorrerá o erro 5,
        'que será tratado com um 'Resume Continua_Pos_Erro_5',
        'ou seja, a rotina continuará sendo processada a partir do label Continua_Pos_Erro_5:
        Set objProdutoCusto = colProdutosProduzidos.Item(tProdutoCusto.sCodProduto)

        'Guarda o custo do produto no seu obj
        objProdutoCusto.dCustoMesAnterior = tProdutoCusto.dCustoMesAnterior
            
'Esse label indica que, em caso de ocorrência do erro 5, o VB deverá continuar
'processando a rotina a partir desse ponto
Continua_Pos_Erro_5:
        
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90547
        
    Loop

    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ProdProduzidos_Custos_MesAnt_Le = SUCESSO

    Exit Function

Erro_ProdProduzidos_Custos_MesAnt_Le:

    ProdProduzidos_Custos_MesAnt_Le = gErr

    Select Case gErr
        
        'Esse ocorrerá caso o produto não tenha sido produzido no mês atual
        'isso significa que ele não será encontrado em colProdutosProduzidos
        'o que gerará o erro 5
        Case 5
            
            'Nesse caso, continua o processamento da função normalmente
            'a partir do ponto onde se encontra o label Continua_Pos_Erro_5:
            Resume Continua_Pos_Erro_5
            
        Case 90544
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 90545, 90546, 90547
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152128)

    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ProdProduzidos_Custos_Mes_Le(ByVal iMes As Integer, ByVal iAno As Integer, ByVal colProdutosProduzidos As Collection) As Long
'Função criada em 13/06/2001 por Luiz Gustavo de Freitas Nogueira
'Usada para preenchimento do grid na tela 'Custo de Produção'
'Essa função lê no BD todos os produtos produzidos dentro de iMes / iAno
'E preenche uma coleção global à tela com os produtos lidos

Dim lComando As Long
Dim lErro As Long
Dim sMes As String
Dim dtDataInicial As Date
Dim dtDataFinal As Date
Dim tProdutoCusto As typeProdutoCusto
Dim objProdutoCusto As ClassProdutoCusto

On Error GoTo Erro_ProdProduzidos_Custos_Mes_Le
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90540
    
    'Monta a data inicial
    'A data inicial deve ser o primeiro dia de iMes em iAno
    dtDataInicial = CDate(1 & "/" & iMes & "/" & iAno)
    
    'Monta a data final
    'A data final deve ser o último dia de iMes em iAno
    dtDataFinal = CDate(Dias_Mes(iMes, iAno) & "/" & iMes & "/" & iAno)
        
    'Converte iMes em string para ser utilizado na sintaxe do Comando SELECT
    sMes = CStr(iMes)
    
    'Inicializa as variáveis que receberão strings
    tProdutoCusto.sCodProduto = String(STRING_PRODUTO, 0)
    tProdutoCusto.sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
    tProdutoCusto.sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)

    'Lê no BD os Produtos de giFilialEmpresa
    'que tiveram movimento de estoque apropriados pelo custo real de produção
    'dentro do período de DataInicial e DataFinal
    lErro = Comando_Executar(lComando, "SELECT DISTINCT Produtos.Codigo, Produtos.Descricao, Produtos.SiglaUMEstoque, SldMesEst.CustoProducao" + sMes + " FROM Produtos, MovimentoEstoque, SldMesEst WHERE SldMesEst.FilialEmpresa = MovimentoEstoque.FilialEmpresa AND SldMesEst.Produto =  Produtos.Codigo AND MovimentoEstoque.Produto = Produtos.Codigo AND  SldMesEst.FilialEmpresa =? AND SldMesEst.Ano =? AND MovimentoEstoque.Data >=? AND MovimentoEstoque.Data <=? AND  MovimentoEstoque.Apropriacao =? ORDER BY Produtos.Codigo", tProdutoCusto.sCodProduto, tProdutoCusto.sDescProduto, tProdutoCusto.sSiglaUMEstoque, tProdutoCusto.dCusto, giFilialEmpresa, iAno, dtDataInicial, dtDataFinal, APROPR_CUSTO_REAL)
    If lErro <> AD_SQL_SUCESSO Then gError 90541
    
    'Busca o primeiro registro lido
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90542
    
    'Enquanto houverem produtos
    Do While lErro = AD_SQL_SUCESSO

        'Instancia um obj que receberá os dados do produto lido
        Set objProdutoCusto = New ClassProdutoCusto

        'Guarda no obj os dados do produto: Código, Descrição e UMEstoque, Custo
        objProdutoCusto.sCodProduto = tProdutoCusto.sCodProduto
        objProdutoCusto.sDescProduto = tProdutoCusto.sDescProduto
        objProdutoCusto.sSiglaUMEstoque = tProdutoCusto.sSiglaUMEstoque
        objProdutoCusto.dCusto = tProdutoCusto.dCusto
            
        'Guarda o obj na coleção
        colProdutosProduzidos.Add objProdutoCusto, objProdutoCusto.sCodProduto
        
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90543
        
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ProdProduzidos_Custos_Mes_Le = SUCESSO

    Exit Function

Erro_ProdProduzidos_Custos_Mes_Le:

    ProdProduzidos_Custos_Mes_Le = gErr

    Select Case gErr
        
        Case 90540
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 90541, 90542, 90543
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152129)

    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function OrdemDeProducaoBaixada_Le_ComItens(ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'carrega do bd a ordem de producao baixada inclusive os seus itens

' *** Função corrigida em 09/08/2001 por Luiz G.F. Nogueira ***
' *** A correção foi feita para que a função leia também o valor do campo Beneficiamento ***

Dim lErro As Long
Dim iIndice As Integer
Dim dtDataEmissao As Date
Dim iNumItens As Integer, lCodPrestador As Long
Dim iNumItensBaixados As Integer
Dim tItemOP As typeItemOP
Dim objItemOP As ClassItemOP
Dim alComando(0 To 2) As Long
Dim tItemOPGrade As typeItemOP
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim iTipo As Integer
Dim iTipoTerc As Integer, lCodTerc As Long, iFilialTerc As Integer, iEscaninhoTerc As Integer
Dim iIgnoraEst As Integer

On Error GoTo Erro_OrdemDeProducaoBaixada_Le_ComItens

    For iIndice = LBound(alComando) To UBound(alComando)
        'Abertura comando
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 82794
    Next

    'Leitura da tabela de Ordens de Produção
    lErro = Comando_Executar(alComando(0), "SELECT DataEmissao, NumItens, NumItensBaixados, CodPrestador, Tipo, TipoTerc, CodTerc, FilialTerc, EscaninhoTerc, IgnoraEst FROM OrdensDeProducaoBaixadas WHERE FilialEmpresa = ? AND Codigo = ? ", dtDataEmissao, iNumItens, iNumItensBaixados, lCodPrestador, iTipo, iTipoTerc, lCodTerc, iFilialTerc, iEscaninhoTerc, iIgnoraEst, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 82795

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82796

    If lErro <> AD_SQL_SUCESSO Then gError 82797

    objOrdemDeProducao.dtDataEmissao = dtDataEmissao
    objOrdemDeProducao.iNumItens = iNumItens
    objOrdemDeProducao.iNumItensBaixados = iNumItensBaixados
    objOrdemDeProducao.lCodPrestador = lCodPrestador
    objOrdemDeProducao.iTipo = iTipo
    objOrdemDeProducao.iTipoTerc = iTipoTerc
    objOrdemDeProducao.lCodTerc = lCodTerc
    objOrdemDeProducao.iFilialTerc = iFilialTerc
    objOrdemDeProducao.iEscaninhoTerc = iEscaninhoTerc
    objOrdemDeProducao.iIgnoraEst = iIgnoraEst

    tItemOP.sProduto = String(STRING_PRODUTO, 0)
    tItemOP.sCcl = String(STRING_CCL, 0)
    tItemOP.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemOP.sVersao = String(STRING_KIT_VERSAO, 0)
    tItemOP.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

    'Leitura da tabela de Itens de Ordem de Produção Grade
    lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, Item, Produto, Descricao, FilialPedido, CodPedido, Almoxarifado, CCL, SiglaUM, Quantidade, DataInicioProd, DataFimProd,  Prioridade, QuantidadeProd, Situacao, Destinacao, Beneficiamento, NumIntEquipamento, Versao, ItensOPGrade.Tipo FROM ItensOPGrade, Produtos WHERE ItensOPGrade.FilialEmpresa = ? AND ItensOPGrade.Codigo = ? AND ItensOPGrade.Produto = Produtos.Codigo ORDER BY Item", _
        tItemOP.lNumIntDoc, tItemOP.iItem, tItemOP.sProduto, tItemOP.sDescricao, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.iAlmoxarifado, tItemOP.sCcl, tItemOP.sSiglaUM, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.dQuantidadeProd, tItemOP.iSituacao, tItemOP.iDestinacao, tItemOP.iBeneficiamento, tItemOP.lNumIntEquipamento, tItemOP.sVersao, tItemOP.iTipo, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 126555

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126556

    'Enquanto existirem itens para uma determinada ordem de produção
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objItemOP = New ClassItemOP

        objItemOP.lNumIntDoc = tItemOP.lNumIntDoc
        objItemOP.iItem = tItemOP.iItem
        objItemOP.sProduto = tItemOP.sProduto
        objItemOP.sDescricao = tItemOP.sDescricao
        objItemOP.iAlmoxarifado = tItemOP.iAlmoxarifado
        objItemOP.sCcl = tItemOP.sCcl
        objItemOP.sSiglaUM = tItemOP.sSiglaUM
        objItemOP.dQuantidade = tItemOP.dQuantidade
        objItemOP.dtDataInicioProd = tItemOP.dtDataInicioProd
        objItemOP.dtDataFimProd = tItemOP.dtDataFimProd
        objItemOP.iPrioridade = tItemOP.iPrioridade
        objItemOP.dQuantidadeProd = tItemOP.dQuantidadeProd
        objItemOP.iSituacao = tItemOP.iSituacao
        objItemOP.iDestinacao = tItemOP.iDestinacao
        objItemOP.iBeneficiamento = tItemOP.iBeneficiamento
        objItemOP.iFilialPedido = tItemOP.iFilialPedido
        objItemOP.lCodPedido = tItemOP.lCodPedido
        objItemOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
        objItemOP.sCodigo = objOrdemDeProducao.sCodigo
        objItemOP.iPossuiGrade = MARCADO
        
        objItemOP.lNumIntEquipamento = tItemOP.lNumIntEquipamento
        objItemOP.sVersao = tItemOP.sVersao
        objItemOP.iTipo = tItemOP.iTipo

        'Inserido por Jorge Specian - 11/05/2005
        '---------------------------------------
        lErro = CF("OrdemDeProducao_Le_Operacoes", objItemOP)
        If lErro <> SUCESSO And lErro <> 137016 Then gError 137011
        '---------------------------------------
        
        objOrdemDeProducao.colItens.Add objItemOP

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126558

    Loop


    tItemOP.sProduto = String(STRING_PRODUTO, 0)
    tItemOP.sCcl = String(STRING_CCL, 0)
    tItemOP.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemOP.sVersao = String(STRING_KIT_VERSAO, 0)
    tItemOP.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

    'Leitura da tabela de Itens de Ordem de Produção
    lErro = Comando_Executar(alComando(1), "SELECT NumIntItemOPGrade, NumIntDoc, Item, Produto, Descricao, FilialPedido, CodPedido, Almoxarifado, CCL, SiglaUM, Quantidade, DataInicioProd, DataFimProd,  Prioridade, QuantidadeProd, Situacao, Destinacao, Beneficiamento, NumIntEquipamento,NumIntEquipamento2, Versao, ItensOrdemProducaoBaixadas.Tipo, ItensOrdemProducaoBaixadas.ProduzLogo FROM ItensOrdemProducaoBaixadas, Produtos WHERE ItensOrdemProducaoBaixadas.FilialEmpresa = ? AND ItensOrdemProducaoBaixadas.Codigo = ? AND ItensOrdemProducaoBaixadas.Produto = Produtos.Codigo ORDER BY Item", _
        tItemOP.lNumIntItemOP, tItemOP.lNumIntDoc, tItemOP.iItem, tItemOP.sProduto, tItemOP.sDescricao, tItemOP.iFilialPedido, tItemOP.lCodPedido, tItemOP.iAlmoxarifado, tItemOP.sCcl, tItemOP.sSiglaUM, tItemOP.dQuantidade, tItemOP.dtDataInicioProd, tItemOP.dtDataFimProd, tItemOP.iPrioridade, tItemOP.dQuantidadeProd, tItemOP.iSituacao, tItemOP.iDestinacao, tItemOP.iBeneficiamento, tItemOP.lNumIntEquipamento, tItemOP.lNumIntEquipamento2, tItemOP.sVersao, tItemOP.iTipo, tItemOP.iProduzLogo, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 82798

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82799

    'Enquanto existirem itens para uma determinada ordem de produção
    Do While lErro <> AD_SQL_SEM_DADOS

        If tItemOP.lNumIntItemOP <> 0 Then

            Set objItemRomaneioGrade = New ClassItemRomaneioGrade
            
            objItemRomaneioGrade.sProduto = tItemOP.sProduto
            objItemRomaneioGrade.sDescricao = tItemOP.sDescricao
            objItemRomaneioGrade.iAlmoxarifado = tItemOP.iAlmoxarifado
            objItemRomaneioGrade.sUMEstoque = tItemOP.sSiglaUMEstoque
            objItemRomaneioGrade.dQuantidade = tItemOP.dQuantidade
            objItemRomaneioGrade.sVersao = tItemOP.sVersao
            objItemRomaneioGrade.lNumIntDoc = tItemOP.lNumIntDoc

            For Each objItemOP In objOrdemDeProducao.colItens
            
                If objItemOP.lNumIntDoc = tItemOP.lNumIntItemOP Then

                    objItemOP.colItensRomaneioGrade.Add objItemRomaneioGrade
                    Exit For
                    
                End If
                
            Next

        Else

            Set objItemOP = New ClassItemOP
            
            objItemOP.lNumIntDoc = tItemOP.lNumIntDoc
            objItemOP.iItem = tItemOP.iItem
            objItemOP.sProduto = tItemOP.sProduto
            objItemOP.sDescricao = tItemOP.sDescricao
            objItemOP.iAlmoxarifado = tItemOP.iAlmoxarifado
            objItemOP.sCcl = tItemOP.sCcl
            objItemOP.sSiglaUM = tItemOP.sSiglaUM
            objItemOP.dQuantidade = tItemOP.dQuantidade
            objItemOP.dtDataInicioProd = tItemOP.dtDataInicioProd
            objItemOP.dtDataFimProd = tItemOP.dtDataFimProd
            objItemOP.iPrioridade = tItemOP.iPrioridade
            objItemOP.dQuantidadeProd = tItemOP.dQuantidadeProd
            objItemOP.iSituacao = tItemOP.iSituacao
            objItemOP.iDestinacao = tItemOP.iDestinacao
            objItemOP.iBeneficiamento = tItemOP.iBeneficiamento
            objItemOP.iFilialPedido = tItemOP.iFilialPedido
            objItemOP.lCodPedido = tItemOP.lCodPedido
            objItemOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
            objItemOP.sCodigo = objOrdemDeProducao.sCodigo
            objItemOP.iPossuiGrade = DESMARCADO

            objItemOP.lNumIntEquipamento = tItemOP.lNumIntEquipamento
            objItemOP.lNumIntEquipamento2 = tItemOP.lNumIntEquipamento2
            objItemOP.sVersao = tItemOP.sVersao
            objItemOP.iTipo = tItemOP.iTipo
            objItemOP.iProduzLogo = tItemOP.iProduzLogo
    
            'Inserido por Jorge Specian - 11/05/2005
            '---------------------------------------
            lErro = CF("OrdemDeProducao_Le_Operacoes", objItemOP)
            If lErro <> SUCESSO And lErro <> 137016 Then gError 137012
            '---------------------------------------
    
            objOrdemDeProducao.colItens.Add objItemOP

        End If
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82800

    Loop

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    OrdemDeProducaoBaixada_Le_ComItens = SUCESSO

    Exit Function

Erro_OrdemDeProducaoBaixada_Le_ComItens:

    OrdemDeProducaoBaixada_Le_ComItens = gErr

    Select Case gErr

        Case 82794
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 82795, 82796
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAOBAIXADAS", gErr)

        Case 82797 'Ordem de Producao Baixada nao encontrada

        Case 82798, 82799, 82800
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAOBAIXADAS", gErr)

        Case 126555, 126556, 126558
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOPGRADE1", gErr, objOrdemDeProducao.iFilialEmpresa, objOrdemDeProducao.sCodigo)

        Case 137011, 137012
            'erros tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152130)

    End Select

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Function MovimentoEstoque_Le_Primeira_Data(ByVal iFilialEmpresa As Integer, ByVal objItemMovEstoque As ClassItemMovEstoque) As Long
'Função criada por Luiz G.F. Nogueira em 13/08/01
'Lê a primeira data de movimento cadastrado e a coloca em objItemMovEstoque

Dim lErro As Long
Dim lComando As Long
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovimentoEstoque_Le_Primeira_Data

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 90705

    'le a ultima data cadastrada
    lErro = Comando_Executar(lComando, "SELECT MIN(Data) FROM MovimentoEstoque WHERE FilialEmpresa = ?", tItemMovEstoque.dtData, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 90706

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90707

    'se o movimento não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Or tItemMovEstoque.dtData = DATA_NULA Then gError 90708
    
    objItemMovEstoque.dtData = tItemMovEstoque.dtData
    
    Call Comando_Fechar(lComando)

    MovimentoEstoque_Le_Primeira_Data = SUCESSO
    
    Exit Function

Erro_MovimentoEstoque_Le_Primeira_Data:

    MovimentoEstoque_Le_Primeira_Data = gErr

    Select Case gErr

        Case 90705
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 90706, 90707
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, iFilialEmpresa)

        Case 90708
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152131)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Rotina criada por Daniel
'Revisão: Luiz G.F.Nogueira
'Data da inclusão no MATSelect: 14/09/01
'Pendências: não
Public Function Embalagem_Le_Sigla(ByVal objEmbalagem As ClassEmbalagem) As Long
'Lê a embalagem a partir da sigla passada como parametro em objEmbalagem

Dim lErro As Long
Dim lComando As Long
Dim tEmbalagem As typeEmbalagem

On Error GoTo Erro_Embalagem_Le_Sigla

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 95085
    
    'Inicializa as strings
    tEmbalagem.sDescricao = String(STRING_EMBALAGEM_DESCRICAO, 0)
    tEmbalagem.sSigla = String(STRING_EMBALAGEM_SIGLA, 0)
    
    'Busca no BD uma embalagem com a sigla passada
    lErro = Comando_Executar(lComando, "SELECT Codigo, Descricao, Sigla, Capacidade, Peso FROM Embalagens WHERE Sigla = ?", tEmbalagem.iCodigo, tEmbalagem.sDescricao, tEmbalagem.sSigla, tEmbalagem.dCapacidade, tEmbalagem.dPeso, objEmbalagem.sSigla)
    If lErro <> AD_SQL_SUCESSO Then gError 95086
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 95087
    
    'Se encontrou ==> erro
    If lErro <> AD_SQL_SUCESSO Then gError 95088
    
    'Guarda no obj os dados lidos do BD
    objEmbalagem.iCodigo = tEmbalagem.iCodigo
    objEmbalagem.sDescricao = tEmbalagem.sDescricao
    objEmbalagem.sSigla = tEmbalagem.sSigla
    objEmbalagem.dCapacidade = tEmbalagem.dCapacidade
    objEmbalagem.dPeso = tEmbalagem.dPeso
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Embalagem_Le_Sigla = SUCESSO
    
    Exit Function
    
Erro_Embalagem_Le_Sigla:

    Embalagem_Le_Sigla = gErr
    
    Select Case gErr
    
        Case 95085
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 95086, 95087
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)
            
        Case 95088
            'Erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152132)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Rotina criada por Daniel
'Revisão: Luiz G.F.Nogueira
'Data da inclusão no MATSelect: 14/09/01
'Pendências: não
Public Function ProdutoEmbalagem_Le_Produto(ByVal objProduto As ClassProduto) As Long
'Lê as embalagens associadas a um produto a partir do código de produto passado como parâmetro em objProduto
'Devolve as embalagens na coleção colProdutoEmbalagem dentro do objProduto

Dim lErro As Long
Dim lComando As Long
Dim tProdutoEmbalagem As typeProdutoEmbalagem
Dim objProdutoEmbalagem As ClassProdutoEmbalagem

On Error GoTo Erro_ProdutoEmbalagem_Le_Produto

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 95013
    
    With tProdutoEmbalagem
        
        'Inicializa a colecao existente dentro do obj
        Set objProduto.colProdutoEmbalagem = New Collection
    
        'Inicializa as strings
        .sProduto = String(STRING_PRODUTOEMBALAGEM_PRODUTO, 0)
        .sNomeProdEmb = String(STRING_PRODUTOEMBALAGEM_NOMEPRODEMB, 0)
        .sUMEmbalagem = String(STRING_PRODUTOEMBALAGEM_UMEMBALAGEM, 0)
        .sUMPeso = String(STRING_PRODUTOEMBALAGEM_UMPESO, 0)
        
        'Lê dados de ProdutoEmbalagem a partir de Produto
        lErro = Comando_Executar(lComando, "SELECT Produto, Embalagem, SeqGrid, Padrao, NomeProdEmb, UMEmbalagem, Capacidade, UMPeso, PesoLiqTotal, PesoBruto FROM ProdutoEmbalagem WHERE Produto = ? ORDER BY SeqGrid", .sProduto, .iEmbalagem, .iSeqGrid, .iPadrao, .sNomeProdEmb, .sUMEmbalagem, .dCapacidade, .sUMPeso, .dPesoLiqTotal, .dPesoBruto, objProduto.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 95014
    
    End With
    
    'Buscar o Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 95015

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 95016
    
    'Busca todos os registros filtrados pelo SELECT
    While lErro = AD_SQL_SUCESSO
        
        'Inicializa o objProdutoEmbalagem
        Set objProdutoEmbalagem = New ClassProdutoEmbalagem
        
        'Preenche o obj com os dados lidos no BD
        With objProdutoEmbalagem
                
            .iEmbalagem = tProdutoEmbalagem.iEmbalagem
            .iSeqGrid = tProdutoEmbalagem.iSeqGrid
            .iPadrao = tProdutoEmbalagem.iPadrao
            .sNomeProdEmb = tProdutoEmbalagem.sNomeProdEmb
            .sUMEmbalagem = tProdutoEmbalagem.sUMEmbalagem
            .dCapacidade = tProdutoEmbalagem.dCapacidade
            .sUMPeso = tProdutoEmbalagem.sUMPeso
            .dPesoLiqTotal = tProdutoEmbalagem.dPesoLiqTotal
            .dPesoBruto = tProdutoEmbalagem.dPesoBruto
        
        End With
            
        'Adiciona o Códgo
        objProduto.sCodigo = tProdutoEmbalagem.sProduto
        
        'adiciona na colecao
        objProduto.colProdutoEmbalagem.Add objProdutoEmbalagem
        
        'busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 95017
        
    Wend
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    ProdutoEmbalagem_Le_Produto = SUCESSO
    
    Exit Function

Erro_ProdutoEmbalagem_Le_Produto:

    ProdutoEmbalagem_Le_Produto = gErr
    
    Select Case gErr
    
        Case 95013
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 95014, 95015, 95017
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOEMBALAGEM", gErr)
        
        Case 95016
            'Erro tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152133)
            
    End Select
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function MovimentoEstoque_Le_Producoes(ByVal objItemMovEst As ClassItemMovEstoque, ByVal colItemMovEst As Collection) As Long
'le as produções para este produto para o codigo de FilialEmpresa/O.P. passado como parametro em objItemMovEst e devolve-os em colItemMovEst

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim sSiglaUM As String, dtData As Date
Dim dQuantidade As Double
Dim objItemMovEst1 As ClassItemMovEstoque
Dim colApropriacaoInsumo As Collection

On Error GoTo Erro_MovimentoEstoque_Le_Producoes

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92481

    sSiglaUM = String(STRING_UM_SIGLA, 0)

    'le as produções para o produto/ Codigo da O.P./FilialEmpresa
    lErro = Comando_Executar(lComando, "SELECT Data, NumIntDoc, SiglaUM, Quantidade FROM MovimentoEstoque WHERE FilialEmpresa = ? AND Produto = ? AND CodigoOP = ? AND (TipoMov = ? Or TipoMov =?)", dtData, lNumIntDoc, sSiglaUM, dQuantidade, objItemMovEst.iFilialEmpresa, objItemMovEst.sProduto, objItemMovEst.sOPCodigo, MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3)
    If lErro <> AD_SQL_SUCESSO Then gError 92482

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92483

    Do While lErro = AD_SQL_SUCESSO

        Set objItemMovEst1 = New ClassItemMovEstoque
        
        objItemMovEst1.dtData = dtData
        objItemMovEst1.lNumIntDoc = lNumIntDoc
        objItemMovEst1.sSiglaUM = sSiglaUM
        objItemMovEst1.dQuantidade = dQuantidade
        
        Set colApropriacaoInsumo = New Collection
        
        'Le as Apropriações do Item
        lErro = CF("ApropriacaoInsumo_Le_NumIntDocOrigem", lNumIntDoc, colApropriacaoInsumo)
        If lErro <> SUCESSO Then gError 92653
                        
        Set objItemMovEst1.colApropriacaoInsumo = colApropriacaoInsumo
        
        colItemMovEst.Add objItemMovEst1
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92651
        
    Loop

    Call Comando_Fechar(lComando)

    MovimentoEstoque_Le_Producoes = SUCESSO
    
    Exit Function

Erro_MovimentoEstoque_Le_Producoes:

    MovimentoEstoque_Le_Producoes = gErr

    Select Case gErr

        Case 92653
        
        Case 92481
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92482, 92483, 92651
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE6", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.sProduto, objItemMovEst.sOPCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152134)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function
        
Function EstoqueMes_Le_HorasMaquina(ByVal objEstoqueMes As ClassEstoqueMes) As Long
'Le na tabela EstoqueMes o total de horas maquina e custos de insumos utilizados na fabricação do mes/ano passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim sMes As String
Dim dCustoMatPrim As Double
Dim lHorasMaquina As Long
Dim dGastosDiretos As Double
Dim dGastosIndiretos As Double

On Error GoTo Erro_EstoqueMes_Le_HorasMaquina

    sMes = CStr(objEstoqueMes.iMes)

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92500

    lErro = Comando_Executar(lComando, "SELECT HorasMaquina , CustoMatPrim, GastosIndiretos, GastosDiretos FROM SldMesEst WHERE  FilialEmpresa = ? AND Ano = ? AND Mes = ?", lHorasMaquina, dCustoMatPrim, dGastosIndiretos, dGastosDiretos, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
    If lErro <> AD_SQL_SUCESSO Then gError 92501

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92502

    If lErro = AD_SQL_SEM_DADOS Then gError 92503

    objEstoqueMes.lHorasMaquina = lHorasMaquina
    objEstoqueMes.dCustoMatPrim = dCustoMatPrim
    objEstoqueMes.dGastosDiretos = dGastosDiretos
    objEstoqueMes.dGastosIndiretos = dGastosIndiretos

    Call Comando_Fechar(lComando)

    EstoqueMes_Le_HorasMaquina = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_HorasMaquina:

    EstoqueMes_Le_HorasMaquina = gErr

    Select Case gErr

        Case 92500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92501, 92502
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
        
        Case 92503
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152135)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovEstoque_Le_HorasMaq_CustoMPrim(ByVal colEstoqueMes As Collection, ByVal iMes As Integer, ByVal iAno As Integer, ByVal colEstoqueMesProduto As Collection) As Long
'Apura o total de horas maquina e custo das matérias primas para cada filial e coloca em EstoqueMes
'colEstoqueMesProduto contém os produtos para os quais foram informados os gastos para produzi-los e portanto não necessita calcular ou incluir as materias primas/horas trabalhadas utilizadas por estes na producao

Dim lErro As Long
Dim lErro1 As Long
Dim alComando(1 To 2) As Long
Dim dtDataFinal As Date
Dim dtDataInicial As Date
Dim dCustoMatPrim As Double
Dim dCustoInsumos As Double
Dim iFilialEmpresaAtual As Integer
Dim objEstoqueMes As ClassEstoqueMes
Dim iAchou As Integer
Dim iIndice As Integer
Dim lTotalHorasMaquina As Long
Dim dTotalCustoMatPrim As Double
Dim objItemMovEst As New ClassItemMovEstoque
Dim tItemMovEst As typeItemMovEstoque
Dim colItemMovEst As Collection
Dim colApropriacaoInsumo As Collection
Dim objOrdemProducao As ClassOrdemDeProducao
Dim lNumIntDoc As Long
Dim iCalculouCustoMatPrim As Integer
Dim objEstoqueMesProduto As ClassEstoqueMesProduto
Dim objProduto As New ClassProduto
Dim dQuantFator1 As Double, dQuantFator2 As Double, dQuantFator3 As Double, dQuantFator4 As Double, dQuantFator5 As Double, dQuantFator6 As Double
Dim dTotQuantFator1 As Double, dTotQuantFator2 As Double, dTotQuantFator3 As Double, dTotQuantFator4 As Double, dTotQuantFator5 As Double, dTotQuantFator6 As Double
Dim bCalculaAprop As Boolean, objItemOP As ClassItemOP

On Error GoTo Erro_MovEstoque_Le_HorasMaq_CustoMPrim

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92527
    Next

    'Preenche as datas inicial e final
    dtDataInicial = CDate(1 & "/" & iMes & "/" & iAno)
    dtDataFinal = DateAdd("m", 1, dtDataInicial) - 1
    
    'zera a quantidade total produzida de cada produto no mes
    For Each objEstoqueMesProduto In colEstoqueMesProduto
        objEstoqueMesProduto.dQuantidade = 0
    Next
    
    With tItemMovEst
    
        .sOPCodigo = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUM = String(STRING_BUFFER_MAX_TEXTO, 0)
    
        lErro = Comando_Executar(alComando(1), "SELECT FilialEmpresa, HorasMaquina, NumIntDoc, NumIntDocOrigem, CodigoOP, Produto, SiglaUM, Quantidade, Data, Hora FROM MovimentoEstoque WHERE Data >=? AND Data <=? AND (TipoMov = ? Or TipoMov = ?) And NumIntDocEst = 0 ORDER BY FilialEmpresa, Data, NumIntDoc", .iFilialEmpresa, .lHorasMaquina, .lNumIntDoc, .lNumIntDocOrigem, .sOPCodigo, .sProduto, .sSiglaUM, .dQuantidade, .dtData, .dHora, dtDataInicial, dtDataFinal, MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3)
        If lErro <> AD_SQL_SUCESSO Then gError 92528

    End With

    lErro1 = Comando_BuscarPrimeiro(alComando(1))
    If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 92529

    iFilialEmpresaAtual = tItemMovEst.iFilialEmpresa
    
    'Guarda o valor em lErro para não afetar o Loop _
    OBS -> lErro1 é uma "Flag" usada para saber se há pelo menos um Movto.
    lErro = lErro1

    Do While lErro = AD_SQL_SUCESSO
    
        With tItemMovEst
            objItemMovEst.iFilialEmpresa = .iFilialEmpresa
            objItemMovEst.lNumIntDoc = .lNumIntDoc
            objItemMovEst.lNumIntDocOrigem = .lNumIntDocOrigem
            objItemMovEst.sOPCodigo = .sOPCodigo
            objItemMovEst.sProduto = .sProduto
            objItemMovEst.sSiglaUM = .sSiglaUM
            objItemMovEst.dQuantidade = .dQuantidade
            objItemMovEst.dtData = .dtData
            objItemMovEst.dtHora = .dHora
        End With
    
        iCalculouCustoMatPrim = 0
    
        dCustoMatPrim = 0
    
        If iFilialEmpresaAtual <> tItemMovEst.iFilialEmpresa Then
    
            lErro = Armazena_HorasMaquina_CustoMatPrim(colEstoqueMes, iFilialEmpresaAtual, lTotalHorasMaquina, dTotalCustoMatPrim, iMes, iAno, dTotQuantFator1, dTotQuantFator2, dTotQuantFator3, dTotQuantFator4, dTotQuantFator5, dTotQuantFator6)
            If lErro <> AD_SQL_SUCESSO Then gError 92532
    
            iFilialEmpresaAtual = tItemMovEst.iFilialEmpresa
            lTotalHorasMaquina = 0
            dTotalCustoMatPrim = 0
            dTotQuantFator1 = 0
            dTotQuantFator2 = 0
            dTotQuantFator3 = 0
            dTotQuantFator4 = 0
            dTotQuantFator5 = 0
            dTotQuantFator6 = 0
                    
        End If
    
        iAchou = 0
    
        'se o produto é um dos que foram informados os gastos ==> não participa da acumulação de horas maquina ou custo de materias primas
        For Each objEstoqueMesProduto In colEstoqueMesProduto
            If tItemMovEst.iFilialEmpresa = objEstoqueMesProduto.iFilialEmpresa And tItemMovEst.sProduto = objEstoqueMesProduto.sProduto Then
                iAchou = 1
                Exit For
            End If
        Next
    
        'se não foi informado pelo usuário os gastos para produzir o produto em questão ==> tem que calcular o custo das materias primas utilizadas e as horas trabalhadas.
        If iAchou = 0 Then
    
            Set colApropriacaoInsumo = New Collection
            
            lNumIntDoc = tItemMovEst.lNumIntDoc
            
            'Le as Apriações do Item
            lErro = CF("ApropriacaoInsumo_Le_NumIntDocOrigem", lNumIntDoc, colApropriacaoInsumo)
            If lErro <> SUCESSO Then gError 92662
                            
            Set objItemMovEst.colApropriacaoInsumo = colApropriacaoInsumo
        
            If colApropriacaoInsumo.Count > 0 Then
            
                lErro = CF("Calcula_Custo_MatPrim_Apropriacao", objItemMovEst, dCustoMatPrim, dCustoInsumos)
                If lErro <> SUCESSO Then gError 92665
            
                iCalculouCustoMatPrim = 1
            
            Else
            
                bCalculaAprop = False
                'se o produto não tem apropriação de insumos associado ==> a O.P. precisa estar baixada
                Set objOrdemProducao = New ClassOrdemDeProducao
            
                objOrdemProducao.iFilialEmpresa = objItemMovEst.iFilialEmpresa
                objOrdemProducao.sCodigo = objItemMovEst.sOPCodigo
            
                'pesquisa se a O.P. associada está baixada
                lErro = CF("OrdemProducao_Le1", objOrdemProducao)
                If lErro <> SUCESSO And lErro <> 94578 And lErro <> 94579 Then gError 92666
        
                'se a O.P. não estiver cadastrada ==> erro
                If lErro = 94579 Then gError 92667
                
                If lErro = 94578 Then bCalculaAprop = True
                
                If Not bCalculaAprop Then
                
                    Set objItemOP = New ClassItemOP
                    
                    objItemOP.lNumIntDoc = objItemMovEst.lNumIntDocOrigem
                
                    lErro = CF("ItemOP_Le_NumIntDoc", objItemOP, True)
                    If lErro <> SUCESSO Then gError 92666
                    
                    If objItemOP.iSituacao = ITEMOP_SITUACAO_BAIXADA Then bCalculaAprop = True
                
                End If
        
                'se a O.P. estiver baixada
                If bCalculaAprop Then
            
                    Set colItemMovEst = New Collection
                    
                    'pesquisa as produções para este produto associada a esta O.P.
                    lErro = CF("MovimentoEstoque_Le_Producoes", objItemMovEst, colItemMovEst)
                    If lErro <> SUCESSO Then gError 92654
                    
                    'pesquisa o custo das materias primas e insumos associadas à produção
                    lErro = CF("MovEstoque_Le_CustoMatPrim", objItemMovEst, dCustoMatPrim, dCustoInsumos, colItemMovEst)
                    If lErro <> SUCESSO And lErro <> 92489 Then gError 92531
                    
                    iCalculouCustoMatPrim = 1
        
                End If
        
            End If
    
            If iCalculouCustoMatPrim = 1 Then
        
                'Calcular pesos
                lErro = CF("Producao_CalculaFatores", objItemMovEst, dQuantFator1, dQuantFator2, dQuantFator3, dQuantFator4, dQuantFator5, dQuantFator6)
                If lErro <> SUCESSO Then gError 92654
        
                dTotQuantFator1 = dTotQuantFator1 + dQuantFator1
                dTotQuantFator2 = dTotQuantFator2 + dQuantFator2
                dTotQuantFator3 = dTotQuantFator3 + dQuantFator3
                dTotQuantFator4 = dTotQuantFator4 + dQuantFator4
                dTotQuantFator5 = dTotQuantFator5 + dQuantFator5
                dTotQuantFator6 = dTotQuantFator6 + dQuantFator6
                
                lTotalHorasMaquina = lTotalHorasMaquina + tItemMovEst.lHorasMaquina
                dTotalCustoMatPrim = dTotalCustoMatPrim + dCustoMatPrim
        
            End If
    
        Else
            
            'acumula a quantidade produzida do produto em questão em objEstoqueMesProduto
            lErro = Acumula_QuantProduzida_Produto(objEstoqueMesProduto, objItemMovEst)
            If lErro <> SUCESSO Then gError 92895
    
        End If
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92530
    
    Loop
    
    'Se tem pelo menos 1 movimento => Armazena as HorasMaquina...
    If lErro1 <> AD_SQL_SEM_DADOS Then
        
        lErro = Armazena_HorasMaquina_CustoMatPrim(colEstoqueMes, iFilialEmpresaAtual, lTotalHorasMaquina, dTotalCustoMatPrim, iMes, iAno, dTotQuantFator1, dTotQuantFator2, dTotQuantFator3, dTotQuantFator4, dTotQuantFator5, dTotQuantFator6)
        If lErro <> AD_SQL_SUCESSO Then gError 92533
        
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    MovEstoque_Le_HorasMaq_CustoMPrim = SUCESSO

    Exit Function

Erro_MovEstoque_Le_HorasMaq_CustoMPrim:

    MovEstoque_Le_HorasMaq_CustoMPrim = gErr

    Select Case gErr

        Case 92527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92528, 92529, 92530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 92531, 92532, 92533, 92654, 92662, 92665, 92666, 92895

        Case 92667
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OP_INEXISTENTE_MOVESTOQUE", gErr, objItemMovEst.sOPCodigo, objItemMovEst.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152136)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function Acumula_QuantProduzida_Produto(ByVal objEstoqueMesProduto As ClassEstoqueMesProduto, ByVal objItemMovEst As ClassItemMovEstoque) As Long
'acumula a quantidade produto do produto em questão em objEstoqueMesProduto

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim dFator As Double

On Error GoTo Erro_Acumula_QuantProduzida_Produto

    objProduto.sCodigo = objItemMovEst.sProduto
                
    'le o produto associado a requisição para descobrir se a natureza do produto é materia prima.
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 92892
    
    If lErro <> SUCESSO Then gError 92893

    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemMovEst.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 92894
        
    objItemMovEst.dQuantidadeEst = objItemMovEst.dQuantidade * dFator

    objEstoqueMesProduto.dQuantidade = objEstoqueMesProduto.dQuantidade + objItemMovEst.dQuantidadeEst

    Acumula_QuantProduzida_Produto = SUCESSO
    
    Exit Function
    
Erro_Acumula_QuantProduzida_Produto:

    Select Case gErr

        Case 92892, 92894

        Case 92893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152137)
            
    End Select
    
    Exit Function

End Function

Private Function Armazena_HorasMaquina_CustoMatPrim(ByVal colEstoqueMes As Collection, ByVal iFilialEmpresaAtual As Integer, ByVal lTotalHorasMaquina As Long, ByVal dTotalCustoMatPrim As Double, ByVal iMes As Integer, ByVal iAno As Integer, ByVal dTotQuantFator1 As Double, ByVal dTotQuantFator2 As Double, ByVal dTotQuantFator3 As Double, ByVal dTotQuantFator4 As Double, ByVal dTotQuantFator5 As Double, ByVal dTotQuantFator6 As Double) As Long

Dim iAchou As Integer
Dim lErro As Long
Dim objEstoqueMes As ClassEstoqueMes

On Error GoTo Erro_Armazena_HorasMaquina_CustoMatPrim

    iAchou = 0

    For Each objEstoqueMes In colEstoqueMes
    
        If objEstoqueMes.iFilialEmpresa = iFilialEmpresaAtual Then
        
            With objEstoqueMes
                
                .lHorasMaquina = lTotalHorasMaquina
                .dCustoMatPrim = dTotalCustoMatPrim
                .dQuantFator1 = dTotQuantFator1
                .dQuantFator2 = dTotQuantFator2
                .dQuantFator3 = dTotQuantFator3
                .dQuantFator4 = dTotQuantFator4
                .dQuantFator5 = dTotQuantFator5
                .dQuantFator6 = dTotQuantFator6
                
            End With
            
            iAchou = 1
            
        End If
    Next

    If iAchou = 0 Then
    
        Set objEstoqueMes = New ClassEstoqueMes
        With objEstoqueMes
            .iFilialEmpresa = iFilialEmpresaAtual
            .lHorasMaquina = lTotalHorasMaquina
            .dCustoMatPrim = dTotalCustoMatPrim
            .iMes = iMes
            .iAno = iAno
            .dQuantFator1 = dTotQuantFator1
            .dQuantFator2 = dTotQuantFator2
            .dQuantFator3 = dTotQuantFator3
            .dQuantFator4 = dTotQuantFator4
            .dQuantFator5 = dTotQuantFator5
            .dQuantFator6 = dTotQuantFator6
        End With
        colEstoqueMes.Add objEstoqueMes
    End If

    Armazena_HorasMaquina_CustoMatPrim = SUCESSO
    
    Exit Function

Erro_Armazena_HorasMaquina_CustoMatPrim:

    Armazena_HorasMaquina_CustoMatPrim = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152138)

    End Select
    
    Exit Function
    
End Function

Function EstoqueMes_Le_Anos1(ByVal colAnos As Collection) As Long
'Le todos os anos da tabela EstoqueMes

Dim lErro As Long
Dim sComando_SQL As String
Dim lComando As Long
Dim iAno As Integer

On Error GoTo Erro_EstoqueMes_Le_Anos1

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92561

    lErro = Comando_Executar(lComando, "SELECT DISTINCT Ano FROM EstoqueMes ORDER BY Ano", iAno)
    If lErro <> AD_SQL_SUCESSO Then gError 92562

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92563

    If lErro = AD_SQL_SEM_DADOS Then gError 92564

    Do While lErro <> AD_SQL_SEM_DADOS

        colAnos.Add iAno

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92565

    Loop

    Call Comando_Fechar(lComando)

    EstoqueMes_Le_Anos1 = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_Anos1:

    EstoqueMes_Le_Anos1 = gErr

    Select Case gErr

        Case 92561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 92562, 92563, 92565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", gErr)

        Case 92564
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152139)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function EstoqueMes_Le_Meses1(ByVal iAno As Integer, ByVal colMeses As Collection) As Long
'Le todos os meses da tabela EstoqueMes com Ano = iAno

Dim lErro As Long
Dim sComando_SQL As String
Dim lComando As Long
Dim iMes As Integer

On Error GoTo Erro_EstoqueMes_Le_Meses1

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92565

    lErro = Comando_Executar(lComando, "SELECT DISTINCT Mes FROM EstoqueMes WHERE Ano = ? ORDER BY Mes", iMes, iAno)
    If lErro <> AD_SQL_SUCESSO Then gError 92566

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92567

    If lErro = AD_SQL_SEM_DADOS Then gError 92568

    Do While lErro <> AD_SQL_SEM_DADOS

        colMeses.Add iMes

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92569

    Loop

    Call Comando_Fechar(lComando)

    EstoqueMes_Le_Meses1 = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_Meses1:

    EstoqueMes_Le_Meses1 = gErr

    Select Case gErr

        Case 92565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92566, 92567, 92569
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", gErr)
        
        Case 92568
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152140)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function EstoqueMes_Le_FiliaisEmpresa(ByVal iAno As Integer, ByVal iMes As Integer, ByVal colEstoqueMes As Collection) As Long
'Le todas as filiaisEmpresa da tabela EstoqueMes para o Ano/Mes em questão

Dim lErro As Long
Dim lComando As Long
Dim sNomeFilialEmpresa
Dim iFilialEmpresa As Integer
Dim objEstoqueMes As ClassEstoqueMes
Dim dGastosDiretos As Double
Dim dGastosIndiretos As Double
Dim iCustoProdApuracao As Integer
Dim dCustoFator1 As Double, dCustoFator2 As Double, dCustoFator3 As Double, dCustoFator4 As Double, dCustoFator5 As Double, dCustoFator6 As Double

On Error GoTo Erro_EstoqueMes_Le_FiliaisEmpresa

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92570

    sNomeFilialEmpresa = String(STRING_FILIAL_NOME, 0)

    lErro = Comando_Executar(lComando, "SELECT EstoqueMes.FilialEmpresa, Nome, GastosDiretos, GastosIndiretos, CustoProdApurado, CustoFator1, CustoFator2, CustoFator3, CustoFator4, CustoFator5, CustoFator6 FROM EstoqueMes, FiliaisEmpresa WHERE Ano = ? AND Mes = ? AND EstoqueMes.FilialEmpresa = FiliaisEmpresa.FilialEmpresa ORDER BY EstoqueMes.FilialEmpresa", iFilialEmpresa, sNomeFilialEmpresa, dGastosDiretos, dGastosIndiretos, iCustoProdApuracao, dCustoFator1, dCustoFator2, dCustoFator3, dCustoFator4, dCustoFator5, dCustoFator6, iAno, iMes)
    If lErro <> AD_SQL_SUCESSO Then gError 92571

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92572

    If lErro = AD_SQL_SEM_DADOS Then gError 92573

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objEstoqueMes = New ClassEstoqueMes

        objEstoqueMes.iFilialEmpresa = iFilialEmpresa
        objEstoqueMes.sNomeFilialEmpresa = sNomeFilialEmpresa
        objEstoqueMes.dGastosDiretos = dGastosDiretos
        objEstoqueMes.dGastosIndiretos = dGastosIndiretos
        objEstoqueMes.iCustoProdApurado = iCustoProdApuracao
        objEstoqueMes.dCustoFator1 = dCustoFator1
        objEstoqueMes.dCustoFator2 = dCustoFator2
        objEstoqueMes.dCustoFator3 = dCustoFator3
        objEstoqueMes.dCustoFator4 = dCustoFator4
        objEstoqueMes.dCustoFator5 = dCustoFator5
        objEstoqueMes.dCustoFator6 = dCustoFator6

        colEstoqueMes.Add objEstoqueMes

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92574

    Loop

    Call Comando_Fechar(lComando)

    EstoqueMes_Le_FiliaisEmpresa = SUCESSO

    Exit Function

Erro_EstoqueMes_Le_FiliaisEmpresa:

    EstoqueMes_Le_FiliaisEmpresa = gErr

    Select Case gErr

        Case 92570
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92571, 92572, 92574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES1", gErr)
        
        Case 92573
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152141)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Calcula_Custo_MatPrim_Apropriacao(ByVal objItemMovEst As ClassItemMovEstoque, dCustoMatPrim As Double, dCustoInsumos As Double) As Long
'calcula o custo das materias primas e dos insumos para a produção em questão em função das apropriações associadas.

Dim lComando As Long
Dim objItemMovEst1 As ClassItemMovEstoque
Dim lErro As Long
Dim objProduto1 As ClassProduto
Dim dFator As Double
Dim objApropriacaoInsumo As ClassApropriacaoInsumosProd

On Error GoTo Erro_Calcula_Custo_MatPrim_Apropriacao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 92663

    dCustoMatPrim = 0
    
    For Each objApropriacaoInsumo In objItemMovEst.colApropriacaoInsumo
        
        Set objItemMovEst1 = New ClassItemMovEstoque
        
        objItemMovEst1.lNumIntDoc = objApropriacaoInsumo.lNumIntReqProd
        
        'Carrega os dados do movimento de requisição de produção com o NumIntDoc passado como parametro
        lErro = CF("MovimentoEstoque_Le_NumIntDoc", lComando, objItemMovEst1)
        If lErro <> SUCESSO And lErro <> 83748 Then gError 92507

        'se a requisição não existir ==> erro
        If lErro = 83748 Then gError 92508
        
        'se a requisição estiver estornada ==> erro
        If objItemMovEst1.lNumIntDocEst <> 0 Then gError 92509

        Set objProduto1 = New ClassProduto
        
        objProduto1.sCodigo = objItemMovEst1.sProduto
                    
        'le o produto associado a requisição para descobrir se a natureza do produto é materia prima.
        lErro = CF("Produto_Le", objProduto1)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 92510
        
        If lErro <> SUCESSO Then gError 92511

        'se nao for insumo fornecido por 3os p/beneficiamento
        If objItemMovEst1.iTipoMov <> MOV_EST_REQ_PRODUCAO_BENEF3 Or (objProduto1.iCompras = PRODUTO_PRODUZIVEL) Then
        
            lErro = CF("UM_Conversao_Trans", objProduto1.iClasseUM, objItemMovEst1.sSiglaUM, objProduto1.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 92512
                
            objItemMovEst1.dQuantidadeEst = objItemMovEst1.dQuantidade * dFator
                
            'apura o custo pelo proporção do material apropriado em relação ao material requisitado
            dCustoInsumos = dCustoInsumos + (objApropriacaoInsumo.dQuantidade / objItemMovEst1.dQuantidadeEst) * objItemMovEst1.dCusto
            If objProduto1.iNatureza = NATUREZA_PROD_MATERIA_PRIMA Then
                dCustoMatPrim = dCustoMatPrim + (objApropriacaoInsumo.dQuantidade / objItemMovEst1.dQuantidadeEst) * objItemMovEst1.dCusto
            End If

        End If
        
    Next
    
    Call Comando_Fechar(lComando)
    
    Calcula_Custo_MatPrim_Apropriacao = SUCESSO
    
    Exit Function

Erro_Calcula_Custo_MatPrim_Apropriacao:

    Calcula_Custo_MatPrim_Apropriacao = gErr
    
    Select Case gErr
    
        Case 92507, 92510, 92512
            
        Case 92508
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APROPRIACAO_REQUISICAO_INEXISTENTE", gErr, objApropriacaoInsumo.sProduto, objApropriacaoInsumo.lNumIntReqProd)
        
        Case 92509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APROPRIACAO_REQUISICAO_ESTORNADA", gErr, objApropriacaoInsumo.sProduto, objApropriacaoInsumo.lNumIntReqProd)
        
        Case 92511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objItemMovEst1.sProduto)
        
        Case 92663
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152142)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function OrdemProducao_Le1(ByVal objOrdemProducao As ClassOrdemDeProducao, Optional ByVal bLeBaixada As Boolean = False) As Long
'Lê a Ordem de Produção filtrando pelo Codigo e FilialEmpresa passados em objOrdemProducao

Dim lErro As Long
Dim lComando As Long
Dim tOrdemProducao As typeOrdemProducao

On Error GoTo Erro_OrdemProducao_Le1

    'Abertura de comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94573

    With tOrdemProducao
        
        .sCodigo = String(STRING_OPCODIGO, 0)
        .sOPGeradora = String(STRING_OPCODIGO, 0)

        lErro = Comando_Executar(lComando, "SELECT DataEmissao, NumItens, NumItensBaixados, CodPrestador, FilialEmpresa, Tipo, Codigo, OPGeradora, TipoTerc, CodTerc, FilialTerc, EscaninhoTerc, IgnoraEst FROM OrdensDeProducao WHERE FilialEmpresa = ? AND Codigo = ?", .dtDataEmissao, .iNumItens, .iNumItensBaixados, .lCodPrestador, .iFilialEmpresa, .iTipo, .sCodigo, .sOPGeradora, .iTipoTerc, .lCodTerc, .iFilialTerc, .iEscaninhoTerc, .iIgnoraEst, objOrdemProducao.iFilialEmpresa, objOrdemProducao.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 94574

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94575

        If lErro = AD_SQL_SEM_DADOS Then
        
            lErro = Comando_Executar(lComando, "SELECT DataEmissao, NumItens, NumItensBaixados, CodPrestador, FilialEmpresa, Tipo, Codigo, OPGeradora, TipoTerc, CodTerc, FilialTerc, EscaninhoTerc, IgnoraEst  FROM OrdensDeProducaoBaixadas WHERE FilialEmpresa = ? AND Codigo = ?", .dtDataEmissao, .iNumItens, .iNumItensBaixados, .lCodPrestador, .iFilialEmpresa, .iTipo, .sCodigo, .sOPGeradora, .iTipoTerc, .lCodTerc, .iFilialTerc, .iEscaninhoTerc, .iIgnoraEst, objOrdemProducao.iFilialEmpresa, objOrdemProducao.sCodigo)
            If lErro <> AD_SQL_SUCESSO Then gError 94576
    
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94577
        
            If lErro = AD_SQL_SUCESSO And Not bLeBaixada Then gError 94578
        
            If lErro <> AD_SQL_SUCESSO Then gError 94579
            
        End If

        objOrdemProducao.iNumItens = .iNumItens
        objOrdemProducao.dtDataEmissao = .dtDataEmissao
        objOrdemProducao.iNumItensBaixados = .iNumItensBaixados
        objOrdemProducao.lCodPrestador = .lCodPrestador
        objOrdemProducao.iFilialEmpresa = .iFilialEmpresa
        objOrdemProducao.iTipo = .iTipo
        objOrdemProducao.sCodigo = .sCodigo
        objOrdemProducao.sOPGeradora = .sOPGeradora
        objOrdemProducao.iTipoTerc = .iTipoTerc
        objOrdemProducao.lCodTerc = .lCodTerc
        objOrdemProducao.iFilialTerc = .iFilialTerc
        objOrdemProducao.iEscaninhoTerc = .iEscaninhoTerc
        objOrdemProducao.iIgnoraEst = .iIgnoraEst

    End With

    Call Comando_Fechar(lComando)

    OrdemProducao_Le1 = SUCESSO

    Exit Function

Erro_OrdemProducao_Le1:

    OrdemProducao_Le1 = gErr

    Select Case gErr

        Case 94573
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 94574, 94575
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAO", Err)

        Case 94578, 94579 'Tratado na Rotina Chamadora

        Case 94576, 94577
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMPRODUCAOBAIXADA", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 152143)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Rotina_CMP_EstoqueMes_CriticaLock(ByVal lComando As Long, ByVal iAno As Integer, ByVal iMes As Integer, ByVal colEstoqueMes As Collection) As Long
'Verifica se EstoqueMes tem Custo Medio Produção não apurado e
'está fechado. Faz lock exclusive. Chamada EM TRANSAÇÃO.

Dim lErro As Long
Dim iCustoProdApurado As Integer
Dim dGastosDiretos As Double
Dim dGastosIndiretos As Double
Dim iFilialEmpresa As Integer
Dim objEstoqueMes As ClassEstoqueMes, lComando1 As Long
Dim dCustoFator1 As Double, dCustoFator2 As Double, dCustoFator3 As Double, dCustoFator4 As Double, dCustoFator5 As Double, dCustoFator6 As Double

On Error GoTo Erro_Rotina_CMP_EstoqueMes_CriticaLock

    'Abertura de comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 201212
    
    'Seleciona EstoqueMes
    lErro = Comando_ExecutarLockado(lComando1, "SELECT FilialEmpresa, CustoProdApurado, GastosDiretos, GastosIndiretos, CustoFator1, CustoFator2, CustoFator3, CustoFator4, CustoFator5, CustoFator6 FROM EstoqueMes WHERE Ano=? AND Mes=?", iFilialEmpresa, iCustoProdApurado, dGastosDiretos, dGastosIndiretos, dCustoFator1, dCustoFator2, dCustoFator3, dCustoFator4, dCustoFator5, dCustoFator6, iAno, iMes)
    If lErro <> AD_SQL_SUCESSO Then gError 92601

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92602

    If lErro = AD_SQL_SEM_DADOS Then gError 92603 'não encontrou EstoqueMes

    Do While lErro = AD_SQL_SUCESSO

        'Faz lock em EstoqueMes
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 92604

        Set objEstoqueMes = New ClassEstoqueMes
        
        objEstoqueMes.iFilialEmpresa = iFilialEmpresa
        objEstoqueMes.iAno = iAno
        objEstoqueMes.iMes = iMes
        objEstoqueMes.dGastosDiretos = dGastosDiretos
        objEstoqueMes.dGastosIndiretos = dGastosIndiretos
        objEstoqueMes.iCustoProdApurado = iCustoProdApurado
        objEstoqueMes.dCustoFator1 = dCustoFator1
        objEstoqueMes.dCustoFator2 = dCustoFator2
        objEstoqueMes.dCustoFator3 = dCustoFator3
        objEstoqueMes.dCustoFator4 = dCustoFator4
        objEstoqueMes.dCustoFator5 = dCustoFator5
        objEstoqueMes.dCustoFator6 = dCustoFator6

        colEstoqueMes.Add objEstoqueMes

        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92605

    Loop

    Call Comando_Fechar(lComando1)
    
    Rotina_CMP_EstoqueMes_CriticaLock = SUCESSO

    Exit Function

Erro_Rotina_CMP_EstoqueMes_CriticaLock:

    Rotina_CMP_EstoqueMes_CriticaLock = gErr

    Select Case gErr

        Case 201212
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 92601, 92602, 92605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEMES2", gErr, iAno, iMes)

        Case 92603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE3", gErr, iAno, iMes)

        Case 92604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_ESTOQUEMES", gErr, iFilialEmpresa, iAno, iMes)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152144)

    End Select

    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Private Function Produto_ObterRelQtdSldIniTerc(ByVal sProduto As String, ByVal iAlmoxarifado As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, dQuantInicial As Double, dValorInicial As Double, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
''IMPORTANTE: especifica para classe ClassRelMesEstAlm de apoio a execucao de relatorios
'Obtém a quantidade inicial de terceiros e o valor inicial de terceiros do Produto/Almoxarifado correspondentes até a data anterior a data inicial
''ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long
Dim dQuantInicial1 As Double
Dim dValorInicial1 As Double
Dim dtDataInicial1 As Date
Dim iIndice As Integer
Dim iAno As Integer
Dim iMes As Integer
Dim alComando(0 To 1) As Long

On Error GoTo Erro_Produto_ObterRelQtdSldIniTerc

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
        If alComando(iIndice) = 0 Then gError 93711
    Next
    
    'obter a quantidade e valor inicial na implantacao do produto no almoxarifado
    lErro = Comando_Executar(alComando(1), "SELECT QuantInicialConsig3 + QuantInicialDemo3 + QuantInicialConserto3 + QuantInicialOutras3 + QuantInicialBenef3, ValorInicialConsig3 + ValorInicialDemo3 + ValorInicialConserto3 + ValorInicialOutras3 + ValorInicialBenef3, DataInicial FROM EstoqueProduto WHERE Produto = ? AND Almoxarifado = ?", dQuantInicial, dValorInicial, dtDataInicial1, sProduto, iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 93712
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93713
    
    If lErro = AD_SQL_SEM_DADOS Then gError 93714

    'se a data de inicializacao do produto no estoque for menor que a data final do periodo solicitado
    '==> os valores não estão disponíveis
    If dtDataFinal <> DATA_NULA And dtDataInicial1 > dtDataFinal Then
        dQuantInicial = 0
        dValorInicial = 0
    Else
        
        If dtDataInicial <> DATA_NULA Then
            
            'obter os totais movimentados até a data anterior a inicial
            lErro = Comando_Executar(alComando(0), "SELECT SUM((QuantEntConsig3 + QuantEntDemo3 + QuantEntConserto3 + QuantEntOutros3 + QuantEntBenef3) - (QuantSaiConsig3 + QuantSaiDemo3 + QuantSaiConserto3 + QuantSaiOutros3 + QuantSaiBenef3)), SUM((ValorEntConsig3 + ValorEntDemo3 + ValorEntConserto3 + ValorEntOutros3 + ValorEntBenef3) - (ValorSaiConsig3 + ValorSaiDemo3 + ValorSaiConserto3 + ValorSaiOutros3 + ValorSaiBenef3)) FROM SldDiaEstAlm WHERE Produto = ? AND Almoxarifado = ? AND Data < ?", dQuantInicial1, dValorInicial1, sProduto, iAlmoxarifado, dtDataInicial)
            If lErro <> AD_SQL_SUCESSO Then gError 93715
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93716
        
            dQuantInicial = dQuantInicial + dQuantInicial1
            dValorInicial = dValorInicial + dValorInicial1
        
            iAno = Year(dtDataInicial)
            iMes = Month(dtDataInicial)
            
            'Se é um produto produzido e a ultima apuracao de custo de producao for menor do que ano/mes inicial do periodo solicitado ==> o saldo em valor não está disponível
            If iApropriacaoCusto = APROPR_CUSTO_REAL And (iAnoEstApurado < iAno Or (iAno = iAnoEstApurado And iMesEstApurado < iMes)) Then dValorInicial = 0
            
        End If
        
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
            
            
    Produto_ObterRelQtdSldIniTerc = SUCESSO

    Exit Function

Erro_Produto_ObterRelQtdSldIniTerc:

    Produto_ObterRelQtdSldIniTerc = gErr

    Select Case gErr

        Case 93711
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 93715, 93716
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM1", gErr, iAlmoxarifado, sProduto)

        Case 93712, 93713
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO", gErr, sProduto, iAlmoxarifado)

        Case 93714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", gErr, sProduto, iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152145)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function Produto_ObterRelQtdSldIniFil(ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, dQuantInicial As Double, dValorInicial As Double, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, Optional ByVal bExibeErroSemEST As Boolean = True) As Long
'Obtém a quantidade inicial e o valor inicial do Produto/Filial correspondentes até a data anterior a data inicial
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long
Dim dQuantInicial1 As Double
Dim dValorInicial1 As Double
Dim dQuantInicial1Total As Double
Dim dValorInicial1Total As Double
Dim dtDataInicial1 As Date
Dim iIndice As Integer
Dim iAno As Integer
Dim iMes As Integer
Dim dQuantInicialTerc As Double
Dim dValorInicialTerc As Double
Dim alComando(0 To 1) As Long


On Error GoTo Erro_Produto_ObterRelQtdSldIniFil

    For iIndice = 0 To 1
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
        If alComando(iIndice) = 0 Then gError 93717
    Next
    
    'obter a quantidade e valor inicial na implantacao do produto na filial
    lErro = Comando_Executar(alComando(1), "SELECT QuantidadeInicial, SaldoInicial, DataInicial FROM EstoqueProduto, Almoxarifado WHERE Codigo=Almoxarifado AND Produto = ? AND FilialEmpresa = ?", dQuantInicial1, dValorInicial1, dtDataInicial1, sProduto, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 93718
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93719
    
    If lErro = AD_SQL_SEM_DADOS Then gError 93720
        
    dQuantInicial1Total = 0
    dValorInicial1Total = 0
         
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'se a data de inicializacao do produto no estoque for menor que a data final do periodo solicitado
        '==> os valores não estão disponíveis
        If dtDataFinal <> DATA_NULA And dtDataInicial1 > dtDataFinal Then
            dQuantInicial1 = 0
            dValorInicial1 = 0
        End If
        
        dQuantInicial1Total = dQuantInicial1Total + dQuantInicial1
        dValorInicial1Total = dValorInicial1Total + dValorInicial1
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93725
        
    Loop

    dQuantInicial = dQuantInicial1Total
    dValorInicial = dValorInicial1Total

    If dtDataInicial <> DATA_NULA Then
        
        'obter os totais movimentados até a data anterior a inicial
        lErro = Comando_Executar(alComando(0), "SELECT SUM(QuantEntrada - QuantEntConsig3 - QuantEntDemo3 - QuantEntConserto3 - QuantEntOutros3 - QuantEntBenef3) - SUM(QuantSaida - QuantSaiConsig3 - QuantSaiDemo3 - QuantSaiConserto3 - QuantSaiOutros3 - QuantSaiBenef3) , SUM(ValorEntrada - ValorEntConsig3 - ValorEntDemo3 - ValorEntConserto3 - ValorEntOutros3 - ValorEntBenef3) - SUM(ValorSaida - ValorSaiConsig3 - ValorSaiDemo3 - ValorSaiConserto3 - ValorSaiOutros3 - ValorSaiBenef3) FROM SldDiaEst WHERE Produto = ? AND FilialEmpresa = ? AND Data < ?", dQuantInicial1, dValorInicial1, sProduto, iFilialEmpresa, dtDataInicial)
        If lErro <> AD_SQL_SUCESSO Then gError 93721

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93722
    
        dQuantInicial = dQuantInicial1Total + dQuantInicial1
        dValorInicial = dValorInicial1Total + dValorInicial1
    
    End If
        
    iAno = Year(dtDataInicial)
    iMes = Month(dtDataInicial)
    
    'Se é um produto produzido e a ultima apuracao de custo de producao for menor do que ano/mes inicial do periodo solicitado ==> o saldo em valor não está disponível
    If iApropriacaoCusto = APROPR_CUSTO_REAL And (iAnoEstApurado < iAno Or (iAno = iAnoEstApurado And iMesEstApurado < iMes)) Then dValorInicial = -1
            
    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next
               
    Produto_ObterRelQtdSldIniFil = SUCESSO

    Exit Function

Erro_Produto_ObterRelQtdSldIniFil:

    Produto_ObterRelQtdSldIniFil = gErr

    Select Case gErr
        
        Case 93723
        
        Case 93717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 93721, 93722
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST", gErr, iFilialEmpresa, sProduto)

        Case 93718, 93719, 93725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO2", gErr, sProduto)

        Case 93720
            If bExibeErroSemEST Then Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO1", gErr, sProduto, iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152146)

    End Select

    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function Produto_ObterRelQtdSldIniTercFil(ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, dQuantInicial As Double, dValorInicial As Double, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
''IMPORTANTE: especifica para classe ClassRelMesEstFil de apoio a execucao de relatorios
'Obtém a quantidade inicial de terceiros e o valor inicial de terceiros do Produto/Filial correspondentes até a data anterior a data inicial

Dim lErro As Long
Dim dQuantInicial1 As Double
Dim dValorInicial1 As Double
Dim dtDataInicial1 As Date
Dim dQuantInicial1Total As Double
Dim dValorInicial1Total As Double
Dim iIndice As Integer
Dim iAno As Integer
Dim iMes As Integer
Dim alComando(0 To 1) As Long

On Error GoTo Erro_Produto_ObterRelQtdSldIniTercFil

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
        If alComando(iIndice) = 0 Then gError 93733
    Next
    
    'obter a quantidade e valor inicial na implantacao do produto na Filial
    lErro = Comando_Executar(alComando(1), "SELECT QuantInicialConsig3 + QuantInicialDemo3 + QuantInicialConserto3 + QuantInicialOutras3 + QuantInicialBenef3, ValorInicialConsig3 + ValorInicialDemo3 + ValorInicialConserto3 + ValorInicialOutras3 + ValorInicialBenef3, DataInicial FROM EstoqueProduto, Almoxarifado WHERE Codigo=Almoxarifado AND Produto = ? AND FilialEmpresa = ?", dQuantInicial, dValorInicial, dtDataInicial1, sProduto, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 93734
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93735
    
    If lErro = AD_SQL_SEM_DADOS Then gError 93736
    
    dQuantInicial1Total = 0
    dValorInicial1Total = 0
         
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'se a data de inicializacao do produto no estoque for menor que a data final do periodo solicitado
        '==> os valores não estão disponíveis
        If dtDataFinal <> DATA_NULA And dtDataInicial1 > dtDataFinal Then
            dQuantInicial = 0
            dValorInicial = 0
        End If
        
        dQuantInicial1Total = dQuantInicial1Total + dQuantInicial
        dValorInicial1Total = dValorInicial1Total + dValorInicial
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93739
              
    Loop
          
    If dtDataInicial <> DATA_NULA Then
        'obter os totais movimentados até a data anterior a inicial
        lErro = Comando_Executar(alComando(0), "SELECT SUM((QuantEntConsig3 + QuantEntDemo3 + QuantEntConserto3 + QuantEntOutros3 + QuantEntBenef3) - (QuantSaiConsig3 + QuantSaiDemo3 + QuantSaiConserto3 + QuantSaiOutros3 + QuantSaiBenef3)), SUM((ValorEntConsig3 + ValorEntDemo3 + ValorEntConserto3 + ValorEntOutros3 + ValorEntBenef3) - (ValorSaiConsig3 + ValorSaiDemo3 + ValorSaiConserto3 + ValorSaiOutros3 + ValorSaiBenef3)) FROM SldDiaEst WHERE Produto = ? AND FilialEmpresa = ? AND Data < ?", dQuantInicial1, dValorInicial1, sProduto, iFilialEmpresa, dtDataInicial)
        If lErro <> AD_SQL_SUCESSO Then gError 93737

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93738
    
        dQuantInicial = dQuantInicial1Total + dQuantInicial1
        dValorInicial = dValorInicial1Total + dValorInicial1
    
    End If
    
    iAno = Year(dtDataInicial)
    iMes = Month(dtDataInicial)
    
    'Se é um produto produzido e a ultima apuracao de custo de producao for menor do que ano/mes inicial do periodo solicitado ==> o saldo em valor não está disponível
    If iApropriacaoCusto = APROPR_CUSTO_REAL And (iAnoEstApurado < iAno Or (iAno = iAnoEstApurado And iMesEstApurado < iMes)) Then dValorInicial = 0
            
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
                
    Produto_ObterRelQtdSldIniTercFil = SUCESSO

    Exit Function

Erro_Produto_ObterRelQtdSldIniTercFil:

    Produto_ObterRelQtdSldIniTercFil = gErr

    Select Case gErr

        Case 93733
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 93737, 93738
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST", gErr, iFilialEmpresa, sProduto)

        Case 93734, 93735, 93739
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO2", gErr, sProduto)

        Case 93736
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO1", gErr, sProduto, iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152147)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function MovEstoque_Le_Requisicoes_OP(ByVal lNumIntDocOrigem As Long, ByVal colItemMovEstReqOp As Collection) As Long
'le as requisiçoes associadas a OP representada por NumIntDocOrigem e coloca-as em colItemMovEstReqOp
'???? ATENCAO, PARA FACILITAR CALCULOS SEGUINTES A QTDE VAI ESTAR NA U.M. DE ESTOQUE

Dim alComando(1) As Long
Dim lErro As Long
Dim tItemMovEst As typeItemMovEstoque
Dim dFator As Double
Dim objItemMovEst As ClassItemMovEstoque
Dim iIndice As Integer

On Error GoTo Erro_MovEstoque_Le_Requisicoes_OP

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92486
    Next

    tItemMovEst.sProduto = String(STRING_PRODUTO, 0)
    tItemMovEst.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemMovEst.sSiglaUMEst = String(STRING_UM_SIGLA, 0)

    With tItemMovEst

        'le os movimentos de estoque de requisição de produto associados ao itemOP passado como parametro referentes aos produtos com natureza materia prima e que não estejam estornados e devolve a soma dos seus custos
        lErro = Comando_Executar(alComando(0), "SELECT MovimentoEstoque.TipoMov, MovimentoEstoque.Produto, MovimentoEstoque.Data, MovimentoEstoque.Hora, MovimentoEstoque.NumIntDoc, MovimentoEstoque.Quantidade, MovimentoEstoque.SiglaUM, MovimentoEstoque.Custo, Produtos.Natureza, Produtos.ClasseUM, Produtos.SiglaUMEstoque FROM ItensOrdemProducao, MovimentoEstoque, Produtos WHERE ItensOrdemProducao.NumIntDoc = ? AND (MovimentoEstoque.TipoMov = ? Or MovimentoEstoque.TipoMov=?) AND MovimentoEstoque.NumIntDocOrigem = ItensOrdemProducao.NumIntDoc AND MovimentoEstoque.Produto = Produtos.Codigo AND MovimentoEstoque.NumIntDocEst = 0 ORDER BY MovimentoEstoque.Produto, MovimentoEstoque.Data, MovimentoEstoque.Hora", _
        .iTipoMov, .sProduto, .dtData, .dHora, .lNumIntDoc, .dQuantidade, .sSiglaUM, .dCusto, .iNaturezaProduto, .iClasseUM, .sSiglaUMEst, lNumIntDocOrigem, MOV_EST_REQ_PRODUCAO, MOV_EST_REQ_PRODUCAO_BENEF3)
        If lErro <> AD_SQL_SUCESSO Then gError 92487

    End With

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92488

    'se não há movimento para o produto em ItensOrdemProducao => faz o mesmo select em ItensOrdemProducaoBaixados
    If lErro = AD_SQL_SEM_DADOS Then

        tItemMovEst.sProduto = String(STRING_PRODUTO, 0)
        tItemMovEst.sSiglaUM = String(STRING_UM_SIGLA, 0)
        tItemMovEst.sSiglaUMEst = String(STRING_UM_SIGLA, 0)

        With tItemMovEst
            'le os movimentos de estoque de requisição de produto associados ao itemOP passado como parametro referentes aos produtos com natureza materia prima e que não estejam estornados e devolve a soma dos seus custos
            lErro = Comando_Executar(alComando(0), "SELECT MovimentoEstoque.TipoMov, MovimentoEstoque.Produto, MovimentoEstoque.Data, MovimentoEstoque.Hora, MovimentoEstoque.NumIntDoc, MovimentoEstoque.Quantidade, MovimentoEstoque.SiglaUM, MovimentoEstoque.Custo, Produtos.Natureza, Produtos.ClasseUM, Produtos.SiglaUMEstoque FROM ItensOrdemProducaoBaixadas, MovimentoEstoque, Produtos WHERE ItensOrdemProducaoBaixadas.NumIntDoc = ? AND (MovimentoEstoque.TipoMov = ? Or MovimentoEstoque.TipoMov=?) AND MovimentoEstoque.NumIntDocOrigem = ItensOrdemProducaoBaixadas.NumIntDoc AND MovimentoEstoque.Produto = Produtos.Codigo AND MovimentoEstoque.NumIntDocEst = 0 ORDER BY MovimentoEstoque.Produto, MovimentoEstoque.Data, MovimentoEstoque.Hora", _
            .iTipoMov, .sProduto, .dtData, .dHora, .lNumIntDoc, .dQuantidade, .sSiglaUM, .dCusto, .iNaturezaProduto, .iClasseUM, .sSiglaUMEst, lNumIntDocOrigem, MOV_EST_REQ_PRODUCAO, MOV_EST_REQ_PRODUCAO_BENEF3)
            If lErro <> AD_SQL_SUCESSO Then gError 94558
        End With

        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94559

    End If
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objItemMovEst = New ClassItemMovEstoque
    
        lErro = CF("UM_Conversao_Trans", tItemMovEst.iClasseUM, tItemMovEst.sSiglaUM, tItemMovEst.sSiglaUMEst, dFator)
        If lErro <> SUCESSO Then gError 92661
    
        objItemMovEst.iTipoMov = tItemMovEst.iTipoMov
        objItemMovEst.dQuantidadeEst = tItemMovEst.dQuantidade * dFator
        objItemMovEst.sProduto = tItemMovEst.sProduto
        objItemMovEst.dtData = tItemMovEst.dtData
        objItemMovEst.dtHora = tItemMovEst.dHora
        objItemMovEst.lNumIntDoc = tItemMovEst.lNumIntDoc
        objItemMovEst.dQuantidade = objItemMovEst.dQuantidadeEst '???? ATENCAO, PARA FACILITAR CALCULOS SEGUINTES A QTDE VAI ESTAR NA U.M. DE ESTOQUE
        objItemMovEst.sSiglaUM = tItemMovEst.sSiglaUM
        objItemMovEst.dCusto = tItemMovEst.dCusto
        objItemMovEst.iNaturezaProduto = tItemMovEst.iNaturezaProduto
        objItemMovEst.iClasseUM = tItemMovEst.iClasseUM
        objItemMovEst.sSiglaUMEst = tItemMovEst.sSiglaUMEst
   
        colItemMovEstReqOp.Add objItemMovEst
   
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92518
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    MovEstoque_Le_Requisicoes_OP = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Le_Requisicoes_OP:

    MovEstoque_Le_Requisicoes_OP = gErr

    Select Case gErr

        Case 92486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92487, 92488, 92518, 94558, 94559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152148)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Calcula_Total_Producoes_OP(ByVal objItemMovEst As ClassItemMovEstoque, ByVal colItemMovEst As Collection, dQuantProducoesEst As Double) As Long
'calcula o total das produçoes para a O.P. em questao

Dim dFator As Double
Dim lErro As Long
Dim objProduto1 As New ClassProduto
Dim objItemMovEst1 As ClassItemMovEstoque

On Error GoTo Erro_Calcula_Total_Producoes_OP

    objProduto1.sCodigo = objItemMovEst.sProduto
                
    'le o produto associado a produção em questão
    lErro = CF("Produto_Le", objProduto1)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 92655
    
    If lErro <> SUCESSO Then gError 92656

    lErro = CF("UM_Conversao_Trans", objProduto1.iClasseUM, objItemMovEst.sSiglaUM, objProduto1.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 92657

    objItemMovEst.dQuantidadeEst = objItemMovEst.dQuantidade * dFator

    'calcula o total das produções para a O.P. em questão que não tenham apropriação. Este valor será posteriormente utilizado para proporcionalizar o custo dos insumos/materia prima em função da quantidade produzida
    For Each objItemMovEst1 In colItemMovEst

        'se a produção não tinha apropriação associada ==> acumula sua quantidade para depois poder calcular os custos em função da proporção da quantidade produzida em questão em relação ao total das produções sem apropriação
        If objItemMovEst1.colApropriacaoInsumo.Count = 0 Then

            lErro = CF("UM_Conversao_Trans", objProduto1.iClasseUM, objItemMovEst1.sSiglaUM, objProduto1.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 92658
            
            objItemMovEst1.dQuantidadeEst = objItemMovEst1.dQuantidade * dFator
            
            'total das quantidades produzidas sem apropriação vinculada
            dQuantProducoesEst = dQuantProducoesEst + objItemMovEst1.dQuantidadeEst
            
        End If
            
    Next
    
    Calcula_Total_Producoes_OP = SUCESSO
    
    Exit Function
    
Erro_Calcula_Total_Producoes_OP:

    Calcula_Total_Producoes_OP = gErr
    
    Select Case gErr

        Case 92655, 92657, 92658
        
        Case 92656
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto1.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152149)

    End Select
    
    Exit Function

End Function

Private Sub Calcula_Quant_Nao_Apropriada_Req2(ByVal objItemMovEst As ClassItemMovEstoque, ByVal colItemMovEstReqOp As Collection, ByVal dQuantProducoesEst As Double, ByVal colItemMovEst As Collection, dQuantNaoApropriada As Double, ByVal dQuantNaoApropriadaAteData As Double)
'ratear a dQuantNaoApropriadaAteData levando em conta o total de producoes ainda sem apropriacao na data do movto

Dim dQtdeProdSemApropNaData As Double, iCont As Integer
Dim objItemMovEst1 As ClassItemMovEstoque

        iCont = 0
        
        For Each objItemMovEst1 In colItemMovEst
        
            If objItemMovEst1.dtData = objItemMovEst.dtData And objItemMovEst1.colApropriacaoInsumo.Count = 0 Then
                
                dQtdeProdSemApropNaData = dQtdeProdSemApropNaData + objItemMovEst1.dQuantidadeEst
                iCont = iCont + 1
                    
            End If
            
        Next
        
        If iCont <> 1 Then
        
            dQuantNaoApropriada = dQuantNaoApropriada * objItemMovEst.dQuantidadeEst / dQtdeProdSemApropNaData
        
        End If
        
End Sub

Private Function Calcula_Quant_Nao_Apropriada_Req(ByVal objItemMovEst As ClassItemMovEstoque, ByVal colItemMovEstReqOp As Collection, ByVal dQuantProducoesEst As Double, ByVal colItemMovEst As Collection) As Long
'calcula a quantidade não apropriada de cada requisição e o total não apropriado para cada produto

Dim sProduto As String
Dim iIndice As Integer
Dim iIndiceInicio As Integer, iIndiceFim As Integer
Dim dQuantNaoApropriada As Double, dQuantNaoApropriadaAteData As Double
Dim objItemMovEst1 As ClassItemMovEstoque
Dim objItemMovEst2 As ClassItemMovEstoque
Dim objApropriacaoInsumosProd As ClassApropriacaoInsumosProd
Dim lErro As Long, dQtdeProdSemApropNaData As Double

On Error GoTo Erro_Calcula_Quant_Nao_Apropriada_Req

    If colItemMovEstReqOp.Count > 0 Then
        sProduto = colItemMovEstReqOp.Item(1).sProduto
    End If
    
    iIndiceInicio = 1
    iIndiceFim = 0
    dQuantNaoApropriada = 0
    dQuantNaoApropriadaAteData = 0
    
    'Para cada requisição relativa a O.P. da produção em questão
    'descobre o total de requisições do produto que não estavam apropriadas e guarda junta a cada requisição (guarda em dQuantidadeOpEst)
    'guarda junto a cada requisição (dQuantidadeOpEst) somente a quantidade ainda não apropriada
    For Each objItemMovEst2 In colItemMovEstReqOp
        
        iIndiceFim = iIndiceFim + 1
        
        'guarda junta a cada requisição a quantidade total não apropriada
        If sProduto <> objItemMovEst2.sProduto Then
    
            'esta é a quantidade que vai ser apropriada do produto para esta produção
            dQuantNaoApropriada = dQuantNaoApropriada * (objItemMovEst.dQuantidadeEst / dQuantProducoesEst)
            
            If dQuantNaoApropriada > dQuantNaoApropriadaAteData Then
            
                'ratear a dQuantNaoApropriadaAteData
                Call Calcula_Quant_Nao_Apropriada_Req2(objItemMovEst, colItemMovEstReqOp, dQuantProducoesEst, colItemMovEst, dQuantNaoApropriada, dQuantNaoApropriadaAteData)
            
            End If
            
            For iIndice = iIndiceInicio To iIndiceFim - 1
                colItemMovEstReqOp.Item(iIndice).dQuantidadeOPEst = dQuantNaoApropriada
            Next
            
            dQuantNaoApropriada = 0
            dQuantNaoApropriadaAteData = 0
            iIndiceInicio = iIndiceFim
            sProduto = objItemMovEst2.sProduto
        
        End If
        
        'pesquisar para cada produção associada a O.P. em questão se alguma tem apropriação que diz respeito a requisição em questão
        'se tiver retira a quantidade da requisição pois so queremos ficar com a quantidade não apropriada.
        'retirar do total das requisições as apropriações associadas as outras produções, já que vai somente considerar o total das requisições não alocadas (proporcional a producao)
        For Each objItemMovEst1 In colItemMovEst
    
            'pega todas as produções para o produto/O.P. com exceção da em questão
            If objItemMovEst1.lNumIntDoc <> objItemMovEst.lNumIntDoc Then
    
                'pega as apropriações que se referirem a resquisição em questão
                For Each objApropriacaoInsumosProd In objItemMovEst1.colApropriacaoInsumo
                
                    If objApropriacaoInsumosProd.lNumIntReqProd = objItemMovEst2.lNumIntDoc Then
                    
                        'guarda em dQuantdadeEst sómente a quantidade da requisição que não foi apropriada
                        objItemMovEst2.dQuantidadeEst = objItemMovEst2.dQuantidadeEst - objApropriacaoInsumosProd.dQuantidade
                        
                    End If
                Next
                
            End If
            
        Next
    
        dQuantNaoApropriada = dQuantNaoApropriada + objItemMovEst2.dQuantidadeEst
        If objItemMovEst2.dtData <= objItemMovEst.dtData Then dQuantNaoApropriadaAteData = dQuantNaoApropriadaAteData + objItemMovEst2.dQuantidadeEst
    
    Next
    
    'esta é a quantidade que vai ser apropriada do produto para esta produção
    dQuantNaoApropriada = dQuantNaoApropriada * (objItemMovEst.dQuantidadeEst / dQuantProducoesEst)
    
    If dQuantNaoApropriada > dQuantNaoApropriadaAteData Then
    
        'ratear a dQuantNaoApropriadaAteData
        Call Calcula_Quant_Nao_Apropriada_Req2(objItemMovEst, colItemMovEstReqOp, dQuantProducoesEst, colItemMovEst, dQuantNaoApropriada, dQuantNaoApropriadaAteData)
    
    End If
    
    If colItemMovEstReqOp.Count > 0 Then
    
        For iIndice = iIndiceInicio To iIndiceFim
            colItemMovEstReqOp.Item(iIndice).dQuantidadeOPEst = dQuantNaoApropriada
        Next
    
    End If

    Calcula_Quant_Nao_Apropriada_Req = SUCESSO
    
    Exit Function
    
Erro_Calcula_Quant_Nao_Apropriada_Req:

    Calcula_Quant_Nao_Apropriada_Req = gErr
    
    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152150)

    End Select
    
    Exit Function

End Function

Function MovEstoque_Le_CustoMatPrim(ByVal objItemMovEst As ClassItemMovEstoque, dCustoMatPrim As Double, dCustoInsumos As Double, ByVal colItemMovEst As Collection, Optional ByVal iMostrarMensagemErro As Integer = 0) As Long
'le o custo das materias primas necessárias para a produção em questão
'o parametro iMostrarMensagemErro indica que: se se trata da rotina de calculo do custo real ==> deve dar erro
'se se trata da rotina que calcula o custo total da materia prima ==> não deve dar erro

Dim lErro As Long
Dim tItemMovEst As typeItemMovEstoque
Dim dQuantProducoesEst As Double
Dim colItemMovEstReqOp As New Collection
Dim objItemMovEst2 As ClassItemMovEstoque
Dim dQuantApropriar As Double
Dim dQuantApropriada As Double
Dim sProduto As String
Dim objApropriacaoInsumos As ClassApropriacaoInsumosProd
Dim dQuantDeveriaApropriar As Double
Dim objProduto1 As New ClassProduto, dCustoMatPrimAprop As Double

On Error GoTo Erro_MovEstoque_Le_CustoMatPrim
    
    dCustoMatPrim = 0
    dCustoInsumos = 0
    
    'devolve o total de producoes para a O.P. em questão
    lErro = Calcula_Total_Producoes_OP(objItemMovEst, colItemMovEst, dQuantProducoesEst)
    If lErro <> SUCESSO Then gError 92671
    
    'le as requisiçoes associadas a OP representada por NumIntDocOrigem e coloca-as em colItemMovEstReqOp
    lErro = MovEstoque_Le_Requisicoes_OP(objItemMovEst.lNumIntDocOrigem, colItemMovEstReqOp)
    If lErro <> SUCESSO Then gError 92672
    
    'calcula a quantidade não apropriada de cada requisição e o total não apropriado para cada produto
    lErro = Calcula_Quant_Nao_Apropriada_Req(objItemMovEst, colItemMovEstReqOp, dQuantProducoesEst, colItemMovEst)
    If lErro <> SUCESSO Then gError 92673
    
    'Para cada requisição associada a O.P. em questão
    For Each objItemMovEst2 In colItemMovEstReqOp
    
        If sProduto <> objItemMovEst2.sProduto Then
        
            'se se trata da rotina de calculo do custo real ==> deve dar erro
            'se se trata da rotina que calcula o custo total da materia prima ==> não deve dar erro
            'se a quantidade apropriada for diferente da quantidade total a apropriar ==> erro
            If iMostrarMensagemErro = MOSTRA_MENSAGEM_ERRO_QUANTAPROPRIADA And sProduto <> "" And Abs(dQuantApropriada - dQuantDeveriaApropriar) > QTDE_ESTOQUE_DELTA Then gError 92489
        
            sProduto = objItemMovEst2.sProduto
            dQuantApropriada = 0
            dQuantDeveriaApropriar = objItemMovEst2.dQuantidadeOPEst
            
        End If
    
        'se a requisição já foi custeada (ou for de produto produzido) e for anterior à produção ==> tenta alocar quantidade
        If (objItemMovEst2.iNaturezaProduto <> NATUREZA_PROD_MATERIA_PRIMA Or objItemMovEst2.dCusto <> 0) And objItemMovEst2.dtData <= objItemMovEst.dtData Then
    
            'se a quantidade disponivel para a apropriacao for maior do que a quantidade necessária ==> apropria a quantidade necessária
            If objItemMovEst2.dQuantidadeEst > objItemMovEst2.dQuantidadeOPEst - dQuantApropriada Then
                
                dQuantApropriar = objItemMovEst2.dQuantidadeOPEst - dQuantApropriada
                
            Else
            
                dQuantApropriar = objItemMovEst2.dQuantidadeEst
                
            End If
            
            'se há alguma quantidade a apropriar
            If dQuantApropriar > 0 Then
            
                Set objApropriacaoInsumos = New ClassApropriacaoInsumosProd
            
                objApropriacaoInsumos.lNumIntDocOrigem = objItemMovEst.lNumIntDoc
                objApropriacaoInsumos.lNumIntReqProd = objItemMovEst2.lNumIntDoc
                objApropriacaoInsumos.sProduto = objItemMovEst2.sProduto
                objApropriacaoInsumos.dQuantidade = dQuantApropriar
                
                lErro = CF("ApropriacaoInsumosProd_Inserir1", objApropriacaoInsumos)
                If lErro <> SUCESSO Then gError 92680
            
                dQuantApropriada = dQuantApropriada + dQuantApropriar
                    
                If objItemMovEst2.iTipoMov = MOV_EST_REQ_PRODUCAO_BENEF3 Then
                
                    objProduto1.sCodigo = objItemMovEst2.sProduto
                                
                    'le o produto associado a requisição para descobrir se a natureza do produto é materia prima.
                    lErro = CF("Produto_Le", objProduto1)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 106529
                    
                    If lErro <> SUCESSO Then gError 106530

                End If
                
                'se nao for insumo fornecido por 3os p/beneficiamento
                If objItemMovEst2.iTipoMov <> MOV_EST_REQ_PRODUCAO_BENEF3 Or (objProduto1.iCompras = PRODUTO_PRODUZIVEL) Then
                
                    dCustoMatPrimAprop = (objItemMovEst2.dCusto / objItemMovEst2.dQuantidade) * dQuantApropriar
                    
                    'acumula no custo dos insumos do custo da quantidade que foi apropriada
                    dCustoInsumos = dCustoInsumos + dCustoMatPrimAprop
                    
                    'acumula no custo das materias primas o custo da quantidade que foi apropriada, se o produto se tratar de uma materia prima
                    If objItemMovEst2.iNaturezaProduto = NATUREZA_PROD_MATERIA_PRIMA Then
                        
                        dCustoMatPrim = dCustoMatPrim + dCustoMatPrimAprop
                        
                    End If
    
                End If
    
            End If
    
        End If
    
    Next

    MovEstoque_Le_CustoMatPrim = SUCESSO
    
    Exit Function

Erro_MovEstoque_Le_CustoMatPrim:

    MovEstoque_Le_CustoMatPrim = gErr

    Select Case gErr

        Case 92671, 92672, 92673, 92674, 92680, 106510
        
        Case 92489
            lErro = Rotina_Erro(vbOKOnly, "QUANTIDADE_APROPRIADA_MENOR_REQUERIDA", gErr, objItemMovEst.lCodigo, objItemMovEst.iFilialEmpresa, objItemMovEst.sOPCodigo)
        
        Case 106529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objItemMovEst2.sProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152151)

    End Select

    Exit Function

End Function

Function Produto_ObterRelQtdSldIniFilTercEsc(ByVal sProduto As String, ByVal iFilialEmpresa As Integer, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, dQuantInicial As Double, dValorInicial As Double, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal iNumEscaninho As Integer) As Long
'IMPORTANTE: especifica para classe ClassRelMesEstFil3 de apoio a execucao de relatorios
'Obtém a quantidade inicial e o valor inicial de ou em 3os do Produto/Filial p/o escaninho correspondentes até a data anterior a data inicial

Dim lErro As Long
Dim dQuantInicial1 As Double
Dim dValorInicial1 As Double
Dim dtDataInicial1 As Date
Dim dQuantInicial1Total As Double
Dim dValorInicial1Total As Double
Dim iIndice As Integer
Dim iAno As Integer
Dim iMes As Integer, sCampos As String, sQuantEnt As String, sQuantSaida As String, sValorEntrada As String, sValorSaida As String
Dim alComando(0 To 1) As Long

On Error GoTo Erro_Produto_ObterRelQtdSldIniFilTercEsc

    For iIndice = 0 To 1
        If alComando(iIndice) = 0 Then
            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
            If alComando(iIndice) = 0 Then gError 93746
        End If
    Next
    
    Select Case iNumEscaninho
        
        Case ESCANINHO_CONSERTO_NOSSO
            sCampos = "QuantInicialConserto, ValorInicialConserto"
            
        Case ESCANINHO_CONSIG_NOSSO
            sCampos = "QuantInicialConsig, ValorInicialConsig"
            
        Case ESCANINHO_DEMO_NOSSO
            sCampos = "QuantInicialDemo, ValorInicialDemo"
            
        Case ESCANINHO_OUTROS_NOSSO
            sCampos = "QuantInicialOutras, ValorInicialOutras"
            
        Case ESCANINHO_BENEF_NOSSO
            sCampos = "QuantInicialBenef, ValorInicialBenef"
            
        Case ESCANINHO_CONSERTO_3
            sCampos = "QuantInicialConserto3, ValorInicialConserto3"
            
        Case ESCANINHO_CONSIG_3
            sCampos = "QuantInicialConsig3, ValorInicialConsig3"
            
        Case ESCANINHO_DEMO_3
            sCampos = "QuantInicialDemo3, ValorInicialDemo3"
            
        Case ESCANINHO_OUTROS_3
            sCampos = "QuantInicialOutras3, ValorInicialOutras3"
            
        Case ESCANINHO_BENEF_3
            sCampos = "QuantInicialBenef3, ValorInicialBenef3"
           
    End Select
    
    'obter a quantidade e valor inicial na implantacao do produto na Filial
    lErro = Comando_Executar(alComando(1), "SELECT " & sCampos & ", DataInicial FROM EstoqueProduto, Almoxarifado WHERE Codigo=Almoxarifado AND Produto = ? AND FilialEmpresa = ?", dQuantInicial, dValorInicial, dtDataInicial1, sProduto, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 93747
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93748
    
    If lErro = AD_SQL_SEM_DADOS Then gError 93749
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'se a data de inicializacao do produto no estoque for menor que a data final do periodo solicitado
        '==> os valores não estão disponíveis
        If dtDataFinal <> DATA_NULA And dtDataInicial1 > dtDataFinal Then
            dQuantInicial = 0
            dValorInicial = 0
        End If
            
        dQuantInicial1Total = dQuantInicial1Total + dQuantInicial
        dValorInicial1Total = dValorInicial1Total + dValorInicial
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93739
       
    Loop
                
    If dtDataInicial <> DATA_NULA Then
    
        Select Case iNumEscaninho
    
            Case ESCANINHO_CONSERTO_NOSSO
                sQuantEnt = "QuantEntConserto"
                sQuantSaida = "QuantSaiConserto"
                sValorEntrada = "ValorEntConserto"
                sValorSaida = "ValorSaiConserto"
                
            Case ESCANINHO_CONSIG_NOSSO
                sQuantEnt = "QuantEntConsig"
                sQuantSaida = "QuantSaiConsig"
                sValorEntrada = "ValorEntConsig"
                sValorSaida = "ValorSaiConsig"
                
            Case ESCANINHO_DEMO_NOSSO
                sQuantEnt = "QuantEntDemo"
                sQuantSaida = "QuantSaiDemo"
                sValorEntrada = "ValorEntDemo"
                sValorSaida = "ValorSaiDemo"
                
            Case ESCANINHO_OUTROS_NOSSO
                sQuantEnt = "QuantEntOutros"
                sQuantSaida = "QuantSaiOutros"
                sValorEntrada = "ValorEntOutros"
                sValorSaida = "ValorSaiOutros"
                
            Case ESCANINHO_BENEF_NOSSO
                sQuantEnt = "QuantEntBenef"
                sQuantSaida = "QuantSaiBenef"
                sValorEntrada = "ValorEntBenef"
                sValorSaida = "ValorSaiBenef"
                
            Case ESCANINHO_CONSERTO_3
                sQuantEnt = "QuantEntConserto3"
                sQuantSaida = "QuantSaiConserto3"
                sValorEntrada = "ValorEntConserto3"
                sValorSaida = "ValorSaiConserto3"
                
            Case ESCANINHO_CONSIG_3
                sQuantEnt = "QuantEntConsig3"
                sQuantSaida = "QuantSaiConsig3"
                sValorEntrada = "ValorEntConsig3"
                sValorSaida = "ValorSaiConsig3"
                
            Case ESCANINHO_DEMO_3
                sQuantEnt = "QuantEntDemo3"
                sQuantSaida = "QuantSaiDemo3"
                sValorEntrada = "ValorEntDemo3"
                sValorSaida = "ValorSaiDemo3"
            
            Case ESCANINHO_OUTROS_3
                sQuantEnt = "QuantEntOutros3"
                sQuantSaida = "QuantSaiOutros3"
                sValorEntrada = "ValorEntOutros3"
                sValorSaida = "ValorSaiOutros3"
            
            Case ESCANINHO_BENEF_3
                sQuantEnt = "QuantEntBenef3"
                sQuantSaida = "QuantSaiBenef3"
                sValorEntrada = "ValorEntBenef3"
                sValorSaida = "ValorSaiBenef3"
        
        End Select
                   
        'obter os totais movimentados até a data anterior a inicial
        lErro = Comando_Executar(alComando(0), "SELECT SUM(" & sQuantEnt & ") - SUM(" & sQuantSaida & "), SUM(" & sValorEntrada & ") - SUM(" & sValorSaida & ") FROM SldDiaEst WHERE Produto = ? AND FilialEmpresa = ? AND Data < ?", dQuantInicial1, dValorInicial1, sProduto, iFilialEmpresa, dtDataInicial)
        If lErro <> AD_SQL_SUCESSO Then gError 93750
    
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 93751
    
        dQuantInicial = dQuantInicial1Total + dQuantInicial1
        dValorInicial = dValorInicial1Total + dValorInicial1

        iAno = Year(dtDataInicial)
        iMes = Month(dtDataInicial)
        
        'Se é um produto produzido e a ultima apuracao de custo de producao for menor do que ano/mes inicial do periodo solicitado ==> o saldo em valor não está disponível
        If iApropriacaoCusto = APROPR_CUSTO_REAL And (iAnoEstApurado < iAno Or (iAno = iAnoEstApurado And iMesEstApurado < iMes)) Then dValorInicial = -1
         
    Else
    
        dQuantInicial = dQuantInicial1Total
        dValorInicial = dValorInicial1Total
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
             
    Produto_ObterRelQtdSldIniFilTercEsc = SUCESSO
    
    Exit Function

Erro_Produto_ObterRelQtdSldIniFilTercEsc:

    Produto_ObterRelQtdSldIniFilTercEsc = gErr

    Select Case gErr

        Case 93746
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 93747, 93748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO2", gErr, sProduto)

        Case 93749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO1", gErr, sProduto, iFilialEmpresa)

        Case 93750, 93751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAEST", gErr, iFilialEmpresa, sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152152)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Function ProdutoEmbalagem_Le(objProdutoEmbalagem As ClassProdutoEmbalagem) As Long
'Le a UMEmbalagem e a capacidade a partir do Produto e da Embalagem

Dim lErro As Long
Dim lComando As Long
Dim tProdEmb As typeProdutoEmbalagem

On Error GoTo Erro_ProdutoEmbalagem_Le

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 96147

    With tProdEmb
    
        .sNomeProdEmb = String(STRING_PRODUTOEMBALAGEM_NOMEPRODEMB, 0)
        .sUMEmbalagem = String(STRING_UM_SIGLA, 0)
        .sUMPeso = String(STRING_UM_SIGLA, 0)
        
    End With
    
    'Lê as embalagens do Produto
    lErro = Comando_Executar(lComando, "SELECT SeqGrid, Padrao, NomeProdEmb, UMEmbalagem, Capacidade, UMPeso, PesoLiqTotal, PesoBruto FROM ProdutoEmbalagem WHERE Produto = ? AND Embalagem = ?", _
        tProdEmb.iSeqGrid, tProdEmb.iPadrao, tProdEmb.sNomeProdEmb, tProdEmb.sUMEmbalagem, tProdEmb.dCapacidade, tProdEmb.sUMPeso, tProdEmb.dPesoLiqTotal, tProdEmb.dPesoBruto, objProdutoEmbalagem.sProduto, objProdutoEmbalagem.iEmbalagem)
    If lErro <> AD_SQL_SUCESSO Then gError 96148

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 96149
    
    If lErro <> AD_SQL_SUCESSO Then gError 96156
    
    With tProdEmb
        
        objProdutoEmbalagem.iSeqGrid = .iSeqGrid
        objProdutoEmbalagem.iPadrao = .iPadrao
        objProdutoEmbalagem.sNomeProdEmb = .sNomeProdEmb
        objProdutoEmbalagem.sUMEmbalagem = .sUMEmbalagem
        objProdutoEmbalagem.dCapacidade = .dCapacidade
        objProdutoEmbalagem.sUMPeso = .sUMPeso
        objProdutoEmbalagem.dPesoLiqTotal = .dPesoLiqTotal
        objProdutoEmbalagem.dPesoBruto = .dPesoBruto
        
        'calculado
        objProdutoEmbalagem.dPesoEmbalagem = objProdutoEmbalagem.dPesoBruto - objProdutoEmbalagem.dPesoLiqTotal
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ProdutoEmbalagem_Le = SUCESSO

    Exit Function

Erro_ProdutoEmbalagem_Le:

    ProdutoEmbalagem_Le = gErr

    Select Case gErr
       
        Case 96147
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 96148, 96149
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOEMBALAGEM", gErr, objProdutoEmbalagem.sProduto)
 
        Case 96156
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152153)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Revisado - OK Luiz 19/02/02
Public Function ProdutoHistorico_Le_Todos(ByVal colProdutoHistorico As Collection) As Long
'Lê todos os registros da tabela ProdutoHistorico e devolve-os na coleção
'colProdutoOPHistorico RETORNA(Output) os registros lidos

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tProdutoHistorico As typeProdutoHistorico
Dim objProdutoHistorico As ClassProdutoHistorico

On Error GoTo Erro_ProdutoHistorico_Le_Todos

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97821
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT ProdutoHistorico.DataAtualizacao, ProdutoHistorico.CodProduto, ProdutoHistorico.DescProduto FROM ProdutoHistorico"
    
    With tProdutoHistorico
        
        'Inicializa as Strings do Type
        .sCodigoProd = String(STRING_PRODUTO, 0)
        .sDescProduto = String(STRING_PRODUTO_DESCRICAO, 0)
        
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .dtDataAtualizacao, .sCodigoProd, .sDescProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 97822
    
    End With
        
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97823

    'Se nao encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 97820
    
    'Enquanto tem Produto => adiciona na colecao
    Do While lErro <> AD_SQL_SEM_DADOS
                                    
        'Inicializa o obj
        Set objProdutoHistorico = New ClassProdutoHistorico
        
        With objProdutoHistorico
    
            .dtDataAtualizacao = tProdutoHistorico.dtDataAtualizacao
            .sCodProduto = tProdutoHistorico.sCodigoProd
            .sDescProduto = tProdutoHistorico.sDescProduto
            
        End With
    
        'Adiciona na Colecao
        colProdutoHistorico.Add objProdutoHistorico
    
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97824
        
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    ProdutoHistorico_Le_Todos = SUCESSO
    
    Exit Function
    
Erro_ProdutoHistorico_Le_Todos:

    ProdutoHistorico_Le_Todos = gErr

    Select Case gErr
        
        Case 97821
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97822, 97823, 97824
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOHISTORICO", gErr)
            
        Case 97820 'Sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152154)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

'Rotina Criada por Daniel em 18/02/02
Function Produto_Critica_Usu(objProduto As ClassProduto) As Long
'Como é Versao2 => apenas retorna SUCESSO
'Em versões USU ela Critica o uso do Produto em:
'CotacaoServico, CompServGr, ServicoItemServico, TabPrecoItensGR, SolServServico (G.RIO)

    Produto_Critica_Usu = SUCESSO

End Function


'*** ALTERACAO POR TULIO EM 29/05/02 ***
Public Function CodigosBarra_Le_Produto(ByVal objProduto As ClassProduto) As Long
'Le os codigos de barra do produto passado como parametro e retorna na colecao
'de codigos de barra que o mesmo parametro possui os codigos de barra associados
'objProduto eh parametro de INPUT E OUTPUT

Dim lErro As Long
Dim lComando As Long
Dim sCodBarras As String

On Error GoTo Erro_CodigosBarra_Le_Produto

    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 101731

    'inicializa o buffer de recebimento
    sCodBarras = String(STRING_PRODUTO_CODIGO_BARRAS, 0)

    'obtem o conjunto de codigos de barra relativos ao produto em questao
    lErro = Comando_Executar(lComando, "SELECT CodBarras FROM ProdutoCodBarras WHERE CodProduto = ?", sCodBarras, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 101732
    
    'se posiciona no 1o registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101733
    
    'enquanto houverem registros no conjunto
    'obs. ela nao retorna erro se nao achar propositalmente...
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'adiciona na colecao o dito cujo
        objProduto.colCodBarras.Add sCodBarras
    
        'tenta ler o proximo do conjunto
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101734
    
    Loop
    
    'fecha comando
    Call Comando_Fechar(lComando)
    
    CodigosBarra_Le_Produto = SUCESSO
       
    Exit Function

Erro_CodigosBarra_Le_Produto:

    CodigosBarra_Le_Produto = gErr

    Select Case gErr
        
        Case 101732, 101733, 101734
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "ProdutoCodBarras")
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152155)

    End Select

    'fecha comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function
'*** FIM ALTERACAO POR TULIO EM 29/05/02 ***
Function Kit_Le_Padrao(ByVal objKit As ClassKit) As Long
'Le a versao padrao do kit

Dim lErro As Long
Dim lComando As Long
Dim tKit As typeKit

On Error GoTo Erro_Kit_Le_Padrao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106301

    With tKit
        .sVersao = String(STRING_KIT_VERSAO, 0)
        .sObservacao = String(STRING_KIT_OBSERVACAO, 0)

        lErro = Comando_Executar(lComando, "SELECT Versao, Observacao, Data, Situacao, VersaoFormPreco, PesoFator1, PesoFator2, PesoFator3, PesoFator4, PesoFator5, PesoFator6 FROM Kit WHERE ProdutoRaiz = ? AND Situacao = ?", .sVersao, .sObservacao, .dtData, .iSituacao, .iVersaoFormPreco, .dPesoFator1, .dPesoFator2, .dPesoFator3, .dPesoFator4, .dPesoFator5, .dPesoFator6, objKit.sProdutoRaiz, KIT_SITUACAO_PADRAO)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 106302

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106303

    If lErro = AD_SQL_SEM_DADOS Then gError 106304

    With tKit
        objKit.sVersao = .sVersao
        objKit.dtData = .dtData
        objKit.sObservacao = .sObservacao
        objKit.iSituacao = .iSituacao
        objKit.iVersaoFormPreco = .iVersaoFormPreco
        objKit.dPesoFator1 = .dPesoFator1
        objKit.dPesoFator2 = .dPesoFator2
        objKit.dPesoFator3 = .dPesoFator3
        objKit.dPesoFator4 = .dPesoFator4
        objKit.dPesoFator5 = .dPesoFator5
        objKit.dPesoFator6 = .dPesoFator6
    End With

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Kit_Le_Padrao = SUCESSO

    Exit Function

Erro_Kit_Le_Padrao:

    Kit_Le_Padrao = gErr

    Select Case gErr

        Case 106301
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 106302, 106303
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_KIT", gErr)

        Case 106304

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152156)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

End Function

Public Function Kit_Le_Produziveis(ByVal objKit As ClassKit, ByVal colKits As Collection) As Long
'Função que le todos os kits ativos e o padrao retornado na coleção

Dim lErro As Long
Dim lComando As Long
Dim tKit As typeKit

On Error GoTo Erro_Kit_Le_Produziveis
    
    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106330

    tKit.sObservacao = String(STRING_KIT_OBSERVACAO, 0)
    tKit.sProdutoRaiz = String(STRING_PRODUTO, 0)
    tKit.sVersao = String(STRING_KIT_VERSAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT ProdutoRaiz ,Versao, Data, Observacao, Situacao, VersaoFormPreco FROM Kit WHERE ProdutoRaiz = ? AND Situacao <> ?", tKit.sProdutoRaiz, tKit.sVersao, tKit.dtData, tKit.sObservacao, tKit.iSituacao, tKit.iVersaoFormPreco, objKit.sProdutoRaiz, KIT_SITUACAO_INATIVO)
    If lErro <> AD_SQL_SUCESSO Then gError 106331

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106332

    If lErro = AD_SQL_SEM_DADOS Then gError 106333
    
    'Passa para a collection
    Do While lErro = AD_SQL_SUCESSO

        Set objKit = New ClassKit
        
        objKit.sProdutoRaiz = tKit.sProdutoRaiz
        objKit.sVersao = tKit.sVersao
        objKit.dtData = tKit.dtData
        objKit.sObservacao = tKit.sObservacao
        objKit.iSituacao = tKit.iSituacao
        objKit.iVersaoFormPreco = tKit.iVersaoFormPreco
        
        colKits.Add objKit

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106334

    Loop
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    Kit_Le_Produziveis = SUCESSO

    Exit Function

Erro_Kit_Le_Produziveis:

    Kit_Le_Produziveis = gErr

    Select Case gErr

        Case 106330
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 106331, 106332, 106334
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_KIT", gErr)

        Case 106333

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152157)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Alteracao Daniel em 11/07/2002
Public Function Kit_Verifica_Movimentacao(ByVal sProdutoRaiz As String, ByVal sVersao As String) As Long
'Verifica se já existe ordem de producao para o produto "sProdutoRaiz" e para a versao "sVersao"

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim sOP As String
Dim lComando1 As Long

On Error GoTo Erro_Kit_Verifica_Movimentacao

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106220
    
    'Abertura do Comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 106373
    
    'Inicializa a string
    sOP = String(STRING_OPCODIGO, 0)
    
    'Guarda o comando SQL
    sSelect = "SELECT Codigo FROM ItensOrdemProducao WHERE Produto = ? AND Versao = ?"

    'Prepara a execucao do comando SQL
    lErro = Comando_Executar(lComando, sSelect, sOP, sProdutoRaiz, sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 106221
    
    'Busca o Primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106222
    
    'Se encontrou => Erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 106223
    
    'Inicializa a string
    sOP = String(STRING_OPCODIGO, 0)
    
    'Guarda o comando SQL
    sSelect = "SELECT Codigo FROM ItensOrdemProducaoBaixadas WHERE Produto = ? AND Versao = ?"

    'Prepara a execucao do comando SQL
    lErro = Comando_Executar(lComando, sSelect, sOP, sProdutoRaiz, sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 106370
    
    'Busca o Primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106371
    
    'Se encontrou => Erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 106372
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Kit_Verifica_Movimentacao = SUCESSO
    
    Exit Function

Erro_Kit_Verifica_Movimentacao:
    
    Kit_Verifica_Movimentacao = gErr

    Select Case gErr

        Case 106220, 106373
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 106221, 106222, 106370, 106371
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO", gErr)
            
        Case 106223, 106372
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_COM_MOVIMENTACAO", gErr)
            'ERRO_KIT_COM_MOVIMENTACAO = "O Kit já foi movimentado e não poderá ser excluído. Para inutilizá-lo grave-o com Situacao = Inativo."
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152158)
        
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

End Function
'**********************************************************
'Alteracao Daniel em 11/07/2002
Public Function Kit_Verifica_Outras_Versoes(ByVal sProdutoRaiz As String, ByVal sVersao As String, iQtdVersoesAtivas As Integer, iQtdVersoesInativas As Integer, sProxVersao As String) As Long
'Verifica a existencia de OUTRAS versoes do kit(sProdutoRaiz)

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim iSituacao As Integer, sProxVersaoTemp As String

On Error GoTo Erro_Kit_Verifica_Outras_Versoes

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106214

    'Inicializacao da string
    sProxVersaoTemp = String(STRING_KIT_VERSAO, 0)
    
    'Guarda a Selecao SQL
    sSelect = "SELECT Versao, Situacao FROM Kit WHERE ProdutoRaiz = ? AND Versao <> ?"
    
    'Prepara a Execucao do Comando SQL
    lErro = Comando_Executar(lComando, sSelect, sProxVersaoTemp, iSituacao, sProdutoRaiz, sVersao)
    If lErro <> AD_SQL_SUCESSO Then gError 106215
    
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106216
    
    'Se nao encontrou outra versao para o kit
    If lErro = AD_SQL_SEM_DADOS Then gError 106217
        
    'Enquanto existir versoes para o kit
    Do While (lErro <> AD_SQL_SEM_DADOS)
    
        'Se for uma versao ativa => Incrementa o numero de versoes ativas encontradas
        If iSituacao = KIT_SITUACAO_ATIVO Or iSituacao = KIT_SITUACAO_PADRAO Then
        
            iQtdVersoesAtivas = iQtdVersoesAtivas + 1
            sProxVersao = sProxVersaoTemp
            
        ElseIf iSituacao = KIT_SITUACAO_INATIVO Then
        
            iQtdVersoesInativas = iQtdVersoesInativas + 1
            
        End If
        
        'Busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106218
    
    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Kit_Verifica_Outras_Versoes = SUCESSO
    
    Exit Function
    
Erro_Kit_Verifica_Outras_Versoes:

    Kit_Verifica_Outras_Versoes = gErr
    
    Select Case gErr
    
        Case 106214
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 106215, 106216, 106218
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_KIT", gErr)

        Case 106217
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152159)
        
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Function AlmoxarifadoPadrao_Le_CodNomeRed(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, sNomeReduzidoAlmox As String, iAlmoxarifadoPadrao As Integer) As Long
'Retorna o Nome reduzido e o código do Almoxarifado Padrão do Produto em Questão com a Filial Empresa

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxarifadoPadrao_Le_CodNomeRed

    'se o almoxarifado padrão não foi sugerido na tela - - > Busca o padrão do produto
    lErro = CF("AlmoxarifadoPadrao_Le", iFilialEmpresa, sProduto, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 118755
    
    'se não encontrou o Registro então  -- -  - - > Erro
    If lErro = 23796 Then gError 118756
    
    'se o Almoxarifado Padrão estiver Preenchido le seu nome reduzido
    If iAlmoxarifadoPadrao <> 0 Then
        
        objAlmoxarifado.iCodigo = iAlmoxarifadoPadrao
        
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 118757
        
        'se não encontrou o registro  - - - -> ERRO
        If lErro = 25056 Then gError 118758
            
        sNomeReduzidoAlmox = objAlmoxarifado.sNomeReduzido
    
    End If
            
    AlmoxarifadoPadrao_Le_CodNomeRed = SUCESSO
        
    Exit Function
        
Erro_AlmoxarifadoPadrao_Le_CodNomeRed:
    
    AlmoxarifadoPadrao_Le_CodNomeRed = gErr
    
    Select Case gErr
         
        Case 118755, 118757
            
        Case 118756
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, sProduto)
            
        Case 118758
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, iAlmoxarifadoPadrao)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 152160)
        
    End Select
    
    Exit Function
        
End Function

'Função criada por Mauricio MAciel em 16/04/03
Function PrecoCalculado_Le(ByVal objPrecoCalculado As ClassPrecoCalculado) As Long
'Lê a tabela de preço a partir da chave fornecida em objTabelaPreco. Retorna os dados em objTabelaPreco

Dim lErro As Long
Dim lComando As Long
Dim tPrecoCalculado As typePrecoCalculado

On Error GoTo Erro_PrecoCalculado_Le

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 116822

    If objPrecoCalculado.dtDataReferencia <> DATA_NULA Then
    
        'Pesquisa no BD Preco Calculado
        lErro = Comando_Executar(lComando, "SELECT PrecoCalculado, DataVigencia, DataReferencia FROM PrecoCalculado WHERE FilialEmpresa = ? AND CodProduto = ? AND DataReferencia <= ? AND CodTabela = ? ORDER BY DataReferencia DESC", tPrecoCalculado.dPrecoCalculado, tPrecoCalculado.dtDataVigencia, tPrecoCalculado.dtDataReferencia, objPrecoCalculado.iFilialEmpresa, objPrecoCalculado.sCodProduto, objPrecoCalculado.dtDataReferencia, objPrecoCalculado.iCodTabela)
        If lErro <> AD_SQL_SUCESSO Then gError 116823

    Else
    
        'Pesquisa no BD Preco Calculado, sem data referencia especifica
        lErro = Comando_Executar(lComando, "SELECT PrecoCalculado, DataVigencia, DataReferencia FROM PrecoCalculado WHERE FilialEmpresa = ? AND CodProduto = ? AND CodTabela = ? ORDER BY DataReferencia DESC", tPrecoCalculado.dPrecoCalculado, tPrecoCalculado.dtDataVigencia, tPrecoCalculado.dtDataReferencia, objPrecoCalculado.iFilialEmpresa, objPrecoCalculado.sCodProduto, objPrecoCalculado.iCodTabela)
        If lErro <> AD_SQL_SUCESSO Then gError 116824

    End If

    'Verifica se Ítem existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 116825

    'Não encontrou o Ítem
    If lErro = AD_SQL_SEM_DADOS Then gError 116826

    'Preenche objPrecoCalculado
    objPrecoCalculado.dPrecoCalculado = tPrecoCalculado.dPrecoCalculado
    objPrecoCalculado.dtDataVigencia = tPrecoCalculado.dtDataVigencia
    objPrecoCalculado.dtDataReferencia = tPrecoCalculado.dtDataReferencia
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    PrecoCalculado_Le = SUCESSO

    Exit Function

Erro_PrecoCalculado_Le:

    PrecoCalculado_Le = gErr

    Select Case gErr

        Case 116822
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 116823, 116824, 116825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRECOCALCULADO", gErr, objPrecoCalculado.iFilialEmpresa, objPrecoCalculado.sCodProduto)

        Case 116826
            'Não encontrou a Ítem da Tabela de Preço. Erro tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152161)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Le_Filhos_Grade(ByVal objProduto As ClassProduto, ByVal colItemRomaneio As Collection)
'Busca no BD os produtos filhos do Produto passado.
'Os filhos só serão retornado caso estejam com os itens de categoria da grade passada preenchidos
'Os produtos deverão existir na Filial Empresa em questão.

Dim lComando As Long
Dim lErro As Long
Dim tItemRomaneioGrade As typeItemRomaneioGrade
Dim objItemGrade As New ClassItemRomaneioGrade
Dim tCategProdItem As typeCategoriaProdutoItem
Dim objCategItem As New ClassCategoriaProdutoItem
Dim sProduto As String
Dim bAchou As Boolean

On Error GoTo Erro_Produto_Le_Filhos_Grade

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86299

    'Inicializa os buffers de leitura de texto
    sProduto = String(STRING_PRODUTO, 0)
    tCategProdItem.sCategoria = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    tCategProdItem.sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    tItemRomaneioGrade.sProduto = String(STRING_PRODUTO, 0)
    tItemRomaneioGrade.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
    tItemRomaneioGrade.sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)

    'Busca o próximo produto do mesmo nível que o passado pois ele vai determinar
    'o intervalo para a busca dos filhos
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Produtos WHERE Codigo > ? AND Nivel = ? ORDER BY Codigo", sProduto, objProduto.sCodigo, objProduto.iNivel)
    If lErro <> AD_SQL_SUCESSO Then gError 86300

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86301

    'Se não existe nenhum produto do mesmo nível
    If lErro <> AD_SQL_SUCESSO Then
        'BUsca todos os Filhos desse produto nessa filial que possuem os itens para a grade informada
        lErro = Comando_Executar(lComando, "SELECT ProdutoCategoria.Produto, Produtos.Descricao, ProdutoCategoria.Categoria, ProdutoCategoria.Item, Produtos.SiglaUMEstoque FROM ProdutosFilial, Produtos,ProdutoCategoria,GradeCategoria, CategoriaProdutoItem WHERE ProdutoCategoria.Produto IN (SELECT Produto FROM ProdutoCategoria WHERE Categoria IN (SELECT Categoria FROM GradeCategoria WHERE Grade = ? ) GROUP BY Produto HAVING COUNT(*) = (SELECT COUNT(*) FROM GradeCategoria WHERE Grade = ?) ) " _
                                 & " AND ProdutosFilial.Produto = Produtos.Codigo AND GradeCategoria.Grade = ? AND GradeCategoria.Categoria = ProdutoCategoria.Categoria AND Nivel = ? AND ProdutoCategoria.Produto > ? AND ProdutoCategoria.Produto = Produtos.Codigo AND ProdutosFilial.FilialEmpresa = ? AND Produtos.ClasseUM = ? AND CategoriaProdutoItem.Categoria = ProdutoCategoria.Categoria AND CategoriaProdutoItem.Item = ProdutoCategoria.Item ORDER BY Posicao, CategoriaProdutoItem.Ordem", tItemRomaneioGrade.sProduto, tItemRomaneioGrade.sDescricao, tCategProdItem.sCategoria, tCategProdItem.sItem, tItemRomaneioGrade.sSiglaUMEstoque, objProduto.sGrade, objProduto.sGrade, objProduto.sGrade, objProduto.iNivel + 1, objProduto.sCodigo, giFilialEmpresa, objProduto.iClasseUM)

    Else
        'BUsca todos os Filhos desse produto nessa filial que possuem os itens para a grade informada
        lErro = Comando_Executar(lComando, "SELECT ProdutoCategoria.Produto, Produtos.Descricao, ProdutoCategoria.Categoria, ProdutoCategoria.Item, Produtos.SiglaUMEstoque FROM ProdutosFilial, Produtos,ProdutoCategoria,GradeCategoria, CategoriaProdutoItem WHERE ProdutoCategoria.Produto IN (SELECT Produto FROM ProdutoCategoria WHERE Categoria IN (SELECT Categoria FROM GradeCategoria WHERE Grade = ? ) GROUP BY Produto HAVING COUNT(*) = (SELECT COUNT(*) FROM GradeCategoria WHERE Grade = ?) ) " _
                                 & " AND ProdutosFilial.Produto = Produtos.Codigo AND GradeCategoria.Grade = ? AND GradeCategoria.Categoria = ProdutoCategoria.Categoria AND Nivel = ? AND ProdutoCategoria.Produto BETWEEN ? AND ? AND ProdutoCategoria.Produto = Produtos.Codigo AND ProdutosFilial.FilialEmpresa = ? AND Produtos.ClasseUM = ? AND CategoriaProdutoItem.Categoria = ProdutoCategoria.Categoria AND CategoriaProdutoItem.Item = ProdutoCategoria.Item ORDER BY Posicao, CategoriaProdutoItem.Ordem", tItemRomaneioGrade.sProduto, tItemRomaneioGrade.sDescricao, tCategProdItem.sCategoria, tCategProdItem.sItem, tItemRomaneioGrade.sSiglaUMEstoque, objProduto.sGrade, objProduto.sGrade, objProduto.sGrade, objProduto.iNivel + 1, objProduto.sCodigo, sProduto, giFilialEmpresa, objProduto.iClasseUM)

    End If
    If lErro <> AD_SQL_SUCESSO Then gError 86302

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86303
    
    'Se não possuir nenhum produto filho com grade informada ===> erro
    If lErro <> AD_SQL_SUCESSO Then gError 86304

    'Enquanto o Produto possuir itens com a grade completa
    Do While lErro = AD_SQL_SUCESSO
        'Se estiver lendo um novo item
        
        Call Busca_ItemGrade(colItemRomaneio, tItemRomaneioGrade.sProduto, objItemGrade, bAchou)
        
        If Not bAchou Then
            'Cria o Item de gRade
            Set objItemGrade = New ClassItemRomaneioGrade

            objItemGrade.sProduto = tItemRomaneioGrade.sProduto
            objItemGrade.sDescricao = tItemRomaneioGrade.sDescricao
            objItemGrade.sUMEstoque = tItemRomaneioGrade.sSiglaUMEstoque
            
            'Adiciona na coleção de itens
            colItemRomaneio.Add objItemGrade

        End If

        'Guarda a opção de Item-Categoria do produto
        Set objCategItem = New ClassCategoriaProdutoItem

        objCategItem.sCategoria = tCategProdItem.sCategoria
        objCategItem.sItem = tCategProdItem.sItem

        'Guarda a Informação on item da Grade
        objItemGrade.colCategoria.Add objCategItem

        'BUsca o Proximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86305

    Loop
    
    Call Comando_Fechar(lComando)
    
    Produto_Le_Filhos_Grade = SUCESSO
    
    Exit Function

Erro_Produto_Le_Filhos_Grade:
    
    Produto_Le_Filhos_Grade = gErr
    
    Select Case gErr
    
        Case 86299
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 86300, 86301, 86302, 86303, 86305
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
        
        Case 86304
            'O produto não tem Filhos com a grade informada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152162)
    
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoGrade_Possui_Filhos(objProduto As ClassProduto, bPossuiFilhos As Boolean) As Long
'Verifica se o Produto de Grade passado tem filhos analíticos de grade da
'Na giFilialEmpresa e com a mesma ClasseUM

Dim lComando As Long
Dim lErro As Long
Dim iQuantFilhos As Integer
Dim sProduto As String

On Error GoTo Erro_ProdutoGrade_Possui_Filhos

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86314

    'Inicializa os buffers de leitura de texto
    sProduto = String(STRING_PRODUTO, 0)
    bPossuiFilhos = False
    
    'Busca o próximo produto do mesmo nível que o passado pois ele vai determinar
    'o intervalo para a busca dos filhos
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Produtos WHERE Codigo > ? AND Nivel = ? ORDER BY Codigo", sProduto, objProduto.sCodigo, objProduto.iNivel)
    If lErro <> AD_SQL_SUCESSO Then gError 86315

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86316

    'Se não existe nenhum produto do mesmo nível
    If lErro <> AD_SQL_SUCESSO Then
        'BUsca todos os Filhos desse produto nessa filial que possuem os itens para a grade informada
        lErro = Comando_Executar(lComando, "SELECT COUNT(Produtos.Codigo) FROM ProdutosFilial, Produtos,ProdutoCategoria,GradeCategoria WHERE ProdutoCategoria.Produto IN (SELECT Produto FROM ProdutoCategoria WHERE Categoria IN (SELECT Categoria FROM GradeCategoria WHERE Grade = ? ) GROUP BY Produto HAVING COUNT(*) = (SELECT COUNT(*) FROM GradeCategoria WHERE Grade = ?) ) AND ProdutosFilial.Produto = Produtos.Codigo AND GradeCategoria.Grade = ? AND GradeCategoria.Categoria = ProdutoCategoria.Categoria AND Nivel = ? AND ProdutoCategoria.Produto > ? AND ProdutoCategoria.Produto = Produtos.Codigo AND ProdutosFilial.FilialEmpresa = ? AND Produtos.ClasseUM = ? ", _
        iQuantFilhos, objProduto.sGrade, objProduto.sGrade, objProduto.sGrade, objProduto.iNivel + 1, objProduto.sCodigo, giFilialEmpresa, objProduto.iClasseUM)

    Else
        'BUsca todos os Filhos desse produto nessa filial que possuem os itens para a grade informada
        lErro = Comando_Executar(lComando, "SELECT COUNT(Produtos.Codigo) FROM ProdutosFilial, Produtos,ProdutoCategoria,GradeCategoria WHERE ProdutoCategoria.Produto IN (SELECT Produto FROM ProdutoCategoria WHERE Categoria IN (SELECT Categoria FROM GradeCategoria WHERE Grade = ? ) GROUP BY Produto HAVING COUNT(*) = (SELECT COUNT(*) FROM GradeCategoria WHERE Grade = ?) ) AND ProdutosFilial.Produto = Produtos.Codigo AND GradeCategoria.Grade = ? AND GradeCategoria.Categoria = ProdutoCategoria.Categoria AND Nivel = ? AND ProdutoCategoria.Produto BETWEEN ? AND ? AND ProdutoCategoria.Produto = Produtos.Codigo AND ProdutosFilial.FilialEmpresa = ? AND Produtos.ClasseUM = ?", _
        iQuantFilhos, objProduto.sGrade, objProduto.sGrade, objProduto.sGrade, objProduto.iNivel + 1, objProduto.sCodigo, sProduto, giFilialEmpresa, objProduto.iClasseUM)

    End If
    If lErro <> AD_SQL_SUCESSO Then gError 86317

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 86318

    If iQuantFilhos > 0 Then bPossuiFilhos = True
    
    Call Comando_Fechar(lComando)
    
    ProdutoGrade_Possui_Filhos = SUCESSO
    
    Exit Function

Erro_ProdutoGrade_Possui_Filhos:
    
    ProdutoGrade_Possui_Filhos = gErr
    
    Select Case gErr
    
        Case 86314
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 86315, 86316, 86317, 86318
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
        
        Case 86304
            'O produto não tem Filhos com a grade informada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152163)
    
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function



Function Produto_Le_PaiGrade(ByVal objProduto As ClassProduto, ByRef sProdutoPai As String) As Long
'Retorna o Código do Produto Pai de Grade para o objProduto passado
'Caso ele não tenha um pai de grade, ou apesar do pai ter uma grade, ele não
'seja um filho de grade (categorias nao preenchidas) o sProdutoPai retorna em Branco

Dim lComando As Long
Dim lErro As Long
Dim sCodigo As String
Dim sGrade As String
Dim objProdutoAux As New ClassProduto
Dim colItensRomaneioGrade As New Collection
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim iClasseUM As Integer

On Error GoTo Erro_Produto_Le_PaiGrade

    sProdutoPai = ""

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86323
    
    sCodigo = String(STRING_PRODUTO, 0)
    sGrade = String(STRING_GRADE_CODIGO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, Grade, CLasseUM FROM Produtos WHERE Codigo < ? AND Nivel = ? ORDER BY Codigo DESC", sCodigo, sGrade, iClasseUM, objProduto.sCodigo, objProduto.iNivel - 1)
    If lErro <> AD_SQL_SUCESSO Then gError 86324

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86325
    
    If (lErro = AD_SQL_SUCESSO) And Len(Trim(sGrade)) > 0 Then
            
        objProdutoAux.sCodigo = sCodigo
        objProdutoAux.sGrade = sGrade
        objProdutoAux.iNivel = objProduto.iNivel - 1
        objProdutoAux.iClasseUM = iClasseUM
    
        lErro = CF("Produto_Le_Filhos_Grade", objProdutoAux, colItensRomaneioGrade)
        If lErro <> SUCESSO Then Error 86326
    
        For Each objItemRomaneio In colItensRomaneioGrade
            If objItemRomaneio.sProduto = objProduto.sCodigo Then
                sProdutoPai = sCodigo
                Exit For
            End If
        Next
    
    End If
    
    Call Comando_Fechar(lComando)
    
    Produto_Le_PaiGrade = SUCESSO
    
    Exit Function

Erro_Produto_Le_PaiGrade:

    Produto_Le_PaiGrade = gErr
    
    Select Case gErr
    
        Case 86323
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 86324, 86325
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objProduto.sCodigo)

        Case 86326
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152164)

    End Select

    Exit Function

End Function


Private Sub Busca_ItemGrade(colItemRomaneio As Collection, sProduto As String, objItemGrade As ClassItemRomaneioGrade, bAchou As Boolean)

Dim objItemCol As ClassItemRomaneioGrade

    bAchou = False
    
    For Each objItemCol In colItemRomaneio
        If objItemCol.sProduto = sProduto Then
            Set objItemGrade = objItemCol
            bAchou = True
            Exit For
        End If
    Next
    
    Exit Sub
    
End Sub


Function Produto_Le_ProdutoPai(ByVal sProdutoFilho As String, ByVal objProdutoPai As ClassProduto) As Long
'Lê o Produto pai, e retorna no objProdutoPai, o pai do Produto Passado

Dim sProduto1 As String
Dim iNivel As Integer
Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Produto_Le_ProdutoPai

    'retorna o nivel do Produto
    lErro = Mascara_Produto_ObterNivel(sProdutoFilho, iNivel)
    If lErro <> SUCESSO Then gError 86409

    'se a Produto tem um "pai"
    If iNivel > 1 Then

        sProduto1 = String(STRING_PRODUTO, 0)

        'descobre a Produto "pai"
        lErro = Mascara_RetornaProdutoNoNivel(iNivel - 1, sProdutoFilho, sProduto1)
        If lErro <> SUCESSO Then gError 86410

        'verifica se a Produto "pai" está cadastrado
        objProdutoPai.sCodigo = sProduto1

        lErro = CF("Produto_Le", objProdutoPai)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 86411

        'nao encontrou a Produto
        If lErro = 28030 Then gError 86412

    End If

    Produto_Le_ProdutoPai = SUCESSO

    Exit Function

Erro_Produto_Le_ProdutoPai:

    Produto_Le_ProdutoPai = gErr

    Select Case gErr

        Case 86409

            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_PRODUTO_OBTERNIVEL", gErr, sProdutoFilho)

        Case 86410
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTONONIVEL", gErr, sProdutoFilho, iNivel - 1)

        Case 86411, 86412

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152165)

    End Select

    Exit Function

End Function

Function SldMesEst_Le_CustoMedio2(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, ByVal sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'le os dados de custo medio de SldMesEst relativos ao Ano, FilialEmpresa, Produto e mes passados como parametro
    
Dim lErro As Long
Dim tSldMesEst As typeSldMesEst
Dim dQuantAtual As Double
Dim dValorAtual As Double
Dim iIndice As Integer

On Error GoTo Erro_SldMesEst_Le_CustoMedio2
    
    'seleciona o registro de SldMesEst relativo ao Ano, FilialEmpresa e Produto passados como parametro
    lErro = Comando_Executar(lComando, "SELECT QuantInicialCusto, ValorInicialCusto, SaldoQuantCusto1, SaldoValorCusto1, SaldoQuantCusto2, SaldoValorCusto2, SaldoQuantCusto3, SaldoValorCusto3, SaldoQuantCusto4, SaldoValorCusto4, SaldoQuantCusto5, SaldoValorCusto5, SaldoQuantCusto6, SaldoValorCusto6, SaldoQuantCusto7, SaldoValorCusto7, SaldoQuantCusto8, SaldoValorCusto8, SaldoQuantCusto9, SaldoValorCusto9, SaldoQuantCusto10, SaldoValorCusto10, SaldoQuantCusto11, SaldoValorCusto11, SaldoQuantCusto12, SaldoValorCusto12 FROM SldMesEst WHERE Ano=? AND FilialEmpresa=? AND Produto=?", _
                              tSldMesEst.dQuantInicialCusto, tSldMesEst.dValorInicialCusto, tSldMesEst.adSaldoQuantCusto(1), tSldMesEst.adSaldoValorCusto(1), tSldMesEst.adSaldoQuantCusto(2), tSldMesEst.adSaldoValorCusto(2), tSldMesEst.adSaldoQuantCusto(3), tSldMesEst.adSaldoValorCusto(3), tSldMesEst.adSaldoQuantCusto(4), tSldMesEst.adSaldoValorCusto(4), tSldMesEst.adSaldoQuantCusto(5), tSldMesEst.adSaldoValorCusto(5), tSldMesEst.adSaldoQuantCusto(6), tSldMesEst.adSaldoValorCusto(6), tSldMesEst.adSaldoQuantCusto(7), tSldMesEst.adSaldoValorCusto(7), tSldMesEst.adSaldoQuantCusto(8), tSldMesEst.adSaldoValorCusto(8), tSldMesEst.adSaldoQuantCusto(9), tSldMesEst.adSaldoValorCusto(9), tSldMesEst.adSaldoQuantCusto(10), tSldMesEst.adSaldoValorCusto(10), _
                              tSldMesEst.adSaldoQuantCusto(11), tSldMesEst.adSaldoValorCusto(11), tSldMesEst.adSaldoQuantCusto(12), tSldMesEst.adSaldoValorCusto(12), iAno, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 36519
    
    'le o SldMesEst
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 36520
    
    dQuantAtual = tSldMesEst.dQuantInicialCusto
    dValorAtual = tSldMesEst.dValorInicialCusto

    For iIndice = 1 To iMes
    
        dQuantAtual = dQuantAtual + tSldMesEst.adSaldoQuantCusto(iIndice)
        dValorAtual = dValorAtual + tSldMesEst.adSaldoValorCusto(iIndice)
        
    Next
    
    If Abs(dQuantAtual) > QTDE_ESTOQUE_DELTA Then
    
        dCustoMedio = dValorAtual / dQuantAtual
        
    Else
    
        dCustoMedio = 0
        
    End If

    SldMesEst_Le_CustoMedio2 = SUCESSO

    Exit Function

Erro_SldMesEst_Le_CustoMedio2:
    
    SldMesEst_Le_CustoMedio2 = Err
    
    Select Case Err
    
        Case 36519, 36520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESEST", Err, iAno, iFilialEmpresa, sProduto)
            
        Case 36573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDMESEST", Err, iAno, iFilialEmpresa, sProduto)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152166)
        
    End Select
    
    Exit Function

End Function


Function OrdemServicoProd_Le(ByVal objOrdemServicoProd As ClassOrdemServicoProd) As Long
'Lê a Ordem de Serviço pelo Código e pela FilialEmpresa

Dim lErro As Long
Dim lComando As Long
Dim tOrdemServicoProd As typeOrdemServicoProd

On Error GoTo Erro_OrdemServicoProd_Le

    If objOrdemServicoProd.iFilialEmpresa = 0 Then objOrdemServicoProd.iFilialEmpresa = giFilialEmpresa
    
    'Abertura de comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 86559

    With tOrdemServicoProd
        
        lErro = Comando_Executar(lComando, "SELECT DataEmissao, PrestadorServico, Status FROM OrdemServicoProd WHERE FilialEmpresa = ? AND Codigo = ?", .dtDataEmissao, .lCodPrestador, .iStatus, objOrdemServicoProd.iFilialEmpresa, objOrdemServicoProd.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 86560

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86561
        
        If lErro = AD_SQL_SEM_DADOS Then gError 86562

        objOrdemServicoProd.lCodPrestador = .lCodPrestador
        objOrdemServicoProd.dtDataEmissao = .dtDataEmissao
        objOrdemServicoProd.iStatus = .iStatus

    End With

    Call Comando_Fechar(lComando)

    OrdemServicoProd_Le = SUCESSO

    Exit Function

Erro_OrdemServicoProd_Le:

    OrdemServicoProd_Le = gErr

    Select Case gErr

        Case 86559
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 86560, 86561
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDEMSERVICOPRODUCAO", gErr)

        Case 86562 'Tratado na Rotina Chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152167)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensOrdemServicoProd_Le(ByVal objOrdemServicoProd As ClassOrdemServicoProd) As Long
'Lê os Itens Ordem de Produção a partir do OPCódigo e da FilialEmpresa
'Retorna os Itens na coleção

Dim lErro As Long
Dim lComando As Long
Dim objItemOSProd As ClassItemOSProd
Dim tItemOSProd As typeItemOSProd

On Error GoTo Erro_ItensOrdemServicoProd_Le

    'Abertura de comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86566

    With tItemOSProd
        
        .sProduto = String(STRING_PRODUTO, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)

        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Quantidade ,QuantidadeProd , Item, Produto, SiglaUM, Status FROM ItensOrdemServicoProd WHERE FilialEmpresa = ? AND Codigo = ? ", _
        .lNumIntDoc, .dQuantidade, .dQuantidadeProd, .iItem, .sProduto, .sSiglaUM, .iStatus, objOrdemServicoProd.iFilialEmpresa, objOrdemServicoProd.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 86567

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86568

        While lErro <> AD_SQL_SEM_DADOS

            Set objItemOSProd = New ClassItemOSProd

            objItemOSProd.dQuantidade = .dQuantidade
            objItemOSProd.dQuantidadeProd = .dQuantidadeProd
            objItemOSProd.iItem = .iItem
            objItemOSProd.lNumIntDoc = .lNumIntDoc
            objItemOSProd.sProduto = .sProduto
            objItemOSProd.sSiglaUM = .sSiglaUM
            objItemOSProd.iStatus = .iStatus

            objOrdemServicoProd.colItens.Add objItemOSProd

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86569

        Wend

    End With

    Call Comando_Fechar(lComando)

    ItensOrdemServicoProd_Le = SUCESSO

    Exit Function

Erro_ItensOrdemServicoProd_Le:

    ItensOrdemServicoProd_Le = gErr

    Select Case gErr

        Case 86566
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 86567, 86568, 86569
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOSPROD", gErr, objOrdemServicoProd.sCodigo, objOrdemServicoProd.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152168)

    End Select

    Call Comando_Fechar(lComando)

End Function


Function ItensOrdemServicoProdCons_Le(ByVal objOrdemServicoProd As ClassOrdemServicoProd) As Long
'Lê os Itens Ordem de Produção a partir do OPCódigo e da FilialEmpresa
'Retorna os Itens na coleção

Dim lErro As Long
Dim lComando As Long
Dim objItemOSProdCons As ClassItemOSProdCons
Dim tItemOSProdCons As typeItemOSProdCons

On Error GoTo Erro_ItensOrdemServicoProdCons_Le

    'Abertura de comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86570

    With tItemOSProdCons
        
        .sProduto = String(STRING_PRODUTO, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)

        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Quantidade , Item, Produto, SiglaUM, Data FROM ItensOrdemServicoProdCons WHERE FilialEmpresa = ? AND Codigo = ? ", _
        .lNumIntDoc, .dQuantidade, .iItem, .sProduto, .sSiglaUM, .dtData, objOrdemServicoProd.iFilialEmpresa, objOrdemServicoProd.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 86571

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86572

        While lErro <> AD_SQL_SEM_DADOS

            Set objItemOSProdCons = New ClassItemOSProdCons

            objItemOSProdCons.dQuantidade = .dQuantidade
            objItemOSProdCons.iItem = .iItem
            objItemOSProdCons.lNumIntDoc = .lNumIntDoc
            objItemOSProdCons.sProduto = .sProduto
            objItemOSProdCons.sSiglaUM = .sSiglaUM
            objItemOSProdCons.dtData = .dtData

            objOrdemServicoProd.colItensCons.Add objItemOSProdCons

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86573

        Wend

    End With

    Call Comando_Fechar(lComando)

    ItensOrdemServicoProdCons_Le = SUCESSO

    Exit Function

Erro_ItensOrdemServicoProdCons_Le:

    ItensOrdemServicoProdCons_Le = gErr

    Select Case gErr

        Case 86570
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 86571, 86572, 86573
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOSPRODCONS", gErr, objOrdemServicoProd.sCodigo, objOrdemServicoProd.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152169)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensOPGrade_Le_NumIntDoc(objItemOPGrade As ClassItemOP) As Long

Dim tItemOPGrade As typeItemOP
Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_ItensOPGrade_Le_NumIntDoc

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 126561

    'Se o elemento é um pai de grade
    tItemOPGrade.sProduto = String(STRING_PRODUTO, 0)
    tItemOPGrade.sVersao = String(STRING_KIT_VERSAO, 0)

    'tenta selecionar item  da OP Grade
    lErro = Comando_Executar(lComando, "SELECT Produto, NumIntItemOP, FilialPedido, CodPedido, Almoxarifado, Quantidade, QuantidadeProd, Versao FROM ItensOPGrade WHERE NumIntDoc = ?", _
            tItemOPGrade.sProduto, tItemOPGrade.lNumIntItemOP, tItemOPGrade.iFilialPedido, tItemOPGrade.lCodPedido, tItemOPGrade.iAlmoxarifado, tItemOPGrade.dQuantidade, tItemOPGrade.dQuantidadeProd, tItemOPGrade.sVersao, objItemOPGrade.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 126562

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126563

    If lErro = AD_SQL_SEM_DADOS Then gError 126564

    objItemOPGrade.sProduto = tItemOPGrade.sProduto
    objItemOPGrade.lNumItemOP = tItemOPGrade.lNumIntItemOP
    objItemOPGrade.iFilialPedido = tItemOPGrade.iFilialPedido
    objItemOPGrade.lCodPedido = tItemOPGrade.lCodPedido
    objItemOPGrade.iAlmoxarifado = tItemOPGrade.iAlmoxarifado
    objItemOPGrade.dQuantidade = tItemOPGrade.dQuantidade
    objItemOPGrade.dQuantidadeProd = tItemOPGrade.dQuantidadeProd
    objItemOPGrade.sVersao = tItemOPGrade.sVersao

    Call Comando_Fechar(lComando)

    ItensOPGrade_Le_NumIntDoc = SUCESSO

    Exit Function

Erro_ItensOPGrade_Le_NumIntDoc:

    ItensOPGrade_Le_NumIntDoc = gErr

    Select Case gErr

        Case 126561
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126562, 126563
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOPGRADE1", gErr, objItemOPGrade.lNumIntDoc)

        Case 126564

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152170)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensOrdemProducao_Le_NumItemOPProduto(objItemOPGrade As ClassItemOP) As Long

Dim tItemOPGrade As typeItemOP
Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_ItensOrdemProducao_Le_NumItemOPProduto

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 126565

    'Se o elemento é um pai de grade
    tItemOPGrade.sProduto = String(STRING_PRODUTO, 0)
    tItemOPGrade.sVersao = String(STRING_KIT_VERSAO, 0)
    
    '##############################################################
    'INSERIDO POR WAGNER
    tItemOPGrade.sJustificativaBloqueio = String(STRING_MAX_NOME_ARQUIVO, 0)
    '##############################################################

    'tenta selecionar item  da OP Grade
    '######################################################
    'ALTERADO POR WAGNER
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialPedido, CodPedido, Almoxarifado, Quantidade, QuantidadeProd, Versao, NumIntDocPai, JustificativaBloqueio FROM ItensOrdemProducao WHERE NumIntItemOPGrade = ? And Produto = ?", _
            tItemOPGrade.lNumIntDoc, tItemOPGrade.iFilialPedido, tItemOPGrade.lCodPedido, tItemOPGrade.iAlmoxarifado, tItemOPGrade.dQuantidade, tItemOPGrade.dQuantidadeProd, tItemOPGrade.sVersao, tItemOPGrade.lNumIntDocPai, tItemOPGrade.sJustificativaBloqueio, objItemOPGrade.lNumItemOP, objItemOPGrade.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 126566
    '######################################################

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126567

    If lErro = AD_SQL_SEM_DADOS Then

        'tenta selecionar item  da OP Grade
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, FilialPedido, CodPedido, Almoxarifado, Quantidade, QuantidadeProd, Versao FROM ItensOrdemProducaoBaixadas WHERE NumIntItemOPGrade = ? And Produto = ?", _
                tItemOPGrade.lNumIntDoc, tItemOPGrade.iFilialPedido, tItemOPGrade.lCodPedido, tItemOPGrade.iAlmoxarifado, tItemOPGrade.dQuantidade, tItemOPGrade.dQuantidadeProd, tItemOPGrade.sVersao, objItemOPGrade.lNumItemOP, objItemOPGrade.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 126707
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126708
    
        If lErro = AD_SQL_SEM_DADOS Then gError 126568

    End If

    objItemOPGrade.sProduto = objItemOPGrade.sProduto
    objItemOPGrade.lNumItemOP = objItemOPGrade.lNumItemOP
    objItemOPGrade.iFilialPedido = tItemOPGrade.iFilialPedido
    objItemOPGrade.lCodPedido = tItemOPGrade.lCodPedido
    objItemOPGrade.iAlmoxarifado = tItemOPGrade.iAlmoxarifado
    objItemOPGrade.dQuantidade = tItemOPGrade.dQuantidade
    objItemOPGrade.dQuantidadeProd = tItemOPGrade.dQuantidadeProd
    objItemOPGrade.sVersao = tItemOPGrade.sVersao

    '###################################################
    'INSERIDO POR WAGNER
    objItemOPGrade.lNumIntDocPai = tItemOPGrade.lNumIntDocPai
    objItemOPGrade.sJustificativaBloqueio = tItemOPGrade.sJustificativaBloqueio
    '###################################################

    Call Comando_Fechar(lComando)

    ItensOrdemProducao_Le_NumItemOPProduto = SUCESSO

    Exit Function

Erro_ItensOrdemProducao_Le_NumItemOPProduto:

    ItensOrdemProducao_Le_NumItemOPProduto = gErr

    Select Case gErr

        Case 126565
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126566, 126567
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO4", gErr, objItemOPGrade.lNumItemOP, objItemOPGrade.sProduto)

        Case 126568

        Case 126707, 126708
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAOBAIXADAS1", gErr, objItemOPGrade.lNumItemOP, objItemOPGrade.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152171)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensOPGrade_Le(ByVal colItensOPGrade As Collection, ByVal lNumIntItemOP As Long) As Long

Dim tItemOPGrade As typeItemOP
Dim lComando As Long
Dim lErro As Long
Dim objItemOPGrade As ClassItemOP

On Error GoTo Erro_ItensOPGrade_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 126580

    'Se o elemento é um pai de grade
    tItemOPGrade.sProduto = String(STRING_PRODUTO, 0)
    tItemOPGrade.sVersao = String(STRING_KIT_VERSAO, 0)

    'tenta selecionar item  da OP Grade
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Produto, NumIntItemOP, FilialPedido, CodPedido, Almoxarifado, Quantidade, QuantidadeProd, Versao FROM ItensOPGrade WHERE NumIntItemOP = ?", _
            tItemOPGrade.lNumIntDoc, tItemOPGrade.sProduto, tItemOPGrade.lNumIntItemOP, tItemOPGrade.iFilialPedido, tItemOPGrade.lCodPedido, tItemOPGrade.iAlmoxarifado, tItemOPGrade.dQuantidade, tItemOPGrade.dQuantidadeProd, tItemOPGrade.sVersao, lNumIntItemOP)
    If lErro <> AD_SQL_SUCESSO Then gError 126581

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126582

    Do While lErro = AD_SQL_SUCESSO

        Set objItemOPGrade = New ClassItemOP

        objItemOPGrade.sProduto = tItemOPGrade.sProduto
        objItemOPGrade.lNumIntDoc = tItemOPGrade.lNumIntDoc
        objItemOPGrade.lNumItemOP = lNumIntItemOP
        objItemOPGrade.iFilialPedido = tItemOPGrade.iFilialPedido
        objItemOPGrade.lCodPedido = tItemOPGrade.lCodPedido
        objItemOPGrade.iAlmoxarifado = tItemOPGrade.iAlmoxarifado
        objItemOPGrade.dQuantidade = tItemOPGrade.dQuantidade
        objItemOPGrade.dQuantidadeProd = tItemOPGrade.dQuantidadeProd
        objItemOPGrade.sVersao = tItemOPGrade.sVersao

        colItensOPGrade.Add objItemOPGrade

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126583

    Loop

    Call Comando_Fechar(lComando)

    ItensOPGrade_Le = SUCESSO

    Exit Function

Erro_ItensOPGrade_Le:

    ItensOPGrade_Le = gErr

    Select Case gErr

        Case 126580
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126581, 126582, 126583
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOPGRADE", gErr, lNumIntItemOP)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152172)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function MovimentoEstoqueGrade_Le(ByVal objMovEstoque As ClassMovEstoque) As Long
'Le os Ítens do MovimentodeEstoqueGrade relacionados a FilialEmpresa, Codigo e TipoMov

Dim lErro As Long
Dim lComando As Long
Dim tItemMovEstoque As typeItemMovEstoque
Dim sComando_SQL As String

On Error GoTo Erro_MovimentoEstoqueGrade_Le


    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 117628

    With tItemMovEstoque

        .sCcl = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sProduto = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sProdutoDesc = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sSiglaUM = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilEst = String(STRING_BUFFER_MAX_TEXTO, 0)
        .sContaContabilAplic = String(STRING_BUFFER_MAX_TEXTO, 0)

        If objMovEstoque.iTipoMov = MOV_EST_PRODUCAO Or objMovEstoque.iTipoMov = MOV_EST_REQ_PRODUCAO Or objMovEstoque.iTipoMov = MOV_EST_REQ_PRODUCAO_BENEF3 Or objMovEstoque.iTipoMov = MOV_EST_PRODUCAO_BENEF3 Then

            sComando_SQL = "SELECT X.ContaContabilEst, X.ContaContabilAplic, X.Data, X.NumIntDoc, X.FilialEmpresa, X.Codigo, X.Produto, Produtos.Descricao, X.SiglaUM, X.TipoMov, X.NumIntDocOrigem, X.Ccl, X.NumIntDocEst, X.HorasMaquina, X.DataInicioProducao FROM MovimentoEstoqueGrade As X, Produtos WHERE Produtos.Codigo = X.Produto AND X.FilialEmpresa = ? AND X.Codigo = ? AND X.NumIntDocEst = 0 AND (X.TipoMov = ? Or X.TipoMov = ?) ORDER BY X.NumIntDoc"

            If objMovEstoque.iTipoMov = MOV_EST_PRODUCAO Then

                lErro = Comando_Executar(lComando, sComando_SQL, .sContaContabilEst, .sContaContabilAplic, .dtData, .lNumIntDoc, .iFilialEmpresa, .lCodigo, .sProduto, .sProdutoDesc, .sSiglaUM, .iTipoMov, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .lHorasMaquina, .dtDataInicioProducao, objMovEstoque.iFilialEmpresa, objMovEstoque.lCodigo, MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3)
                If lErro <> AD_SQL_SUCESSO Then gError 117629

            Else

                lErro = Comando_Executar(lComando, sComando_SQL, .sContaContabilEst, .sContaContabilAplic, .dtData, .lNumIntDoc, .iFilialEmpresa, .lCodigo, .sProduto, .sProdutoDesc, .sSiglaUM, .iTipoMov, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .lHorasMaquina, .dtDataInicioProducao, objMovEstoque.iFilialEmpresa, objMovEstoque.lCodigo, MOV_EST_REQ_PRODUCAO, MOV_EST_REQ_PRODUCAO_BENEF3)
                If lErro <> AD_SQL_SUCESSO Then gError 117630

            End If

            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 117631

            Do While lErro <> AD_SQL_SEM_DADOS

                objMovEstoque.colItens.Add .lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, .lHorasMaquina, Nothing, Nothing, .dtDataInicioProducao

                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 117632

            Loop

        Else

            '#############################################################
            'Alterado por Wagner - Correção de Erro REPLICAR_ACERTO
            lErro = Comando_Executar(lComando, "SELECT X.FilialEmpresa, X.ContaContabilEst, X.ContaContabilAplic, X.Data, X.NumIntDoc, X.Produto, Descricao, X.SiglaUM, X.NumIntDocOrigem, X.Ccl, X.NumIntDocEst, X.TipoMov FROM MovimentoEstoqueGrade As X, Produtos WHERE Produtos.Codigo = X.Produto AND X.FilialEmpresa = ? AND X.Codigo = ? AND X.NumIntDocEst = 0 And X.TipoMov =? ORDER BY NumIntDoc", _
            .iFilialEmpresa, .sContaContabilEst, .sContaContabilAplic, .dtData, .lNumIntDoc, .sProduto, .sProdutoDesc, .sSiglaUM, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .iTipoMov, objMovEstoque.iFilialEmpresa, objMovEstoque.lCodigo, objMovEstoque.iTipoMov)
            If lErro <> AD_SQL_SUCESSO Then gError 117633
            '#############################################################

            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 117634

            Do While lErro <> AD_SQL_SEM_DADOS

                objMovEstoque.colItens.Add .lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, 0, Nothing, Nothing, DATA_NULA

                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 117635

            Loop

        End If

    End With

    Call Comando_Fechar(lComando)

    MovimentoEstoqueGrade_Le = SUCESSO

    Exit Function

Erro_MovimentoEstoqueGrade_Le:

    MovimentoEstoqueGrade_Le = gErr

    Select Case gErr

        Case 117628
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 117629, 117630, 117631, 117632, 117633, 117634
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUEGRADE", gErr, objMovEstoque.iFilialEmpresa, objMovEstoque.lCodigo, objMovEstoque.iTipoMov)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152173)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MovimentoEstoqueGrade_Le_NumIntDoc(ByVal lComando As Long, ByVal objItemMovEst As ClassItemMovEstoque) As Long
'Carrega os dados do movimento com o NumIntDoc passado como parametro
    
Dim lErro As Long
Dim tItemMovEstoque As typeItemMovEstoque

On Error GoTo Erro_MovimentoEstoqueGrade_Le_NumIntDoc

    tItemMovEstoque.sProduto = String(STRING_PRODUTO, 0)
    tItemMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemMovEstoque.sCcl = String(STRING_CCL, 0)
    tItemMovEstoque.sOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
    tItemMovEstoque.sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)
    tItemMovEstoque.sContaContabilAplic = String(STRING_CONTA, 0)
    tItemMovEstoque.sContaContabilEst = String(STRING_CONTA, 0)
    
    'le os movimento de estoque referente ao numintdoc  passado como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT FilialEmpresa, Codigo, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao, Hora FROM MovimentoEstoqueGrade WHERE NumIntDoc = ?", 0, _
    tItemMovEstoque.iFilialEmpresa, tItemMovEstoque.lCodigo, tItemMovEstoque.dCusto, tItemMovEstoque.iApropriacao, tItemMovEstoque.sProduto, tItemMovEstoque.sSiglaUM, tItemMovEstoque.dQuantidade, tItemMovEstoque.iAlmoxarifado, tItemMovEstoque.iTipoMov, tItemMovEstoque.lNumIntDocOrigem, tItemMovEstoque.iTipoNumIntDocOrigem, tItemMovEstoque.dtData, tItemMovEstoque.sCcl, tItemMovEstoque.lNumIntDocEst, tItemMovEstoque.lCliente, tItemMovEstoque.lFornecedor, tItemMovEstoque.sOPCodigo, tItemMovEstoque.sDocOrigem, tItemMovEstoque.sContaContabilEst, tItemMovEstoque.sContaContabilAplic, tItemMovEstoque.lHorasMaquina, tItemMovEstoque.dtDataInicioProducao, tItemMovEstoque.dHora, objItemMovEst.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 117639

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 117640

    'se o movimento não está cadastrado
    If lErro = AD_SQL_SEM_DADOS Then gError 117641

    objItemMovEst.iFilialEmpresa = tItemMovEstoque.iFilialEmpresa
    objItemMovEst.lCodigo = tItemMovEstoque.lCodigo
    objItemMovEst.dCusto = tItemMovEstoque.dCusto
    objItemMovEst.iApropriacao = tItemMovEstoque.iApropriacao
    objItemMovEst.sProduto = tItemMovEstoque.sProduto
    objItemMovEst.sSiglaUM = tItemMovEstoque.sSiglaUM
    objItemMovEst.dQuantidade = tItemMovEstoque.dQuantidade
    objItemMovEst.iAlmoxarifado = tItemMovEstoque.iAlmoxarifado
    objItemMovEst.iTipoMov = tItemMovEstoque.iTipoMov
    objItemMovEst.lNumIntDocOrigem = tItemMovEstoque.lNumIntDocOrigem
    objItemMovEst.iTipoNumIntDocOrigem = tItemMovEstoque.iTipoNumIntDocOrigem
    objItemMovEst.dtData = tItemMovEstoque.dtData
    objItemMovEst.sCcl = tItemMovEstoque.sCcl
    objItemMovEst.lNumIntDocEst = tItemMovEstoque.lNumIntDocEst
    objItemMovEst.lCliente = tItemMovEstoque.lCliente
    objItemMovEst.lFornecedor = tItemMovEstoque.lFornecedor
    objItemMovEst.sOPCodigo = tItemMovEstoque.sOPCodigo
    objItemMovEst.sDocOrigem = tItemMovEstoque.sDocOrigem
    objItemMovEst.sContaContabilEst = tItemMovEstoque.sContaContabilEst
    objItemMovEst.sContaContabilAplic = tItemMovEstoque.sContaContabilAplic
    objItemMovEst.lHorasMaquina = tItemMovEstoque.lHorasMaquina
    objItemMovEst.dtDataInicioProducao = tItemMovEstoque.dtDataInicioProducao
    objItemMovEst.dtHora = tItemMovEstoque.dHora
        
    MovimentoEstoqueGrade_Le_NumIntDoc = SUCESSO
    
    Exit Function

Erro_MovimentoEstoqueGrade_Le_NumIntDoc:

    MovimentoEstoqueGrade_Le_NumIntDoc = gErr

    Select Case gErr

        Case 117639, 117640
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUEGRADE1", gErr, objItemMovEst.lNumIntDoc)

        Case 117641
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152174)

    End Select

    Exit Function

End Function

Function Customizacoes_MovEstoque(ByVal objMovEst As ClassMovEstoque) As Long

Dim lErro As Long
Dim lComando As Long
Dim tMovEstoque As typeMovEst
Dim dHora As Double

On Error GoTo Erro_Customizacoes_MovEstoque

    Customizacoes_MovEstoque = SUCESSO

    Exit Function

Erro_Customizacoes_MovEstoque:

    Customizacoes_MovEstoque = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152175)

    End Select

    Exit Function

End Function

Public Function TP_Almoxarifado_Filial_Le_Aux(objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado, Optional iCria As Integer = 1) As Long
'Lê Almoxarifado de giFilialEmpresa com Código ou NomeRed em objAlmoxarifadoMaskEdBox.Text
'Devolve em objAlmoxarifado. Coloca NomeReduzido em .Text

Dim sAlmoxarifado As String
Dim Almoxarifado As Object
Dim eTipoAlmoxarifado As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iFilialEmpresa1 As Integer

On Error GoTo TP_Almoxarifado_Filial_Le_Aux

    iFilialEmpresa1 = giFilialEmpresa
    lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa1)
    If lErro <> SUCESSO Then Error 30899
                
    Set Almoxarifado = objAlmoxarifadoMaskEdBox
    sAlmoxarifado = Trim(Almoxarifado.Text)
    eTipoAlmoxarifado = Tipo_Campo(sAlmoxarifado)

    Select Case eTipoAlmoxarifado

    Case TIPO_STRING

        objAlmoxarifado.sNomeReduzido = sAlmoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then Error 25131
        If lErro <> SUCESSO Then Error 25136
        If objAlmoxarifado.iFilialEmpresa <> iFilialEmpresa1 Then Error 25115
        
        Almoxarifado.Text = sAlmoxarifado

    Case TIPO_CODIGO

        objAlmoxarifado.iCodigo = CInt(sAlmoxarifado)
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 25142
        If lErro <> SUCESSO Then Error 25143
        If objAlmoxarifado.iFilialEmpresa <> iFilialEmpresa1 Then Error 25118

        Almoxarifado.Text = objAlmoxarifado.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 25144

    Case TIPO_DECIMAL

        Error 25145

    Case TIPO_NAO_POSITIVO

        Error 25146

    End Select

    TP_Almoxarifado_Filial_Le_Aux = SUCESSO

    Exit Function

TP_Almoxarifado_Filial_Le_Aux:

    TP_Almoxarifado_Filial_Le_Aux = Err

    Select Case Err
        
        Case 25115, 25118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_OUTRA_FILIAL", Err, sAlmoxarifado)
            Almoxarifado.SetFocus
        
        Case 25131, 25142, 30899 'Tratados nas rotinas chamadas
        
            Almoxarifado.SetFocus

        Case 25136  'Almoxarifado com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Almoxarifado não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", objAlmoxarifado.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Almoxarifadoes
                    lErro = Chama_Tela("Almoxarifado", objAlmoxarifado)
                Else
                    'Segura o foco
                    Almoxarifado.SetFocus
                End If
            Else
                'Segura o foco
                Almoxarifado.SetFocus
            End If
        
        Case 25143  'Almoxarifado com Código não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Almoxarifado não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", objAlmoxarifado.iCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Almoxarifados
                    lErro = Chama_Tela("Almoxarifado", objAlmoxarifado)
                Else
                    'Segura o foco
                    Almoxarifado.SetFocus
                End If
            Else
                'Segura o foco
                Almoxarifado.SetFocus
            End If
        
        Case 25144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sAlmoxarifado)
            Almoxarifado.SetFocus

        Case 25145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sAlmoxarifado)
            Almoxarifado.SetFocus

        Case 25146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sAlmoxarifado)
            Almoxarifado.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152176)

    End Select

End Function

Public Function TP_Almoxarifado_Filial_Grid_Aux(objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado) As Long
'Lê Almoxarifado de giFilialEmpresa com Código ou NomeRed em objAlmoxarifadoMaskEdBox.Text
'Devolve em objAlmoxarifado. Coloca NomeReduzido em .Text
'Para ser usada em Grids. Não segura foco nem chama tela.

Dim sAlmoxarifado As String
Dim Almoxarifado As Object
Dim eTipoAlmoxarifado As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iFilialEmpresa1 As Integer

On Error GoTo TP_Almoxarifado_Filial_Grid_Aux

    iFilialEmpresa1 = giFilialEmpresa
    lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa1)
    If lErro <> SUCESSO Then Error 30899
                
    Set Almoxarifado = objAlmoxarifadoMaskEdBox
    sAlmoxarifado = Trim(Almoxarifado.Text)
    eTipoAlmoxarifado = Tipo_Campo(sAlmoxarifado)

    Select Case eTipoAlmoxarifado

    Case TIPO_STRING

        objAlmoxarifado.sNomeReduzido = sAlmoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then Error 25147
        If lErro <> SUCESSO Then Error 25148
        If objAlmoxarifado.iFilialEmpresa <> iFilialEmpresa1 Then Error 25149
        
        Almoxarifado.Text = sAlmoxarifado

    Case TIPO_CODIGO

        objAlmoxarifado.iCodigo = CInt(sAlmoxarifado)
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 25150
        If lErro <> SUCESSO Then Error 25151
        If objAlmoxarifado.iFilialEmpresa <> iFilialEmpresa1 Then Error 25152

        Almoxarifado.Text = objAlmoxarifado.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 25153

    Case TIPO_DECIMAL

        Error 25154

    Case TIPO_NAO_POSITIVO

        Error 25155

    End Select

    TP_Almoxarifado_Filial_Grid_Aux = SUCESSO

    Exit Function

TP_Almoxarifado_Filial_Grid_Aux:

    TP_Almoxarifado_Filial_Grid_Aux = Err

    Select Case Err
        
        Case 25147, 25150, 30899 'Tratados nas rotinas chamadas
        
        Case 25148  'Almoxarifado com NomeReduzido não cadastrado
        
        Case 25149, 25152
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_OUTRA_FILIAL", Err, sAlmoxarifado)
        
        Case 25151  'Almoxarifado com Código não cadastrado
         
        Case 25153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sAlmoxarifado)

        Case 25154
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sAlmoxarifado)

        Case 25155
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sAlmoxarifado)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152177)

    End Select

End Function

Public Function TP_Almoxarifado_Filial_Produto_Grid_Aux(sProduto As String, objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado) As Long
'Lê Almoxarifado de giFilialEmpresa com Código ou NomeRed em objAlmoxarifadoMaskEdBox.Text
'Critica se Almoxarifado tem o produto com código sProduto cadastrado.
'Devolve em objAlmoxarifado. Coloca NomeReduzido em .Text
'Para ser usada em Grids. Não segura foco nem chama tela.

Dim sAlmoxarifado As String
Dim Almoxarifado As Object
Dim eTipoAlmoxarifado As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim colAlmoxarifado As New Collection
Dim vAlmoxarifado As Variant
Dim iTemProduto As Integer
Dim iFilialEmpresa1 As Integer

On Error GoTo TP_Almoxarifado_Filial_Produto_Grid_Aux

    iFilialEmpresa1 = giFilialEmpresa
    lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa1)
    If lErro <> SUCESSO Then Error 30899
                
    Set Almoxarifado = objAlmoxarifadoMaskEdBox
    sAlmoxarifado = Trim(Almoxarifado.Text)
    eTipoAlmoxarifado = Tipo_Campo(sAlmoxarifado)

    Select Case eTipoAlmoxarifado

    Case TIPO_STRING

        objAlmoxarifado.sNomeReduzido = sAlmoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then Error 25156
        If lErro <> SUCESSO Then Error 25157
        If objAlmoxarifado.iFilialEmpresa <> iFilialEmpresa1 Then Error 25158
        
        'Lê Almoxarifados que tem sProduto cadastrado
        lErro = CF("AlmoxarifadosProduto_Le", sProduto, colAlmoxarifado)
        If lErro <> SUCESSO Then Error 25159
       
        'Verifica se objAlmoxarifado tem o Produto
        iTemProduto = 0
        For Each vAlmoxarifado In colAlmoxarifado
        
            If vAlmoxarifado = objAlmoxarifado.iCodigo Then
                iTemProduto = 1
                Exit For
            End If
        
        Next
        
        'Se não tem, erro
        If iTemProduto = 0 Then Error 25160
        
        Almoxarifado.Text = sAlmoxarifado

    Case TIPO_CODIGO

        objAlmoxarifado.iCodigo = CInt(sAlmoxarifado)
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 25161
        If lErro <> SUCESSO Then Error 25162
        If objAlmoxarifado.iFilialEmpresa <> iFilialEmpresa1 Then Error 25163

        'Lê Almoxarifados que tem sProduto cadastrado
        lErro = CF("AlmoxarifadosProduto_Le", sProduto, colAlmoxarifado)
        If lErro <> SUCESSO Then Error 25164
       
        'Verifica se objAlmoxarifado tem o Produto
        iTemProduto = 0
        For Each vAlmoxarifado In colAlmoxarifado
        
            If vAlmoxarifado = objAlmoxarifado.iCodigo Then
                iTemProduto = 1
                Exit For
            End If
        
        Next
        
        'Se não tem, erro
        If iTemProduto = 0 Then Error 25165
        
        Almoxarifado.Text = objAlmoxarifado.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 25166

    Case TIPO_DECIMAL

        Error 25167

    Case TIPO_NAO_POSITIVO

        Error 25168

    End Select

    TP_Almoxarifado_Filial_Produto_Grid_Aux = SUCESSO

    Exit Function

TP_Almoxarifado_Filial_Produto_Grid_Aux:

    TP_Almoxarifado_Filial_Produto_Grid_Aux = Err

    Select Case Err
        
        Case 25156, 25161, 30899 'Tratados nas rotinas chamadas
        
        Case 25157  'Almoxarifado com NomeReduzido não cadastrado
       
        Case 25158, 25163
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_OUTRA_FILIAL", Err, sAlmoxarifado)
        
        Case 25159, 25164 'Tratados na rotina chamada
        
        Case 25160, 25165
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_TEM_PRODUTO", Err, objAlmoxarifado.iCodigo, sProduto)

        Case 25162  'Almoxarifado com Código não cadastrado
         
        Case 25166
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sAlmoxarifado)

        Case 25167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sAlmoxarifado)

        Case 25168
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sAlmoxarifado)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152178)

    End Select

End Function

Public Function TP_Almoxarifado_Filial_Le_ComCodigo_Aux(objAlmoxarifadoMaskEdBox As Object, objAlmoxarifado As ClassAlmoxarifado, Optional iCria As Integer = 0) As Long
'Lê Almoxarifado com Código, NomeRed ou Codigo-NomeRed em objAlmoxarifadoMaskEdBox.Text
'Devolve em objAlmoxarifado. Coloca Codigo-NomeReduzido em .Text

Dim sAlmoxarifado As String, iCodigo As Integer
Dim Almoxarifado As Object
Dim eTipoAlmoxarifado As enumTipo
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iFilialEmpresa1 As Integer

On Error GoTo TP_Almoxarifado_Filial_Le_ComCodigo_Aux

    iFilialEmpresa1 = giFilialEmpresa
    lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa1)
    If lErro <> SUCESSO Then Error 30899
                
    Set Almoxarifado = objAlmoxarifadoMaskEdBox
    sAlmoxarifado = Trim(Almoxarifado.Text)
    
    'incluido p/tratar codigo-nomereduzido
    iCodigo = Codigo_Extrai(sAlmoxarifado)
    If iCodigo <> 0 Then
        sAlmoxarifado = CStr(iCodigo)
    End If
    
    eTipoAlmoxarifado = Tipo_Campo(sAlmoxarifado)

    Select Case eTipoAlmoxarifado

    Case TIPO_STRING

        objAlmoxarifado.sNomeReduzido = sAlmoxarifado
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then Error 25169
        If lErro <> SUCESSO Then Error 25170
        If objAlmoxarifado.iFilialEmpresa <> iFilialEmpresa1 Then Error 25171
       
        Almoxarifado.Text = CStr(objAlmoxarifado.iCodigo) & SEPARADOR & sAlmoxarifado

    Case TIPO_CODIGO

        objAlmoxarifado.iCodigo = iCodigo
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 25172
        If lErro <> SUCESSO Then Error 25173
        If objAlmoxarifado.iFilialEmpresa <> iFilialEmpresa1 Then Error 25174

        Almoxarifado.Text = sAlmoxarifado & SEPARADOR & objAlmoxarifado.sNomeReduzido
    
    Case TIPO_OVERFLOW

        Error 25175

    Case TIPO_DECIMAL

        Error 25176

    Case TIPO_NAO_POSITIVO

        Error 25177

    End Select

    TP_Almoxarifado_Filial_Le_ComCodigo_Aux = SUCESSO

    Exit Function

TP_Almoxarifado_Filial_Le_ComCodigo_Aux:

    TP_Almoxarifado_Filial_Le_ComCodigo_Aux = Err

    Select Case Err
        
        Case 25169, 25172, 30899 'Tratados nas rotinas chamadas
        
            Almoxarifado.SetFocus

        Case 25170  'Almoxarifado com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Almoxarifado não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", objAlmoxarifado.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Almoxarifadoes
                    lErro = Chama_Tela("Almoxarifado", objAlmoxarifado)
                Else
                    'Segura o foco
                    Almoxarifado.SetFocus
                End If
            Else
                
                lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO1", Err, objAlmoxarifado.sNomeReduzido)
                
                'Segura o foco
                Almoxarifado.SetFocus
            End If
        
        Case 25171, 25174
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_OUTRA_FILIAL", Err, sAlmoxarifado)
                
        Case 25173  'Almoxarifado com Código não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Almoxarifado não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", objAlmoxarifado.iCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Almoxarifados
                    lErro = Chama_Tela("Almoxarifado", objAlmoxarifado)
                Else
                    'Segura o foco
                    Almoxarifado.SetFocus
                End If
            Else
                
                lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO", Err, objAlmoxarifado.iCodigo)
                
                'Segura o foco
                Almoxarifado.SetFocus
            End If
        
        Case 25175
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTEIRO_OVERFLOW", Err, sAlmoxarifado)
            Almoxarifado.SetFocus

        Case 25176
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_INTEIRO", Err, sAlmoxarifado)
            Almoxarifado.SetFocus

        Case 25177
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", Err, sAlmoxarifado)
            Almoxarifado.SetFocus
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152179)

    End Select

End Function

Function TabelaPrecoItem_Le1(ByVal objTabelaPrecoItem As ClassTabelaPrecoItem, Optional dtDataLeitura As Date = DATA_NULA) As Long
'Lê a tabela de preço a partir da chave fornecida em objTabelaPreco. Retorna os dados em objTabelaPreco
'Funcao criada pois existe versao customizada

Dim lErro As Long

On Error GoTo Erro_TabelaPrecoItem_Le1

    lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
    If lErro <> SUCESSO And lErro <> 28014 Then gError 126786

    If lErro = 28014 Then gError 126784

    TabelaPrecoItem_Le1 = SUCESSO

    Exit Function

Erro_TabelaPrecoItem_Le1:

    TabelaPrecoItem_Le1 = gErr

    Select Case gErr

        Case 126784, 126786

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152180)

    End Select

    Exit Function

End Function

Function Customizacao_FilialEmpresa(iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_Customizacao_FilialEmpresa

    Customizacao_FilialEmpresa = SUCESSO

    Exit Function

Erro_Customizacao_FilialEmpresa:

    Customizacao_FilialEmpresa = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152181)

    End Select

    Exit Function

End Function




'###################################################
Function ImportTabelaPrecoItem_Le(ByVal colTabelaPrecoItem As Collection) As Long
'Rotina de Importação da Tabela de Preços
'Inserido por Wagner 19/10/2004

Dim lErro As Long
Dim lComando As Long
Dim tTabelaPrecoItem As typeTabelaPrecoItem
Dim objTabelaPrecoItem As ClassTabelaPrecoItem
Dim sProdutoConvertido As String, iPreenchido As Integer

On Error GoTo Erro_ImportTabelaPrecoItem_Le

    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 131020

    'inicializa o buffer de recebimento
    tTabelaPrecoItem.sCodProduto = String(STRING_PRODUTO, 0)
    tTabelaPrecoItem.sObservacao = String(STRING_TABELA_OBSERVACAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Produto, Preco, FilialEmpresa, CodTabela, Observacao, DataVigencia FROM ImportTabelaPrecoItem", tTabelaPrecoItem.sCodProduto, tTabelaPrecoItem.dPreco, tTabelaPrecoItem.iFilialEmpresa, tTabelaPrecoItem.iCodTabela, tTabelaPrecoItem.sObservacao, tTabelaPrecoItem.dtDataVigencia)
    If lErro <> AD_SQL_SUCESSO Then gError 131021

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131022
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objTabelaPrecoItem = New ClassTabelaPrecoItem
    
        lErro = CF("Produto_Formata", tTabelaPrecoItem.sCodProduto, sProdutoConvertido, iPreenchido)
        If lErro <> SUCESSO Then gError 27682
    
        objTabelaPrecoItem.sCodProduto = sProdutoConvertido
        objTabelaPrecoItem.dPreco = tTabelaPrecoItem.dPreco
        objTabelaPrecoItem.iFilialEmpresa = tTabelaPrecoItem.iFilialEmpresa
        objTabelaPrecoItem.iCodTabela = tTabelaPrecoItem.iCodTabela
        objTabelaPrecoItem.sObservacao = tTabelaPrecoItem.sObservacao
        objTabelaPrecoItem.dtDataVigencia = tTabelaPrecoItem.dtDataVigencia
        
        colTabelaPrecoItem.Add objTabelaPrecoItem
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 131023
    
    Loop
    
    Call Comando_Fechar(lComando)

    ImportTabelaPrecoItem_Le = SUCESSO

    Exit Function

Erro_ImportTabelaPrecoItem_Le:

    ImportTabelaPrecoItem_Le = gErr

    Select Case gErr
    
        Case 131020
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 131021 To 131023
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTTABELAPRECO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152182)

    End Select
    
    Call Comando_Fechar(lComando)

End Function

'Inserido por Wagner 10/12/2004
Function Kit_Valida_Quantidade(ByVal objKit As ClassKit) As Long

Dim lErro As Long, sUMRaiz As String
Dim objProdutoKit As ClassProdutoKit
Dim dFator As Double
Dim dQtdProdutoRaiz As Double
Dim dQtdTotalSemRaiz As Double
Dim iClasseProdutoRaiz As Integer
Dim bClasseDiferente As Boolean
Dim objProduto As ClassProduto
Dim vbMsgResult As VbMsgBoxResult

On Error GoTo Erro_Kit_Valida_Quantidade

    dQtdProdutoRaiz = 0
    dQtdTotalSemRaiz = 0
    bClasseDiferente = False

    For Each objProdutoKit In objKit.colComponentes
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objProdutoKit.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 131681
        If lErro = 28030 Then gError 131682
    
        If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
            
            sUMRaiz = objProdutoKit.sUnidadeMed
            dQtdProdutoRaiz = objProdutoKit.dQuantidade
            iClasseProdutoRaiz = objProduto.iClasseUM
        
        Else
            
            If objProdutoKit.iNivel = (KIT_NIVEL_RAIZ + 1) Then
            
                If iClasseProdutoRaiz <> objProduto.iClasseUM Then
                    bClasseDiferente = True
                    Exit For
                End If
                
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, sUMRaiz, dFator)
                If lErro <> SUCESSO Then gError 131684
    
                dQtdTotalSemRaiz = dQtdTotalSemRaiz + (objProdutoKit.dQuantidade * dFator)
                
            End If
        
        End If
    
    Next
    
    If Not bClasseDiferente Then
    
        If dQtdProdutoRaiz - dQtdTotalSemRaiz > QTDE_ESTOQUE_DELTA Then
        
            vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_QTD_RAIZ_MAIOR")
            
            If vbMsgResult = vbNo Then gError 131683
        
        End If
    
    End If
        
    Exit Function

Erro_Kit_Valida_Quantidade:

    Kit_Valida_Quantidade = gErr

    Select Case gErr
    
        Case 131681
        
        Case 131682
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
    
        Case 131683, 131684
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152183)

    End Select
    
    Exit Function

End Function

Function Kit_Valida(ByVal objKit As ClassKit) As Long

Dim lErro As Long, sUMBase As String
Dim objProdutoKit As ClassProdutoKit
Dim dFator As Double
Dim dQtdBase As Double
Dim dQtd As Double
Dim iClasseUMBase As Integer
Dim objProduto As ClassProduto
Dim sProdutoBase As String
Dim sReferenciaCircular As String, sCaminho As String
Dim iNivel As Integer

On Error GoTo Erro_Kit_Valida

    'Verifica recursividade direta
    For Each objProdutoKit In objKit.colComponentes
    
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objProdutoKit.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
        If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
            sUMBase = objProdutoKit.sUnidadeMed
            dQtdBase = objProdutoKit.dQuantidade
            iClasseUMBase = objProduto.iClasseUM
            sProdutoBase = objProdutoKit.sProduto
        Else
            If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
                If sProdutoBase = objProdutoKit.sProduto Then
                    lErro = CF("UM_Conversao_Trans", iClasseUMBase, objProdutoKit.sUnidadeMed, sUMBase, dFator)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    dQtd = dQtd + (objProdutoKit.dQuantidade * dFator)
                End If
            End If
        End If
    
    Next
    
    'Um produto não pode ser insumo dele mesmo em igual ou maior quantidade
    If dQtd - dQtdBase > -QTDE_ESTOQUE_DELTA Then gError 206610
    
    For Each objProdutoKit In objKit.colComponentes
        If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
        
            sReferenciaCircular = ""
            sCaminho = ""
            iNivel = 0
            lErro = Kit_Valida_Recursividade(objProdutoKit, sProdutoBase, iClasseUMBase, sUMBase, dQtdBase, sReferenciaCircular, iNivel, sCaminho)
            If lErro <> SUCESSO And lErro <> 206606 And lErro <> 206607 And lErro <> 206608 Then gError ERRO_SEM_MENSAGEM 'Referência Circular Indireta (Um dos insumos utiliza o produto base em maior quantidade)
        
            sReferenciaCircular = "Prod: " & Trim(objKit.sProdutoRaiz) & " Versão: " & objKit.sVersao & sReferenciaCircular
        
            If lErro = 206606 Then gError 206611 'Limite de recursividade = 50
            If lErro = 206608 Then gError 206612 'Referencia Circular ao Produto Base
        
        Else
        
            If objProdutoKit.sProdutoRaiz <> objProdutoKit.sProduto Then gError 212051
            
        End If
        
    Next
    
    Exit Function

Erro_Kit_Valida:

    Kit_Valida = gErr

    Select Case gErr
    
        Case 206610
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODRAIZ_IGUAL_QTD_PRODKIT", gErr, Trim(sProdutoBase))
        
        Case 206611
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_RECURSIVIDADE", gErr)
        
        Case 206612
            Call Rotina_Erro(vbOKOnly, "ERRO_REFERENCIA_CIRCULAR", gErr, sReferenciaCircular)
        
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206613)

    End Select
    
    Exit Function

End Function

Private Function Kit_Valida_Recursividade(ByVal objProdutoKitFilho As ClassProdutoKit, ByVal sProdutoBase As String, ByVal iClasseUMBase As Integer, ByVal sUMBase As String, ByVal dQtdBase As Double, sReferenciaCircular As String, ByVal iNivel As Integer, ByVal sCaminho As String) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objKit As New ClassKit
Dim dFatorAcerto As Double, dFator As Double
Dim objProdutoKit As ClassProdutoKit
Dim dQtd As Double

On Error GoTo Erro_Kit_Valida_Recursividade

    objProduto.sCodigo = objProdutoKitFilho.sProduto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

    If objProduto.iCompras = PRODUTO_PRODUZIVEL And objProdutoKitFilho.dQuantidade > 0.009 Then
    
        iNivel = iNivel + 1
        If iNivel > 50 Then gError 206606
               
        objKit.sVersao = objProdutoKitFilho.sVersaoKitComp
        objKit.sProdutoRaiz = objProdutoKitFilho.sProduto
               
        If objKit.sVersao = "" Then

            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError ERRO_SEM_MENSAGEM

            'Se não existe versão padrão p/ o Kit -> Erro.
            If lErro <> SUCESSO Then gError 206607
       
        End If
        
        sReferenciaCircular = sReferenciaCircular & "-> Prod: " & Trim(objKit.sProdutoRaiz) & " Versão: " & objKit.sVersao
        sCaminho = sCaminho & "-> Prod: " & Trim(objKit.sProdutoRaiz) & " Versão: " & objKit.sVersao
    
        'tem que ler o próprio cara e os seus filhos
        lErro = CF("ProdutoKit_Le_PrimeiroNivel_Versao", objKit)
        If lErro <> SUCESSO And lErro <> 103203 Then gError ERRO_SEM_MENSAGEM
        
        For Each objProdutoKit In objKit.colComponentes
              
            If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProdutoKitFilho.sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                dFatorAcerto = objProdutoKitFilho.dQuantidade / (objProdutoKit.dQuantidade * dFator)
            Else
                If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
                    
                    objProdutoKit.dQuantidade = objProdutoKit.dQuantidade * dFatorAcerto
                    
                    If sProdutoBase = objProdutoKit.sProduto Then
                        lErro = CF("UM_Conversao_Trans", iClasseUMBase, objProdutoKit.sUnidadeMed, sUMBase, dFator)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        dQtd = dQtd + (objProdutoKit.dQuantidade * dFator)
                    End If
                
                End If
            End If
        
        Next
        
        'Um produto não pode ser insumo dele mesmo em igual ou maior quantidade
        If dQtd - dQtdBase > -QTDE_ESTOQUE_DELTA Then gError 206608

        For Each objProdutoKit In objKit.colComponentes
            If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
            
                lErro = Kit_Valida_Recursividade(objProdutoKit, sProdutoBase, iClasseUMBase, sUMBase, dQtdBase, sReferenciaCircular, iNivel, sCaminho)
                If lErro <> SUCESSO And lErro <> 206606 And lErro <> 206607 And lErro <> 206608 Then gError ERRO_SEM_MENSAGEM 'Referência Circular Indireta (Um dos insumos utiliza o produto base em maior quantidade)
            
                If lErro <> SUCESSO Then gError lErro
            End If
        Next

    End If
    
    sReferenciaCircular = ""
        
    Exit Function

Erro_Kit_Valida_Recursividade:

    Kit_Valida_Recursividade = gErr

    Select Case gErr
    
        Case 206606, 206607, 206608
        
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206609)

    End Select
    
    Exit Function

End Function

Function Kit_Le_FormPreco(ByVal objKit As ClassKit) As Long
'Le a versao do kit que deve ser usada para formacao de precos

Dim lErro As Long
Dim lComando As Long
Dim tKit As typeKit

On Error GoTo Erro_Kit_Le_FormPreco

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106301

    With tKit
        .sVersao = String(STRING_KIT_VERSAO, 0)
        .sObservacao = String(STRING_KIT_OBSERVACAO, 0)

        lErro = Comando_Executar(lComando, "SELECT Versao, Observacao, Data, Situacao, VersaoFormPreco, PesoFator1, PesoFator2, PesoFator3, PesoFator4, PesoFator5, PesoFator6 FROM Kit WHERE ProdutoRaiz = ? AND VersaoFormPreco = ?", .sVersao, .sObservacao, .dtData, .iSituacao, .iVersaoFormPreco, .dPesoFator1, .dPesoFator2, .dPesoFator3, .dPesoFator4, .dPesoFator5, .dPesoFator6, objKit.sProdutoRaiz, 1)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 106302

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106303

    If lErro = AD_SQL_SEM_DADOS Then gError 106304

    With tKit
        objKit.sVersao = .sVersao
        objKit.dtData = .dtData
        objKit.sObservacao = .sObservacao
        objKit.iSituacao = .iSituacao
        objKit.iVersaoFormPreco = .iVersaoFormPreco
        objKit.dPesoFator1 = .dPesoFator1
        objKit.dPesoFator2 = .dPesoFator2
        objKit.dPesoFator3 = .dPesoFator3
        objKit.dPesoFator4 = .dPesoFator4
        objKit.dPesoFator5 = .dPesoFator5
        objKit.dPesoFator6 = .dPesoFator6
    End With

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Kit_Le_FormPreco = SUCESSO

    Exit Function

Erro_Kit_Le_FormPreco:

    Kit_Le_FormPreco = gErr

    Select Case gErr

        Case 106301
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 106302, 106303
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_KIT", gErr)

        Case 106304

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152184)

    End Select

   'Fechamento comando
    Call Comando_Fechar(lComando)

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function TestesQualidade_Le(ByVal objTestesQualidade As ClassTestesQualidade) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTestesQualidade As typeTestesQualidade

On Error GoTo Erro_TestesQualidade_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130106

    'Alocação de espaço no buffer
    With tTestesQualidade
        .sNomeReduzido = String(STRING_TESTEQUALIDADE_NOMERED, 0)
        .sEspecificacao = String(STRING_TESTEQUALIDADE_ESPECIFICACAO, 0)
        .sMetodoUsado = String(STRING_TESTEQUALIDADE_METODOUSADO, 0)
        .sObservacao = String(STRING_TESTEQUALIDADE_OBSERVACAO, 0)
    End With
    
    'Le a tabelaTestesQualidade
    lErro = Comando_Executar(lComando, "SELECT Codigo, NomeReduzido, Especificacao, TipoResultado, LimiteDe, LimiteAte, " & _
                "MetodoUsado, Observacao, NoCertificado FROM TestesQualidade WHERE Codigo= ? ", _
                tTestesQualidade.iCodigo, tTestesQualidade.sNomeReduzido, tTestesQualidade.sEspecificacao, tTestesQualidade.iTipoResultado, _
                tTestesQualidade.dLimiteDe, tTestesQualidade.dLimiteAte, tTestesQualidade.sMetodoUsado, tTestesQualidade.sObservacao, tTestesQualidade.iNoCertificado, _
                objTestesQualidade.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 130107

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130108

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 130109

    With tTestesQualidade
        objTestesQualidade.iCodigo = .iCodigo
        objTestesQualidade.sNomeReduzido = .sNomeReduzido
        objTestesQualidade.sEspecificacao = .sEspecificacao
        objTestesQualidade.iTipoResultado = .iTipoResultado
        objTestesQualidade.dLimiteDe = .dLimiteDe
        objTestesQualidade.dLimiteAte = .dLimiteAte
        objTestesQualidade.sMetodoUsado = .sMetodoUsado
        objTestesQualidade.sObservacao = .sObservacao
        objTestesQualidade.iNoCertificado = .iNoCertificado
    End With
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    TestesQualidade_Le = SUCESSO

    Exit Function

Erro_TestesQualidade_Le:

    TestesQualidade_Le = gErr

    Select Case gErr

        Case 130106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130107, 130108
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TESTESQUALIDADE", gErr)

        Case 130109 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152185)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoTeste_Testa_Relacionamentos(ByVal objProdutoTeste As ClassProdutoTeste) As Long
'Testa se a associacao entre o Produto e o Teste passados no ObjProdutoTeste existe em alguma tabela

Dim lErro As Long
Dim lComando As Long
Dim iTeste As Integer

On Error GoTo Erro_ProdutoTeste_Testa_Relacionamentos

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 130285
        
    'verifica se já foi usado o teste para o produto em algum lote
    lErro = Comando_Executar(lComando, "SELECT ProdutoTeste.TesteCodigo From ProdutoTeste, RastreamentoLoteTeste, RastreamentoLote WHERE RastreamentoLoteTeste.NumIntRastroLote = RastreamentoLote.NumIntDoc AND ProdutoTeste.TesteCodigo = RastreamentoLoteTeste.TesteCodigo AND ProdutoTeste.Produto = RastreamentoLote.Produto AND ProdutoTeste.Produto = ? AND ProdutoTeste.TesteCodigo = ?", _
        iTeste, objProdutoTeste.sProduto, objProdutoTeste.iTesteCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 130286
        
    'Buscar o Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130287

    'Se encontrou => erro
    If lErro = AD_SQL_SUCESSO Then gError 130288

    Call Comando_Fechar(lComando)
        
    ProdutoTeste_Testa_Relacionamentos = SUCESSO
    
    Exit Function

Erro_ProdutoTeste_Testa_Relacionamentos:

    ProdutoTeste_Testa_Relacionamentos = gErr
    
    Select Case gErr
    
        Case 130285
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 130286, 130287
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOTESTE_USO", gErr)
            
        Case 130288
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTOTESTE_RELACIONADO_LOTE", gErr, objProdutoTeste.iTesteCodigo, objProdutoTeste.sProduto)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152186)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function ProdutoTeste_Le(ByVal objProdutoTeste As ClassProdutoTeste) As Long
'Le dados da associacao do Produto com o Teste

Dim lErro As Long
Dim lComando As Long
Dim tProdTeste As typeProdutoTeste

On Error GoTo Erro_ProdutoTeste_Le

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130289

    With tProdTeste
    
        .sTesteEspecificacao = String(STRING_TESTEQUALIDADE_ESPECIFICACAO, 0)
        .sTesteMetodoUsado = String(STRING_TESTEQUALIDADE_METODOUSADO, 0)
        .sTesteObservacao = String(STRING_TESTEQUALIDADE_OBSERVACAO, 0)
        
        'Lê os testes do Produto
        lErro = Comando_Executar(lComando, "SELECT SeqGrid, TesteEspecificacao, TesteTipoResultado, TesteLimiteDe, TesteLimiteAte, TesteMetodoUsado, TesteObservacao, TesteNoCertificado FROM ProdutoTeste WHERE Produto = ? AND TesteCodigo = ?", _
            .iSeqGrid, .sTesteEspecificacao, .iTesteTipoResultado, .dTesteLimiteDe, .dTesteLimiteAte, .sTesteMetodoUsado, .sTesteObservacao, .iTesteNoCertificado, objProdutoTeste.sProduto, objProdutoTeste.iTesteCodigo)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 130290

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130291
    
    If lErro <> AD_SQL_SUCESSO Then gError 130292
    
    With tProdTeste
        
        objProdutoTeste.iSeqGrid = .iSeqGrid
        objProdutoTeste.sTesteEspecificacao = .sTesteEspecificacao
        objProdutoTeste.iTesteTipoResultado = .iTesteTipoResultado
        objProdutoTeste.dTesteLimiteDe = .dTesteLimiteDe
        objProdutoTeste.dTesteLimiteAte = .dTesteLimiteAte
        objProdutoTeste.sTesteMetodoUsado = .sTesteMetodoUsado
        objProdutoTeste.sTesteObservacao = .sTesteObservacao
        objProdutoTeste.iTesteNoCertificado = .iTesteNoCertificado
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ProdutoTeste_Le = SUCESSO

    Exit Function

Erro_ProdutoTeste_Le:

    ProdutoTeste_Le = gErr

    Select Case gErr
       
        Case 130289
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130290, 130291
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOTESTE", gErr, objProdutoTeste.sProduto)
 
        Case 130292
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152187)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ProdutoTeste_Le_Produto(ByVal objProduto As ClassProduto) As Long
'Lê os teste associados a um produto a partir do código de produto passado como parâmetro em objProduto
'Devolve os testes na coleção colProdutoTeste dentro do objProduto

Dim lErro As Long
Dim lComando As Long
Dim tProdutoTeste As typeProdutoTeste
Dim objProdutoTeste As ClassProdutoTeste

On Error GoTo Erro_ProdutoTeste_Le_Produto

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 130293
    
    'Inicializa a colecao existente dentro do obj
    Set objProduto.colProdutoTeste = New Collection
    
    With tProdutoTeste
        
        'Inicializa as strings
        .sTesteEspecificacao = String(STRING_TESTEQUALIDADE_ESPECIFICACAO, 0)
        .sTesteMetodoUsado = String(STRING_TESTEQUALIDADE_METODOUSADO, 0)
        .sTesteObservacao = String(STRING_TESTEQUALIDADE_OBSERVACAO, 0)
        
        'Lê os testes do Produto
        lErro = Comando_Executar(lComando, "SELECT TesteCodigo, SeqGrid, TesteEspecificacao, TesteTipoResultado, TesteLimiteDe, TesteLimiteAte, TesteMetodoUsado, TesteObservacao, TesteNoCertificado FROM ProdutoTeste WHERE Produto = ? ORDER BY SeqGrid", _
            .iTesteCodigo, .iSeqGrid, .sTesteEspecificacao, .iTesteTipoResultado, .dTesteLimiteDe, .dTesteLimiteAte, .sTesteMetodoUsado, .sTesteObservacao, .iTesteNoCertificado, objProduto.sCodigo)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 130294
    
    'Buscar o Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130295

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 130296
    
    'Busca todos os registros filtrados pelo SELECT
    While lErro = AD_SQL_SUCESSO
        
        'Inicializa o objProdutoTeste
        Set objProdutoTeste = New ClassProdutoTeste
        
        'Preenche o obj com os dados lidos no BD
        With tProdutoTeste
            
            objProdutoTeste.sProduto = objProduto.sCodigo
            objProdutoTeste.iTesteCodigo = .iTesteCodigo
            objProdutoTeste.iSeqGrid = .iSeqGrid
            objProdutoTeste.sTesteEspecificacao = .sTesteEspecificacao
            objProdutoTeste.iTesteTipoResultado = .iTesteTipoResultado
            objProdutoTeste.dTesteLimiteDe = .dTesteLimiteDe
            objProdutoTeste.dTesteLimiteAte = .dTesteLimiteAte
            objProdutoTeste.sTesteMetodoUsado = .sTesteMetodoUsado
            objProdutoTeste.sTesteObservacao = .sTesteObservacao
            objProdutoTeste.iTesteNoCertificado = .iTesteNoCertificado
            
        End With
            
        'adiciona na colecao
        objProduto.colProdutoTeste.Add objProdutoTeste
        
        'busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130297
        
    Wend
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    ProdutoTeste_Le_Produto = SUCESSO
    
    Exit Function

Erro_ProdutoTeste_Le_Produto:

    ProdutoTeste_Le_Produto = gErr
    
    Select Case gErr
    
        Case 130293
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 130294, 130295, 130297
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOTESTE", gErr)
        
        Case 130296
            'Erro tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152188)
            
    End Select
    
    'Fechamento dos comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function TesteQualidade_Le_NomeReduzido(ByVal objTestesQualidade As ClassTestesQualidade) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTestesQualidade As typeTestesQualidade

On Error GoTo Erro_TesteQualidade_Le_NomeReduzido

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130106

    'Alocação de espaço no buffer
    With tTestesQualidade
        .sEspecificacao = String(STRING_TESTEQUALIDADE_ESPECIFICACAO, 0)
        .sMetodoUsado = String(STRING_TESTEQUALIDADE_METODOUSADO, 0)
        .sObservacao = String(STRING_TESTEQUALIDADE_OBSERVACAO, 0)
    End With
    
    'Le a tabelaTestesQualidade
    lErro = Comando_Executar(lComando, "SELECT Codigo, Especificacao, TipoResultado, LimiteDe, LimiteAte, " & _
                "MetodoUsado, Observacao, NoCertificado FROM TestesQualidade WHERE NomeReduzido = ?", _
                tTestesQualidade.iCodigo, tTestesQualidade.sEspecificacao, tTestesQualidade.iTipoResultado, _
                tTestesQualidade.dLimiteDe, tTestesQualidade.dLimiteAte, tTestesQualidade.sMetodoUsado, tTestesQualidade.sObservacao, tTestesQualidade.iNoCertificado, _
                objTestesQualidade.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 130107

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130108

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError 130109

    With tTestesQualidade
        objTestesQualidade.iCodigo = .iCodigo
        objTestesQualidade.sEspecificacao = .sEspecificacao
        objTestesQualidade.iTipoResultado = .iTipoResultado
        objTestesQualidade.dLimiteDe = .dLimiteDe
        objTestesQualidade.dLimiteAte = .dLimiteAte
        objTestesQualidade.sMetodoUsado = .sMetodoUsado
        objTestesQualidade.sObservacao = .sObservacao
        objTestesQualidade.iNoCertificado = .iNoCertificado
    End With
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    TesteQualidade_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_TesteQualidade_Le_NomeReduzido:

    TesteQualidade_Le_NomeReduzido = gErr

    Select Case gErr

        Case 130106
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130107, 130108
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TESTESQUALIDADE", gErr)

        Case 130109 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152189)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RastreamentoLote_Le_Testes(ByVal objRastreamentoLote As ClassRastreamentoLote) As Long
'preenche a objRastreamentoLote.colTeste com os testes associados a objRastreamentoLote.NumIntDoc

Dim lErro As Long, objTeste As ClassRastreamentoLoteTeste
Dim tTeste As typeRastreamentoLoteTeste, lComando As Long

On Error GoTo Erro_RastreamentoLote_Le_Testes

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130310
    
    With tTeste
        
        .sTesteEspecificacao = String(STRING_TESTEQUALIDADE_ESPECIFICACAO, 0)
        .sTesteMetodoUsado = String(STRING_TESTEQUALIDADE_METODOUSADO, 0)
        .sTesteObservacao = String(STRING_TESTEQUALIDADE_OBSERVACAO, 0)
        .sRegistroAnaliseID = String(STRING_RESULTADO_ANALISE_ID, 0)
        .sResultadoValor = String(STRING_RESULTADO_VALOR, 0)
        .sResultadoObservacao = String(STRING_RESULTADO_OBS, 0)
        
        lErro = Comando_Executar(lComando, "SELECT TesteCodigo, SeqGrid, TesteEspecificacao, TesteTipoResultado, TesteLimiteDe, TesteLimiteAte, " & _
            "TesteMetodoUsado , TesteObservacao, RegistroAnaliseID, RegistroAnaliseData, ResultadoNaoConforme, ResultadoValor, ResultadoObservacao, TesteNoCertificado FROM RastreamentoLoteTeste WHERE NumIntRastroLote = ? ORDER BY SeqGrid", _
            .iTesteCodigo, .iSeqGrid, .sTesteEspecificacao, .iTesteTipoResultado, .dTesteLimiteDe, .dTesteLimiteAte, _
            .sTesteMetodoUsado, .sTesteObservacao, .sRegistroAnaliseID, .dtRegistroAnaliseData, .iResultadoNaoConforme, .sResultadoValor, .sResultadoObservacao, .iTesteNoCertificado, objRastreamentoLote.lNumIntDoc)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 130311
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130312
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objTeste = New ClassRastreamentoLoteTeste
        
        With objTeste
        
            .lNumIntRastroLote = objRastreamentoLote.lNumIntDoc
            .iTesteCodigo = tTeste.iTesteCodigo
            .iSeqGrid = tTeste.iSeqGrid
            .sTesteEspecificacao = tTeste.sTesteEspecificacao
            .iTesteTipoResultado = tTeste.iTesteTipoResultado
            .dTesteLimiteDe = tTeste.dTesteLimiteDe
            .dTesteLimiteAte = tTeste.dTesteLimiteAte
            .sTesteMetodoUsado = tTeste.sTesteMetodoUsado
            .sTesteObservacao = tTeste.sTesteObservacao
            .sRegistroAnaliseID = tTeste.sRegistroAnaliseID
            .dtRegistroAnaliseData = tTeste.dtRegistroAnaliseData
            .iResultadoNaoConforme = tTeste.iResultadoNaoConforme
            .sResultadoValor = tTeste.sResultadoValor
            .sResultadoObservacao = tTeste.sResultadoObservacao
            .iTesteNoCertificado = tTeste.iTesteNoCertificado
            
        End With
        
        objRastreamentoLote.colTeste.Add objTeste
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130313
    
    Loop
            
    Call Comando_Fechar(lComando)
    
    RastreamentoLote_Le_Testes = SUCESSO
     
    Exit Function
    
Erro_RastreamentoLote_Le_Testes:

    RastreamentoLote_Le_Testes = gErr
     
    Select Case gErr
          
        Case 130310
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 130311, 130312, 130313
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TESTESLOTERASTRO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152190)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function UnidadesDeMedidas_Le(ByVal objUnidadeDeMedida As ClassUnidadeDeMedida) As Long
'Lê a Unidade de Medida a partir de uma Sigla, independente da ClasseUM
'Criado por Jorge Specian - 30/03/2005

Dim lErro As Long
Dim lComando As Long
Dim iClasse As Integer
Dim sSigla As String
Dim sNome As String
Dim dQuantidade As Double
Dim sSiglaUMBase As String

On Error GoTo Erro_UnidadesDeMedidas_Le

    'Inicializar comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 134460

    sSigla = String(STRING_UM_SIGLA, 0)
    sNome = String(STRING_CLASSEUM_NOME, 0)
    sSiglaUMBase = String(STRING_UM_SIGLA, 0)

    'Executar comando SQL
    lErro = Comando_Executar(lComando, "SELECT Classe, Sigla, Nome, Quantidade, SiglaUMBase FROM UnidadesDeMedida WHERE Sigla = ? ", iClasse, sSigla, sNome, dQuantidade, sSiglaUMBase, objUnidadeDeMedida.sSigla)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134461

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 134462

    'Se não encontrou
    If lErro = AD_SQL_SEM_DADOS Then gError 134463

    objUnidadeDeMedida.iClasse = iClasse
    objUnidadeDeMedida.sNome = sNome
    objUnidadeDeMedida.dQuantidade = dQuantidade
    objUnidadeDeMedida.sSiglaUMBase = sSiglaUMBase

    Call Comando_Fechar(lComando)

    UnidadesDeMedidas_Le = SUCESSO
    
    Exit Function

Erro_UnidadesDeMedidas_Le:

    UnidadesDeMedidas_Le = gErr

    Select Case gErr

        Case 134460
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 134461, 134462
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_UNIDADESDEMEDIDA", gErr)

        Case 134463
            'sem dados
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152191)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ProdutoKit_Le_QuantInsumos(objItemOP As ClassItemOP, colItensMovEstoque As colItensMovEstoque) As Long
'descobre a quantidade dos insumos que serao necessarios para produzir a quantidade passada pelo parametro objItemOP e devolve em colItensMovEstoque.

Dim objProdutoKitRaiz As New ClassProdutoKit
Dim objProdutoKitProdutos As New ClassProdutoKitProdutos
Dim objKit As New ClassKit
Dim lErro As Long
Dim dQuantidadeConv As Double, dFator As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdutoKit_Le_QuantInsumos

    objProdutoKitRaiz.sProdutoRaiz = objItemOP.sProduto
    objProdutoKitRaiz.sVersao = objItemOP.sVersao

    lErro = CF("ProdutoKit_Le_Raiz", objProdutoKitRaiz)
    If lErro <> SUCESSO And lErro <> 34875 Then gError 126596
    
    If lErro = SUCESSO Then
    
        objProdutoKitProdutos.sProdutoRaiz = objItemOP.sProduto
        objProdutoKitProdutos.sVersao = objItemOP.sVersao
        objProdutoKitProdutos.dQuantidade = objItemOP.dQuantidade
        objProdutoKitProdutos.iClasseUM = objProdutoKitRaiz.iClasseUM
        objProdutoKitProdutos.sUnidadeMed = objItemOP.sSiglaUM
        
        
        'Obtem uma Colecao com os itens do kit
        lErro = CF("OrdemProducao_Le_Col_Empenho2", objProdutoKitProdutos, objKit.colComponentes, False)
        If lErro <> SUCESSO And lErro <> 106393 Then gError 126597
        
        'Se nao Encontrou => Erro
        If lErro = SUCESSO Then
        
            lErro = CF("UM_Conversao_Trans", objProdutoKitRaiz.iClasseUM, objItemOP.sSiglaUM, objProdutoKitRaiz.sUnidadeMed, dFator)
            If lErro <> SUCESSO Then gError 126598
        
            For Each objProdutoKitProdutos In objKit.colComponentes
        
                If objProdutoKitProdutos.iComposicao = PRODUTOKIT_COMPOSICAO_FIXA Then
        
                    dQuantidadeConv = objProdutoKitProdutos.dQuantidade
        
                ElseIf objProdutoKitProdutos.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
        
                    dQuantidadeConv = (objItemOP.dQuantidade * dFator / objProdutoKitRaiz.dQuantidade) * (objProdutoKitProdutos.dQuantidade)
        
                End If
        
                'Carrega sProdutoFormatado em objProduto
                objProduto.sCodigo = objProdutoKitProdutos.sProduto
    
                'Tenta ler o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 126617
        
                If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    colItensMovEstoque.Add 0, 0, 0, 0, objProdutoKitProdutos.sProduto, "", objProdutoKitProdutos.sUnidadeMed, dQuantidadeConv, 0, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                End If
    
            Next

        End If

    End If

    ProdutoKit_Le_QuantInsumos = SUCESSO

    Exit Function

Erro_ProdutoKit_Le_QuantInsumos:

    ProdutoKit_Le_QuantInsumos = gErr

    Select Case gErr

        Case 126596, 126597, 126598, 126617

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 126632)

    End Select

    Exit Function

End Function

Function Produto_Testa_Quant_Requerida(objItemOP As ClassItemOP, sProdMaiorReq As String) As Long
'descobre a quantidade dos insumos que serao necessarios para produzir a quantidade passada pelo parametro objItemOP e devolve em colItensMovEstoque.

    Produto_Testa_Quant_Requerida = SUCESSO

End Function

Function ReqProducao_Total_Produto_OP(ByVal objProduto As ClassProduto, ByVal objItemMovEst As ClassItemMovEstoque) As Long
'totaliza as requisicoes de um insumo de um produto de uma op

Dim lErro As Long
Dim lComando As Long
Dim sSiglaUM As String
Dim dQuantidade As Double
Dim dFator As Double


On Error GoTo Erro_ReqProducao_Total_Produto_OP

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 126602

    sSiglaUM = String(STRING_UM_SIGLA, 0)

    'Leitura da tabela de Ordens de Produção
    lErro = Comando_Executar(lComando, "SELECT ME.Quantidade, ME.SiglaUM FROM MovimentoEstoque AS ME, ItensOrdemProducao AS IOP WHERE IOP.NumIntDoc = ME.NumIntDocOrigem AND IOP.FilialEmpresa = ? AND (ME.TipoMov = ? OR ME.TipoMov = ?) And ME.TipoNumIntDocOrigem = ? AND IOP.Codigo = ? AND ME.Produto = ? AND IOP.Produto = ?", _
        dQuantidade, sSiglaUM, objItemMovEst.iFilialEmpresa, MOV_EST_REQ_PRODUCAO, MOV_EST_REQ_PRODUCAO_BENEF3, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, objItemMovEst.sOPCodigo, objItemMovEst.sProduto, objItemMovEst.sProdutoOP)
    If lErro <> AD_SQL_SUCESSO Then gError 126603

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126604
    
    objItemMovEst.dQuantidadeEst = 0
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 126605
    
        objItemMovEst.dQuantidadeEst = objItemMovEst.dQuantidadeEst + (dQuantidade * dFator)
    
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126606
    
    Loop

    Call Comando_Fechar(lComando)

    ReqProducao_Total_Produto_OP = SUCESSO

    Exit Function

Erro_ReqProducao_Total_Produto_OP:

    ReqProducao_Total_Produto_OP = gErr

    Select Case gErr

        Case 126602
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126603, 126604, 126606
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 126605

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 126607)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Quant_Requerida_Superior(objMovEstoque As ClassMovEstoque, sReqMaiorOP As String) As Long
'verifica se  a quantidade requerida dos insumos supera o que esta especificado nos kits

    Produto_Quant_Requerida_Superior = SUCESSO

End Function

Public Function Escaninhos_Le_Todos(ByVal colEscaninhos As Collection) As Long
'Le todos os escaninhos

Dim lErro As Long
Dim lComando As Long
Dim tEscaninho As typeEscaninho
Dim objEscaninho As ClassEscaninho

On Error GoTo Erro_Escaninhos_Le_Todos

    'Abertura do comando
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 119705

    'inicializa a string p/ receber o nome do escaninho
    tEscaninho.sNome = String(STRING_NOME_ESCANINHO, 0)

    'Lê os dados dos Escaninhos que podem se de terceiros / nosso
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, RastroEstoqueInicial FROM Escaninhos ", tEscaninho.iCodigo, tEscaninho.sNome, tEscaninho.iRastroEstoqueInicial)
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

   Escaninhos_Le_Todos = SUCESSO

    Exit Function

Erro_Escaninhos_Le_Todos:

    Escaninhos_Le_Todos = gErr

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

'######################################################################
'Inserido por Wagner 31/05/2006
Function MovEstoque_Le_ItemNF_Grade(ByVal objItemMovEstoque As ClassItemMovEstoque, ByVal colMovEstoque As Collection) As Long

Dim lErro As Long
Dim objItemNF As New ClassItemNF
Dim objItemMovEst As ClassItemMovEstoque
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_MovEstoque_Le_ItemNF_Grade

    objItemNF.lNumIntDoc = objItemMovEstoque.lNumIntDocOrigem
    
    lErro = CF("ItensNFGrade_Le", objItemNF)
    If lErro <> SUCESSO Then gError 180021
    
    For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
    
        Set objItemMovEst = New ClassItemMovEstoque
        
        objItemMovEst.lNumIntDocOrigem = objItemRomaneio.lNumIntDoc
        objItemMovEst.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
        objItemMovEst.iFilialEmpresa = objItemMovEstoque.iFilialEmpresa
    
        lErro = CF("MovEstoque_Le_ItemNF", objItemMovEst, colMovEstoque)
        If lErro <> SUCESSO Then gError 180022
        
    Next
    
    MovEstoque_Le_ItemNF_Grade = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Le_ItemNF_Grade:

    MovEstoque_Le_ItemNF_Grade = gErr
    
    Select Case gErr
    
        Case 180021, 180022

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180020)

    End Select
    
    Exit Function
    
End Function
'#########################################################################

'#########################################################################
'Inserido por Wagner 12/06/2006
Function FornecedoresProdutoFF_Le_Produto(colFornecedorProdutoFF As Collection, ByVal objFilialFornecedor As ClassFilialFornecedor) As Long

Dim lErro As Long
Dim alComando(0 To 3) As Long
Dim iIndice As Integer
Dim objFornecedorProdutoFF As New ClassFornecedorProdutoFF
Dim tFornecedorProdutoFF As typeFornecedorProdutoFF

On Error GoTo Erro_FornecedoresProdutoFF_Le_Produto

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If alComando(iIndice) = 0 Then gError 180191
    Next

    tFornecedorProdutoFF.sProduto = String(STRING_PRODUTO, 0)
    tFornecedorProdutoFF.sProdutoFornecedor = String(STRING_PRODUTO, 0)

    'Lê a tabela FornecedorProdutoFF de acordo com o Produto e FilialEmpresa, ordenados por Fornecedor e FilialForn
    lErro = Comando_Executar(alComando(1), "SELECT FilialEmpresa,Produto,Fornecedor,FilialForn,ProdutoFornecedor,LoteMinimo,Nota FROM FornecedorProdutoFF WHERE FilialEmpresa=? AND Fornecedor=? AND FilialForn = ? ORDER BY Produto", tFornecedorProdutoFF.iFilialEmpresa, tFornecedorProdutoFF.sProduto, tFornecedorProdutoFF.lFornecedor, tFornecedorProdutoFF.iFilialForn, tFornecedorProdutoFF.sProdutoFornecedor, tFornecedorProdutoFF.dLoteMinimo, tFornecedorProdutoFF.iNota, giFilialEmpresa, objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 180192

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180193

    'Se nao encontrou registro em FornecedorProdutoFF
    If lErro = AD_SQL_SEM_DADOS Then gError 180194

    'Set colFilFornFilEmp = New Collection
    Set colFornecedorProdutoFF = New Collection
    
    Do While lErro = AD_SQL_SUCESSO

        'Guarda na coleção os regitros lidos de FornecedorProdutoFF
        'Set objFilFornFilEmp = New ClassFilFornFilEmp
        Set objFornecedorProdutoFF = New ClassFornecedorProdutoFF
        
        'Preenche objFilFornFilEmp com os dados lidos do BD
        objFornecedorProdutoFF.dLoteMinimo = tFornecedorProdutoFF.dLoteMinimo
        objFornecedorProdutoFF.iFilialEmpresa = tFornecedorProdutoFF.iFilialEmpresa
        objFornecedorProdutoFF.iFilialForn = tFornecedorProdutoFF.iFilialForn
        objFornecedorProdutoFF.iNota = tFornecedorProdutoFF.iNota
        objFornecedorProdutoFF.iPadrao = tFornecedorProdutoFF.iPadrao
        objFornecedorProdutoFF.lFornecedor = tFornecedorProdutoFF.lFornecedor
        objFornecedorProdutoFF.sProduto = tFornecedorProdutoFF.sProduto
        objFornecedorProdutoFF.sProdutoFornecedor = tFornecedorProdutoFF.sProdutoFornecedor
        
        'Adiciona na colecao
        colFornecedorProdutoFF.Add objFornecedorProdutoFF
        
        'Le o próximo registro existente no banco de dados
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180195

    Loop

    'Fecha o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FornecedoresProdutoFF_Le_Produto = SUCESSO

    Exit Function

Erro_FornecedoresProdutoFF_Le_Produto:

    FornecedoresProdutoFF_Le_Produto = gErr

    Select Case gErr

        Case 180191
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180192, 180193, 180195
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORPRODUTOFF", gErr)

        Case 180194
            'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180195)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function
'#################################################################################

Function Produto_Customiza_Form_Load(objTela As Object) As Long

Dim lErro As Long
Dim lComando As Long
Dim tMovEstoque As typeMovEst
Dim dHora As Double

On Error GoTo Erro_Produto_Customiza_Form_Load

    Produto_Customiza_Form_Load = SUCESSO

    Exit Function

Erro_Produto_Customiza_Form_Load:

    Produto_Customiza_Form_Load = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178506)

    End Select

    Exit Function

End Function


Function Empenho_Le_1(ByVal objEmpenho As ClassEmpenho) As Long
'Lê os dados de Empenho à partir da chave FilialEmpresa, candidata fornecida em objEmpenho, com os seguintes campos preenchidos:
'objEmpenho.iFilialEmpresa, objEmpenho.sCodigoOP, objEmpenho.iItemOP, objEmpenho.sProduto, objEmpenho.iAlmoxarifado

Dim lErro As Long
Dim lComando As Long
Dim tEmpenho As typeEmpenho

On Error GoTo Erro_Empenho_Le_1

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178589

    'Pesquisa no BD o Empenho passado como parâmetro
    lErro = Comando_Executar(lComando, "SELECT SUM(Quantidade), SUM(QuantidadeRequisitada) FROM Empenho WHERE NumIntDocItemOP = ? AND FilialEmpresa = ? AND Empenho.Produto = ?", _
        tEmpenho.dQuantidade, tEmpenho.dQuantidadeRequisitada, objEmpenho.lNumIntDocItemOP, objEmpenho.iFilialEmpresa, objEmpenho.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 178590

    'Verifica se o Empenho existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178591

    'Não encontrou o Empenho ==> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 178592

    'Preenche objEmpenho
    objEmpenho.dQuantidade = tEmpenho.dQuantidade
    objEmpenho.dQuantidadeRequisitada = tEmpenho.dQuantidadeRequisitada
    
    'Fechamento comando
    Call Comando_Fechar(lComando)

    Empenho_Le_1 = SUCESSO

    Exit Function

Erro_Empenho_Le_1:

    Empenho_Le_1 = gErr

    Select Case gErr

        Case 178589
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178590, 178591
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPENHO2", gErr)

        Case 178592

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178593)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItensOP_Baixada_Le(ByVal objOrdemProducao As ClassOrdemDeProducao) As Long
'Lê os Itens Ordem de Produção a partir do OPCódigo e da FilialEmpresa
'Retorna os Itens na coleção

Dim lErro As Long
Dim lComando As Long
Dim tOrdemProducao As typeOrdemProducao
Dim objItemOP As ClassItemOP
Dim tItemOP As typeItemOP

On Error GoTo Erro_ItensOP_Baixada_Le

    'Abertura de comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178685

    With tItemOP

        .sCcl = String(STRING_CCL, 0)
        .sProduto = String(STRING_PRODUTO, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)
        .sVersao = String(STRING_KIT_VERSAO, 0)

        '##############################################################
        'INSERIDO POR WAGNER
        .sJustificativaBloqueio = String(STRING_MAX_NOME_ARQUIVO, 0)
        .sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        '##############################################################

        '##############################################################
        'ALTERADO POR WAGNER
        lErro = Comando_Executar(lComando, "SELECT I.Quantidade ,I.QuantidadeProd ,I.DataFimProd ,I.DataInicioProd, I.Almoxarifado, I.Destinacao, I.FilialPedido, I.Item, I.NumIntOrigem, I.Prioridade, I.Situacao, I.CodPedido, I.NumIntDoc, I.Ccl, I.Produto, I.SiglaUM, I.NumIntEquipamento, I.Versao, I.NumIntDocPai, I.JustificativaBloqueio, P.Descricao FROM ItensOrdemProducaoBaixadas AS I, Produtos AS P WHERE I.Produto=P.Codigo AND I.FilialEmpresa = ? AND I.Codigo = ? ", _
        .dQuantidade, .dQuantidadeProd, .dtDataFimProd, .dtDataInicioProd, .iAlmoxarifado, .iDestinacao, .iFilialPedido, .iItem, .lNumIntOrigem, .iPrioridade, .iSituacao, .lCodPedido, .lNumIntDoc, .sCcl, .sProduto, .sSiglaUM, .lNumIntEquipamento, .sVersao, .lNumIntDocPai, .sJustificativaBloqueio, .sDescricao, objOrdemProducao.iFilialEmpresa, objOrdemProducao.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 178686
        '##############################################################

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178687

        While lErro <> AD_SQL_SEM_DADOS

            Set objItemOP = New ClassItemOP

            objItemOP.dQuantidade = .dQuantidade
            objItemOP.dQuantidadeProd = .dQuantidadeProd
            objItemOP.dtDataFimProd = .dtDataFimProd
            objItemOP.dtDataInicioProd = .dtDataInicioProd
            objItemOP.iAlmoxarifado = .iAlmoxarifado
            objItemOP.iDestinacao = .iDestinacao
            objItemOP.iFilialPedido = .iFilialPedido
            objItemOP.iItem = .iItem
            objItemOP.lNumIntOrigem = .lNumIntOrigem
            objItemOP.iPrioridade = .iPrioridade
            objItemOP.iSituacao = .iSituacao
            objItemOP.lCodPedido = .lCodPedido
            objItemOP.lNumIntDoc = .lNumIntDoc
            objItemOP.sCcl = .sCcl
            objItemOP.sProduto = .sProduto
            objItemOP.sSiglaUM = .sSiglaUM
            
            objItemOP.lNumIntEquipamento = .lNumIntEquipamento
            objItemOP.sVersao = .sVersao
            
            objItemOP.lNumIntDocPai = .lNumIntDocPai
            objItemOP.sJustificativaBloqueio = .sJustificativaBloqueio
            objItemOP.sDescricao = .sDescricao

            objOrdemProducao.colItens.Add objItemOP

            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178688

        Wend

    End With

    'Verifica se há algum Item para nessa Ordem de Produção
    If objOrdemProducao.colItens.Count = 0 Then gError 178689

    Call Comando_Fechar(lComando)

    ItensOP_Baixada_Le = SUCESSO

    Exit Function

Erro_ItensOP_Baixada_Le:

    ItensOP_Baixada_Le = gErr

    Select Case gErr

        Case 178685
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178686 To 178688
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAOBAIXADAS", gErr)

        Case 178689

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 178690)

    End Select

    Call Comando_Fechar(lComando)

End Function

Function ItensOP_Le_PV(ByVal lNumIntItemOP As Long, lCodPV As Long, iFilialPV As Integer) As Long
'procura o retornar o codigo do pedido de venda e filial de pedido de venda associado ao item da op. Pesquisa no item e nos seus ancestrais.

Dim lErro As Long
Dim lComando As Long
Dim tOrdemProducao As typeOrdemProducao
Dim objItemOP As ClassItemOP
Dim tItemOP As typeItemOP
Dim lNumIntDocPai As Long

On Error GoTo Erro_ItensOP_Le_PV

    'Abertura de comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 178691

    lErro = Comando_Executar(lComando, "SELECT CodPedido, FilialPedido, NumIntDocPai FROM ItensOrdemProducao WHERE NumIntDoc = ?", lCodPV, iFilialPV, lNumIntDocPai, lNumIntItemOP)
    If lErro <> AD_SQL_SUCESSO Then gError 178692

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178693

    If lErro = AD_SQL_SEM_DADOS Then

        lErro = Comando_Executar(lComando, "SELECT CodPedido, FilialPedido, NumIntDocPai FROM ItensOrdemProducaoBaixadas WHERE NumIntDoc = ?", lCodPV, iFilialPV, lNumIntDocPai, lNumIntItemOP)
        If lErro <> AD_SQL_SUCESSO Then gError 178694

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 178695

        If lErro = AD_SQL_SEM_DADOS Then gError 178696

    End If

    If lCodPV = 0 And lNumIntDocPai <> 0 Then
    
        lErro = ItensOP_Le_PV(lNumIntDocPai, lCodPV, iFilialPV)
        If lErro <> SUCESSO Then gError 178697

    End If

    Call Comando_Fechar(lComando)

    ItensOP_Le_PV = SUCESSO

    Exit Function

Erro_ItensOP_Le_PV:

    ItensOP_Le_PV = gErr

    Select Case gErr

        Case 178691
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 178692, 178693
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOP", gErr)

        Case 178694, 178695
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDENSPRODUCAOBAIXADAS", gErr)

        Case 178696, 178697

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 178690)

    End Select

    Call Comando_Fechar(lComando)

End Function

Function ProdutoCategoria_Le_Item(ByVal objProdutoCategoria As ClassProdutoCategoria, ByVal objCategoriaProdutoItem As ClassCategoriaProdutoItem) As Long
'retorna o item do produto/categoria

Dim lErro As Long
Dim lComando As Long
Dim tCatProdItem As typeCategoriaProdutoItem


On Error GoTo Erro_ProdutoCategoria_Le_Item

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 183545
    
    tCatProdItem.sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    tCatProdItem.sCategoria = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    tCatProdItem.sDescricao = String(STRING_CATEGORIAPRODUTO_DESCRICAO, 0)

    lErro = Comando_Executar(lComando, "SELECT CategoriaProdutoItem.Categoria, CategoriaProdutoItem.Item, CategoriaProdutoItem.Ordem, CategoriaProdutoItem.Descricao, Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Valor7, Valor8 FROM ProdutoCategoria, CategoriaProdutoItem WHERE ProdutoCategoria.Categoria = CategoriaProdutoItem.Categoria AND ProdutoCategoria.Produto = ? AND ProdutoCategoria.Categoria = ? AND ProdutoCategoria.Item = CategoriaProdutoItem.Item", _
    tCatProdItem.sCategoria, tCatProdItem.sItem, tCatProdItem.iOrdem, tCatProdItem.sDescricao, tCatProdItem.dvalor1, tCatProdItem.dvalor2, tCatProdItem.dvalor3, tCatProdItem.dvalor4, tCatProdItem.dvalor5, tCatProdItem.dvalor6, tCatProdItem.dvalor7, tCatProdItem.dvalor8, objProdutoCategoria.sProduto, objProdutoCategoria.sCategoria)
    If lErro <> AD_SQL_SUCESSO Then gError 183546
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 183547
        
    If lErro = AD_SQL_SEM_DADOS Then gError 183548
    
    objCategoriaProdutoItem.sCategoria = tCatProdItem.sCategoria
    objCategoriaProdutoItem.sItem = tCatProdItem.sItem
    objCategoriaProdutoItem.iOrdem = tCatProdItem.iOrdem
    objCategoriaProdutoItem.sDescricao = tCatProdItem.sCategoria
    objCategoriaProdutoItem.dvalor1 = tCatProdItem.dvalor1
    objCategoriaProdutoItem.dvalor2 = tCatProdItem.dvalor2
    objCategoriaProdutoItem.dvalor3 = tCatProdItem.dvalor3
    objCategoriaProdutoItem.dvalor4 = tCatProdItem.dvalor4
    objCategoriaProdutoItem.dvalor5 = tCatProdItem.dvalor5
    objCategoriaProdutoItem.dvalor6 = tCatProdItem.dvalor6
    objCategoriaProdutoItem.dvalor7 = tCatProdItem.dvalor7
    objCategoriaProdutoItem.dvalor8 = tCatProdItem.dvalor8
    
    Call Comando_Fechar(lComando)
    
    ProdutoCategoria_Le_Item = SUCESSO

    Exit Function

Erro_ProdutoCategoria_Le_Item:

    ProdutoCategoria_Le_Item = gErr

    Select Case gErr

        Case 183545
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 183546, 183547
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTO", Err, objProdutoCategoria.sProduto, objProdutoCategoria.sCategoria)
        
        Case 183548
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 183549)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function MovEstoqueItens_Le_Desc2(ByVal objMovEstoque As ClassMovEstoque, ByVal lNumIntNF As Long) As Long
'Lê todos os Itens de movimento de estoque associados aos itensdenfiscal a partir da FilialEmpresa e TipoMov e retorna no colItens do objeto
'Ordena pelo Item.

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim dtData As Date
Dim tItemMovEstoque As typeItemMovEstoque
Dim objItemNF As ClassItemNF
Dim iIndice As Integer
Dim lNumIntDocOrigem As Long
Dim lNumIntItemNF As Long
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_MovEstoqueItens_Le_Desc2

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 186960
    Next

    With tItemMovEstoque


        'Pesquisa os Ítens do Recebimento
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM ItensNFiscal WHERE NumIntNF = ? ORDER BY Item ", 0, lNumIntItemNF, lNumIntNF)
        If lErro <> AD_SQL_SUCESSO Then gError 186957

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186958

        Do While lErro = AD_SQL_SUCESSO

            Set objItemNF = New ClassItemNF
            
            objItemNF.lNumIntDoc = lNumIntItemNF
    
            lErro = CF("ItensNFGrade_Le", objItemNF)
            If lErro <> SUCESSO Then gError 186961
    
            If objItemNF.iPossuiGrade <> MARCADO Then
    
                .sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
                .sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
                .sProdutoOP = String(STRING_PRODUTO, 0)
                .sOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
                .sProduto = String(STRING_PRODUTO, 0)
                .sSiglaUM = String(STRING_UM_SIGLA, 0)
                .sCcl = String(STRING_CCL, 0)
                .sContaContabilEst = String(STRING_CONTA, 0)
                .sContaContabilAplic = String(STRING_CONTA, 0)
                .sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)
    
                lErro = Comando_Executar(alComando(2), "SELECT MovimentoEstoque.ContaContabilEst, MovimentoEstoque.ContaContabilAplic, MovimentoEstoque.Data, MovimentoEstoque.NumIntDoc, MovimentoEstoque.Custo, MovimentoEstoque.Apropriacao, MovimentoEstoque.Produto, Produtos.Descricao, MovimentoEstoque.SiglaUM, MovimentoEstoque.Quantidade, MovimentoEstoque.Almoxarifado, Almoxarifado.NomeReduzido, MovimentoEstoque.NumIntDocOrigem, MovimentoEstoque.Ccl, MovimentoEstoque.NumIntDocEst, MovimentoEstoque.TipoMov,MovimentoEstoque.DocOrigem,MovimentoEstoque.Cliente,MovimentoEstoque.Fornecedor FROM MovimentoEstoque, Produtos, Almoxarifado WHERE Almoxarifado.Codigo = MovimentoEstoque.Almoxarifado AND Produtos.Codigo = MovimentoEstoque.Produto AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.TipoNumIntDocOrigem = ?  AND MovimentoEstoque.NumIntDocOrigem = ? AND MovimentoEstoque.TipoMov = ? ORDER BY NumIntDoc DESC", _
                .sContaContabilEst, .sContaContabilAplic, dtData, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .iTipoMov, .sDocOrigem, .lCliente, .lFornecedor, objMovEstoque.iFilialEmpresa, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL, lNumIntItemNF, objMovEstoque.iTipoMov)
                If lErro <> AD_SQL_SUCESSO Then gError 186949
    
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186950
    
                If lErro = AD_SQL_SUCESSO Then
                
                    objMovEstoque.dtData = dtData
        
                    objMovEstoque.colItens.Add .lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, 0, Nothing, Nothing, DATA_NULA
                    objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).sDocOrigem = .sDocOrigem
                    objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lCliente = .lCliente
                    objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lFornecedor = .lFornecedor
                    objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).dtData = dtData
                    objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).iFilialEmpresa = objMovEstoque.iFilialEmpresa
    
                End If
    
            Else

                For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                    
                    .sProdutoDesc = String(STRING_PRODUTO_DESCRICAO, 0)
                    .sAlmoxarifadoNomeRed = String(STRING_ALMOXARIFADO_NOME_REDUZIDO, 0)
                    .sProdutoOP = String(STRING_PRODUTO, 0)
                    .sOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
                    .sProduto = String(STRING_PRODUTO, 0)
                    .sSiglaUM = String(STRING_UM_SIGLA, 0)
                    .sCcl = String(STRING_CCL, 0)
                    .sContaContabilEst = String(STRING_CONTA, 0)
                    .sContaContabilAplic = String(STRING_CONTA, 0)
                    .sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)

                    
                    lErro = Comando_Executar(alComando(3), "SELECT MovimentoEstoque.ContaContabilEst, MovimentoEstoque.ContaContabilAplic, MovimentoEstoque.Data, MovimentoEstoque.NumIntDoc, MovimentoEstoque.Custo, MovimentoEstoque.Apropriacao, MovimentoEstoque.Produto, Produtos.Descricao, MovimentoEstoque.SiglaUM, MovimentoEstoque.Quantidade, MovimentoEstoque.Almoxarifado, Almoxarifado.NomeReduzido, MovimentoEstoque.NumIntDocOrigem, MovimentoEstoque.Ccl, MovimentoEstoque.NumIntDocEst, MovimentoEstoque.TipoMov,MovimentoEstoque.DocOrigem,MovimentoEstoque.Cliente,MovimentoEstoque.Fornecedor FROM MovimentoEstoque, Produtos, Almoxarifado WHERE Almoxarifado.Codigo = MovimentoEstoque.Almoxarifado AND Produtos.Codigo = MovimentoEstoque.Produto AND MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.TipoNumIntDocOrigem = ?  AND MovimentoEstoque.NumIntDocOrigem = ? AND MovimentoEstoque.TipoMov = ? ORDER BY NumIntDoc DESC", _
                    .sContaContabilEst, .sContaContabilAplic, dtData, .lNumIntDoc, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .iTipoMov, .sDocOrigem, .lCliente, .lFornecedor, objMovEstoque.iFilialEmpresa, MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE, objItemRomaneio.lNumIntDoc, objMovEstoque.iTipoMov)
                    If lErro <> AD_SQL_SUCESSO Then gError 186949
    
                    lErro = Comando_BuscarPrimeiro(alComando(3))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186950
                    
                    If lErro = AD_SQL_SUCESSO Then
                        objMovEstoque.dtData = dtData
            
                        objMovEstoque.colItens.Add .lNumIntDoc, .iTipoMov, .dCusto, .iApropriacao, .sProduto, .sProdutoDesc, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .lNumIntDocOrigem, .sCcl, .lNumIntDocEst, .sOPCodigo, .sProdutoOP, .sContaContabilAplic, .sContaContabilEst, 0, Nothing, Nothing, DATA_NULA
                        objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).sDocOrigem = .sDocOrigem
                        objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lCliente = .lCliente
                        objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).lFornecedor = .lFornecedor
                        objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).dtData = dtData
                        objMovEstoque.colItens.Item(objMovEstoque.colItens.Count).iFilialEmpresa = objMovEstoque.iFilialEmpresa
                    End If
                
                Next
            
            End If
    
            'Não pode estar aqui porque o item pode não controlar estoque e tem que ler o movimentos dos demais para desfazer
            'If objMovEstoque.colItens.Count = 0 Then gError 186952
    
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 186959

        Loop

    End With
    
    'A rotina chamadora vai ser alterada para apenas não excluir os movimentos se não achar
    If objMovEstoque.colItens.Count = 0 Then gError 186952

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    MovEstoqueItens_Le_Desc2 = SUCESSO

    Exit Function

Erro_MovEstoqueItens_Le_Desc2:

    MovEstoqueItens_Le_Desc2 = gErr

    Select Case gErr

        Case 186949 To 186951
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case 186952, 186961

        Case 186957
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", Err)

        Case 186960
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 186953)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Produto_ObterRelMesFat_SemDev(ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, dValorFaturado As Double, dDesconto As Double, dQuantidadePedida As Double, dQuantidadeFaturada As Double, dQuantidadeDevolvida As Double, lComando As Long) As Long
'IMPORTANTE: especifica para classe ClassRelMesFat de apoio a execucao de relatorios
'Obtém as quantidades pedida e faturada, o desconto e valor total faturado por Produto entre datas
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long

On Error GoTo Erro_Produto_ObterRelMesFat_SemDev

    If lComando = 0 Then
        lComando = Comando_AbrirExt(GL_lConexaoRel)
        If lComando = 0 Then Error 43314
    End If
    
    If dtDataInicial <> DATA_NULA And dtDataFinal <> DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT SUM(ValorFaturado), SUM(TotalDescontos), SUM(QuantPedida), SUM(QuantFaturada), SUM(QuantDevolvida) FROM SldDiaFat WHERE FilialEmpresa = ? AND Produto = ? AND Data >= ? AND Data <= ?", dValorFaturado, dDesconto, dQuantidadePedida, dQuantidadeFaturada, dQuantidadeDevolvida, iFilialEmpresa, sProduto, dtDataInicial, dtDataFinal)
                   
    ElseIf dtDataInicial <> DATA_NULA And dtDataFinal = DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT SUM(ValorFaturado), SUM(TotalDescontos), SUM(QuantPedida), SUM(QuantFaturada), SUM(QuantDevolvida) FROM SldDiaFat WHERE FilialEmpresa = ? AND Produto = ? AND Data >= ?", dValorFaturado, dDesconto, dQuantidadePedida, dQuantidadeFaturada, dQuantidadeDevolvida, iFilialEmpresa, sProduto, dtDataInicial)
    
    ElseIf dtDataInicial = DATA_NULA And dtDataFinal <> DATA_NULA Then
        lErro = Comando_Executar(lComando, "SELECT SUM(ValorFaturado), SUM(TotalDescontos), SUM(QuantPedida), SUM(QuantFaturada), SUM(QuantDevolvida) FROM SldDiaFat WHERE FilialEmpresa = ? AND Produto = ? AND Data <= ?", dValorFaturado, dDesconto, dQuantidadePedida, dQuantidadeFaturada, dQuantidadeDevolvida, iFilialEmpresa, sProduto, dtDataFinal)
    
    Else
        lErro = Comando_Executar(lComando, "SELECT SUM(ValorFaturado), SUM(TotalDescontos), SUM(QuantPedida), SUM(QuantFaturada), SUM(QuantDevolvida) FROM SldDiaFat WHERE FilialEmpresa = ? AND Produto = ?", dValorFaturado, dDesconto, dQuantidadePedida, dQuantidadeFaturada, dQuantidadeDevolvida, iFilialEmpresa, sProduto)
        
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 43315
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43316
    
    If lErro = AD_SQL_SEM_DADOS Then
        
        dValorFaturado = 0
        dDesconto = 0
        dQuantidadePedida = 0
        dQuantidadeFaturada = 0
        dQuantidadeDevolvida = 0

    End If

    Produto_ObterRelMesFat_SemDev = SUCESSO
    
    Exit Function
    
Erro_Produto_ObterRelMesFat_SemDev:

    Produto_ObterRelMesFat_SemDev = Err
    
    Select Case Err

        Case 43314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 43315, 43316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAFAT_PROD", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 152062)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function ProdutoKit_Le_InfoUsu(ByVal objProdutoKit As ClassProdutoKit) As Long
    ProdutoKit_Le_InfoUsu = SUCESSO
End Function

'##################################################################
'Inserido por Wagner 23/02/07
Function RastreamentoLote_Le_Customizado(ByVal objRastroLote As ClassRastreamentoLote) As Long
    RastreamentoLote_Le_Customizado = SUCESSO
End Function

Function RastroLoteSaldo_Obter_QuantOP_Customizado(ByVal lNumIntDocLote As Long, dQuantOP As Double) As Long
    RastroLoteSaldo_Obter_QuantOP_Customizado = SUCESSO
End Function
'##################################################################

Public Function ProdutoCodBarras_Le(ByVal sCodBarras As String, ByVal objProduto As ClassProduto) As Long
'Le o codigo do produto associado ao codigo de barras

Dim lErro As Long
Dim lComando As Long
Dim sCodigo As String

On Error GoTo Erro_ProdutoCodBarras_Le

    'abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 193962

    'inicializa o buffer de recebimento
    sCodigo = String(STRING_PRODUTO, 0)

    'obtem o conjunto de codigos de barra relativos ao produto em questao
    lErro = Comando_Executar(lComando, "SELECT CodProduto FROM ProdutoCodBarras WHERE CodBarras = ?", sCodigo, sCodBarras)
    If lErro <> AD_SQL_SUCESSO Then gError 193963
    
    'se posiciona no 1o registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 193964
    
    If lErro = AD_SQL_SEM_DADOS Then gError 193965
    
    objProduto.sCodigo = sCodigo
    
    'fecha comando
    Call Comando_Fechar(lComando)
    
    ProdutoCodBarras_Le = SUCESSO
       
    Exit Function

Erro_ProdutoCodBarras_Le:

    ProdutoCodBarras_Le = gErr

    Select Case gErr
        
        Case 193962
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 193963, 193964
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, "ProdutoCodBarras")
        
        Case 193965
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193966)

    End Select

    'fecha comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function RastreamentoLoteLoc_Le(ByVal objRastreamentoLoteLoc As ClassRastreamentoLoteLoc) As Long

Dim lErro As Long
Dim lComando As Long
Dim tRastreamentoLoteLoc As typeRastreamentoLoteLoc

On Error GoTo Erro_RastreamentoLoteLoc_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 198706

    'Alocação de espaço no buffer
    tRastreamentoLoteLoc.sLocalizacao = String(STRING_RASTRO_LOCALIZACAO, 0)

    'Le a tabelaRastreamentoLoteLoc
    lErro = Comando_Executar(lComando, "SELECT Localizacao FROM RastreamentoLoteLoc WHERE Localizacao= ? ", _
                tRastreamentoLoteLoc.sLocalizacao, _
                objRastreamentoLoteLoc.sLocalizacao)
    If lErro <> AD_SQL_SUCESSO Then gError 198707

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 198708

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS

    objRastreamentoLoteLoc.sLocalizacao = tRastreamentoLoteLoc.sLocalizacao

    'Fecha Comando
    Call Comando_Fechar(lComando)

    RastreamentoLoteLoc_Le = SUCESSO

    Exit Function

Erro_RastreamentoLoteLoc_Le:

    RastreamentoLoteLoc_Le = gErr

    Select Case gErr

        Case 198706
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 198707, 198708
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTELOC", gErr)

        Case ERRO_LEITURA_SEM_DADOS 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198710)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ItemOP_Le_Cust(ByVal objItemOP As ClassItemOP) As Long
    ItemOP_Le_Cust = SUCESSO
End Function

Function ProdutoEmbalagem_Le1(objProdutoEmbalagem As ClassProdutoEmbalagem, sEmbalagem As String) As Long
'Le a UMEmbalagem e a capacidade a partir do Produto e da Embalagem

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_ProdutoEmbalagem_Le1

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 199114

    sEmbalagem = String(STRING_PRODUTO, 0)
    
    'Lê a embalagem padrao do Produto
    lErro = Comando_Executar(lComando, "SELECT E.Produto FROM ProdutoEmbalagem AS PE, Embalagens AS E WHERE PE.Embalagem = E.Codigo AND PE.Padrao = 1 AND PE.Produto = ?", _
        sEmbalagem, objProdutoEmbalagem.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 199115

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199116
    
    If lErro <> AD_SQL_SUCESSO Then gError 199117
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ProdutoEmbalagem_Le1 = SUCESSO

    Exit Function

Erro_ProdutoEmbalagem_Le1:

    ProdutoEmbalagem_Le1 = gErr

    Select Case gErr
       
        Case 199114
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 199115, 199116
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOEMBALAGEM", gErr, objProdutoEmbalagem.sProduto)
 
        Case 199117
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199118)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Sub Chama_Tela_ProducaoSaida(sNomeTela As String)
'Não faz nada
End Sub

Function Almoxarifado_Le_Todos(ByVal colAlmoxFilial As Collection) As Long
'Le todos os almoxaridados e coloca-os em colAlmoxFilial

Dim lComando As Long
Dim lErro As Long
Dim tAlmoxarifado As typeAlmoxarifado
Dim objAlmoxarifado As ClassAlmoxarifado

On Error GoTo Erro_Almoxarifado_Le_Todos

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 199272

    With tAlmoxarifado

        'Prepara variáveis string para receber conteúdo do BD
        .sDescricao = String(STRING_ALMOXARIFADO_DESCRICAO, 0)
        .sNomeReduzido = String(STRING_ALMOXARIFADO_NOMEREDUZIDO, 0)
    
        'Seleciona os dados da tabela Almoxarifado para a Filial em questão
        lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa, Descricao, Endereco, NomeReduzido FROM Almoxarifado ", _
        .iCodigo, .iFilialEmpresa, .sDescricao, .lEndereco, .sNomeReduzido)
        If lErro <> AD_SQL_SUCESSO Then gError 199273
    
        'Le o primeiro almoxarifado
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199274
    
        Do While lErro = SUCESSO
    
            Set objAlmoxarifado = New ClassAlmoxarifado
    
            'Passa dados lidos para objAlmoxarifado
            objAlmoxarifado.iCodigo = .iCodigo
            objAlmoxarifado.iFilialEmpresa = .iFilialEmpresa
            objAlmoxarifado.sDescricao = .sDescricao
            objAlmoxarifado.lEndereco = .lEndereco
            objAlmoxarifado.sNomeReduzido = .sNomeReduzido
            
            colAlmoxFilial.Add objAlmoxarifado

            'Ler o proximo almoxarifado
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 199275

        Loop

    End With

    'Fecha comando
    Call Comando_Fechar(lComando)

    Almoxarifado_Le_Todos = SUCESSO

    Exit Function

Erro_Almoxarifado_Le_Todos:

    Almoxarifado_Le_Todos = gErr

    Select Case gErr

        Case 199272
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 199273, 199274, 199275
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO2", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199276)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TRANSF_ProdutoLinha_Preenche(ByVal objTela As Object, objProduto As ClassProduto) As Long

    TRANSF_ProdutoLinha_Preenche = SUCESSO

End Function

Function Produto_ObterCustoRelMesFat(ByVal sProduto As String, ByVal dtDataInicial As Date, ByVal dtDataFinal As Date, ByVal iFilialEmpresa As Integer, dValorCusto As Double) As Long
'IMPORTANTE: especifica para classe ClassRelMesFat de apoio a execucao de relatorios
'Obtém as quantidades pedida e faturada, o desconto e valor total faturado por Produto entre datas
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim lErro As Long
Dim sSelect As String
Dim lComando As Long

On Error GoTo Erro_Produto_ObterCustoRelMesFat

    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 204869
    
    sSelect = "SELECT SUM(M.Custo) From MovimentoEstoque AS M, NFiscal As N, TiposDocInfo AS T , ItensNFiscal As I "
    sSelect = sSelect & "WHERE N.Status <> 7 AND T.Faturamento = 1 AND T.Codigo = N.TipoNFiscal AND "
    sSelect = sSelect & "M.TipoNumIntDocOrigem = 1 AND M.NumIntDocOrigem = I.NumIntDoc AND I.NumIntNF = N.NumINtDoc "
    
    
    If dtDataInicial <> DATA_NULA And dtDataFinal <> DATA_NULA Then
        lErro = Comando_Executar(lComando, sSelect & " AND N.FilialEmpresa = ? AND I.Produto = ? AND N.DataEmissao >= ? AND N.DataEmissao <= ?", dValorCusto, iFilialEmpresa, sProduto, dtDataInicial, dtDataFinal)
                   
    ElseIf dtDataInicial <> DATA_NULA And dtDataFinal = DATA_NULA Then
        lErro = Comando_Executar(lComando, sSelect & " AND N.FilialEmpresa = ? AND I.Produto = ? AND N.DataEmissao >= ?", dValorCusto, iFilialEmpresa, sProduto, dtDataInicial)
    
    ElseIf dtDataInicial = DATA_NULA And dtDataFinal <> DATA_NULA Then
        lErro = Comando_Executar(lComando, sSelect & " AND  N.FilialEmpresa = ? AND I.Produto = ? AND N.DataEmissao <= ?", dValorCusto, iFilialEmpresa, sProduto, dtDataFinal)
    
    Else
        lErro = Comando_Executar(lComando, sSelect & " AND N.FilialEmpresa = ? AND I.Produto = ?", dValorCusto, iFilialEmpresa, sProduto)
        
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 204870
        
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 204871
    
    If lErro = AD_SQL_SEM_DADOS Then
        
'        MsgBox ("nao encotrou dados Filial = " & iFilialEmpresa & " Produto = " & sProduto & " DataInicial = " & CStr(dtDataInicial) & " DataFinal = " & CStr(dtDataFinal))
        
        dValorCusto = 0

    End If

'    MsgBox ("valorcusto retornado = " & dValorCusto)
'    MsgBox ("sSelect = " & sSelect)

    Call Comando_Fechar(lComando)

    Produto_ObterCustoRelMesFat = SUCESSO
    
    Exit Function
    
Erro_Produto_ObterCustoRelMesFat:

    Produto_ObterCustoRelMesFat = gErr
    
    Select Case gErr

        Case 204869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 204870, 204871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 204872)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function MovEstoque_Le_MovSRV(ByVal iFilialEmpresa As Integer, ByVal sCodigoOS As String, ByVal colItemMovEst As Collection) As Long
'le as movimentecoes de estoque de saida de material para conserto associadas a OS

Dim alComando(1) As Long
Dim lErro As Long
Dim tItemMovEst As typeItemMovEstoque
Dim dFator As Double
Dim objItemMovEst As ClassItemMovEstoque
Dim iIndice As Integer
Dim sDescricao As String

On Error GoTo Erro_MovEstoque_Le_MovSRV

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 210542
    Next

    tItemMovEst.sProduto = String(STRING_PRODUTO, 0)
    tItemMovEst.sSiglaUM = String(STRING_UM_SIGLA, 0)
    sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)

    With tItemMovEst

        'le os movimentos de estoque de requisição de produto associados ao itemOP passado como parametro referentes aos produtos com natureza materia prima e que não estejam estornados e devolve a soma dos seus custos
        lErro = Comando_Executar(alComando(0), "SELECT MovimentoEstoque.TipoMov, MovimentoEstoque.Produto, MovimentoEstoque.Data, MovimentoEstoque.Hora, MovimentoEstoque.NumIntDoc, MovimentoEstoque.Quantidade, MovimentoEstoque.SiglaUM, MovimentoEstoque.Custo, Produtos.Descricao FROM MovimentoEstoque, Produtos WHERE MovimentoEstoque.FilialEmpresa = ? AND MovimentoEstoque.TipoMov = ? AND MovimentoEstoque.Produto = Produtos.Codigo AND MovimentoEstoque.CodigoOP = ? ORDER BY MovimentoEstoque.Produto, MovimentoEstoque.Data, MovimentoEstoque.Hora", _
        .iTipoMov, .sProduto, .dtData, .dHora, .lNumIntDoc, .dQuantidade, .sSiglaUM, .dCusto, sDescricao, iFilialEmpresa, MOV_EST_INT2_SAIDA_CONSERTO, sCodigoOS)
        If lErro <> AD_SQL_SUCESSO Then gError 210543

    End With

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210544

    Do While lErro = AD_SQL_SUCESSO
    
        Set objItemMovEst = New ClassItemMovEstoque
    
        objItemMovEst.iTipoMov = tItemMovEst.iTipoMov
        objItemMovEst.dQuantidade = tItemMovEst.dQuantidade
        objItemMovEst.sProduto = tItemMovEst.sProduto
        objItemMovEst.dtData = tItemMovEst.dtData
        objItemMovEst.dtHora = tItemMovEst.dHora
        objItemMovEst.lNumIntDoc = tItemMovEst.lNumIntDoc
        objItemMovEst.sSiglaUM = tItemMovEst.sSiglaUM
        objItemMovEst.dCusto = tItemMovEst.dCusto
        objItemMovEst.iNaturezaProduto = tItemMovEst.iNaturezaProduto
        objItemMovEst.sDescricao = sDescricao
   
        colItemMovEst.Add objItemMovEst
   
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210545
    
    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    MovEstoque_Le_MovSRV = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_Le_MovSRV:

    MovEstoque_Le_MovSRV = gErr

    Select Case gErr

        Case 210542
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 210543 To 210545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210546)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function RastroLoteSaldo_ObterNaData(ByVal objRastroLoteSaldo As ClassRastreamentoLoteSaldo, ByVal dtDataInicial As Date, alComando() As Long) As Long
'obtem o saldo inicial do lote-almoxarifado na data informada

Dim lErro As Long, iIndice As Integer, lNumIntDocLote As Long, dQuantidadeEst As Double, dFator As Double
Dim tRastroLoteSaldo As typeRastreamentoLoteSaldo, objTipoMovEstoque As New ClassTipoMovEst
Dim iTipoMov As Integer, dQuantidade As Double, sSiglaUM As String, sSiglaUMEstoque As String, iClasseUM As Integer

On Error GoTo Erro_RastroLoteSaldo_ObterNaData

    If alComando(0) = 0 Then
    
        For iIndice = LBound(alComando) To UBound(alComando)
        
            alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
            If alComando(0) = 0 Then gError 130543
            
        Next
        
    End If
    
    lNumIntDocLote = objRastroLoteSaldo.lNumIntDocLote
    If lNumIntDocLote = 0 Then
    
        With objRastroLoteSaldo
            lErro = Comando_Executar(alComando(0), "SELECT NumIntDoc FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", lNumIntDocLote, .sProduto, .sLote, .iFilialOP)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 130544
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130545
    
        'Não encontrou o lote ==> Erro
        If lErro = AD_SQL_SEM_DADOS Then gError 130546
        
    End If
    
    'obter o saldo atual
    With tRastroLoteSaldo
        lErro = Comando_Executar(alComando(1), "SELECT QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantEmpenhada, QuantPedida, " & _
            "QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, " & _
            "QuantOutras3, QuantOutras, QuantOP, QuantBenef, QuantBenef3 FROM RastreamentoLoteSaldo WHERE NumIntDocLote = ? AND Almoxarifado = ?", _
            .dQuantDispNossa, .dQuantReservada, .dQuantReservadaConsig, .dQuantEmpenhada, .dQuantPedida, _
            .dQuantRecIndl, .dQuantIndOutras, .dQuantDefeituosa, .dQuantConsig3, .dQuantConsig, .dQuantDemo3, .dQuantDemo, .dQuantConserto3, .dQuantConserto, _
            .dQuantOutras3, .dQuantOutras, .dQuantOP, .dQuantBenef, .dQuantBenef3, lNumIntDocLote, objRastroLoteSaldo.iAlmoxarifado)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 130547
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130548

    'Não encontrou o lote ==> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 130549
    
    'copiar dados
    With tRastroLoteSaldo
        objRastroLoteSaldo.dQuantDispNossa = .dQuantDispNossa
        objRastroLoteSaldo.dQuantReservada = .dQuantReservada
        objRastroLoteSaldo.dQuantReservadaConsig = .dQuantReservadaConsig
        objRastroLoteSaldo.dQuantEmpenhada = .dQuantEmpenhada
        objRastroLoteSaldo.dQuantPedida = .dQuantPedida
        objRastroLoteSaldo.dQuantRecIndl = .dQuantRecIndl
        objRastroLoteSaldo.dQuantIndOutras = .dQuantIndOutras
        objRastroLoteSaldo.dQuantDefeituosa = .dQuantDefeituosa
        objRastroLoteSaldo.dQuantConsig3 = .dQuantConsig3
        objRastroLoteSaldo.dQuantConsig = .dQuantConsig
        objRastroLoteSaldo.dQuantDemo3 = .dQuantDemo3
        objRastroLoteSaldo.dQuantDemo = .dQuantDemo
        objRastroLoteSaldo.dQuantConserto3 = .dQuantConserto3
        objRastroLoteSaldo.dQuantConserto = .dQuantConserto
        objRastroLoteSaldo.dQuantOutras3 = .dQuantOutras3
        objRastroLoteSaldo.dQuantOutras = .dQuantOutras
        objRastroLoteSaldo.dQuantOP = .dQuantOP
        objRastroLoteSaldo.dQuantBenef = .dQuantBenef
        objRastroLoteSaldo.dQuantBenef3 = .dQuantBenef3
    End With
        
    'descontar os movimentos até a data informada
    sSiglaUM = String(STRING_UM_SIGLA, 0)
    sSiglaUMEstoque = String(STRING_UM_SIGLA, 0)
    lErro = Comando_Executar(alComando(2), "SELECT MovimentoEstoque.TipoMov, RastreamentoMovto.Quantidade, RastreamentoMovto.SiglaUM, Produtos.SiglaUMEstoque, Produtos.ClasseUM FROM Produtos, RastreamentoMovto, MovimentoEstoque WHERE Produtos.Codigo = RastreamentoMovto.Produto AND RastreamentoMovto.NumIntDocOrigem = MovimentoEstoque.NumIntDoc AND RastreamentoMovto.TipoDocOrigem = ? AND RastreamentoMovto.NumIntDocLote = ? AND MovimentoEstoque.Data >= ? AND MovimentoEstoque.NumIntDocEst = 0 ORDER BY MovimentoEstoque.Data DESC", _
        iTipoMov, dQuantidade, sSiglaUM, sSiglaUMEstoque, iClasseUM, 0, lNumIntDocLote, dtDataInicial)
    If lErro <> AD_SQL_SUCESSO Then gError 130550
    
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130551
    
    Do While lErro = AD_SQL_SUCESSO
    
        'obter qtde do movimento na unidade de estoque do produto
        'Converte a unidade de medida da quantidade do Movimento de Estoque para a UM de Estoque
        lErro = CF("UM_Conversao_Trans", iClasseUM, sSiglaUM, sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 130552
        
        dQuantidadeEst = dQuantidade * dFator
        
        'para desfazer os movimentos
        dQuantidadeEst = -dQuantidadeEst
        
        objTipoMovEstoque.iCodigo = iTipoMov
        lErro = CF("TipoMovEstoque_Le", objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 130553
    
        'se a movimentação é referente a um conserto
        If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_ADICIONACONSERTO Then
            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                objRastroLoteSaldo.dQuantConserto = objRastroLoteSaldo.dQuantConserto + dQuantidadeEst
            Else
                objRastroLoteSaldo.dQuantConserto3 = objRastroLoteSaldo.dQuantConserto3 + dQuantidadeEst
            End If
        End If
        
        'se é uma movimentação referente a um conserto
        If objTipoMovEstoque.iAtualizaConserto = TIPOMOV_EST_SUBTRAICONSERTO Then
            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                objRastroLoteSaldo.dQuantConserto = objRastroLoteSaldo.dQuantConserto - dQuantidadeEst
            Else
                objRastroLoteSaldo.dQuantConserto3 = objRastroLoteSaldo.dQuantConserto3 - dQuantidadeEst
            End If
        End If
        
        'se a movimentação é referente a uma demonstração
        If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_ADICIONADEMO Then
            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                objRastroLoteSaldo.dQuantDemo = objRastroLoteSaldo.dQuantDemo + dQuantidadeEst
            Else
                objRastroLoteSaldo.dQuantDemo3 = objRastroLoteSaldo.dQuantDemo3 + dQuantidadeEst
            End If
        End If
        
        'se a movimentação é referente a uma demonstração
        If objTipoMovEstoque.iAtualizaDemo = TIPOMOV_EST_SUBTRAIDEMO Then
            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                objRastroLoteSaldo.dQuantDemo = objRastroLoteSaldo.dQuantDemo - dQuantidadeEst
            Else
                objRastroLoteSaldo.dQuantDemo3 = objRastroLoteSaldo.dQuantDemo3 - dQuantidadeEst
            End If
        End If
        
        'se a movimentação é referente a material em consignação
        If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_ADICIONACONSIGNACAO Then
            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                objRastroLoteSaldo.dQuantConsig = objRastroLoteSaldo.dQuantConsig + dQuantidadeEst
            Else
                objRastroLoteSaldo.dQuantConsig3 = objRastroLoteSaldo.dQuantConsig3 + dQuantidadeEst
            End If
        End If
        
        'se a movimentação é referente a material em consignação
        If objTipoMovEstoque.iAtualizaConsig = TIPOMOV_EST_SUBTRAICONSIGNACAO Then
            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                objRastroLoteSaldo.dQuantConsig = objRastroLoteSaldo.dQuantConsig - dQuantidadeEst
            Else
                objRastroLoteSaldo.dQuantConsig3 = objRastroLoteSaldo.dQuantConsig3 - dQuantidadeEst
            End If
        End If
        
        'se a movimentação é referente a outras movimentações de material
        If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_ADICIONAOUTRAS Then
            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                objRastroLoteSaldo.dQuantOutras = objRastroLoteSaldo.dQuantOutras + dQuantidadeEst
            Else
                objRastroLoteSaldo.dQuantOutras3 = objRastroLoteSaldo.dQuantOutras3 + dQuantidadeEst
            End If
        End If
        
        'se a movimentação é referente a outras movimentações de material
        If objTipoMovEstoque.iAtualizaOutras = TIPOMOV_EST_SUBTRAIOUTRAS Then
            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                objRastroLoteSaldo.dQuantOutras = objRastroLoteSaldo.dQuantOutras - dQuantidadeEst
            Else
                objRastroLoteSaldo.dQuantOutras3 = objRastroLoteSaldo.dQuantOutras3 - dQuantidadeEst
            End If
        End If
        
        'se a movimentação é referente a material em beneficiamento
        If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_ADICIONABENEF Then
            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                objRastroLoteSaldo.dQuantBenef = objRastroLoteSaldo.dQuantBenef + dQuantidadeEst
            Else
                objRastroLoteSaldo.dQuantBenef3 = objRastroLoteSaldo.dQuantBenef3 + dQuantidadeEst
            End If
        End If
        
        'se é uma movimentação referente a um conserto
        If objTipoMovEstoque.iAtualizaBenef = TIPOMOV_EST_SUBTRAIBENEF Then
            If objTipoMovEstoque.iProdutoDeTerc = TIPOMOV_EST_PRODUTONOSSO Then
                objRastroLoteSaldo.dQuantBenef = objRastroLoteSaldo.dQuantBenef - dQuantidadeEst
            Else
                objRastroLoteSaldo.dQuantBenef3 = objRastroLoteSaldo.dQuantBenef3 - dQuantidadeEst
            End If
        End If

        'se a movimentação é referente a outras movimentações de material indisponivel
        If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_ADICIONAINDOUTRAS Then
            objRastroLoteSaldo.dQuantIndOutras = objRastroLoteSaldo.dQuantIndOutras + dQuantidadeEst
        End If
        
        'se a movimentação é referente a outras movimentações de material indisponivel
        If objTipoMovEstoque.iAtualizaIndOutras = TIPOMOV_EST_SUBTRAIINDOUTRAS Then
            objRastroLoteSaldo.dQuantIndOutras = objRastroLoteSaldo.dQuantIndOutras - dQuantidadeEst
        End If
        
        'se a movimentação é referente a nosso material disponivel
        If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_ADICIONANOSSADISP Then
            objRastroLoteSaldo.dQuantDispNossa = objRastroLoteSaldo.dQuantDispNossa + dQuantidadeEst
        End If
        
        'se a movimentação é referente a nosso material disponivel
        If objTipoMovEstoque.iAtualizaNossaDisp = TIPOMOV_EST_SUBTRAINOSSADISP Then
            objRastroLoteSaldo.dQuantDispNossa = objRastroLoteSaldo.dQuantDispNossa - dQuantidadeEst
        End If
        
        'se a movimentação é referente a material defeituoso
        If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_ADICIONADEFEITUOSA Then
            objRastroLoteSaldo.dQuantDefeituosa = objRastroLoteSaldo.dQuantDefeituosa + dQuantidadeEst
        End If
        
        'se a movimentação é referente a material defeituoso
        If objTipoMovEstoque.iAtualizaDefeituosa = TIPOMOV_EST_SUBTRAIDEFEITUOSA Then
            objRastroLoteSaldo.dQuantDefeituosa = objRastroLoteSaldo.dQuantDefeituosa - dQuantidadeEst
        End If
        
        'se a movimentação é referente a material recebido e indisponível
        If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_ADICIONARECEBINDISP Then
            objRastroLoteSaldo.dQuantRecIndl = objRastroLoteSaldo.dQuantRecIndl + dQuantidadeEst
        End If
        
        'se a movimentação é referente a material recebido e indisponível
        If objTipoMovEstoque.iAtualizaRecebIndisp = TIPOMOV_EST_SUBTRAIRECEBINDISP Then
            objRastroLoteSaldo.dQuantRecIndl = objRastroLoteSaldo.dQuantRecIndl - dQuantidadeEst
        End If
        
        'se a movimentação é referente a material em ordem de producao
        If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_ADICIONAOP Then
            objRastroLoteSaldo.dQuantOP = objRastroLoteSaldo.dQuantOP + dQuantidadeEst
        End If
        
        'se a movimentação é referente a material em ordem de producao
        If objTipoMovEstoque.iAtualizaOP = TIPOMOV_EST_SUBTRAIOP Then
            objRastroLoteSaldo.dQuantOP = objRastroLoteSaldo.dQuantOP - dQuantidadeEst
        End If
    
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130554
    
    Loop
    
    RastroLoteSaldo_ObterNaData = SUCESSO
     
    Exit Function
    
Erro_RastroLoteSaldo_ObterNaData:

    RastroLoteSaldo_ObterNaData = gErr
     
    Select Case gErr
          
        Case 130543
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130544, 130545, 130546
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTROLOTE", gErr)
                    
        Case 130547, 130548, 130549
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTROLOTESALDO", gErr)
                    
        Case 130550, 130551, 130554
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTROMOVTO", gErr)
        
        Case 130552, 130553
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130542)
     
    End Select
     
    Exit Function

End Function

