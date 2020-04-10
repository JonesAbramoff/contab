VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'ALTERAÇÕES INPAL
'1 - Márcio - 08/2000 - Incluido preço Base nos ItensPV
'2 - Márcio - 09/2000 - Incluido preço Unitário, Cliente e Vendedor em PrevVenda

Option Explicit

Private Declare Function Conexao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Private Declare Function Conexao_FecharExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Function PVItens_Le_Customizado(ByVal objPV As ClassPedidoDeVenda) As Long
'Lê os valores customizados dos ítens do Pedido de Venda

Dim lErro As Long
Dim lComando As Long
Dim objItemPV As ClassItemPedido
Dim dPrecoBase As Double

On Error GoTo Erro_PVItens_Le_Customizado

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 141984

    For Each objItemPV In objPV.colItensPedido
        
        'Pesquisa os ítens vinculados ao pedido de venda
        lErro = Comando_Executar(lComando, "SELECT PrecoBase FROM ItensPedidoDeVenda WHERE NumIntDoc = ? ", dPrecoBase, objItemPV.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 141985

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141986

        If lErro <> AD_SQL_SUCESSO Then
        
            'Pesquisa os ítens vinculados à Nota Fiscal
            lErro = Comando_Executar(lComando, "SELECT PrecoBase FROM ItensPedidoDeVendaBaixados WHERE NumIntDoc = ? ", dPrecoBase, objItemPV.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 141987
    
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 141988
        
        End If
                
        If lErro = AD_SQL_SUCESSO Then
            objItemPV.dPrecoBase = dPrecoBase
        Else
            objItemPV.dPrecoBase = 0
        End If

    Next

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    PVItens_Le_Customizado = SUCESSO

    Exit Function

Erro_PVItens_Le_Customizado:

    PVItens_Le_Customizado = gErr

    Select Case gErr

        Case 141984
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 141985 To 141988
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSPEDIDODEVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141989)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'2 - Márcio - 09/2000 - Incluido preço Unitário, Cliente e Vendedor em PrevVenda
Function PrevVenda_Le(ByVal objPrevVenda As ClassPrevVenda) As Long
'Lê a Previsão de Venda a partir da chave fornecida em objPrevVenda. Retorna os dados em objPrevVenda

Dim lComando As Long
Dim lErro As Long
Dim tPrevVenda As typePrevVenda

On Error GoTo Erro_PrevVenda_Le

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 34523

    'Inicializa as variáveis com o tamanho do BD
    tPrevVenda.sCodigo = String(STRING_PREVVENDA_CODIGO, 0)
    tPrevVenda.sProduto = String(STRING_PRODUTO, 0)

    'Seleciona a Previsão de Venda cujo código foi passado em objPrevVenda
    lErro = Comando_Executar(lComando, "SELECT Codigo, DataPrevisao, Produto, Quantidade, Valor, DataInicio, DataFim, CodRegiao FROM PrevVenda WHERE FilialEmpresa = ? AND Codigo = ? AND Cliente = ? AND Vendedor = ? AND Produto = ?", tPrevVenda.sCodigo, tPrevVenda.dtDataPrevisao, tPrevVenda.sProduto, tPrevVenda.dQuantidade, tPrevVenda.dValor, tPrevVenda.dtDataInicio, tPrevVenda.dtDataFim, tPrevVenda.iCodRegiao, giFilialEmpresa, objPrevVenda.sCodigo, objPrevVenda.lCliente, objPrevVenda.iVendedor, objPrevVenda.sProduto)
    If lErro <> AD_SQL_SUCESSO Then Error 34524

    'Lê a Previsão de Venda
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 34525

    'Se não encontrou a previsão
    If lErro = AD_SQL_SEM_DADOS Then Error 34526

    'Completa objPrevVenda com os dados selecionados do BD
    With objPrevVenda
        
        .iFilialEmpresa = giFilialEmpresa
        .dtDataPrevisao = tPrevVenda.dtDataPrevisao
        .sProduto = tPrevVenda.sProduto
        .iAlmoxarifado = tPrevVenda.iAlmoxarifado
        .dQuantidade = tPrevVenda.dQuantidade
        .dValor = tPrevVenda.dValor
        .dtDataFim = tPrevVenda.dtDataFim
        .dtDataInicio = tPrevVenda.dtDataInicio
        .iCodRegiao = tPrevVenda.iCodRegiao
        
    End With

    Call Comando_Fechar(lComando)

    PrevVenda_Le = SUCESSO

    Exit Function

Erro_PrevVenda_Le:

    PrevVenda_Le = Err

    Select Case Err

        Case 34523
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 34524, 34525
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDA", Err, objPrevVenda.sCodigo)

        Case 34526 'previsão não encontrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'Funcoes Correspondentes a tela PlanComissoesInpal
'Colocadas aqui em: 20/03 por Tulio

Public Function ComissoesInpalPlan_Le(ByVal objComissoesInpalPlan As ClassComisInpalPlan) As Long
'Le os dados da planilha no BD...
'Deve ser passado como chave o codigo da planilha
'objPlanComissoesInpalPlan RECEBE (INPUT) o código da planilha
'e RETORNA (OUTPUT) as informacoes lidas do BD

Dim lErro As Long
Dim lComando As Long
Dim tComissoesInpalPlan As typeComissoesInpalPlan

On Error GoTo Erro_ComissoesInpalPlan_Le
    
    'inicializa comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 98795

    With tComissoesInpalPlan
    
        'obtem um conjunto unitario de registros da tabela ComissoesInpalPlan
        lErro = Comando_Executar(lComando, "SELECT Codigo, Vendedor, Tecnico, PercComissaoEmissao, PercComissaoBaixa, ComissaoSobreTotal, ComissaoFrete, ComissaoDesp, ComissaoIPI, ComissaoSeguro, AjudaCusto, TipoAjudaCusto FROM ComissoesInpalPlan WHERE Codigo = ?", .lCodigo, .iVendedor, .iTecnico, .dPercComissaoEmissao, .dPercComissaoBaixa, .iComissaoSobreTotal, .iComissaoFrete, .iComissaoDesp, .iComissaoIPI, .iComissaoSeguro, .dAjudaCusto, .iTipoAjudaCusto, objComissoesInpalPlan.lCodigo)

    End With

    If lErro <> AD_SQL_SUCESSO Then gError 98798

    'se posiciona no primeiro registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98799

    'Nao achou a planilha
    If lErro <> AD_SQL_SUCESSO Then gError 98762

    'Coloca o conteudo do type no obj
    With objComissoesInpalPlan
    
        .dAjudaCusto = tComissoesInpalPlan.dAjudaCusto
        .dPercComissaoBaixa = tComissoesInpalPlan.dPercComissaoBaixa
        .dPercComissaoEmissao = tComissoesInpalPlan.dPercComissaoEmissao
        .iComissaoFrete = tComissoesInpalPlan.iComissaoFrete
        .iComissaoDesp = tComissoesInpalPlan.iComissaoDesp
        .iComissaoIPI = tComissoesInpalPlan.iComissaoIPI
        .iComissaoSeguro = tComissoesInpalPlan.iComissaoSeguro
        .iComissaoSobreTotal = tComissoesInpalPlan.iComissaoSobreTotal
        .iTecnico = tComissoesInpalPlan.iTecnico
        .iTipoAjudaCusto = tComissoesInpalPlan.iTipoAjudaCusto
        .iVendedor = tComissoesInpalPlan.iVendedor
        .lCodigo = tComissoesInpalPlan.lCodigo
    
    End With
    
    'fecha comando
    Call Comando_Fechar(lComando)
    
    ComissoesInpalPlan_Le = SUCESSO
    
    Exit Function
    
Erro_ComissoesInpalPlan_Le:

    ComissoesInpalPlan_Le = gErr
    
    Select Case gErr
    
        Case 98762
        
        Case 98795
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 98798, 98799
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESINPALPLAN", gErr, objComissoesInpalPlan.lCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
    
    End Select
    
    'fecha comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ComissoesInpalRegras_Le_CodPlanilha(ByVal lCodPlanilha As Long, ByVal colComissoesRegras As Collection) As Long
'Le as regras de uma planilha no BD
'lNumIntPlanilha RECEBE(INPUT) o numero interno da planilha
'colComissoesRegras RETORNA(OUTPUT) as regras lidas do bd

Dim lErro As Long
Dim lComando As Long
Dim tRegras As typeComissoesInpalRegras
Dim objRegras As ClassComisInpalRegras

On Error GoTo Erro_ComissoesInpalRegras_Le_CodPlanilha
    
    'inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98801
    
    'inicializando strings
    tRegras.sCategoriaProduto = String(STRING_CATEGORIAPRODUTO_CATEGORIA, 0)
    tRegras.sItemCatProduto = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    
    With tRegras
    
        'obtem um conjunto de registros da tabela ComissoesInpalRegras
        lErro = Comando_Executar(lComando, "SELECT NumIntDoc, RegiaoVenda, Cliente, FilialCliente, CategoriaProduto, ItemCatProduto, PercTabelaA, PercTabelaB FROM ComissoesInpalRegras WHERE CodPlanilha = ?", .lNumIntDoc, .iRegiaoVenda, .lCliente, .iFilialCliente, .sCategoriaProduto, .sItemCatProduto, .dPercTabelaA, .dPercTabelaB, lCodPlanilha)

    End With

    If lErro <> AD_SQL_SUCESSO Then gError 98807

    'se posiciona no primeiro registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98808
    
    'Nao achou regra(s) relacionada(s) a planilha em questao
    If lErro <> AD_SQL_SUCESSO Then gError 98785
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'faz com que objregras aponte para uma nova area em memoria..
        Set objRegras = New ClassComisInpalRegras
    
        'coloca o conteudo lido anteriormente no obj
        With objRegras
            .dPercTabelaA = tRegras.dPercTabelaA
            .dPercTabelaB = tRegras.dPercTabelaB
            .iFilialCliente = tRegras.iFilialCliente
            .iRegiaoVenda = tRegras.iRegiaoVenda
            .lCliente = tRegras.lCliente
            .lNumIntDoc = tRegras.lNumIntDoc
            .lCodPlanilha = tRegras.lCodPlanilha
            .sCategoriaProduto = tRegras.sCategoriaProduto
            .sItemCatProduto = tRegras.sItemCatProduto
        End With
        
        'adiciona o obj na colecao para poder voltar com os dados
        colComissoesRegras.Add objRegras
        
        'se posiciona no proximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98810
            
    Loop
    
    'fecha comando
    Call Comando_Fechar(lComando)
    
    ComissoesInpalRegras_Le_CodPlanilha = SUCESSO
    
    Exit Function
    
Erro_ComissoesInpalRegras_Le_CodPlanilha:

    ComissoesInpalRegras_Le_CodPlanilha = gErr
    
    Select Case gErr
    
        Case 98801
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 98807, 98808, 98810
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESINPALREGRAS", gErr)
    
        Case 98785 'SEM DADOS
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    'fecha comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function ComissoesInpalRegras_Le_Tabelas_Lock(ByVal objComissoesInpalRegras As ClassComisInpalRegras) As Long
'Faz os locks nas tabelas relacionadas com a regra
'objComissoesInpalRegras RECEBE (INPUT) os dados das tabelas nas quais o lock
'sera feito...

Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_ComissoesInpalRegras_Le_Tabelas_Lock

    'abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98886
    
    'Se cliente estiver preenchido
    If objComissoesInpalRegras.lCliente > 0 Then
        
        'Se a filial estiver preenchida
        If objComissoesInpalRegras.iFilialCliente > 0 Then
            
            'locka a filial
            lErro = CF("FilialCliente_Lock", objComissoesInpalRegras.lCliente, objComissoesInpalRegras.iFilialCliente)
            If lErro <> SUCESSO Then gError 98887
            
        Else
        
            'Locka o Cliente
            lErro = CF("Cliente_Lock", objComissoesInpalRegras.lCliente)
            If lErro <> SUCESSO Then gError 98888
    
        End If
        
    End If
    
    'Se a regiao estiver preenchida...
    If objComissoesInpalRegras.iRegiaoVenda > 0 Then
        
        'locka a regiao de venda
        lErro = CF("RegiaoVenda_Lock", lComando, objComissoesInpalRegras.iRegiaoVenda)
        If lErro <> SUCESSO Then gError 98889

    End If
    
    'se a categoria de produto estiver preenchida
    If Len(Trim(objComissoesInpalRegras.sCategoriaProduto)) > 0 Then
        
        'locka o item de categoria
        lErro = CF("CategoriaProdutoItem_Lock", objComissoesInpalRegras.sCategoriaProduto, SCodigo_Extrai(objComissoesInpalRegras.sItemCatProduto), lComando)
        If lErro <> SUCESSO Then gError 98890
    
    End If
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    ComissoesInpalRegras_Le_Tabelas_Lock = SUCESSO
    
    Exit Function

Erro_ComissoesInpalRegras_Le_Tabelas_Lock:

    ComissoesInpalRegras_Le_Tabelas_Lock = gErr

    Select Case gErr
    
        Case 98886
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 98887, 98888, 98889, 98890
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Tulio subiu em 24/04/02
Public Function ComissoesInpalPlan_Le_Completa(ByVal objComissoesInpalPlan As ClassComisInpalPlan) As Long
'Le os dados do bd referente a planilha e suas regras...
'objComissoesInpalPlan RETORNA (OUTPUT) os dados do BD

Dim lErro As Long

On Error GoTo Erro_ComissoesInpalPlan_Le_Completa

    'le os dados da planilha
    lErro = CF("ComissoesInpalPlan_Le", objComissoesInpalPlan)
    If lErro <> SUCESSO And lErro <> 98762 Then gError 98791
    
    'se nao achou planilha
    If lErro <> SUCESSO Then gError 98792
        
    'le as regras referentes a planilha lida anteriormente
    lErro = CF("ComissoesInpalRegras_Le_CodPlanilha", objComissoesInpalPlan.lCodigo, objComissoesInpalPlan.colComissoesInpalRegras)
    If lErro <> SUCESSO And lErro <> 98785 Then gError 98793
    
    'se nao achou regra alguma...
    If lErro <> SUCESSO Then gError 98794
        
    ComissoesInpalPlan_Le_Completa = SUCESSO
    
    Exit Function

Erro_ComissoesInpalPlan_Le_Completa:

    ComissoesInpalPlan_Le_Completa = gErr
    
    Select Case gErr
    
        Case 98791, 98792, 98793, 98794
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select
    
    Exit Function

End Function

Public Function ComissoesInpalPlan_Le_Vendedor_Tecnico(ByVal objComissoesInpalPlan As ClassComisInpalPlan) As Long
'Le os dados da planilha no BD... a partir de um vendedor diretou ou de um vendedor indireto
'vendedor + flag direto/indireto (tecnico) = chave candidata da tabela de plancomissoesinpal
'Deve ser passado como chave o codigo da planilha
'objPlanComissoesInpalPlan RECEBE (INPUT) o vendedor e a flag tecnico
'e RETORNA (OUTPUT) as informacoes lidas do BD

Dim lErro As Long
Dim lComando As Long
Dim tComissoesInpalPlan As typeComissoesInpalPlan

On Error GoTo Erro_ComissoesInpalPlan_Le_Vendedor_Tecnico
    
    'inicializa comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 101695

    With tComissoesInpalPlan
    
        'obtem um conjunto unitario de registros da tabela ComissoesInpalPlan
        lErro = Comando_Executar(lComando, "SELECT Codigo, Vendedor, Tecnico, PercComissaoEmissao, PercComissaoBaixa, ComissaoSobreTotal, ComissaoFrete, ComissaoDesp, ComissaoIPI, ComissaoSeguro, AjudaCusto, TipoAjudaCusto FROM ComissoesInpalPlan WHERE Vendedor = ? AND Tecnico = ?", .lCodigo, .iVendedor, .iTecnico, .dPercComissaoEmissao, .dPercComissaoBaixa, .iComissaoSobreTotal, .iComissaoFrete, .iComissaoDesp, .iComissaoIPI, .iComissaoSeguro, .dAjudaCusto, .iTipoAjudaCusto, objComissoesInpalPlan.iVendedor, objComissoesInpalPlan.iTecnico)

    End With

    If lErro <> AD_SQL_SUCESSO Then gError 101696

    'se posiciona no primeiro registro do conjunto obtido anteriormente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 101697

    'Nao achou a planilha
    If lErro <> AD_SQL_SUCESSO Then gError 101698

    'Coloca o conteudo do type no obj
    With objComissoesInpalPlan
    
        .dAjudaCusto = tComissoesInpalPlan.dAjudaCusto
        .dPercComissaoBaixa = tComissoesInpalPlan.dPercComissaoBaixa
        .dPercComissaoEmissao = tComissoesInpalPlan.dPercComissaoEmissao
        .iComissaoFrete = tComissoesInpalPlan.iComissaoFrete
        .iComissaoDesp = tComissoesInpalPlan.iComissaoDesp
        .iComissaoIPI = tComissoesInpalPlan.iComissaoIPI
        .iComissaoSeguro = tComissoesInpalPlan.iComissaoSeguro
        .iComissaoSobreTotal = tComissoesInpalPlan.iComissaoSobreTotal
        .iTecnico = tComissoesInpalPlan.iTecnico
        .iTipoAjudaCusto = tComissoesInpalPlan.iTipoAjudaCusto
        .iVendedor = tComissoesInpalPlan.iVendedor
        .lCodigo = tComissoesInpalPlan.lCodigo
    
    End With
    
    'fecha comando
    Call Comando_Fechar(lComando)
    
    ComissoesInpalPlan_Le_Vendedor_Tecnico = SUCESSO
    
    Exit Function
    
Erro_ComissoesInpalPlan_Le_Vendedor_Tecnico:

    ComissoesInpalPlan_Le_Vendedor_Tecnico = gErr
    
    Select Case gErr
    
        Case 101695
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 101696, 101697
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESINPALPLAN", gErr, objComissoesInpalPlan.lCodigo)
    
        Case 101698
            'nao encontrou
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
    
    End Select
    
    'fecha comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Incluído por Daniel
Function PrevVenda_Le2(ByVal objPrevVenda As ClassPrevVenda) As Long
'Lê a Previsão de Venda a partir da chave fornecida em objPrevVenda.
'Retorna os dados em objPrevVenda
'ATENÇÃO : Deve ser usada apenas para saber se existe uma PrevVenda com o Código/Filial passados
' a inpal tem essa função customizada

Dim lComando As Long
Dim lErro As Long
Dim tPrevVenda As typePrevVenda

On Error GoTo Erro_PrevVenda_Le2

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 108660

    'Inicializa as variáveis com o tamanho do BD
    tPrevVenda.sCodigo = String(STRING_PREVVENDA_CODIGO, 0)
    tPrevVenda.sProduto = String(STRING_PRODUTO, 0)

    'Seleciona a Previsão de Venda cujo código foi passado em objPrevVenda
    lErro = Comando_Executar(lComando, "SELECT Codigo, Produto, CodRegiao FROM PrevVendaMensal WHERE FilialEmpresa = ? AND Codigo = ?", tPrevVenda.sCodigo, tPrevVenda.sProduto, tPrevVenda.iCodRegiao, giFilialEmpresa, objPrevVenda.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 108661

    'Lê a Previsão de Venda
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108662

    'Se não encontrou a previsão
    If lErro = AD_SQL_SEM_DADOS Then gError 108663

    'Completa objPrevVenda com os dados selecionados do BD
    With objPrevVenda
        
        .iFilialEmpresa = giFilialEmpresa
        .sProduto = tPrevVenda.sProduto
        .iCodRegiao = tPrevVenda.iCodRegiao
        
    End With

    Call Comando_Fechar(lComando)

    PrevVenda_Le2 = SUCESSO

    Exit Function

Erro_PrevVenda_Le2:

    PrevVenda_Le2 = gErr

    Select Case gErr

        Case 108660
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 108661, 108662
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PREVVENDA", gErr, objPrevVenda.sCodigo)

        Case 108663 'previsão não encontrada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Call Comando_Fechar(lComando)

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'LEITURA
Public Function RegTick_Le(ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal colRegTick As Collection) As Long

Dim lErro As Long
Dim lComandoExt As Long
Dim lComando As Long
Dim lConexaoExt As Long
Dim tRegTick As typeRegTick
Dim objRegTick As ClassRegTick
Dim lNumIntDocNF As Long
Dim sDSN As String
Dim iLenDSN As Integer
Dim sParamOut As String, sCodItem As String, sTick_SttFim As String, sTick_Status As String, sTick_RecExp As String
Dim iLenParamOut As Integer, lUtltlTick_Indice As Long

On Error GoTo Erro_RegTick_Le

    If dtDataIni < CDate("03/11/2006") Then gError 184124
    
    sDSN = "DSN=SGEBalanca;UID=;PWD=I;ReadOnly=1;Exclusive=1"
    iLenDSN = Len(sDSN)
    sParamOut = String(1024, 0)
    iLenParamOut = 1024
    lUtltlTick_Indice = 0
    
    lConexaoExt = Conexao_AbrirExt(AD_SQL_DRIVER_ODBC, sDSN, iLenDSN, sParamOut, iLenParamOut)
    If lConexaoExt = 0 Then gError 181919

    'Executa a abertura do Comando
    lComandoExt = Comando_AbrirExt(lConexaoExt)
    If lComandoExt = 0 Then gError 181901

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181906
    
    dtDataFim = DateAdd("d", 1, dtDataFim)

    'Alocação de espaço no buffer
    tRegTick.sTick_CampoUsu2 = String(255, 0)
    tRegTick.sCgc = String(255, 0)
    sCodItem = String(255, 0)
    sTick_SttFim = String(255, 0)
    sTick_Status = String(255, 0)
    sTick_RecExp = String(255, 0)
    
    'Le a tabelaRegTick
    lErro = Comando_Executar(lComandoExt, "SELECT Tick_Indice, Tick_DtHrPesoFinal, Tick_LiquidoCorrigido, Tick_PesoLiqCorrUsu , Tick_CampoUsu2, Emiss_CNPJ, Tick_CodItem, Tick_SttFim, Tick_Status, Tick_RecExp FROM RegTick, CadEmiss WHERE Tick_CodEmissor = Emiss_Codigo AND Tick_DtHrPesoFinal >= ? AND Tick_DtHrPesoFinal < ? ORDER BY Tick_Indice ASC, Emiss_DataUltAtlz DESC", _
                tRegTick.lTick_Indice, tRegTick.dtTick_DtHrPesoFinal, tRegTick.dTick_LiquidoCorrigido, tRegTick.dTick_PesoLiqCorrUsu, tRegTick.sTick_CampoUsu2, tRegTick.sCgc, sCodItem, sTick_SttFim, sTick_Status, sTick_RecExp, dtDataIni, dtDataFim)
    If lErro <> AD_SQL_SUCESSO Then gError 181902

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComandoExt)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO_PARCIAL Then gError 181903

    'Sem Dados
    Do While lErro <> AD_SQL_SEM_DADOS
    
        If lUtltlTick_Indice <> tRegTick.lTick_Indice Then

            'para evitar repetir qdo o emissor está com codigo repetido no cadastro
            lUtltlTick_Indice = tRegTick.lTick_Indice
            
            If sCodItem = "001" And sTick_SttFim = "S" And sTick_Status = "F" And sTick_RecExp = "R" Then
            
                lErro = Comando_Executar(lComando, "SELECT NF.NumIntDoc FROM NF_RAIZ AS R, NFiscal AS NF WHERE R.NumIntDocNF = NF.NumIntDoc AND R.Indice_RegTick = ?", lNumIntDocNF, tRegTick.lTick_Indice)
                If lErro <> AD_SQL_SUCESSO Then gError 181907
        
                'Busca Primeiro
                lErro = Comando_BuscarPrimeiro(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181908
                
                If lErro = AD_SQL_SEM_DADOS Then
        
                    Set objRegTick = New ClassRegTick
                
                    objRegTick.lTick_Indice = tRegTick.lTick_Indice
                    objRegTick.dtTick_DtHrPesoFinal = tRegTick.dtTick_DtHrPesoFinal
                    objRegTick.dTick_LiquidoCorrigido = tRegTick.dTick_LiquidoCorrigido
                    objRegTick.sTick_CampoUsu2 = tRegTick.sTick_CampoUsu2
                    objRegTick.dTick_PesoLiqCorrUsu = tRegTick.dTick_PesoLiqCorrUsu
                    objRegTick.sCgc = tRegTick.sCgc
                    
                    colRegTick.Add objRegTick
                    
                End If
    
            End If
        
        End If
        
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComandoExt)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO_PARCIAL Then gError 181904

    Loop

    'Fecha Comando
    Call Comando_Fechar(lComandoExt)
    Call Comando_Fechar(lComando)
    
    Call Conexao_FecharExt(lConexaoExt)

    RegTick_Le = SUCESSO

    Exit Function

Erro_RegTick_Le:

    RegTick_Le = gErr

    Select Case gErr

        Case 181901, 181906
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181902, 181903
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGTICK", gErr)

        Case 184124
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_REGTICK", gErr)
            
        Case 181904 'Sem dados -> Tratado na rotina chamadora
        
        Case 181919

        Case 181907, 181908
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181905)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComandoExt)
    Call Comando_Fechar(lComando)

    Call Conexao_FecharExt(lConexaoExt)

    Exit Function

End Function

Public Function ImportacaoNFRaiz_Le_Data(dtData As Date) As Long

Dim lErro As Long
Dim lComando As Long
Dim dtDataEmissao As Date

On Error GoTo Erro_ImportacaoNFRaiz_Le_Data

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181964
            
    lErro = Comando_Executar(lComando, "SELECT NF.DataEmissao FROM NF_RAIZ AS R, NFiscal AS NF WHERE R.NumIntDocNF = NF.NumIntDoc ORDER BY NF.DataEmissao DESC", dtDataEmissao)
    If lErro <> AD_SQL_SUCESSO Then gError 181965

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181966
        
    If lErro <> AD_SQL_SEM_DADOS Then
        dtData = dtDataEmissao
    Else
        dtData = DATA_NULA
    End If

    Call Comando_Fechar(lComando)

    ImportacaoNFRaiz_Le_Data = SUCESSO

    Exit Function

Erro_ImportacaoNFRaiz_Le_Data:

    ImportacaoNFRaiz_Le_Data = gErr

    Select Case gErr

        Case 181964
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 181965, 181966
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Empresa_Le_Nome_Customizado(sNomeEmpresa As String) As Long
    sNomeEmpresa = FormataCpoTexto("INPAL S.A. Industrias Quimicas", 30)
    Empresa_Le_Nome_Customizado = SUCESSO
End Function

Function CodConvBanco_Obtem_Customizado(ByVal objCarteiraCobrador As ClassCarteiraCobrador, sCodConvBanco As String) As Long
    If objCarteiraCobrador.iNumCarteiraNoBanco = 1 Then
        sCodConvBanco = FormataCpoTexto("000027328001411019", 20)
    Else
        sCodConvBanco = FormataCpoTexto("000027328001431019", 20)
    End If
    CodConvBanco_Obtem_Customizado = SUCESSO
End Function

Function BB_Obter_Dados_DetalheSegP_Customizado(ByVal gobjCobrancaEletronica As ClassCobrancaEletronica, ByVal objOcorrRemParcRem As ClassOcorrRemParcRec, ByVal objCarteiraCobrador As ClassCarteiraCobrador, ByVal objParcelaRec As ClassParcelaReceber, sTipoDocumento As String, sNumDocCobranca As String, sDVAgencia As String, sCodProtesto As String, sNumDiasProtesto As String) As Long
    
    sTipoDocumento = "1"
    sNumDocCobranca = CStr(FormataCpoNum(gobjCobrancaEletronica.lNumTitulo, 6)) & IIf(gobjCobrancaEletronica.iFilialEmpTitulo = 1, "UNC", IIf(gobjCobrancaEletronica.iFilialEmpTitulo = 2, "UN ", "UNP")) & IIf(objParcelaRec.iNumParcela = 1, "A", IIf(objParcelaRec.iNumParcela = 2, "B", "C")) & FormataCpoTexto("", 5)
    sDVAgencia = FormataCpoTexto("", 1)

    sCodProtesto = "3"
    sNumDiasProtesto = FormataCpoNum(0, 2)

    BB_Obter_Dados_DetalheSegP_Customizado = SUCESSO
End Function

Function Itau_Obter_Dados_DetalheSegP_Customizado(ByVal gobjCobrancaEletronica As ClassCobrancaEletronica, sNumDocCobranca As String, ByVal objParcelaRec As ClassParcelaReceber, Optional ByVal objOcorrRemParcRem As ClassOcorrRemParcRec, Optional sInstrucao1 As String, Optional sInstrucao2 As String) As Long
    
    sNumDocCobranca = CStr(FormataCpoNum(gobjCobrancaEletronica.lNumTitulo, 6)) & IIf(gobjCobrancaEletronica.iFilialEmpTitulo = 1, "UNC", IIf(gobjCobrancaEletronica.iFilialEmpTitulo = 2, "UN ", "UNP")) & IIf(objParcelaRec.iNumParcela = 1, "A", IIf(objParcelaRec.iNumParcela = 2, "B", "C"))
    
    If Not (objOcorrRemParcRem Is Nothing) Then
        If objOcorrRemParcRem.iInstrucao1 = 0 Then
            sInstrucao1 = "01"
            sInstrucao2 = "05"
        End If
    End If

    Itau_Obter_Dados_DetalheSegP_Customizado = SUCESSO
End Function

Function CobrancaEletronica_Obter_DadosParcelaRec_Cust(ByVal objCobrancaEletronica As ClassCobrancaEletronica) As Long
    objCobrancaEletronica.iEspecieTitulo = 1
    objCobrancaEletronica.sCGCCedente = "33413527000105"
    CobrancaEletronica_Obter_DadosParcelaRec_Cust = SUCESSO
End Function

Function TipoNFiscal_Testar_Credito(ByVal iTipoNFiscal As Integer, iTestar As Integer) As Long

Dim lErro As Long, iValor As Integer
Dim lComando As Long

On Error GoTo Erro_TipoNFiscal_Testar_Credito

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 138925

    lErro = Comando_Executar(lComando, "SELECT Valor FROM TipoNFiscalNaoTestarCredito WHERE Codigo = ?", iValor, iTipoNFiscal)
    If lErro <> AD_SQL_SUCESSO Then gError 138926

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 138927

    'Sem Dados
    If lErro = AD_SQL_SEM_DADOS Then
        iTestar = 1
    Else
        iTestar = iValor
    End If
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    TipoNFiscal_Testar_Credito = SUCESSO

    Exit Function

Erro_TipoNFiscal_Testar_Credito:

    TipoNFiscal_Testar_Credito = gErr

    Select Case gErr

        Case 138925
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 138926, 138927
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPONFNAOTESTACREDITO", gErr)

        Case 138928 'Sem dados -> Tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function EncComissoesCliProd_Cust(ByVal objVendedor As ClassVendedor, dFator As Double) As Long
Dim lMatricula As Long
    If IsNumeric(objVendedor.sMatricula) Then
        lMatricula = StrParaLong(objVendedor.sMatricula)
    End If
    If lMatricula >= 1261 Then
        dFator = (0.59 / 0.83)
    Else
        dFator = 0.59
    End If
    EncComissoesCliProd_Cust = SUCESSO
End Function
