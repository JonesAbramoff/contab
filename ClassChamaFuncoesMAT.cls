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
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function TipoDeProduto_Le(objTipoDeProduto As ClassTipoDeProduto) As Long
'Lê a tabela TiposDeProduto, a partir do Tipo De Produto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeProduto_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TIPODEPRODUTO_LE", X)
    If lErro <> SUCESSO Then Error 24890
    
    
    TipoDeProduto_Le = X.TipoDeProduto_Le(objTipoDeProduto)

    Exit Function
    
Erro_TipoDeProduto_Le:
    
    TipoDeProduto_Le = Err
    
    Select Case Err

        Case 24890
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145896)
        
    End Select
    
    Exit Function

End Function


Function TipoDeProduto_Le_Categorias(objTipoDeProduto As ClassTipoDeProduto, colCategoriaItem As Collection) As Long
'Lê na tabela de TiposDeProdutoCategoria, todas as categorias de um Tipo de Produto, guardando na coleção colCategoriaItem

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeProduto_Le_Categorias
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TIPODEPRODUTO_LE_CATEGORIAS", X)
    If lErro <> SUCESSO Then Error 24891
    
    
    TipoDeProduto_Le_Categorias = X.TipoDeProduto_Le_Categorias(objTipoDeProduto, colCategoriaItem)

    Exit Function
    
Erro_TipoDeProduto_Le_Categorias:
    
    TipoDeProduto_Le_Categorias = Err
    
    Select Case Err

        Case 24891
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145897)
        
    End Select
    
    Exit Function

End Function


Function TipoDeProduto_Exclui(objTipoDeProduto As ClassTipoDeProduto) As Long
'Exclui o Tipo de Produto
'Se o Tipo de Produto estiver sendo usado em Produtos não poderá ser excluído

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeProduto_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TIPODEPRODUTO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 24893
    
    
    TipoDeProduto_Exclui = X.TipoDeProduto_Exclui(objTipoDeProduto)

    Exit Function
    
Erro_TipoDeProduto_Exclui:
    
    TipoDeProduto_Exclui = Err
    
    Select Case Err

        Case 24893
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145898)
        
    End Select
    
    Exit Function

End Function

Function TipoDeProduto_Grava(objTipoDeProduto As ClassTipoDeProduto) As Long
'Inclui ou altera um Tipo de Produto
'Tabelas:TiposDeProduto e TipoDeProdutoCategoria

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeProduto_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TIPODEPRODUTO_GRAVA", X)
    If lErro <> SUCESSO Then Error 24894
    
    
    TipoDeProduto_Grava = X.TipoDeProduto_Grava(objTipoDeProduto)

    Exit Function
    
Erro_TipoDeProduto_Grava:
    
    TipoDeProduto_Grava = Err
    
    Select Case Err

        Case 24894
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145899)
        
    End Select
    
    Exit Function

End Function

Function TipoDeProduto_Lock(objTipoDeProduto As ClassTipoDeProduto) As Long
'Lê e faz "lock" na tabela TiposDeProduto, a partir do Tipo De Produto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeProduto_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TIPODEPRODUTO_LOCK", X)
    If lErro <> SUCESSO Then Error 19673
    
    
    TipoDeProduto_Lock = X.TipoDeProduto_Lock(objTipoDeProduto)

    Exit Function
    
Erro_TipoDeProduto_Lock:
    
    TipoDeProduto_Lock = Err
    
    Select Case Err

        Case 19673
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145900)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_Grava(objMovEst As ClassMovEstoque) As Long
'grava as informações relativas a um conjunto de movimentações de estoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_GRAVA", X)
    If lErro <> SUCESSO Then Error 25073
    
    
    MovEstoque_Grava = X.MovEstoque_Grava(objMovEst)

    Exit Function
    
Erro_MovEstoque_Grava:
    
    MovEstoque_Grava = Err
    
    Select Case Err

        Case 25073
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145901)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_Grava1(alComando() As Long, objMovEst As ClassMovEstoque) As Long
'grava as informações relativas a um conjunto de movimentações de estoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Grava1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_GRAVA1", X)
    If lErro <> SUCESSO Then Error 44285
    
    
    MovEstoque_Grava1 = X.MovEstoque_Grava1(alComando(), objMovEst)

    Exit Function
    
Erro_MovEstoque_Grava1:
    
    MovEstoque_Grava1 = Err
    
    Select Case Err

        Case 44285
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145902)
        
    End Select
    
    Exit Function

End Function


Function MovEstoque_Grava_Trans(objMovEst As ClassMovEstoque) As Long
'grava as informações relativas a um conjunto de movimentações de estoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Grava_Trans
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_GRAVA_TRANS", X)
    If lErro <> SUCESSO Then Error 25074
    
    
    MovEstoque_Grava_Trans = X.MovEstoque_Grava_Trans(objMovEst)

    Exit Function
    
Erro_MovEstoque_Grava_Trans:
    
    MovEstoque_Grava_Trans = Err
    
    Select Case Err

        Case 25074
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145903)
        
    End Select
    
    Exit Function

End Function

Function UM_Conversao(iClasseUM As Integer, sUnidade1 As String, sUnidade2 As String, dFator As Double) As Long
'Calcula o fator de conversão entre 2 Unidades de Medida
'De sUnidade1 para sUnidade2

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UM_Conversao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_UM_CONVERSAO", X)
    If lErro <> SUCESSO Then Error 25067
    
    
    UM_Conversao = X.UM_Conversao(iClasseUM, sUnidade1, sUnidade2, dFator)

    Exit Function
    
Erro_UM_Conversao:
    
    UM_Conversao = Err
    
    Select Case Err

        Case 25067
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145904)
        
    End Select
    
    Exit Function

End Function
Function AlmoxarifadosProduto_Le(sProduto As String, colAlmoxarifado As Collection) As Long
'Lê os códigos de todos os Almoxarifados que tem o Produto cadastrado,
'mesmo que a quantidade disponível seja zero.
'Entra com código do produto em sProduto, retorna colAlmoxarifado
'Filtra os Almoxarifados do Tipo Normal
Dim X As Object, lErro As Long
    
On Error GoTo Erro_AlmoxarifadosProduto_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ALMOXARIFADOSPRODUTO_LE", X)
    If lErro <> SUCESSO Then Error 25128
    
    
    AlmoxarifadosProduto_Le = X.AlmoxarifadosProduto_Le(sProduto, colAlmoxarifado)

    Exit Function
    
Erro_AlmoxarifadosProduto_Le:
    
    AlmoxarifadosProduto_Le = Err
    
    Select Case Err

        Case 25128
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145905)
        
    End Select
    
    Exit Function

End Function
Function EstoquesProduto_Le(sProduto As String, colEstoque As colEstoqueProduto) As Long
'Lê nas tabelas de EstoqueProduto e Almoxarifado as posições do produto
'Retorna os Almoxarifados com os produtos em questão em colEstoque
'Filtra apenas os Estoques que têm quantidade positiva do produto
'Filtra os Almoxarifados do Tipo Normal
Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoquesProduto_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUESPRODUTO_LE", X)
    If lErro <> SUCESSO Then Error 25122
    
    
    EstoquesProduto_Le = X.EstoquesProduto_Le(sProduto, colEstoque)

    Exit Function
    
Erro_EstoquesProduto_Le:
    
    EstoquesProduto_Le = Err
    
    Select Case Err

        Case 25122
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 145906)
        
    End Select
    
    Exit Function

End Function

Function Almoxarifado_Le(objAlmoxarifado As ClassAlmoxarifado) As Long
'Le no BD os dados do Almoxarifado cujo código foi passado como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Almoxarifado_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ALMOXARIFADO_LE", X)
    If lErro <> SUCESSO Then Error 25061
    
    
    Almoxarifado_Le = X.Almoxarifado_Le(objAlmoxarifado)

    Exit Function
    
Erro_Almoxarifado_Le:
    
    Almoxarifado_Le = Err
    
    Select Case Err

        Case 25061
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145907)
        
    End Select
    
    Exit Function

End Function
Function Almoxarifado_Le_NomeReduzido(objAlmoxarifado As ClassAlmoxarifado) As Long
'Le todos no BD os dados do Almoxarifado cujo NomeReduzido foi passado por parâmetro
Dim X As Object, lErro As Long
    
On Error GoTo Erro_Almoxarifado_Le_NomeReduzido
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ALMOXARIFADO_LE_NOME_REDUZIDO", X)
    If lErro <> SUCESSO Then Error 25062
    
    
    Almoxarifado_Le_NomeReduzido = X.Almoxarifado_Le_NomeReduzido(objAlmoxarifado)

    Exit Function
    
Erro_Almoxarifado_Le_NomeReduzido:
    
    Almoxarifado_Le_NomeReduzido = Err
    
    Select Case Err

        Case 25062
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145908)
        
    End Select
    
    Exit Function

End Function
Function Produtos_Le_Reserva(colProduto As Collection) As Long
'Lê Codigo e Descricao de todos os produtos ATIVOS que fazem RESERVA e coloca-os em colProduto
Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produtos_Le_Reserva
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTOS_LE_RESERVA", X)
    If lErro <> SUCESSO Then Error 25051
    
    
    Produtos_Le_Reserva = X.Produtos_Le_Reserva(colProduto)

    Exit Function
    
Erro_Produtos_Le_Reserva:
    
    Produtos_Le_Reserva = Err
    
    Select Case Err

        Case 25051
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145909)
        
    End Select
    
    Exit Function

End Function

Function Produto_Le_Todos(colProduto As Collection) As Long
'Lê todos os produtos e coloca-os em colProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LE_TODOS", X)
    If lErro <> SUCESSO Then Error 24269
    
    
    Produto_Le_Todos = X.Produto_Le_Todos(colProduto)

    Exit Function
    
Erro_Produto_Le_Todos:
    
    Produto_Le_Todos = Err
    
    Select Case Err

        Case 24269
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145910)
        
    End Select
    
    Exit Function

End Function
Function Inicializa_Mascara_Produto_MaskEd(MaskEdProduto As Object) As Long
'inicializa a mascara de Produto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Inicializa_Mascara_Produto_MaskEd
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INICIALIZA_MASCARA_PRODUTO_MASKED", X)
    If lErro <> SUCESSO Then Error 24272
    
    
    Inicializa_Mascara_Produto_MaskEd = X.Inicializa_Mascara_Produto_MaskEd(MaskEdProduto)

    Exit Function
    
Erro_Inicializa_Mascara_Produto_MaskEd:
    
    Inicializa_Mascara_Produto_MaskEd = Err
    
    Select Case Err

        Case 24272
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145911)
        
    End Select
    
    Exit Function

End Function

Public Function Produto_Perde_Foco(Produto As Object, Desc As Object) As Long
'recebe MaskEdBox do Produto e o label da descrição

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Perde_Foco
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_PERDE_FOCO", X)
    If lErro <> SUCESSO Then Error 24273
    
    
    Produto_Perde_Foco = X.Produto_Perde_Foco(Produto, Desc)

    Exit Function
    
Erro_Produto_Perde_Foco:
    
    Produto_Perde_Foco = Err
    
    Select Case Err

        Case 24273
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145912)
        
    End Select
    
    Exit Function

End Function

Public Function Traz_Produto_MaskEd(sProduto As String, MaskEdProduto As Object, Label As Object) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Traz_Produto_MaskEd
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TRAZ_PRODUTO_MASKED", X)
    If lErro <> SUCESSO Then Error 24274
    
    
    Traz_Produto_MaskEd = X.Traz_Produto_MaskEd(sProduto, MaskEdProduto, Label)

    Exit Function
    
Erro_Traz_Produto_MaskEd:
    
    Traz_Produto_MaskEd = Err
    
    Select Case Err

        Case 24274
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145913)
        
    End Select
    
    Exit Function

End Function
Function Produto_Critica_Estoque(ByVal sProduto As String, objProduto As ClassProduto, iProdutoPreenchido As Integer) As Long
'Verifica se o Produto existe é ATIVO e não é GERENCIAL e participa do ESTOQUE.
'Devolve os dados do Produto em objProduto
'Devolve PRODUTO_PREENCHIDO ou PRODUTO_VAZIO em iProdutoPreenchido
Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Critica_Estoque
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_CRITICA_ESTOQUE", X)
    If lErro <> SUCESSO Then Error 25081
    
    
    Produto_Critica_Estoque = X.Produto_Critica_Estoque(sProduto, objProduto, iProdutoPreenchido)

    Exit Function
    
Erro_Produto_Critica_Estoque:
    
    Produto_Critica_Estoque = Err
    
    Select Case Err

        Case 25081
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145914)
        
    End Select
    
    Exit Function

End Function

Function ProdutoCategorias_Le_Produto(sProduto As String, colProdutoCategoria As Collection) As Long
'preenche uma colecao com objProdutoCategoria(s) associados a sProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutoCategorias_Le_Produto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTOCATEGORIAS_LE_PRODUTO", X)
    If lErro <> SUCESSO Then Error 24598
    
    
    ProdutoCategorias_Le_Produto = X.ProdutoCategorias_Le_Produto(sProduto, colProdutoCategoria)

    Exit Function
    
Erro_ProdutoCategorias_Le_Produto:
    
    ProdutoCategorias_Le_Produto = Err
    
    Select Case Err

        Case 24598
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145915)
        
    End Select
    
    Exit Function

End Function

Function EstoqueProduto_Le(objEstoqueProduto As ClassEstoqueProduto) As Long
'Le as informacoes sobre o produto passado em objEstoqueProduto na tabela EstoqueProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_LE", X)
    If lErro <> SUCESSO Then Error 24692
    
    
    EstoqueProduto_Le = X.EstoqueProduto_Le(objEstoqueProduto)

    Exit Function
    
Erro_EstoqueProduto_Le:
    
    EstoqueProduto_Le = Err
    
    Select Case Err

        Case 24692
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145916)
        
    End Select
    
    Exit Function

End Function

Function ProdutoFilial_Le(objProdutoFilial As ClassProdutoFilial) As Long
'Lê os dados do Produto da Filial à partir da chave fornecida em objProdutoFilial. Retorna os dados em objProdutoFilial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutoFilial_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTOFILIAL_LE", X)
    If lErro <> SUCESSO Then Error 24693
    
    
    ProdutoFilial_Le = X.ProdutoFilial_Le(objProdutoFilial)

    Exit Function
    
Erro_ProdutoFilial_Le:
    
    ProdutoFilial_Le = Err
    
    Select Case Err

        Case 24693
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145917)
        
    End Select
    
    Exit Function

End Function

Function UnidadesDeMedidas_Le_ClasseUM(objClasseUM As ClassClasseUM, colUM As Collection) As Long
'carrega a colUM com objetos do tipo ClassUnidadeDeMedida correspondentes a objClasseUM

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UnidadesDeMedidas_Le_ClasseUM
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_UMS_LE_CLASSE", X)
    If lErro <> SUCESSO Then Error 27499
    
    
    UnidadesDeMedidas_Le_ClasseUM = X.UnidadesDeMedidas_Le_ClasseUM(objClasseUM, colUM)

    Exit Function
    
Erro_UnidadesDeMedidas_Le_ClasseUM:
    
    UnidadesDeMedidas_Le_ClasseUM = Err
    
    Select Case Err

        Case 27499
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145918)
        
    End Select
    
    Exit Function

End Function

Function ClasseUM_Le(objClasseUM As ClassClasseUM) As Long
'Lê na tabela ClasseUM a Descrição e a Sigla da classe, a partir do código da classe

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClasseUM_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSEUM_LE", X)
    If lErro <> SUCESSO Then Error 24870
    
    
    ClasseUM_Le = X.ClasseUM_Le(objClasseUM)

    Exit Function
    
Erro_ClasseUM_Le:
    
    ClasseUM_Le = Err
    
    Select Case Err

        Case 24870
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145919)
        
    End Select
    
    Exit Function

End Function

Function ClasseUM_Grava(objClasseUM As ClassClasseUM, colSiglas As Collection) As Long
'Inclui ou altera uma Classe de Unidade de Medida
'Tabelas:ClasseUM e UnidadeDeMedida

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClasseUM_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSEUM_GRAVA", X)
    If lErro <> SUCESSO Then Error 24871
    
    
    ClasseUM_Grava = X.ClasseUM_Grava(objClasseUM, colSiglas)

    Exit Function
    
Erro_ClasseUM_Grava:
    
    ClasseUM_Grava = Err
    
    Select Case Err

        Case 24871
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145920)
        
    End Select
    
    Exit Function

End Function

Function ClasseUM_Grava_NovasSiglas(objClasseUM As ClassClasseUM, colSiglas As Collection) As Long
'Percorre as Siglas na coleção incluindo-as no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClasseUM_Grava_NovasSiglas
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSEUM_GRAVA_NOVASSIGLAS", X)
    If lErro <> SUCESSO Then Error 24872
    
    
    ClasseUM_Grava_NovasSiglas = X.ClasseUM_Grava_NovasSiglas(objClasseUM, colSiglas)

    Exit Function
    
Erro_ClasseUM_Grava_NovasSiglas:
    
    ClasseUM_Grava_NovasSiglas = Err
    
    Select Case Err

        Case 24872
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145921)
        
    End Select
    
    Exit Function

End Function

Function ClasseUM_Exclui(objClasseUM As ClassClasseUM) As Long
'Exclui a Classe UM
'Se a Classe estiver sendo usada em Produtos ou TiposDeProduto não poderá ser excluída

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClasseUM_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSEUM_EXCLUI", X)
    If lErro <> SUCESSO Then Error 24873
    
    
    ClasseUM_Exclui = X.ClasseUM_Exclui(objClasseUM)

    Exit Function
    
Erro_ClasseUM_Exclui:
    
    ClasseUM_Exclui = Err
    
    Select Case Err

        Case 24873
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145922)
        
    End Select
    
    Exit Function

End Function

Function ClasseUM_Exclui_UM(objClasseUM As ClassClasseUM) As Long
'Exclui as Unidades de Medida de uma Classe

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClasseUM_Exclui_UM
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSEUM_EXCLUI_UM", X)
    If lErro <> SUCESSO Then Error 24874
    
    
    ClasseUM_Exclui_UM = X.ClasseUM_Exclui_UM(objClasseUM)

    Exit Function
    
Erro_ClasseUM_Exclui_UM:
    
    ClasseUM_Exclui_UM = Err
    
    Select Case Err

        Case 24874
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145923)
        
    End Select
    
    Exit Function

End Function

'Function ClasseUM_Exclui2(objClasseUM As ClassClasseUM) As Long
''Retorna SUCESSO se a Classe e a Sigla não estiverem sendo usadas nas tabelas Produtos , TiposDeProduto e ItensPedidoDeVenda
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ClasseUM_Exclui2
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSEUM_EXCLUI2",X)
'    If lErro <> SUCESSO Then Error 24875
'
'
'    ClasseUM_Exclui2 = X.ClasseUM_Exclui2(objClasseUM)
'
'    Exit Function
'
'Erro_ClasseUM_Exclui2:
'
'    ClasseUM_Exclui2 = Err
'
'    Select Case Err
'
'        Case 24875
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145924)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function ClasseUM_Exclui1(objClasseUM As ClassClasseUM) As Long
'''Retorna SUCESSO se a Classe não estiver sendo usada nas tabelas Produtos e TiposDeProduto
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ClasseUM_Exclui1
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSEUM_EXCLUI1",X)
'    If lErro <> SUCESSO Then Error 24876
'
'
'    ClasseUM_Exclui1 = X.ClasseUM_Exclui1(objClasseUM)
'
'    Exit Function
'
'Erro_ClasseUM_Exclui1:
'
'    ClasseUM_Exclui1 = Err
'
'    Select Case Err
'
'        Case 24876
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145925)
'
'    End Select
'
'    Exit Function
'
'End Function
'
Function ClasseUM_Automatico(iProxClasse As Integer) As Long
'Gera o próximo código para Classe de UM
'Chamada dentro de transação

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClasseUM_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSEUM_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 24877
    
    
    ClasseUM_Automatico = X.ClasseUM_Automatico(iProxClasse)

    Exit Function
    
Erro_ClasseUM_Automatico:
    
    ClasseUM_Automatico = Err
    
    Select Case Err

        Case 24877
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145926)
        
    End Select
    
    Exit Function

End Function

Function ClasseUM_Grava_SiglaNaColecao(colUMCopia As Collection, sSigla As String) As Long
'Retorna o Indice da Sigla na coleção ou Zero se não acha-la (na coleção)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClasseUM_Grava_SiglaNaColecao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSEUM_GRAVA_SIGLANACOLECAO", X)
    If lErro <> SUCESSO Then Error 24878
    
    
    ClasseUM_Grava_SiglaNaColecao = X.ClasseUM_Grava_SiglaNaColecao(colUMCopia, sSigla)

    Exit Function
    
Erro_ClasseUM_Grava_SiglaNaColecao:
    
    ClasseUM_Grava_SiglaNaColecao = Err
    
    Select Case Err

        Case 24878
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145927)
        
    End Select
    
    Exit Function

End Function

Function CategoriaProduto_Grava_NovosItens(objCategoriaProduto As ClassCategoriaProduto, colItensCategoria As Collection) As Long
'percorrer os itens na colecao incluindo-os no bd

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaProduto_Grava_NovosItens
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CATEGORIAPRODUTO_GRAVA_NOVOSITENS", X)
    If lErro <> SUCESSO Then Error 24409
    
    
    CategoriaProduto_Grava_NovosItens = X.CategoriaProduto_Grava_NovosItens(objCategoriaProduto, colItensCategoria)

    Exit Function
    
Erro_CategoriaProduto_Grava_NovosItens:
    
    CategoriaProduto_Grava_NovosItens = Err
    
    Select Case Err

        Case 24409
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145928)
        
    End Select
    
    Exit Function

End Function

Function CategoriaProduto_Grava_ItemNaColecao(colItensCategoria As Collection, sItem As String) As Long
'retorna o indice do item na colecao ou zero se nao acha-lo (na colecao)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaProduto_Grava_ItemNaColecao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CATEGORIAPRODUTO_GRAVA_ITEMNACOLECAO", X)
    If lErro <> SUCESSO Then Error 24410
    
    
    CategoriaProduto_Grava_ItemNaColecao = X.CategoriaProduto_Grava_ItemNaColecao(colItensCategoria, sItem)

    Exit Function
    
Erro_CategoriaProduto_Grava_ItemNaColecao:
    
    CategoriaProduto_Grava_ItemNaColecao = Err
    
    Select Case Err

        Case 24410
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145929)
        
    End Select
    
    Exit Function

End Function

Function CategoriaProduto_Grava(objCategoriaProduto As ClassCategoriaProduto, colItensCategoria As Collection) As Long
'inclui ou altera uma categoria de produtos e seus valores
'tabelas:CategoriaProduto e CategoriaProdutoItem

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaProduto_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CATEGORIAPRODUTO_GRAVA", X)
    If lErro <> SUCESSO Then Error 24411
    
    
    CategoriaProduto_Grava = X.CategoriaProduto_Grava(objCategoriaProduto, colItensCategoria)

    Exit Function
    
Erro_CategoriaProduto_Grava:
    
    CategoriaProduto_Grava = Err
    
    Select Case Err

        Case 24411
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145930)
        
    End Select
    
    Exit Function

End Function

Function CategoriaProduto_Exclui(objCategoriaProduto As ClassCategoriaProduto) As Long
'excluir a tabela com todos os seus itens se nao for usada em ProdutoCategoria

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaProduto_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CATEGORIAPRODUTO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 24412
    
    
    CategoriaProduto_Exclui = X.CategoriaProduto_Exclui(objCategoriaProduto)

    Exit Function
    
Erro_CategoriaProduto_Exclui:
    
    CategoriaProduto_Exclui = Err
    
    Select Case Err

        Case 24412
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145931)
        
    End Select
    
    Exit Function

End Function

Function CategoriaProduto_Exclui1(objCategoriaProduto As ClassCategoriaProduto) As Long
'retorna SUCESSO se a categoria nao estiver sendo usada por nenhum produto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaProduto_Exclui1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CATEGORIAPRODUTO_EXCLUI1", X)
    If lErro <> SUCESSO Then Error 24413
    
    
    CategoriaProduto_Exclui1 = X.CategoriaProduto_Exclui1(objCategoriaProduto)

    Exit Function
    
Erro_CategoriaProduto_Exclui1:
    
    CategoriaProduto_Exclui1 = Err
    
    Select Case Err

        Case 24413
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145932)
        
    End Select
    
    Exit Function

End Function

Function CategoriaProduto_ExcluiItens(objCategoriaProduto As ClassCategoriaProduto) As Long
'exclui os itens de uma categoria (se nao for utilizado ??? )

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaProduto_ExcluiItens
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CATEGORIAPRODUTO_EXCLUIITENS", X)
    If lErro <> SUCESSO Then Error 24414
    
    
    CategoriaProduto_ExcluiItens = X.CategoriaProduto_ExcluiItens(objCategoriaProduto)

    Exit Function
    
Erro_CategoriaProduto_ExcluiItens:
    
    CategoriaProduto_ExcluiItens = Err
    
    Select Case Err

        Case 24414
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145933)
        
    End Select
    
    Exit Function

End Function

Function CategoriaProdutoItem_NaoUtilizado(sCategoria As String, sItem As String) As Long
'retorna SUCESSO se nenhum produto tiver o valor de sItem na categoria identificada por sCategoria

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaProdutoItem_NaoUtilizado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CATEGORIAPRODUTOITEM_NAOUTILIZADO", X)
    If lErro <> SUCESSO Then Error 24415
    
    
    CategoriaProdutoItem_NaoUtilizado = X.CategoriaProdutoItem_NaoUtilizado(sCategoria, sItem)

    Exit Function
    
Erro_CategoriaProdutoItem_NaoUtilizado:
    
    CategoriaProdutoItem_NaoUtilizado = Err
    
    Select Case Err

        Case 24415
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145934)
        
    End Select
    
    Exit Function

End Function

Function CategoriaProduto_Le(objCategoriaProduto As ClassCategoriaProduto) As Long
'Le na tabela CategoriaProduto a descrição da categoria

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaProduto_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CATEGORIAPRODUTO_LE", X)
    If lErro <> SUCESSO Then Error 24406
    
    
    CategoriaProduto_Le = X.CategoriaProduto_Le(objCategoriaProduto)

    Exit Function
    
Erro_CategoriaProduto_Le:
    
    CategoriaProduto_Le = Err
    
    Select Case Err

        Case 24406
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145935)
        
    End Select
    
    Exit Function

End Function

Function CategoriasProduto_Le_Todas(colCategorias As Collection) As Long
'Le a tabela CategoriaProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriasProduto_Le_Todas
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CATEGORIASPRODUTO_LE_TODAS", X)
    If lErro <> SUCESSO Then Error 24407
    
    
    CategoriasProduto_Le_Todas = X.CategoriasProduto_Le_Todas(colCategorias)

    Exit Function
    
Erro_CategoriasProduto_Le_Todas:
    
    CategoriasProduto_Le_Todas = Err
    
    Select Case Err

        Case 24407
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145936)
        
    End Select
    
    Exit Function

End Function

Function CategoriaProduto_Le_Itens(objCategoriaProduto As ClassCategoriaProduto, colItensCategoria As Collection) As Long
'Le a tabela de CategoriaProdutoItem

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaProduto_Le_Itens
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CATEGORIAPRODUTO_LE_ITENS", X)
    If lErro <> SUCESSO Then Error 24408
    
    
    CategoriaProduto_Le_Itens = X.CategoriaProduto_Le_Itens(objCategoriaProduto, colItensCategoria)

    Exit Function
    
Erro_CategoriaProduto_Le_Itens:
    
    CategoriaProduto_Le_Itens = Err
    
    Select Case Err

        Case 24408
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145937)
        
    End Select
    
    Exit Function

End Function

Function CategoriaProduto_Le_Item(objCategoriaProdutoItem As ClassCategoriaProdutoItem) As Long
'Le na tabela de CategoriaProdutoItem a Categoria e o Item de uma deteminada Categoria de Produto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaProduto_Le_Item
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CATEGORIAPRODUTO_LE_ITEM", X)
    If lErro <> SUCESSO Then Error 24892
    
    
    CategoriaProduto_Le_Item = X.CategoriaProduto_Le_Item(objCategoriaProdutoItem)

    Exit Function
    
Erro_CategoriaProduto_Le_Item:
    
    CategoriaProduto_Le_Item = Err
    
    Select Case Err

        Case 24892
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145938)
        
    End Select
    
    Exit Function

End Function

Function Kit_Grava(objKit As ClassKit) As Long
'grava os registros correspondentes em ProdutoKit e Kit

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Kit_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_KIT_GRAVA", X)
    If lErro <> SUCESSO Then Error 24895
    
    
    Kit_Grava = X.Kit_Grava(objKit)

    Exit Function
    
Erro_Kit_Grava:
    
    Kit_Grava = Err
    
    Select Case Err

        Case 24895
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145939)
        
    End Select
    
    Exit Function

End Function

Function Kit_Exclui(objKit As ClassKit) As Long
'recebe um produto raiz e exclui suas composições

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Kit_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_KIT_EXCLUI", X)
    If lErro <> SUCESSO Then Error 24896
    
    
    Kit_Exclui = X.Kit_Exclui(objKit)

    Exit Function
    
Erro_Kit_Exclui:
    
    Kit_Exclui = Err
    
    Select Case Err

        Case 24896
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145940)
        
    End Select
    
    Exit Function

End Function

Function Kit_Le(objKit As ClassKit) As Long
'Le a tabela de Kits a partir do produto e da versão passados em objKit

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Kit_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_KIT_LE", X)
    If lErro <> SUCESSO Then Error 24897
    
    
    Kit_Le = X.Kit_Le(objKit)

    Exit Function
    
Erro_Kit_Le:
    
    Kit_Le = Err
    
    Select Case Err

        Case 24897
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145941)
        
    End Select
    
    Exit Function

End Function

Function Kit_Le_MaisRecente(objKit As ClassKit) As Long
'Le uma versão na tabela de Kits de Produtos a partir do produto passado em objKit e devolve os dados do Kit em objKit

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Kit_Le_MaisRecente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_KIT_LE_MAISRECENTE", X)
    If lErro <> SUCESSO Then Error 24898
    
    
    Kit_Le_MaisRecente = X.Kit_Le_MaisRecente(objKit)

    Exit Function
    
Erro_Kit_Le_MaisRecente:
    
    Kit_Le_MaisRecente = Err
    
    Select Case Err

        Case 24898
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145942)
        
    End Select
    
    Exit Function

End Function

Function Kit_Le_Componentes(objKit As ClassKit) As Long
'Preenche objKit.colComponentes a partir do produto e da versão passados em objKit

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Kit_Le_Componentes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_KIT_LE_COMPONENTES", X)
    If lErro <> SUCESSO Then Error 24899
    
    
    Kit_Le_Componentes = X.Kit_Le_Componentes(objKit)

    Exit Function
    
Erro_Kit_Le_Componentes:
    
    Kit_Le_Componentes = Err
    
    Select Case Err

        Case 24899
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145943)
        
    End Select
    
    Exit Function

End Function

Function Estoque_AtualizaMovimento(objMovEst As ClassMovEstoque) As Long
'grava as informações relativas a um conjunto de movimentações de estoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Estoque_AtualizaMovimento
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUE_ATUALIZAMOVIMENTO", X)
    If lErro <> SUCESSO Then Error 20445
    
    
    Estoque_AtualizaMovimento = X.Estoque_AtualizaMovimento(objMovEst)

    Exit Function
    
Erro_Estoque_AtualizaMovimento:
    
    Estoque_AtualizaMovimento = Err
    
    Select Case Err

        Case 20445
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145944)
        
    End Select
    
    Exit Function

End Function

Function Estoque_AtualizaMovimento_Trans(objMovEst As ClassMovEstoque) As Long
'grava as informações relativas a um conjunto de movimentações de estoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Estoque_AtualizaMovimento_Trans
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUE_ATUALIZAMOVIMENTO_TRANS", X)
    If lErro <> SUCESSO Then Error 20446
    
    
    Estoque_AtualizaMovimento_Trans = X.Estoque_AtualizaMovimento_Trans(objMovEst)

    Exit Function
    
Erro_Estoque_AtualizaMovimento_Trans:
    
    Estoque_AtualizaMovimento_Trans = Err
    
    Select Case Err

        Case 20446
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145945)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst_Le_Status(ByVal lComando As Long, objSldMesEst As ClassSldMesEst) As Long
'le os dados de SldMesEst relativos ao Ano, FilialEmpresa e Produto passados como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst_Le_Status
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST_LE_STATUS", X)
    If lErro <> SUCESSO Then Error 20447
    
    
    SldMesEst_Le_Status = X.SldMesEst_Le_Status(lComando, objSldMesEst)

    Exit Function
    
Erro_SldMesEst_Le_Status:
    
    SldMesEst_Le_Status = Err
    
    Select Case Err

        Case 20447
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145946)
        
    End Select
    
    Exit Function

End Function

Function TiposMovEst_Le1(ByVal lComando As Long, objTipoMovEstoque As ClassTipoMovEst) As Long
'Le a tabela de tipos de movimentos de estoque a partir do codigo fornecido em objTipoMovEstoque e devolve os dados neste mesmo obj.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposMovEst_Le1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TIPOSMOVEST_LE1", X)
    If lErro <> SUCESSO Then Error 20448
    
    
    TiposMovEst_Le1 = X.TiposMovEst_Le1(lComando, objTipoMovEstoque)

    Exit Function
    
Erro_TiposMovEst_Le1:
    
    TiposMovEst_Le1 = Err
    
    Select Case Err

        Case 20448
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145947)
        
    End Select
    
    Exit Function

End Function

Function MovimentoEstoque_Le_Estorno(ByVal lComando As Long, objItemMovEst As ClassItemMovEstoque) As Long
'Verifica se o movimento a ser estornado está cadastrado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovimentoEstoque_Le_Estorno
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVIMENTOESTOQUE_LE_ESTORNO", X)
    If lErro <> SUCESSO Then Error 20449
    
    
    MovimentoEstoque_Le_Estorno = X.MovimentoEstoque_Le_Estorno(lComando, objItemMovEst)

    Exit Function
    
Erro_MovimentoEstoque_Le_Estorno:
    
    MovimentoEstoque_Le_Estorno = Err
    
    Select Case Err

        Case 20449
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145948)
        
    End Select
    
    Exit Function

End Function

Function MovimentoEstoque_Excluir(ByVal lComando1 As Long, ByVal lComando2 As Long, objItemMovEst As ClassItemMovEstoque) As Long
'exclui o movimento apontado por lComando2
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovimentoEstoque_Excluir
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVIMENTOESTOQUE_EXCLUIR", X)
    If lErro <> SUCESSO Then Error 20450
    
    
    MovimentoEstoque_Excluir = X.MovimentoEstoque_Excluir(lComando1, lComando2, objItemMovEst)

    Exit Function
    
Erro_MovimentoEstoque_Excluir:
    
    MovimentoEstoque_Excluir = Err
    
    Select Case Err

        Case 20450
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145949)
        
    End Select
    
    Exit Function

End Function

Function MovimentoEstoque_Grava_Estorno(ByVal lComando1 As Long, ByVal lComando2 As Long, objItemMovEst As ClassItemMovEstoque) As Long
'grava no movimento que será estornado o numero de documento de estorno
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovimentoEstoque_Grava_Estorno
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVIMENTOESTOQUE_GRAVA_ESTORNO", X)
    If lErro <> SUCESSO Then Error 20451
    
    
    MovimentoEstoque_Grava_Estorno = X.MovimentoEstoque_Grava_Estorno(lComando1, lComando2, objItemMovEst)

    Exit Function
    
Erro_MovimentoEstoque_Grava_Estorno:
    
    MovimentoEstoque_Grava_Estorno = Err
    
    Select Case Err

        Case 20451
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145950)
        
    End Select
    
    Exit Function

End Function

Function MovimentoEstoque_Inserir(ByVal lComando1 As Long, objItemMovEst As ClassItemMovEstoque) As Long
'insere um movimento de estoque na tabela MovimentoEstoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovimentoEstoque_Inserir
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVIMENTOESTOQUE_INSERIR", X)
    If lErro <> SUCESSO Then Error 20452
    
    
    MovimentoEstoque_Inserir = X.MovimentoEstoque_Inserir(lComando1, objItemMovEst)

    Exit Function
    
Erro_MovimentoEstoque_Inserir:
    
    MovimentoEstoque_Inserir = Err
    
    Select Case Err

        Case 20452
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145951)
        
    End Select
    
    Exit Function

End Function

Function EstoqueProduto_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, objEstoqueProduto As ClassEstoqueProduto, objTipoMovEstoque As ClassTipoMovEst, objItemMovEst As ClassItemMovEstoque, ByVal lComando As Long) As Long
'Atualiza a tabela de EstoqueProduto utilizando os valores passados atraves de objEstoqueProduto.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_Atualiza
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("ESTOQUEPRODUTO_ATUALIZA", X)
    If lErro <> SUCESSO Then Error 20453
    
    
    EstoqueProduto_Atualiza = X.EstoqueProduto_Atualiza(lComando1, lComando2, objEstoqueProduto, objTipoMovEstoque, objItemMovEst, lComando)

    Exit Function
    
Erro_EstoqueProduto_Atualiza:
    
    EstoqueProduto_Atualiza = Err
    
    Select Case Err

        Case 20453
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145952)
        
    End Select
    
    Exit Function

End Function

Function SldDiaEst_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, objSldDiaEst As ClassSldDiaEst) As Long
'Atualiza a tabela de SldDiaEst utilizando os valores passados atraves de objSldDiaEst.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldDiaEst_Atualiza
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDDIAEST_ATUALIZA", X)
    If lErro <> SUCESSO Then Error 20454
    
    
    SldDiaEst_Atualiza = X.SldDiaEst_Atualiza(lComando1, lComando2, objSldDiaEst)

    Exit Function
    
Erro_SldDiaEst_Atualiza:
    
    SldDiaEst_Atualiza = Err
    
    Select Case Err

        Case 20454
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145953)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, objSldDiaEst As ClassSldDiaEst, dCustoMedio As Double) As Long
'Atualiza a tabela de SldMesEst utilizando os valores passados atraves de objSldDiaEst.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst_Atualiza
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST_ATUALIZA", X)
    If lErro <> SUCESSO Then Error 20455
    
    
    SldMesEst_Atualiza = X.SldMesEst_Atualiza(lComando1, lComando2, objSldDiaEst, dCustoMedio)

    Exit Function
    
Erro_SldMesEst_Atualiza:
    
    SldMesEst_Atualiza = Err
    
    Select Case Err

        Case 20455
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145954)
        
    End Select
    
    Exit Function

End Function

Function MovimentoEstoque_Le_Codigo(iFilialEmpresa As Integer, lCodigo As Long) As Long
'Verifica se o filial/codigo já está cadastrado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovimentoEstoque_Le_Codigo
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVIMENTOESTOQUE_LE_CODIGO", X)
    If lErro <> SUCESSO Then Error 36016
    
    
    MovimentoEstoque_Le_Codigo = X.MovimentoEstoque_Le_Codigo(iFilialEmpresa, lCodigo)

    Exit Function
    
Erro_MovimentoEstoque_Le_Codigo:
    
    MovimentoEstoque_Le_Codigo = Err
    
    Select Case Err

        Case 36016
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145955)
        
    End Select
    
    Exit Function

End Function

Function MovimentoEstoque_Le_Ultima_Data(ByVal iFilialEmpresa As Integer, objItemMovEstoque As ClassItemMovEstoque) As Long
'le a ultima data de movimento cadastrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovimentoEstoque_Le_Ultima_Data
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVIMENTOESTOQUE_LE_ULTIMA_DATA", X)
    If lErro <> SUCESSO Then Error 36031
    
    
    MovimentoEstoque_Le_Ultima_Data = X.MovimentoEstoque_Le_Ultima_Data(iFilialEmpresa, objItemMovEstoque)

    Exit Function
    
Erro_MovimentoEstoque_Le_Ultima_Data:
    
    MovimentoEstoque_Le_Ultima_Data = Err
    
    Select Case Err

        Case 36031
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145956)
        
    End Select
    
    Exit Function

End Function

Function Produto_Le_ControleEstoque(ByVal lComando As Long, objProduto As ClassProduto) As Long
'Lê o atributo controle de estoque do produto, cujo codigo foi passado como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_ControleEstoque
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LE_CONTROLEESTOQUE", X)
    If lErro <> SUCESSO Then Error 36035
    
    
    Produto_Le_ControleEstoque = X.Produto_Le_ControleEstoque(lComando, objProduto)

    Exit Function
    
Erro_Produto_Le_ControleEstoque:
    
    Produto_Le_ControleEstoque = Err
    
    Select Case Err

        Case 36035
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145957)
        
    End Select
    
    Exit Function

End Function

Function UM_Le1(ByVal lComando As Long, objUnidadeMedida As ClassUnidadeDeMedida) As Long
'Lê a unidade de medida a partir da Classe e da Sigla -->
'e devolve dados no Objeto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UM_Le1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_UM_LE1", X)
    If lErro <> SUCESSO Then Error 36111
    
    
    UM_Le1 = X.UM_Le1(lComando, objUnidadeMedida)

    Exit Function
    
Erro_UM_Le1:
    
    UM_Le1 = Err
    
    Select Case Err

        Case 36111
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145958)
        
    End Select
    
    Exit Function

End Function

Function ProdutoFilial_Le1(lComando As Long, objProdutoFilial As ClassProdutoFilial) As Long
'Lê os dados do Produto da Filial à partir da chave fornecida em objProdutoFilial. Retorna os dados em objProdutoFilial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutoFilial_Le1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTOFILIAL_LE1", X)
    If lErro <> SUCESSO Then Error 36224
    
    
    ProdutoFilial_Le1 = X.ProdutoFilial_Le1(lComando, objProdutoFilial)

    Exit Function
    
Erro_ProdutoFilial_Le1:
    
    ProdutoFilial_Le1 = Err
    
    Select Case Err

        Case 36224
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145959)
        
    End Select
    
    Exit Function

End Function

Function TipoMovEstoque_Le(objTipoMovEstoque As ClassTipoMovEst) As Long
'Lê o Tipo de MovEstoque a partir do Código

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoMovEstoque_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TIPOMOVESTOQUE_LE", X)
    If lErro <> SUCESSO Then Error 16854
    
    
    TipoMovEstoque_Le = X.TipoMovEstoque_Le(objTipoMovEstoque)

    Exit Function
    
Erro_TipoMovEstoque_Le:
    
    TipoMovEstoque_Le = Err
    
    Select Case Err

        Case 16854
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145960)
        
    End Select
    
    Exit Function

End Function

Function MovEstoqueItens_Le(objMovEstoque As ClassMovEstoque) As Long
'Lê todos os Itens correspondentes a partir da FilialEmpresa, Código e do NumIntDocEst=0 e retorna no colItens do objeto
'Ordena pelo Item.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoqueItens_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUEITENS_LE", X)
    If lErro <> SUCESSO Then Error 16855
    
    
    MovEstoqueItens_Le = X.MovEstoqueItens_Le(objMovEstoque)

    Exit Function
    
Erro_MovEstoqueItens_Le:
    
    MovEstoqueItens_Le = Err
    
    Select Case Err

        Case 16855
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145961)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_Le(objMovEstoque As ClassMovEstoque) As Long
'Lê o MovEstoque a partir do Código.
'A função só pesquisa os MovEstoque com FilialEmpresa=giFilialEmpresa e NumIntDocEst=0

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_LE", X)
    If lErro <> SUCESSO Then Error 16856
    
    
    MovEstoque_Le = X.MovEstoque_Le(objMovEstoque)

    Exit Function
    
Erro_MovEstoque_Le:
    
    MovEstoque_Le = Err
    
    Select Case Err

        Case 16856
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145962)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_Automatico(iFilialEmpresa As Integer, lCodigo As Long) As Long
'retorna o número do próximo MovEstoque disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 16857
    
    
    MovEstoque_Automatico = X.MovEstoque_Automatico(iFilialEmpresa, lCodigo)

    Exit Function
    
Erro_MovEstoque_Automatico:
    
    MovEstoque_Automatico = Err
    
    Select Case Err

        Case 16857
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145963)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_Interno_Grava(objMovEstoque As ClassMovEstoque, objContabil As ClassContabil) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Interno_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_INTERNO_GRAVA", X)
    If lErro <> SUCESSO Then Error 16864
    
    
    MovEstoque_Interno_Grava = X.MovEstoque_Interno_Grava(objMovEstoque, objContabil)

    Exit Function
    
Erro_MovEstoque_Interno_Grava:
    
    MovEstoque_Interno_Grava = Err
    
    Select Case Err

        Case 16864
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145964)
        
    End Select
    
    Exit Function

End Function

Function TiposMovEstoque_Le(colTipoMovEst As colTipoMovEst) As Long
'Lê os Tipos de MovEstoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposMovEstoque_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TIPOSMOVESTOQUE_LE", X)
    If lErro <> SUCESSO Then Error 16865
    
    
    TiposMovEstoque_Le = X.TiposMovEstoque_Le(colTipoMovEst)

    Exit Function
    
Erro_TiposMovEstoque_Le:
    
    TiposMovEstoque_Le = Err
    
    Select Case Err

        Case 16865
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145965)
        
    End Select
    
    Exit Function

End Function

Function MovEstoqueInternoItens_Le(objMovEstoque As ClassMovEstoque) As Long
'Lê todos os Itens correspondentes a partir da FilialEmpresa, Código e do NumIntDocEst=0 e retorna no colItens do objeto
'Ordena pelo Item.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoqueInternoItens_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUEINTERNOITENS_LE", X)
    If lErro <> SUCESSO Then Error 16866
    
    
    MovEstoqueInternoItens_Le = X.MovEstoqueInternoItens_Le(objMovEstoque)

    Exit Function
    
Erro_MovEstoqueInternoItens_Le:
    
    MovEstoqueInternoItens_Le = Err
    
    Select Case Err

        Case 16866
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145966)
        
    End Select
    
    Exit Function

End Function

Function OrdemProducao_Le(objOrdemProducao As ClassOrdemDeProducao) As Long
'Lê a Ordem de Produção pelo Código e pela FilialEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OrdemProducao_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ORDEMPRODUCAO_LE", X)
    If lErro <> SUCESSO Then Error 16867
    
    
    OrdemProducao_Le = X.OrdemProducao_Le(objOrdemProducao)

    Exit Function
    
Erro_OrdemProducao_Le:
    
    OrdemProducao_Le = Err
    
    Select Case Err

        Case 16867
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145967)
        
    End Select
    
    Exit Function

End Function

Function ItensOrdemProducao_Le(objOrdemProducao As ClassOrdemDeProducao) As Long
'Lê os Itens Ordem de Produção a partir do OPCódigo e da FilialEmpresa
'Retorna os Itens na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensOrdemProducao_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ITENSORDEMPRODUCAO_LE", X)
    If lErro <> SUCESSO Then Error 16868
    
    
    ItensOrdemProducao_Le = X.ItensOrdemProducao_Le(objOrdemProducao)

    Exit Function
    
Erro_ItensOrdemProducao_Le:
    
    ItensOrdemProducao_Le = Err
    
    Select Case Err

        Case 16868
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145968)
        
    End Select
    
    Exit Function

End Function

Public Function ItemOP_Le(objItemOP As ClassItemOP) As Long
'Le de ItemOP com filtros objItemOP.iFilialEmpresa , objItemOP.sCodigo , objItemOP.sProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemOP_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ITEMOP_LE", X)
    If lErro <> SUCESSO Then Error 16870
    
    
    ItemOP_Le = X.ItemOP_Le(objItemOP)

    Exit Function
    
Erro_ItemOP_Le:
    
    ItemOP_Le = Err
    
    Select Case Err

        Case 16870
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145969)
        
    End Select
    
    Exit Function

End Function

Public Function MovEstoque_Producao_Grava(objMovEstoque As ClassMovEstoque, objContabil As ClassContabil) As Long
'Grava a producao do movimento de estoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Producao_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_PRODUCAO_GRAVA", X)
    If lErro <> SUCESSO Then Error 16871
    
    
    MovEstoque_Producao_Grava = X.MovEstoque_Producao_Grava(objMovEstoque, objContabil)

    Exit Function
    
Erro_MovEstoque_Producao_Grava:
    
    MovEstoque_Producao_Grava = Err
    
    Select Case Err

        Case 16871
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145970)
        
    End Select
    
    Exit Function

End Function

Function OrdemDeProducao_TrataBaixaItem(sOPCodigo As String, alComando() As Long) As Long
'Deve ser chamada qdo um item da OP acaba de ser baixado, se for o ultimo item nao baixado vai baixar a OP
'De alComando(2) a alComando(8) devem estar "livres"

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OrdemDeProducao_TrataBaixaItem
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ORDEMDEPRODUCAO_TRATABAIXAITEM", X)
    If lErro <> SUCESSO Then Error 16872
    
    
    OrdemDeProducao_TrataBaixaItem = X.OrdemDeProducao_TrataBaixaItem(sOPCodigo, alComando())

    Exit Function
    
Erro_OrdemDeProducao_TrataBaixaItem:
    
    OrdemDeProducao_TrataBaixaItem = Err
    
    Select Case Err

        Case 16872
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145971)
        
    End Select
    
    Exit Function

End Function

Function ItemOP_Le_NumIntDoc(objItemOP As ClassItemOP) As Long
'Lê os dados do item do empenho à partir de NumIntDoc

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemOP_Le_NumIntDoc
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ITEMOP_LE_NUMINTDOC", X)
    If lErro <> SUCESSO Then Error 16895
    
    
    ItemOP_Le_NumIntDoc = X.ItemOP_Le_NumIntDoc(objItemOP)

    Exit Function
    
Erro_ItemOP_Le_NumIntDoc:
    
    ItemOP_Le_NumIntDoc = Err
    
    Select Case Err

        Case 16895
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145972)
        
    End Select
    
    Exit Function

End Function

Function ItemOP_Le_Item(objItemOP As ClassItemOP) As Long
'Lê os dados do Item da Ordem de Produção à partir do Código da OP e Código do Item de uma Filial da Empresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemOP_Le_Item
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ITEMOP_LE_ITEM", X)
    If lErro <> SUCESSO Then Error 16896
    
    
    ItemOP_Le_Item = X.ItemOP_Le_Item(objItemOP)

    Exit Function
    
Erro_ItemOP_Le_Item:
    
    ItemOP_Le_Item = Err
    
    Select Case Err

        Case 16896
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145973)
        
    End Select
    
    Exit Function

End Function

Function Empenho_Le(objEmpenho As ClassEmpenho) As Long
'Lê os dados de Empenho à partir da chave fornecida em objEmpenho
'objEmpenho.iFilialEmpresa e objEmpenho.lCodigo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Empenho_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_EMPENHO_LE", X)
    If lErro <> SUCESSO Then Error 16897
    
    
    Empenho_Le = X.Empenho_Le(objEmpenho)

    Exit Function
    
Erro_Empenho_Le:
    
    Empenho_Le = Err
    
    Select Case Err

        Case 16897
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145974)
        
    End Select
    
    Exit Function

End Function

Function Empenho_Automatico(lCodigo As Long, lComando1 As Long, lComando2 As Long) As Long
'gera codigo para insercao de Empenho. O codigo é sequencial por filialempresa
'TEM QUE SER CHAMADA EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Empenho_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_EMPENHO_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 16898
    
    
    Empenho_Automatico = X.Empenho_Automatico(lCodigo, lComando1, lComando2)

    Exit Function
    
Erro_Empenho_Automatico:
    
    Empenho_Automatico = Err
    
    Select Case Err

        Case 16898
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145975)
        
    End Select
    
    Exit Function

End Function

Function MATConfig_ObterNumInt(sCodigo As String, lNumInt As Long, Optional iQtde As Integer = 1) As Long
'Obter um número automático a partir de um contador identificado por uma chave
'iQtde identifica quantos numeros devem ser "liberados" a partir do lNumInt retornado
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MATConfig_ObterNumInt
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MATCONFIG_OBTERNUMINT", X)
    If lErro <> SUCESSO Then Error 16899
    
    
    MATConfig_ObterNumInt = X.MATConfig_ObterNumInt(sCodigo, lNumInt, iQtde)

    Exit Function
    
Erro_MATConfig_ObterNumInt:
    
    MATConfig_ObterNumInt = Err
    
    Select Case Err

        Case 16899
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145976)
        
    End Select
    
    Exit Function

End Function

Function Empenho_Grava(objEmpenho As ClassEmpenho) As Long
'Grava o Empenho

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Empenho_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_EMPENHO_GRAVA", X)
    If lErro <> SUCESSO Then Error 16900
    
    
    Empenho_Grava = X.Empenho_Grava(objEmpenho)

    Exit Function
    
Erro_Empenho_Grava:
    
    Empenho_Grava = Err
    
    Select Case Err

        Case 16900
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145977)
        
    End Select
    
    Exit Function

End Function

Function Empenho_Inserir_Batch(objEmpenho As ClassEmpenho, lComando As Long) As Long
'Insere um empenho da tabela de empenhos
'DEVE SER CHAMADO EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Empenho_Inserir_Batch
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_EMPENHO_INSERIR_BATCH", X)
    If lErro <> SUCESSO Then Error 16901
    
    
    Empenho_Inserir_Batch = X.Empenho_Inserir_Batch(objEmpenho, lComando)

    Exit Function
    
Erro_Empenho_Inserir_Batch:
    
    Empenho_Inserir_Batch = Err
    
    Select Case Err

        Case 16901
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145978)
        
    End Select
    
    Exit Function

End Function

Function Empenho_Exclui(objEmpenho As ClassEmpenho) As Long
'Exclui o Empenho
'objEmpenho.iFilialEmpresa e objEmpenho.lCodigo tem que estar preenchidos
    'nao deixar excluir se houver qtde requisitada
    'diminuir a qtde empenhada na tabela EstoqueProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Empenho_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_EMPENHO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 16902
    
    
    Empenho_Exclui = X.Empenho_Exclui(objEmpenho)

    Exit Function
    
Erro_Empenho_Exclui:
    
    Empenho_Exclui = Err
    
    Select Case Err

        Case 16902
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145979)
        
    End Select
    
    Exit Function

End Function

Function EstoqueProduto_AlterarEmpenho(objEstoqueProduto As ClassEstoqueProduto, dQtde As Double) As Long
'Aumenta (ou diminui se dQtde for negativo) a qtde empenhada em um produto-almoxarifado
'objEstoqueProduto.sProduto e objEstoqueProduto.iAlmoxarifado tem que estar preenchidos
'DEVE SER CHAMADO EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_AlterarEmpenho
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_ALTERAREMPENHO", X)
    If lErro <> SUCESSO Then Error 16903
    
    
    EstoqueProduto_AlterarEmpenho = X.EstoqueProduto_AlterarEmpenho(objEstoqueProduto, dQtde)

    Exit Function
    
Erro_EstoqueProduto_AlterarEmpenho:
    
    EstoqueProduto_AlterarEmpenho = Err
    
    Select Case Err

        Case 16903
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145980)
        
    End Select
    
    Exit Function

End Function

Function Produto_Exclui(objProduto As ClassProduto) As Long
'Exclui no BD o Produto passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19669
    
    
    Produto_Exclui = X.Produto_Exclui(objProduto)

    Exit Function
    
Erro_Produto_Exclui:
    
    Produto_Exclui = Err
    
    Select Case Err

        Case 19669
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145981)
        
    End Select
    
    Exit Function

End Function

Function Produto_Grava(objProduto As ClassProduto, colTabelaPrecoItem As Collection) As Long
'Grava ou Alterar o Produto passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_GRAVA", X)
    If lErro <> SUCESSO Then Error 19670
    
    
    Produto_Grava = X.Produto_Grava(objProduto, colTabelaPrecoItem)

    Exit Function
    
Erro_Produto_Grava:
    
    Produto_Grava = Err
    
    Select Case Err

        Case 19670
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145982)
        
    End Select
    
    Exit Function

End Function

Function Produto_Le_Categorias(objProduto As ClassProduto, colCategoriaItem As Collection) As Long
'Lê na tabela de ProdutoCategoria, todas as categorias de um  Produto, guardando na coleção colCategoriaItem

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Categorias
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LE_CATEGORIAS", X)
    If lErro <> SUCESSO Then Error 19671
    
    
    Produto_Le_Categorias = X.Produto_Le_Categorias(objProduto, colCategoriaItem)

    Exit Function
    
Erro_Produto_Le_Categorias:
    
    Produto_Le_Categorias = Err
    
    Select Case Err

        Case 19671
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145983)
        
    End Select
    
    Exit Function

End Function

Function Produto_Lock(objProduto As ClassProduto) As Long
'Lê e faz o "lock" no Produto passado por parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LOCK", X)
    If lErro <> SUCESSO Then Error 19672
    
    
    Produto_Lock = X.Produto_Lock(objProduto)

    Exit Function
    
Erro_Produto_Lock:
    
    Produto_Lock = Err
    
    Select Case Err

        Case 19672
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145984)
        
    End Select
    
    Exit Function

End Function

Function Produto_LerBloquear(sProduto As String, lComando As Long) As Long
'Lê e faz o "lock" no Produto passado por parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_LerBloquear
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LERBLOQUEAR", X)
    If lErro <> SUCESSO Then Error 19764
    
    
    Produto_LerBloquear = X.Produto_LerBloquear(sProduto, lComando)

    Exit Function
    
Erro_Produto_LerBloquear:
    
    Produto_LerBloquear = Err
    
    Select Case Err

        Case 19764
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145985)
        
    End Select
    
    Exit Function

End Function

Function UM_Le_Lock(objUnidadeMedida As ClassUnidadeDeMedida) As Long
'Lê e "locka" a unidade de medida a partir da Classe e da SiglaUMBase -->
'e devolve dados no Objeto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UM_Le_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_UM_LE_LOCK", X)
    If lErro <> SUCESSO Then Error 19674
    
    
    UM_Le_Lock = X.UM_Le_Lock(objUnidadeMedida)

    Exit Function
    
Erro_UM_Le_Lock:
    
    UM_Le_Lock = Err
    
    Select Case Err

        Case 19674
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145986)
        
    End Select
    
    Exit Function

End Function

Function Produto_Le_NomeReduzido(objProduto As ClassProduto) As Long
'Lê o Produto com o Nome Reduzido passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_NomeReduzido
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LE_NOMEREDUZIDO", X)
    If lErro <> SUCESSO Then Error 19675
    
    
    Produto_Le_NomeReduzido = X.Produto_Le_NomeReduzido(objProduto)

    Exit Function
    
Erro_Produto_Le_NomeReduzido:
    
    Produto_Le_NomeReduzido = Err
    
    Select Case Err

        Case 19675
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145987)
        
    End Select
    
    Exit Function

End Function

Function Produto_Le_ProdutoMaior(sProduto As String, objProduto As ClassProduto) As Long
'Lê um Produto imediatamente maior que o produto passada como parametro. Devolve os dados em objProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_ProdutoMaior
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LE_PRODUTOMAIOR", X)
    If lErro <> SUCESSO Then Error 19676
    
    
    Produto_Le_ProdutoMaior = X.Produto_Le_ProdutoMaior(sProduto, objProduto)

    Exit Function
    
Erro_Produto_Le_ProdutoMaior:
    
    Produto_Le_ProdutoMaior = Err
    
    Select Case Err

        Case 19676
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145988)
        
    End Select
    
    Exit Function

End Function

Function Produto_Le_TabelaPrecoItem(objProduto As ClassProduto, colTabelaPrecoItem As Collection, iFilialEmpresa As Integer) As Long
'Lê os preços de um produto para a giFilialEmpresa (filial corrente ou EMPRESA_TODA). Retorna os dados em colTabelaPrecoItem

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_TabelaPrecoItem
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LE_TABELAPRECOITEM", X)
    If lErro <> SUCESSO Then Error 19677
    
    
    Produto_Le_TabelaPrecoItem = X.Produto_Le_TabelaPrecoItem(objProduto, colTabelaPrecoItem, iFilialEmpresa)

    Exit Function
    
Erro_Produto_Le_TabelaPrecoItem:
    
    Produto_Le_TabelaPrecoItem = Err
    
    Select Case Err

        Case 19677
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145989)
        
    End Select
    
    Exit Function

End Function

Function Produto_Tem_Filho(sProduto As String, iTemFilho As Integer) As Long
'testa se o Produto em questão tem Produtos abaixo dele, ou seja tem Produtos filho.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Tem_Filho
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_TEM_FILHO", X)
    If lErro <> SUCESSO Then Error 19678
    
    
    Produto_Tem_Filho = X.Produto_Tem_Filho(sProduto, iTemFilho)

    Exit Function
    
Erro_Produto_Tem_Filho:
    
    Produto_Tem_Filho = Err
    
    Select Case Err

        Case 19678
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145990)
        
    End Select
    
    Exit Function

End Function

Function Produto_Critica_ProdutoPai(sProduto As String, sSiglaModulo As String) As Long
'verifica se o Produto tem um pai sintetico

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Critica_ProdutoPai
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_CRITICA_PRODUTOPAI", X)
    If lErro <> SUCESSO Then Error 19679
    
    
    Produto_Critica_ProdutoPai = X.Produto_Critica_ProdutoPai(sProduto, sSiglaModulo)

    Exit Function
    
Erro_Produto_Critica_ProdutoPai:
    
    Produto_Critica_ProdutoPai = Err
    
    Select Case Err

        Case 19679
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145991)
        
    End Select
    
    Exit Function

End Function

Function AlmoxarifadoPadrao_Le(iFilialEmpresa As Integer, sProduto As String, iAlmoxarifadoPadrao As Integer) As Long
'Lê Almoxarifado Padrão do Produto da iFilialEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_AlmoxarifadoPadrao_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ALMOXARIFADOPADRAO_LE", X)
    If lErro <> SUCESSO Then Error 19680
    
    
    AlmoxarifadoPadrao_Le = X.AlmoxarifadoPadrao_Le(iFilialEmpresa, sProduto, iAlmoxarifadoPadrao)

    Exit Function
    
Erro_AlmoxarifadoPadrao_Le:
    
    AlmoxarifadoPadrao_Le = Err
    
    Select Case Err

        Case 19680
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145992)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_Automatico_EmTransacao(iFilialEmpresa As Integer, lCodigo As Long) As Long
'retorna o número do próximo MovEstoque disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Automatico_EmTransacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_AUTOMATICO_EMTRANSACAO", X)
    If lErro <> SUCESSO Then Error 19681
    
    
    MovEstoque_Automatico_EmTransacao = X.MovEstoque_Automatico_EmTransacao(iFilialEmpresa, lCodigo)

    Exit Function
    
Erro_MovEstoque_Automatico_EmTransacao:
    
    MovEstoque_Automatico_EmTransacao = Err
    
    Select Case Err

        Case 19681
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145993)
        
    End Select
    
    Exit Function

End Function

Function CategoriaProdutoItem_Lock(sCategoriaProduto As String, sItemCategoria As String, lComando As Long) As Long
'faz lock shared de uma CategoriaProdutoItem para garantir a sua existencia

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CategoriaProdutoItem_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CATEGORIAPRODUTOITEM_LOCK", X)
    If lErro <> SUCESSO Then Error 19752
    
    
    CategoriaProdutoItem_Lock = X.CategoriaProdutoItem_Lock(sCategoriaProduto, sItemCategoria, lComando)

    Exit Function
    
Erro_CategoriaProdutoItem_Lock:
    
    CategoriaProdutoItem_Lock = Err
    
    Select Case Err

        Case 19752
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145994)
        
    End Select
    
    Exit Function

End Function

Function Almoxarifado_Lock(iAlmoxarifado As Integer, lComando As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Almoxarifado_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ALMOXARIFADO_LOCK", X)
    If lErro <> SUCESSO Then Error 19763
    
    
    Almoxarifado_Lock = X.Almoxarifado_Lock(iAlmoxarifado, lComando)

    Exit Function
    
Erro_Almoxarifado_Lock:
    
    Almoxarifado_Lock = Err
    
    Select Case Err

        Case 19763
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145995)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst_Le_QuantidadeCusto(ByVal lComando As Long, objSldMesEst As ClassSldMesEst) As Long
'le os dados de SldMesEst relativos ao Ano, FilialEmpresa e Produto passados como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst_Le_QuantidadeCusto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST_LE_QUANTIDADECUSTO", X)
    If lErro <> SUCESSO Then Error 36486
    
    
    SldMesEst_Le_QuantidadeCusto = X.SldMesEst_Le_QuantidadeCusto(lComando, objSldMesEst)

    Exit Function
    
Erro_SldMesEst_Le_QuantidadeCusto:
    
    SldMesEst_Le_QuantidadeCusto = Err
    
    Select Case Err

        Case 36486
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145996)
        
    End Select
    
    Exit Function

End Function

Public Function MovEstoque_Grava_Generico(objMovEstoque As ClassMovEstoque, objContabil As ClassContabil) As Long
'Grava o movimento de estoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Grava_Generico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_GRAVA_GENERICO", X)
    If lErro <> SUCESSO Then Error 36496
    
    
    MovEstoque_Grava_Generico = X.MovEstoque_Grava_Generico(objMovEstoque, objContabil)

    Exit Function
    
Erro_MovEstoque_Grava_Generico:
    
    MovEstoque_Grava_Generico = Err
    
    Select Case Err

        Case 36496
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145997)
        
    End Select
    
    Exit Function

End Function

Function EstoqueMes_Le(objEstoqueMes As ClassEstoqueMes) As Long
'Le as informacoes sobre o EstoqueMes através do parametro Filial/Ano/Mes passado em objEstoqueMes

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueMes_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMES_LE", X)
    If lErro <> SUCESSO Then Error 36514
    
    
    EstoqueMes_Le = X.EstoqueMes_Le(objEstoqueMes)

    Exit Function
    
Erro_EstoqueMes_Le:
    
    EstoqueMes_Le = Err
    
    Select Case Err

        Case 36514
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145998)
        
    End Select
    
    Exit Function

End Function

Function EstoqueProduto_Le1(lComando As Long, objEstoqueProduto As ClassEstoqueProduto) As Long
'Le as informacoes sobre o produto passado em objEstoqueProduto na tabela EstoqueProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_Le1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_LE1", X)
    If lErro <> SUCESSO Then Error 36665
    
    
    EstoqueProduto_Le1 = X.EstoqueProduto_Le1(lComando, objEstoqueProduto)

    Exit Function
    
Erro_EstoqueProduto_Le1:
    
    EstoqueProduto_Le1 = Err
    
    Select Case Err

        Case 36665
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145999)
        
    End Select
    
    Exit Function

End Function

Function EstoqueProduto_Atualiza2(ByVal lComando1 As Long, ByVal lComando2 As Long, objEstoqueProduto As ClassEstoqueProduto) As Long
'Atualiza a tabela de EstoqueProduto utilizando os valores passados atraves de objEstoqueProduto.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_Atualiza2
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_ATUALIZA2", X)
    If lErro <> SUCESSO Then Error 36670
    
    
    EstoqueProduto_Atualiza2 = X.EstoqueProduto_Atualiza2(lComando1, lComando2, objEstoqueProduto)

    Exit Function
    
Erro_EstoqueProduto_Atualiza2:
    
    EstoqueProduto_Atualiza2 = Err
    
    Select Case Err

        Case 36670
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146000)
        
    End Select
    
    Exit Function

End Function

Function SldDiaEstAlm_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, objSldDiaEst As ClassSldDiaEst, ByVal iAlmoxarifado As Integer) As Long
'Atualiza a tabela de SldDiaEstAlm utilizando os valores passados atraves de objSldDiaEst e iAlmoxarifado.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldDiaEstAlm_Atualiza
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDDIAESTALM_ATUALIZA", X)
    If lErro <> SUCESSO Then Error 44119
    
    
    SldDiaEstAlm_Atualiza = X.SldDiaEstAlm_Atualiza(lComando1, lComando2, objSldDiaEst, iAlmoxarifado)

    Exit Function
    
Erro_SldDiaEstAlm_Atualiza:
    
    SldDiaEstAlm_Atualiza = Err
    
    Select Case Err

        Case 44119
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146001)
        
    End Select
    
    Exit Function

End Function

Function SldMesEstAlm_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, objSldDiaEst As ClassSldDiaEst, ByVal iAlmoxarifado As Integer) As Long
'Atualiza a tabela de SldMesEstAlm utilizando os valores passados atraves de objSldDiaEst e iAlmoxarifado.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEstAlm_Atualiza
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESESTALM_ATUALIZA", X)
    If lErro <> SUCESSO Then Error 44120
    
    
    SldMesEstAlm_Atualiza = X.SldMesEstAlm_Atualiza(lComando1, lComando2, objSldDiaEst, iAlmoxarifado)

    Exit Function
    
Erro_SldMesEstAlm_Atualiza:
    
    SldMesEstAlm_Atualiza = Err
    
    Select Case Err

        Case 44120
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146002)
        
    End Select
    
    Exit Function

End Function

Function EstoqueMes_Le_Mes(objEstoqueMes As ClassEstoqueMes) As Long
'Le o Ano/Mes que está aberto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueMes_Le_Mes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMES_LE_MES", X)
    If lErro <> SUCESSO Then Error 45116
    
    
    EstoqueMes_Le_Mes = X.EstoqueMes_Le_Mes(objEstoqueMes)

    Exit Function
    
Erro_EstoqueMes_Le_Mes:
    
    EstoqueMes_Le_Mes = Err
    
    Select Case Err

        Case 45116
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146003)
        
    End Select
    
    Exit Function

End Function

Function Almoxarifados_Le_FilialEmpresa(ByVal iFilialEmpresa As Integer, colAlmoxFilial As Collection) As Long
'Le todos os almoxaridados da Filial iFilialEmpresa e coloca-os em colAlmoxFilial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Almoxarifados_Le_FilialEmpresa
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ALMOXARIFADOS_LE_FILIALEMPRESA", X)
    If lErro <> SUCESSO Then Error 44171
    
    
    Almoxarifados_Le_FilialEmpresa = X.Almoxarifados_Le_FilialEmpresa(iFilialEmpresa, colAlmoxFilial)

    Exit Function
    
Erro_Almoxarifados_Le_FilialEmpresa:
    
    Almoxarifados_Le_FilialEmpresa = Err
    
    Select Case Err

        Case 44171
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146004)
        
    End Select
    
    Exit Function

End Function

Function ProdutoFilial_Le2(lComando As Long, objProdutoFilial As ClassProdutoFilial) As Long
'Lê os dados do Produto da Filial à partir da chave fornecida em objProdutoFilial. Retorna os dados em objProdutoFilial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutoFilial_Le2
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTOFILIAL_LE2", X)
    If lErro <> SUCESSO Then Error 44207
    
    
    ProdutoFilial_Le2 = X.ProdutoFilial_Le2(lComando, objProdutoFilial)

    Exit Function
    
Erro_ProdutoFilial_Le2:
    
    ProdutoFilial_Le2 = Err
    
    Select Case Err

        Case 44207
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146005)
        
    End Select
    
    Exit Function

End Function

Function EstoqueProduto_Le_Exclusivo(lComando As Long, objEstoqueProduto As ClassEstoqueProduto) As Long
'le o registro em EstoqueProduto cuja chave (Produto, Almoxarifado) foi passada como parametro em objEstoqueProduto.
'Devolve os dados lidos em objEstoqueProduto.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_Le_Exclusivo
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_LE_EXCLUSIVO", X)
    If lErro <> SUCESSO Then Error 44341
    
    
    EstoqueProduto_Le_Exclusivo = X.EstoqueProduto_Le_Exclusivo(lComando, objEstoqueProduto)

    Exit Function
    
Erro_EstoqueProduto_Le_Exclusivo:
    
    EstoqueProduto_Le_Exclusivo = Err
    
    Select Case Err

        Case 44341
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146006)
        
    End Select
    
    Exit Function

End Function


'mario
Function MovEstoque_Automatico_EmTransacao1(lComando As Long, lComando1 As Long, lComando2 As Long, iFilialEmpresa As Integer, lCodigo As Long) As Long
'retorna o número do próximo MovEstoque disponivel
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Automatico_EmTransacao1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_AUTOMATICO_EMTRANSACAO1", X)
    If lErro <> SUCESSO Then Error 44357
    
    
    MovEstoque_Automatico_EmTransacao1 = X.MovEstoque_Automatico_EmTransacao1(lComando, lComando1, lComando2, iFilialEmpresa, lCodigo)

    Exit Function
    
Erro_MovEstoque_Automatico_EmTransacao1:
    
    MovEstoque_Automatico_EmTransacao1 = Err
    
    Select Case Err

        Case 44357
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146007)
        
    End Select
    
    Exit Function

End Function

'mario
Function Produto_Lock1(lComando As Long, objProduto As ClassProduto) As Long
'Lê o faz o "lock" no Produto passado por parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Lock1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LOCK1", X)
    If lErro <> SUCESSO Then Error 44358
    
    
    Produto_Lock1 = X.Produto_Lock1(lComando, objProduto)

    Exit Function
    
Erro_Produto_Lock1:
    
    Produto_Lock1 = Err
    
    Select Case Err

        Case 44358
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146008)
        
    End Select
    
    Exit Function

End Function

Function CustoMedioProducaoApurado_Le(sProduto As String, dCMPAtual As Double) As Long
'Calcula último Custo Médio de Produção Apurado do Produto passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CustoMedioProducaoApurado_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CUSTOMEDIOPRODUCAOAPURADO_LE", X)
    If lErro <> SUCESSO Then Error 45402
    
    
    CustoMedioProducaoApurado_Le = X.CustoMedioProducaoApurado_Le(sProduto, dCMPAtual)
    
    Exit Function
    
Erro_CustoMedioProducaoApurado_Le:

    CustoMedioProducaoApurado_Le = Err
    
    Select Case Err

        Case 45402
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146009)
        
    End Select
    
    Exit Function
    
End Function

Function EstoquesProduto_Le_Filial(sProduto As String, colEstoqueProduto As colEstoqueProduto) As Long
'??? documentar

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoquesProduto_Le_Filial
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUESPRODUTO_LE_FILIAL", X)
    If lErro <> SUCESSO Then Error 45578
    
    
    EstoquesProduto_Le_Filial = X.EstoquesProduto_Le_Filial(sProduto, colEstoqueProduto)
    
    Exit Function
    
Erro_EstoquesProduto_Le_Filial:

    EstoquesProduto_Le_Filial = Err
    
    Select Case Err

        Case 45578
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146010)
        
    End Select
    
    Exit Function
    
End Function

Function Almoxarifado_Grava(objAlmoxarifado As ClassAlmoxarifado, objEndereco As ClassEndereco) As Long
'Grava o Almoxarifado no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Almoxarifado_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ALMOXARIFADO_GRAVA", X)
    If lErro <> SUCESSO Then Error 45652
    
    
    Almoxarifado_Grava = X.Almoxarifado_Grava(objAlmoxarifado, objEndereco)
    
    Exit Function
    
Erro_Almoxarifado_Grava:

    Almoxarifado_Grava = Err
    
    Select Case Err

        Case 45652
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146011)
        
    End Select
    
    Exit Function
    
End Function

Function Almoxarifado_Exclui(objAlmoxarifado As ClassAlmoxarifado) As Long
'Exclui o Almoxarifado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Almoxarifado_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ALMOXARIFADO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45653
    
    
    Almoxarifado_Exclui = X.Almoxarifado_Exclui(objAlmoxarifado)
    
    Exit Function
    
Erro_Almoxarifado_Exclui:

    Almoxarifado_Exclui = Err
    
    Select Case Err

        Case 45653
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146012)
        
    End Select
    
    Exit Function
    
End Function

Function ClassificacaoABC_Exclui(objClassABC As ClassClassificacaoABC) As Long
'Exclui Classificação ABC e seus ítens do Banco de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClassificacaoABC_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSIFICACAOABC_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45655
    
    
    ClassificacaoABC_Exclui = X.ClassificacaoABC_Exclui(objClassABC)
    
    Exit Function
    
Erro_ClassificacaoABC_Exclui:

    ClassificacaoABC_Exclui = Err
    
    Select Case Err

        Case 45655
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146013)
        
    End Select
    
    Exit Function
    
End Function

Function ItensABC_Exclui(objClassABC As ClassClassificacaoABC) As Long
'Exclui Itens da Classificação ABC do Banco de Dados
'CHAMADA EM TRANSAÇÃO por ClassificacaoABC_Exclui

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensABC_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ITENSABC_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45656
    
    
    ItensABC_Exclui = X.ItensABC_Exclui(objClassABC)
    
    Exit Function
    
Erro_ItensABC_Exclui:

    ItensABC_Exclui = Err
    
    Select Case Err

        Case 45656
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146014)
        
    End Select
    
    Exit Function
    
End Function

Function ClassificacaoABC_Eh_MaisRecente(objClassABC As ClassClassificacaoABC) As Long
'Verifica se a ClassificacaoABC no BD eh a mais recente que atualizou produtos para o TipoProduto dela.
'Passa NumInt, FilialEmpresa e TipoProduto da ClassABC
'Retorna SUCESSO se for + recente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClassificacaoABC_Eh_MaisRecente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSIFICACAOABC_EH_MAISRECENTE", X)
    If lErro <> SUCESSO Then Error 45657
    
    
    ClassificacaoABC_Eh_MaisRecente = X.ClassificacaoABC_Eh_MaisRecente(objClassABC)
    
    Exit Function
    
Erro_ClassificacaoABC_Eh_MaisRecente:

    ClassificacaoABC_Eh_MaisRecente = Err
    
    Select Case Err

        Case 45657
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146015)
        
    End Select
    
    Exit Function
    
End Function

Function ClassificacaoABC_Grava(objClassABC As ClassClassificacaoABC) As Long
'Grava Classificação ABC e seus ítens no Banco de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClassificacaoABC_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSIFICACAOABC_GRAVA", X)
    If lErro <> SUCESSO Then Error 45658
    
    
    ClassificacaoABC_Grava = X.ClassificacaoABC_Grava(objClassABC)
    
    Exit Function
    
Erro_ClassificacaoABC_Grava:

    ClassificacaoABC_Grava = Err
    
    Select Case Err

        Case 45658
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146016)
        
    End Select
    
    Exit Function
    
End Function

Function CurvaABC_LePontos(objClassABC As ClassClassificacaoABC, objCurvaABC As ClassCurvaABC) As Long
'Lê dados da curva ABC no Banco de Dados e guarda os Pontos em objCurvaABC.colPontos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CurvaABC_LePontos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CURVAABC_LEPONTOS", X)
    If lErro <> SUCESSO Then Error 45659
    
    
    CurvaABC_LePontos = X.CurvaABC_LePontos(objClassABC, objCurvaABC)
    
    Exit Function
    
Erro_CurvaABC_LePontos:

    CurvaABC_LePontos = Err
    
    Select Case Err

        Case 45659
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146017)
        
    End Select
    
    Exit Function
    
End Function

Function ClassificacoesABC_Le(colNumIntCodigo As Collection) As Long
'Lê todas as Classificações coloca-as em ColNumIntCodigo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClassificacoesABC_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSIFICACOESABC_LE", X)
    If lErro <> SUCESSO Then Error 45660
    
    
    ClassificacoesABC_Le = X.ClassificacoesABC_Le(colNumIntCodigo)
    
    Exit Function
    
Erro_ClassificacoesABC_Le:

    ClassificacoesABC_Le = Err
    
    Select Case Err

        Case 45660
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146018)
        
    End Select
    
    Exit Function
    
End Function

Function ClassificacaoABC_Le(objClassABC As ClassClassificacaoABC) As Long
'Lê a ClassificacaoABC no BD à partir do Código e Da FilialEmpresa passados em objClassABC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClassificacaoABC_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSIFICACAOABC_LE", X)
    If lErro <> SUCESSO Then Error 45661
    
    
    ClassificacaoABC_Le = X.ClassificacaoABC_Le(objClassABC)
    
    Exit Function
    
Erro_ClassificacaoABC_Le:

    ClassificacaoABC_Le = Err
    
    Select Case Err

        Case 45661
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146019)
        
    End Select
    
    Exit Function
    
End Function

Function ClassificacaoABC_Le_NumInt(objClassABC As ClassClassificacaoABC) As Long
'Lê a ClassificacaoABC no BD à partir do NumInt passado em objClassABC

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClassificacaoABC_Le_NumInt
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSIFICACAOABC_LE_NUMINT", X)
    If lErro <> SUCESSO Then Error 45662
    
    
    ClassificacaoABC_Le_NumInt = X.ClassificacaoABC_Le_NumInt(objClassABC)
    
    Exit Function
    
Erro_ClassificacaoABC_Le_NumInt:

    ClassificacaoABC_Le_NumInt = Err
    
    Select Case Err

        Case 45662
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146020)
        
    End Select
    
    Exit Function
    
End Function

Function Estoque_Grava(objProdutoFilial As ClassProdutoFilial, iControleEstoque As Integer) As Long
'Grava o Estoque do Produto da Filial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Estoque_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUE_GRAVA", X)
    If lErro <> SUCESSO Then Error 45663
    
    
    Estoque_Grava = X.Estoque_Grava(objProdutoFilial, iControleEstoque)
    
    Exit Function
    
Erro_Estoque_Grava:

    Estoque_Grava = Err
    
    Select Case Err

        Case 45663
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146021)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueInicial_Grava(objEstoqueProduto As ClassEstoqueProduto, iAlmoxarifadoPadrao As Integer) As Long
'Grava o estoque inicial do produto
'??? revisar p/evitar repetir codigo, simplificando-o e diminuindo o tamanho desta funcao
'??? nao declarar uma variavel por linha (dim)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueInicial_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEINICIAL_GRAVA", X)
    If lErro <> SUCESSO Then Error 45664
    
    
    EstoqueInicial_Grava = X.EstoqueInicial_Grava(objEstoqueProduto, iAlmoxarifadoPadrao)
    
    Exit Function
    
Erro_EstoqueInicial_Grava:

    EstoqueInicial_Grava = Err
    
    Select Case Err

        Case 45664
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146022)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueInicial_Exclui(objEstoqueProduto As ClassEstoqueProduto) As Long
'Exclui o estoque inicial passado em objEstoqueProduto da tabela EstoqueProduto
'??? nao usar 1 linha p/cada dim

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueInicial_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEINICIAL_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45665
    
    
    EstoqueInicial_Exclui = X.EstoqueInicial_Exclui(objEstoqueProduto)
    
    Exit Function
    
Erro_EstoqueInicial_Exclui:

    EstoqueInicial_Exclui = Err
    
    Select Case Err

        Case 45665
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146023)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueProduto_Le_Almoxarifados(sProduto As String, colAlmoxarifados As Collection) As Long
'Lê todos os Almoxarifados do Produto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_Le_Almoxarifados
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_LE_ALMOXARIFADOS", X)
    If lErro <> SUCESSO Then Error 45666
    
    
    EstoqueProduto_Le_Almoxarifados = X.EstoqueProduto_Le_Almoxarifados(sProduto, colAlmoxarifados)
    
    Exit Function
    
Erro_EstoqueProduto_Le_Almoxarifados:

    EstoqueProduto_Le_Almoxarifados = Err
    
    Select Case Err

        Case 45666
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146024)
        
    End Select
    
    Exit Function
    
End Function

Function SldMesEst_Atualiza_Mes(objTela As Object, iMes As Integer, objSldMesEst As ClassSldMesEst, iCancela As Integer, objEstoque As ClassEstoqueMes) As Long
'Atualiza somente os mes, durante o decorrer do ano

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst_Atualiza_Mes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST_ATUALIZA_MES", X)
    If lErro <> SUCESSO Then Error 45667
    
    
    SldMesEst_Atualiza_Mes = X.SldMesEst_Atualiza_Mes(objTela, iMes, objSldMesEst, iCancela, objEstoque)
    
    Exit Function
    
Erro_SldMesEst_Atualiza_Mes:

    SldMesEst_Atualiza_Mes = Err
    
    Select Case Err

        Case 45667
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146025)
        
    End Select
    
    Exit Function
    
End Function

Function SldMesEst_Atualiza_Ano(objTela As Object, objSldMesEst As ClassSldMesEst, iCancela As Integer) As Long
'Atualiza Ano, abrindo um novo ano(Exercicio) para os produtos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst_Atualiza_Ano
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST_ATUALIZA_ANO", X)
    If lErro <> SUCESSO Then Error 45668
    
    
    SldMesEst_Atualiza_Ano = X.SldMesEst_Atualiza_Ano(objTela, objSldMesEst, iCancela)
    
    Exit Function
    
Erro_SldMesEst_Atualiza_Ano:

    SldMesEst_Atualiza_Ano = Err
    
    Select Case Err

        Case 45668
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146026)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueMes_Atualiza(objSldMesEst As ClassSldMesEst, iMes As Integer) As Long
'Atualiza a tabela de Estoque mes para a insercao de um novo mes
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueMes_Atualiza
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMES_ATUALIZA", X)
    If lErro <> SUCESSO Then Error 45669
    
    
    EstoqueMes_Atualiza = X.EstoqueMes_Atualiza(objSldMesEst, iMes)
    
    Exit Function
    
Erro_EstoqueMes_Atualiza:

    EstoqueMes_Atualiza = Err
    
    Select Case Err

        Case 45669
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146027)
        
    End Select
    
    Exit Function
    
End Function

Function FornecedorProduto_Le_Fornecedores(colFornecedores As Collection, objProduto As ClassProduto) As Long
'Lê todos nomes reduzidos de Fornecedores associados a sProduto e coloca-os em colFornecedores

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FornecedorProduto_Le_Fornecedores
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_FORNECEDORPRODUTO_LE_FORNECEDORES", X)
    If lErro <> SUCESSO Then Error 45670
    
    
    FornecedorProduto_Le_Fornecedores = X.FornecedorProduto_Le_Fornecedores(colFornecedores, objProduto)
    
    Exit Function
    
Erro_FornecedorProduto_Le_Fornecedores:

    FornecedorProduto_Le_Fornecedores = Err
    
    Select Case Err

        Case 45670
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146028)
        
    End Select
    
    Exit Function
    
End Function

Function FornecedorProduto_Le(objFornecedorProduto As ClassFornecedorProduto) As Long
'Lê dados da associacao Fornecedor-Produto à partir da chave fornecida em objFornecedorProduto. Retorna os dados em objFornecedorProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FornecedorProduto_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_FORNECEDORPRODUTO_LE", X)
    If lErro <> SUCESSO Then Error 45671
    
    
    FornecedorProduto_Le = X.FornecedorProduto_Le(objFornecedorProduto)
    
    Exit Function
    
Erro_FornecedorProduto_Le:

    FornecedorProduto_Le = Err
    
    Select Case Err

        Case 45671
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146029)
        
    End Select
    
    Exit Function
    
End Function

Function FornecedorProduto_Grava(objFornecedorProduto As ClassFornecedorProduto, iPadrao As Integer) As Long
'Grava a associação Produto X Fornecedor

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FornecedorProduto_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_FORNECEDORPRODUTO_GRAVA", X)
    If lErro <> SUCESSO Then Error 45672
    
    
    FornecedorProduto_Grava = X.FornecedorProduto_Grava(objFornecedorProduto, iPadrao)
    
    Exit Function
    
Erro_FornecedorProduto_Grava:

    FornecedorProduto_Grava = Err
    
    Select Case Err

        Case 45672
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146030)
        
    End Select
    
    Exit Function
    
End Function

Function FornecedorProduto_Exclui(objFornecedorProduto As ClassFornecedorProduto) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FornecedorProduto_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_FORNECEDORPRODUTO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45673
    
    
    FornecedorProduto_Exclui = X.FornecedorProduto_Exclui(objFornecedorProduto)
    
    Exit Function
    
Erro_FornecedorProduto_Exclui:

    FornecedorProduto_Exclui = Err
    
    Select Case Err

        Case 45673
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146031)
        
    End Select
    
    Exit Function
    
End Function

Function ProdutoFilial_AtualizaFornPadrao(sProduto As String, lForn As Long, iPadrao As Integer, lComando1 As Long, lComando2 As Long) As Long
'iPadrao indica se o fornecedor identificado por lForn está sendo atualizado para se tornar o fornecedor padrao ou nao.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutoFilial_AtualizaFornPadrao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTOFILIAL_ATUALIZAFORNPADRAO", X)
    If lErro <> SUCESSO Then Error 45674
    
    
    ProdutoFilial_AtualizaFornPadrao = X.ProdutoFilial_AtualizaFornPadrao(sProduto, lForn, iPadrao, lComando1, lComando2)
    
    Exit Function
    
Erro_ProdutoFilial_AtualizaFornPadrao:

    ProdutoFilial_AtualizaFornPadrao = Err
    
    Select Case Err

        Case 45674
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146032)
        
    End Select
    
    Exit Function
    
End Function

Function InventarioItens_Le(objInventario As ClassInventario) As Long
'preenche a colecao de itens do inventario passado e a data de objInventario
'os itens sao recuperados por ordem de numintdoc

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InventarioItens_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVENTARIOITENS_LE", X)
    If lErro <> SUCESSO Then Error 45675
    
    
    InventarioItens_Le = X.InventarioItens_Le(objInventario)
    
    Exit Function
    
Erro_InventarioItens_Le:

    InventarioItens_Le = Err
    
    Select Case Err

        Case 45675
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146033)
        
    End Select
    
    Exit Function
    
End Function

Function Inventario_Le(objInventario As ClassInventario) As Long
'Le atributos de objInventario EXCETO colItens
'Faz Select DISTINCT

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Inventario_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVENTARIO_LE", X)
    If lErro <> SUCESSO Then Error 45676
    
    
    Inventario_Le = X.Inventario_Le(objInventario)
    
    Exit Function
    
Erro_Inventario_Le:

    Inventario_Le = Err
    
    Select Case Err

        Case 45676
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146034)
        
    End Select
    
    Exit Function
    
End Function

Function Inventario_Grava(objInventario As ClassInventario, objContabil As ClassContabil) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Inventario_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVENTARIO_GRAVA", X)
    If lErro <> SUCESSO Then Error 45677
    
    
    Inventario_Grava = X.Inventario_Grava(objInventario, objContabil)
    
    Exit Function
    
Erro_Inventario_Grava:

    Inventario_Grava = Err
    
    Select Case Err

        Case 45677
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146035)
        
    End Select
    
    Exit Function
    
End Function

Function SldMesEst_Le_Custos(iMes As Integer, objSldMesEst As ClassSldMesEst) As Long
'Le na tabela SldMesEst os custos relativos aos parametros passados : iMes e objSldMesEst

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst_Le_Custos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST_LE_CUSTOS", X)
    If lErro <> SUCESSO Then Error 45678
    
    
    SldMesEst_Le_Custos = X.SldMesEst_Le_Custos(iMes, objSldMesEst)
    
    Exit Function
    
Erro_SldMesEst_Le_Custos:

    SldMesEst_Le_Custos = Err
    
    Select Case Err

        Case 45678
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146036)
        
    End Select
    
    Exit Function
    
End Function

Function CustoMedioAtual_Le(sProduto As String, dCustoMedioAtual As Double) As Long
'Parametros a serem passados: sProduto
'Le o Ultimo CustoMedio do Ano-Mes aberto e retorna em dCustoMedioAtual

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CustoMedioAtual_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CUSTOMEDIOATUAL_LE", X)
    If lErro <> SUCESSO Then Error 45679
    
    
    CustoMedioAtual_Le = X.CustoMedioAtual_Le(sProduto, dCustoMedioAtual)
    
    Exit Function
    
Erro_CustoMedioAtual_Le:

    CustoMedioAtual_Le = Err
    
    Select Case Err

        Case 45679
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146037)
        
    End Select
    
    Exit Function
    
End Function

Function ItemInventario_Automatico_EmTrans(lCodigo As Long, ByVal iQtde As Integer) As Long
'obtem numero interno para identificar um item de inventario
'procurar em MATConfig por codigo = NUM_PROX_ITEM_INVENTARIO
'TEM QUE SER CHAMADO EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemInventario_Automatico_EmTrans
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ITEMINVENTARIO_AUTOMATICO_EMTRANS", X)
    If lErro <> SUCESSO Then Error 45680
    
    
    ItemInventario_Automatico_EmTrans = X.ItemInventario_Automatico_EmTrans(lCodigo, iQtde)
    
    Exit Function
    
Erro_ItemInventario_Automatico_EmTrans:

    ItemInventario_Automatico_EmTrans = Err
    
    Select Case Err

        Case 45680
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146038)
        
    End Select
    
    Exit Function
    
End Function

Function InvLotePendente_Le(objInvLote As ClassInvLote) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InvLotePendente_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVLOTEPENDENTE_LE", X)
    If lErro <> SUCESSO Then Error 45685
    
    
    InvLotePendente_Le = X.InvLotePendente_Le(objInvLote)
    
    Exit Function
    
Erro_InvLotePendente_Le:

    InvLotePendente_Le = Err
    
    Select Case Err

        Case 45685
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146039)
        
    End Select
    
    Exit Function
    
End Function

Function InvLote_Le(objInvLote As ClassInvLote) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InvLote_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVLOTE_LE", X)
    If lErro <> SUCESSO Then Error 45686
    
    
    InvLote_Le = X.InvLote_Le(objInvLote)
    
    Exit Function
    
Erro_InvLote_Le:

    InvLote_Le = Err
    
    Select Case Err

        Case 45686
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146040)
        
    End Select
    
    Exit Function
    
End Function

Function InventarioPendenteItens_Le(objInventario As ClassInventario) As Long
'preenche a colecao de itens do inventario pendente passado e a data de objInventario
'os itens sao recuperados por ordem de numintdoc

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InventarioPendenteItens_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVENTARIOPENDENTESITENS_LE", X)
    If lErro <> SUCESSO Then Error 45687
    
    
    InventarioPendenteItens_Le = X.InventarioPendenteItens_Le(objInventario)
    
    Exit Function
    
Erro_InventarioPendenteItens_Le:

    InventarioPendenteItens_Le = Err
    
    Select Case Err

        Case 45687
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146041)
        
    End Select
    
    Exit Function
    
End Function

Function InventarioPendente_Le(objInventario As ClassInventario) As Long
'Le atributos de objInventario EXCETO colItens
'Faz Select DISTINCT

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InventarioPendente_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVENTARIOPENDENTE_LE", X)
    If lErro <> SUCESSO Then Error 45688
    
    
    InventarioPendente_Le = X.InventarioPendente_Le(objInventario)
    
    Exit Function
    
Erro_InventarioPendente_Le:

    InventarioPendente_Le = Err
    
    Select Case Err

        Case 45688
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146042)
        
    End Select
    
    Exit Function
    
End Function

Function InventarioPendente_Grava(objInventario As ClassInventario) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InventarioPendente_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVENTARIOPENDENTES_GRAVA", X)
    If lErro <> SUCESSO Then Error 45689
    
    
    InventarioPendente_Grava = X.InventarioPendente_Grava(objInventario)
    
    Exit Function
    
Erro_InventarioPendente_Grava:

    InventarioPendente_Grava = Err
    
    Select Case Err

        Case 45689
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146043)
        
    End Select
    
    Exit Function
    
End Function

Function InventarioPendente_Exclui(objInventario As ClassInventario)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InventarioPendente_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVENTARIOPENDENTE_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45690
    
    
    InventarioPendente_Exclui = X.InventarioPendente_Exclui(objInventario)
    
    Exit Function
    
Erro_InventarioPendente_Exclui:

    InventarioPendente_Exclui = Err
    
    Select Case Err

        Case 45690
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146044)
        
    End Select
    
    Exit Function
    
End Function

Function InvLote_Critica_Atualizado(objInvLote As ClassInvLote, iLoteAtualizado As Integer) As Long
'testa se o Invlote já foi atualizado. Se já foi iLoteAtualizado = LOTE_ATUALIZADO senao retorna 0.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InvLote_Critica_Atualizado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVLOTE_CRITICA_ATUALIZADO", X)
    If lErro <> SUCESSO Then Error 45691
    
    
    InvLote_Critica_Atualizado = X.InvLote_Critica_Atualizado(objInvLote, iLoteAtualizado)
    
    Exit Function
    
Erro_InvLote_Critica_Atualizado:

    InvLote_Critica_Atualizado = Err
    
    Select Case Err

        Case 45691
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146045)
        
    End Select
    
    Exit Function
    
End Function

Function InvLote_Automatico(lLote As Long) As Long
'retorna o número do proximo lote disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InvLote_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVLOTE_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 45692
    
    
    InvLote_Automatico = X.InvLote_Automatico(lLote)
    
    Exit Function
    
Erro_InvLote_Automatico:

    InvLote_Automatico = Err
    
    Select Case Err

        Case 45692
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146046)
        
    End Select
    
    Exit Function
    
End Function

Function InvLotePendente_Exclui(objInvLote As ClassInvLote) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InvLotePendente_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVLOTEPENDENTE_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45693
    
    
    InvLotePendente_Exclui = X.InvLotePendente_Exclui(objInvLote)
    
    Exit Function
    
Erro_InvLotePendente_Exclui:

    InvLotePendente_Exclui = Err
    
    Select Case Err

        Case 45693
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146047)
        
    End Select
    
    Exit Function
    
End Function

Function InvLotePendente_Grava(objInvLote As ClassInvLote) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InvLotePendente_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVLOTEPENDENTE_GRAVA", X)
    If lErro <> SUCESSO Then Error 45694
    
    
    InvLotePendente_Grava = X.InvLotePendente_Grava(objInvLote)
    
    Exit Function
    
Erro_InvLotePendente_Grava:

    InvLotePendente_Grava = Err
    
    Select Case Err

        Case 45694
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146048)
        
    End Select
    
    Exit Function
    
End Function

Function InvLotePendente_Le_Desatualizados(iFilialEmpresa As Integer, colInvLote As Collection, iIncluiLotesAtualizando As Integer) As Long
'Le todos os lotes com IdAtualizacao = LOTE_DESATUALIZADO ou todos pendentes

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InvLotePendente_Le_Desatualizados
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVLOTEPENDENTE_LE_DESATUALIZADOS", X)
    If lErro <> SUCESSO Then Error 45695
    
    
    InvLotePendente_Le_Desatualizados = X.InvLotePendente_Le_Desatualizados(iFilialEmpresa, colInvLote, iIncluiLotesAtualizando)
    
    Exit Function
    
Erro_InvLotePendente_Le_Desatualizados:

    InvLotePendente_Le_Desatualizados = Err
    
    Select Case Err

        Case 45695
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146049)
        
    End Select
    
    Exit Function
    
End Function

Function InvLotePendente_Atualiza(colInvLote As Collection, iIdAtualizacao As Integer) As Long
'atualiza o IdAtualizacao das tabelas Configuracao e LotePendente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InvLotePendente_Atualiza
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVLOTEPENDENTE_ATUALIZA", X)
    If lErro <> SUCESSO Then Error 45696
    
    
    InvLotePendente_Atualiza = X.InvLotePendente_Atualiza(colInvLote, iIdAtualizacao)
    
    Exit Function
    
Erro_InvLotePendente_Atualiza:

    InvLotePendente_Atualiza = Err
    
    Select Case Err

        Case 45696
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146050)
        
    End Select
    
    Exit Function
    
End Function

Function InvLotePendente_Le1(iFilialEmpresa As Integer) As Long
'Lê todos os lotes de uma FilialEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InvLotePendente_Le1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVLOTEPENDENTE_LE1", X)
    If lErro <> SUCESSO Then Error 45697
    
    
    InvLotePendente_Le1 = X.InvLotePendente_Le1(iFilialEmpresa)
    
    Exit Function
    
Erro_InvLotePendente_Le1:

    InvLotePendente_Le1 = Err
    
    Select Case Err

        Case 45697
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146051)
        
    End Select
    
    Exit Function
    
End Function

Function OrdemDeProducao_Le_ComItens(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'carrega do bd a ordem de producao inclusive os seus itens

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OrdemDeProducao_Le_ComItens
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ORDEMDEPRODUCAO_LE_COMITENS", X)
    If lErro <> SUCESSO Then Error 45711
    
    
    OrdemDeProducao_Le_ComItens = X.OrdemDeProducao_Le_ComItens(objOrdemDeProducao)
    
    Exit Function
    
Erro_OrdemDeProducao_Le_ComItens:

    OrdemDeProducao_Le_ComItens = Err
    
    Select Case Err

        Case 45711
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146052)
        
    End Select
    
    Exit Function
    
End Function

Function Empenhos_Excluir_ItemOP(objItemOP As ClassItemOP) As Long
'Exclui os empenhos associados a um item de OP

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Empenhos_Excluir_ItemOP
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_EMPENHOS_EXCLUIR_ITEMOP", X)
    If lErro <> SUCESSO Then Error 45714
    
    
    Empenhos_Excluir_ItemOP = X.Empenhos_Excluir_ItemOP(objItemOP)
    
    Exit Function
    
Erro_Empenhos_Excluir_ItemOP:

    Empenhos_Excluir_ItemOP = Err
    
    Select Case Err

        Case 45714
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146053)
        
    End Select
    
    Exit Function
    
End Function


Function EstoqueProduto_AlterarQtdOP(objEstoqueProduto As ClassEstoqueProduto, dQtde As Double) As Long
'Aumenta (ou diminui se dQtde for negativo) a qtde a ser produzida em um produto-almoxarifado
'objEstoqueProduto.sProduto e objEstoqueProduto.iAlmoxarifado tem que estar preenchidos
'DEVE SER CHAMADO EM TRANSACAO
'??? sugestao: poderia ser criada uma funcao generica pela passagem do nome do campo "qtde" a ser alterado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_AlterarQtdOP
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_ALTERARQTDOP", X)
    If lErro <> SUCESSO Then Error 45716
    
    
    EstoqueProduto_AlterarQtdOP = X.EstoqueProduto_AlterarQtdOP(objEstoqueProduto, dQtde)
    
    Exit Function
    
Erro_EstoqueProduto_AlterarQtdOP:

    EstoqueProduto_AlterarQtdOP = Err
    
    Select Case Err

        Case 45716
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146054)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueProdutoCC_Le(objEstoqueProduto As ClassEstoqueProduto) As Long
'le a conta contabil que está em estoque produto se não encontrou
'le a conta contabil que está em almoxarifado para ser preenchida como Padrão
'Devolve a contacontabil dentro do objEstoqueProduto.sContaContabil

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProdutoCC_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTOCC_LE", X)
    If lErro <> SUCESSO Then Error 45728
    
    
    EstoqueProdutoCC_Le = X.EstoqueProdutoCC_Le(objEstoqueProduto)
    
    Exit Function
    
Erro_EstoqueProdutoCC_Le:

    EstoqueProdutoCC_Le = Err
    
    Select Case Err

        Case 45728
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146055)
        
    End Select
    
    Exit Function
    
End Function

Function TabelaPrecoItem_Le(objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
'Lê a tabela de preço a partir da chave fornecida em objTabelaPreco. Retorna os dados em objTabelaPreco

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPrecoItem_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TABELAPRECOITEM_LE", X)
    If lErro <> SUCESSO Then Error 45747
    
    
    TabelaPrecoItem_Le = X.TabelaPrecoItem_Le(objTabelaPrecoItem)
    
    Exit Function
    
Erro_TabelaPrecoItem_Le:

    TabelaPrecoItem_Le = Err
    
    Select Case Err

        Case 45747
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146056)
        
    End Select
    
    Exit Function
    
End Function

Function TabelaPrecoItem_Exclui(objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
'Exclui a Tabela e Preço do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPrecoItem_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TABELAPRECOITEM_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45751
    
    
    TabelaPrecoItem_Exclui = X.TabelaPrecoItem_Exclui(objTabelaPrecoItem)
    
    Exit Function
    
Erro_TabelaPrecoItem_Exclui:

    TabelaPrecoItem_Exclui = Err
    
    Select Case Err

        Case 45751
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146057)
        
    End Select
    
    Exit Function
    
End Function

Function TabelaPrecoItem_Grava(objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
'Rotina que grava Tabela de Preço Item

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPrecoItem_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TABELAPRECOITEM_GRAVA", X)
    If lErro <> SUCESSO Then Error 45752
    
    
    TabelaPrecoItem_Grava = X.TabelaPrecoItem_Grava(objTabelaPrecoItem)
    
    Exit Function
    
Erro_TabelaPrecoItem_Grava:

    TabelaPrecoItem_Grava = Err
    
    Select Case Err

        Case 45752
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146058)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueProduto_TestaAssociacao(sProduto As String, Almoxarifado As Object) As Long
'Verifica associação entre Produto e almoxarifado na filial corrente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_TestaAssociacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_TESTAASSOCIACAO", X)
    If lErro <> SUCESSO Then Error 45755
    
    
    EstoqueProduto_TestaAssociacao = X.EstoqueProduto_TestaAssociacao(sProduto, Almoxarifado)
    
    Exit Function
    
Erro_EstoqueProduto_TestaAssociacao:

    EstoqueProduto_TestaAssociacao = Err
    
    Select Case Err

        Case 45755
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146059)
        
    End Select
    
    Exit Function
    
End Function

Function ProdutosProduzidosCustos_Grava(iAno As Integer, iMes As Integer, colProdutoCustoAtual As colProdutoCusto) As Long
'Atualiza os Custo de Producao dos produtos na tabela de SldMesEst

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutosProduzidosCustos_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTOSPRODUZIDOSCUSTOS_GRAVA", X)
    If lErro <> SUCESSO Then Error 45756
    
    
    ProdutosProduzidosCustos_Grava = X.ProdutosProduzidosCustos_Grava(iAno, iMes, colProdutoCustoAtual)
    
    Exit Function
    
Erro_ProdutosProduzidosCustos_Grava:

    ProdutosProduzidosCustos_Grava = Err
    
    Select Case Err

        Case 45756
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146060)
        
    End Select
    
    Exit Function
    
End Function

Function ProdutosProduzidosCusto_Le(iMes As Integer, iAno As Integer, colProdutoCusto As Collection) As Long
'Le os Produtos de sldMesEst faz selecao por giFilialEmpresa, iAno, DataInicial, DataFinal, APROPR_CUSTO_REAL
'apartir da ordem de CODIGO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutosProduzidosCusto_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTOSPRODUZIDOSCUSTO_LE", X)
    If lErro <> SUCESSO Then Error 45757
    
    
    ProdutosProduzidosCusto_Le = X.ProdutosProduzidosCusto_Le(iMes, iAno, colProdutoCusto)
    
    Exit Function
    
Erro_ProdutosProduzidosCusto_Le:

    ProdutosProduzidosCusto_Le = Err
    
    Select Case Err

        Case 45757
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146061)
        
    End Select
    
    Exit Function
    
End Function

Function ProdutosProduzidosTodosCustos_Le(iMes As Integer, iAno As Integer, colProdutoCustoAnterior As Collection) As Long
'Pega os Produtos em SldMesEst e faz uma selecao por PRODUTO_PRODUZIVEL, NAO_GERENCIAL, PRODUTO_ATIVO
'e ordena por CODIGO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutosProduzidosTodosCustos_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTOSPRODUZIDOSTODOSCUSTOS_LE", X)
    If lErro <> SUCESSO Then Error 45758
    
    
    ProdutosProduzidosTodosCustos_Le = X.ProdutosProduzidosTodosCustos_Le(iMes, iAno, colProdutoCustoAnterior)
    
    Exit Function
    
Erro_ProdutosProduzidosTodosCustos_Le:

    ProdutosProduzidosTodosCustos_Le = Err
    
    Select Case Err

        Case 45758
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146062)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueMesNaoApurado_Le(objEstoqueMes As ClassEstoqueMes) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueMesNaoApurado_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMESNAOAPURADO_LE", X)
    If lErro <> SUCESSO Then Error 45759
    
    
    EstoqueMesNaoApurado_Le = X.EstoqueMesNaoApurado_Le(objEstoqueMes)
    
    Exit Function
    
Erro_EstoqueMesNaoApurado_Le:

    EstoqueMesNaoApurado_Le = Err
    
    Select Case Err

        Case 45759
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146063)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueMes_Le_Anos(colAnos As Collection) As Long
'Le todos os anos da tabela EstoqueMes com FilialEmpresa = giFilialEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueMes_Le_Anos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMES_LE_ANOS", X)
    If lErro <> SUCESSO Then Error 45760
    
    
    EstoqueMes_Le_Anos = X.EstoqueMes_Le_Anos(colAnos)
    
    Exit Function
    
Erro_EstoqueMes_Le_Anos:

    EstoqueMes_Le_Anos = Err
    
    Select Case Err

        Case 45760
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146064)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueMes_Le_Meses(iAno As Integer, colMeses As Collection) As Long
'Le todos os meses da tabela EstoqueMes com Ano = iAno e FilialEmpresa = giFilialEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueMes_Le_Meses
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMES_LE_MESES", X)
    If lErro <> SUCESSO Then Error 45761
    
    
    EstoqueMes_Le_Meses = X.EstoqueMes_Le_Meses(iAno, colMeses)
    
    Exit Function
    
Erro_EstoqueMes_Le_Meses:

    EstoqueMes_Le_Meses = Err
    
    Select Case Err

        Case 45761
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146065)
        
    End Select
    
    Exit Function
    
End Function

Function ProdutoFilial_Le_Lock(lComando As Long, objProdutoFilial As ClassProdutoFilial) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutoFilial_Le_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTOFILIAL_LE_LOCK", X)
    If lErro <> SUCESSO Then Error 45762
    
    
    ProdutoFilial_Le_Lock = X.ProdutoFilial_Le_Lock(lComando, objProdutoFilial)
    
    Exit Function
    
Erro_ProdutoFilial_Le_Lock:

    ProdutoFilial_Le_Lock = Err
    
    Select Case Err

        Case 45762
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146066)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueMes_Le_Lock(lComando As Long, objEstoqueMes As ClassEstoqueMes) As Long
'Le a tabela EstoqueMes e loca o registro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueMes_Le_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMES_LE_LOCK", X)
    If lErro <> SUCESSO Then Error 45763
    
    
    EstoqueMes_Le_Lock = X.EstoqueMes_Le_Lock(lComando, objEstoqueMes)
    
    Exit Function
    
Erro_EstoqueMes_Le_Lock:

    EstoqueMes_Le_Lock = Err
    
    Select Case Err

        Case 45763
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146067)
        
    End Select
    
    Exit Function
    
End Function

Function MovEstoque_Le_ProdutoMes(iMes As Integer, iAno As Integer, objItemMovEstoque As ClassItemMovEstoque) As Long
'Le o primeiro Movimento de Estoque do Mes-Ano em questão
'com parametros e o sProduto passado dentro do objMovEstoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Le_ProdutoMes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_LE_PRODUTOMES", X)
    If lErro <> SUCESSO Then Error 45764
    
    
    MovEstoque_Le_ProdutoMes = X.MovEstoque_Le_ProdutoMes(iMes, iAno, objItemMovEstoque)
    
    Exit Function
    
Erro_MovEstoque_Le_ProdutoMes:

    MovEstoque_Le_ProdutoMes = Err
    
    Select Case Err

        Case 45764
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146068)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_Almoxarifados_Filial(sProduto As String, colAlmoxarifados As Collection) As Long
'Lê todos os Almoxarifados do Produto para a filial corrente ou EMPRESA_TODA

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_Almoxarifados_Filial
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LE_ALMOXARIFADOS_FILIAL", X)
    If lErro <> SUCESSO Then Error 45836
    
    
    Produto_Le_Almoxarifados_Filial = X.Produto_Le_Almoxarifados_Filial(sProduto, colAlmoxarifados)
    
    Exit Function
    
Erro_Produto_Le_Almoxarifados_Filial:

    Produto_Le_Almoxarifados_Filial = Err
    
    Select Case Err

        Case 45836
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146069)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueProduto_Exclui(objEstoqueProduto As ClassEstoqueProduto) As Long
'Exclui Estoque X Produto passado como parâmetro
'??? nao exclui sldmesest ???? nao testa se tem movto ? Acho que deveria ser a exclusao igual a de estoqueinicial

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45837
    
    
    EstoqueProduto_Exclui = X.EstoqueProduto_Exclui(objEstoqueProduto)
    
    Exit Function
    
Erro_EstoqueProduto_Exclui:

    EstoqueProduto_Exclui = Err
    
    Select Case Err

        Case 45837
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146070)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueProduto_AtualizaAlmoxarifado(objEstoqueProduto As ClassEstoqueProduto, iAlmoxarifadoPadraoFilial As Integer) As Long
'iAlmoxarifadoPadraoFilial com 1(um) ou 0(zero) p/indicar se o almoxarifado é ou nao o padrao p/(filial,produto)
'   permitirá alterar o almoxarifado padrao a nivel de filialEmpresa ou EMPRESA_TODA,
'   (dependendo de como o usuario tiver se logado), e a localizacao fisica no almoxarifado
'       se o registro nao existir em EstoqueProduto sera criado
'   o almoxarifado padrao ficara na tabela ProdutoFilial
'       se nao existir o reg este sera criado
'??? será que deve ter esta funcionalidade toda ?

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_AtualizaAlmoxarifado
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_ATUALIZAALMOXARIFADO", X)
    If lErro <> SUCESSO Then Error 45838
    
    
    EstoqueProduto_AtualizaAlmoxarifado = X.EstoqueProduto_AtualizaAlmoxarifado(objEstoqueProduto, iAlmoxarifadoPadraoFilial)
    
    Exit Function
    
Erro_EstoqueProduto_AtualizaAlmoxarifado:

    EstoqueProduto_AtualizaAlmoxarifado = Err
    
    Select Case Err

        Case 45838
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146071)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_ExisteProduto(iExiste As Integer) As Long
'verifica se tem algum produto cadastrado.
'se existir iExiste = POSSUI_PRODUTO , caso contrario iExiste = NAO_POSSUI_PRODUTO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_ExisteProduto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_EXISTEPRODUTO", X)
    If lErro <> SUCESSO Then Error 44843
    
    
    Produto_ExisteProduto = X.Produto_ExisteProduto(iExiste)
    
    Exit Function
    
Erro_Produto_ExisteProduto:

    Produto_ExisteProduto = Err
    
    Select Case Err

        Case 44843
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146072)
        
    End Select
    
    Exit Function
    
End Function

Function Segmento_Grava_Produto(colSegmentos As Collection) As Long
'grava os registros da colecao na tabela Segmento

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Segmento_Grava_Produto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SEGMENTO_GRAVA_PRODUTO", X)
    If lErro <> SUCESSO Then Error 44844
    
    
    Segmento_Grava_Produto = X.Segmento_Grava_Produto(colSegmentos)
    
    Exit Function
    
Erro_Segmento_Grava_Produto:

    Segmento_Grava_Produto = Err
    
    Select Case Err

        Case 44844
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146073)
        
    End Select
    
    Exit Function
    
End Function

Function Segmento_Grava_Produto_Trans(colSegmentos As Collection) As Long
'grava os registros da colecao na tabela Segmento de Produto
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Segmento_Grava_Produto_Trans
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SEGMENTO_GRAVA_PRODUTO_TRANS", X)
    If lErro <> SUCESSO Then Error 44845
    
    
    Segmento_Grava_Produto_Trans = X.Segmento_Grava_Produto_Trans(colSegmentos)
    
    Exit Function
    
Erro_Segmento_Grava_Produto_Trans:

    Segmento_Grava_Produto_Trans = Err
    
    Select Case Err

        Case 44845
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146074)
        
    End Select
    
    Exit Function
    
End Function

Function Empenho_Le_SemCodigo(objEmpenho As ClassEmpenho) As Long
'Lê os dados de Empenho à partir da chave candidata fornecida em objEmpenho, com os seguintes campos preenchidos:
'objEmpenho.iFilialEmpresa, objEmpenho.sCodigoOP, objEmpenho.iItemOP, objEmpenho.sProduto, objEmpenho.iAlmoxarifado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Empenho_Le_SemCodigo
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_EMPENHO_LE_SEMCODIGO", X)
    If lErro <> SUCESSO Then Error 45879
    
    
    Empenho_Le_SemCodigo = X.Empenho_Le_SemCodigo(objEmpenho)
    
    Exit Function
    
Erro_Empenho_Le_SemCodigo:

    Empenho_Le_SemCodigo = Err
    
    Select Case Err

        Case 45879
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146075)
        
    End Select
    
    Exit Function
    
End Function

Function Kit_Le_Empenho(objKit As ClassKit) As Long
'Lê uma versão mais recente na tabela de Kits de Produtos à partir do produto passado em objKit
'e devolve os dados do Kit em objKit

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Kit_Le_Empenho
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_KIT_LE_EMPENHO", X)
    If lErro <> SUCESSO Then Error 45880
    
    
    Kit_Le_Empenho = X.Kit_Le_Empenho(objKit)
    
    Exit Function
    
Erro_Kit_Le_Empenho:

    Kit_Le_Empenho = Err
    
    Select Case Err

        Case 45880
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146076)
        
    End Select
    
    Exit Function
    
End Function

Function AlmoxarifadoPadrao_Le_NomeReduzido(sProduto As String, sNomeReduzidoAlmox As String) As Long
'Retorna o Nome reduzido do Almoxarifado Padrão do Produto em Questão com a Filial Empresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_AlmoxarifadoPadrao_Le_NomeReduzido
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ALMOXARIFADOPADRAO_LE_NOMEREDUZIDO", X)
    If lErro <> SUCESSO Then Error 45962
    
    
    AlmoxarifadoPadrao_Le_NomeReduzido = X.AlmoxarifadoPadrao_Le_NomeReduzido(sProduto, sNomeReduzidoAlmox)
    
    Exit Function
    
Erro_AlmoxarifadoPadrao_Le_NomeReduzido:

    AlmoxarifadoPadrao_Le_NomeReduzido = Err
    
    Select Case Err

        Case 45962
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146077)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueMes_Insere(objEstoqueMes As ClassEstoqueMes) As Long
'Insere um novo mes aberto na tabela de EstoqueMes
'IMPORTANTE: TEM QUE SER CHAMADO EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueMes_Insere
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMES_INSERE", X)
    If lErro <> SUCESSO Then Error 44967
    
    
    EstoqueMes_Insere = X.EstoqueMes_Insere(objEstoqueMes)
    
    Exit Function
    
Erro_EstoqueMes_Insere:

    EstoqueMes_Insere = Err
    
    Select Case Err

        Case 44967
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146078)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueMes_Le_Aberto(objEstoqueMes As ClassEstoqueMes) As Long
'Le o mes em aberto do Estoque para a filial em questão (giFilialEmpresa).

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueMes_Le_Aberto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMES_LE_ABERTO", X)
    If lErro <> SUCESSO Then Error 55029
    
    
    EstoqueMes_Le_Aberto = X.EstoqueMes_Le_Aberto(objEstoqueMes)
    
    Exit Function
    
Erro_EstoqueMes_Le_Aberto:

    EstoqueMes_Le_Aberto = Err
    
    Select Case Err

        Case 55029
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146079)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueMes_Le_PrimeiroAnoMes(objEstoqueMes As ClassEstoqueMes) As Long
'Le o primeiro ano/mes do Estoque para a filial em questão (giFilialEmpresa).

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueMes_Le_PrimeiroAnoMes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMES_LE_PRIMEIROANOMES", X)
    If lErro <> SUCESSO Then Error 55036
    
    
    EstoqueMes_Le_PrimeiroAnoMes = X.EstoqueMes_Le_PrimeiroAnoMes(objEstoqueMes)
    
    Exit Function
    
Erro_EstoqueMes_Le_PrimeiroAnoMes:

    EstoqueMes_Le_PrimeiroAnoMes = Err
    
    Select Case Err

        Case 55036
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146080)
        
    End Select
    
    Exit Function
    
End Function

Function CustoStandard_Le(sProduto As String, dCustoStandard As Double) As Long
'Parametros a serem passados: sProduto
'Le o CustoStandard do Ano-Mes aberto e retorna em dCustoStandard

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CustoStandard_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CUSTOSTANDARD_LE", X)
    If lErro <> SUCESSO Then Error 55042
    
    
    CustoStandard_Le = X.CustoStandard_Le(sProduto, dCustoStandard)
    
    Exit Function
    
Erro_CustoStandard_Le:

    CustoStandard_Le = Err
    
    Select Case Err

        Case 55042
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146081)
        
    End Select
    
    Exit Function
    
End Function

Public Function Calcula_Custo(dCusto As Double, ByVal iFilialEmpresa As Integer, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstAberto As Integer, ByVal iMesEstAberto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer, ByVal sProduto As String) As Long
'devolve o custo do produto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Calcula_Custo
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CALCULA_CUSTO", X)
    If lErro <> SUCESSO Then Error 55056
    
    
    Calcula_Custo = X.Calcula_Custo(dCusto, iFilialEmpresa, iApropriacaoCusto, iAnoEstAberto, iMesEstAberto, iAnoEstApurado, iMesEstApurado, sProduto)
    
    Exit Function
    
Erro_Calcula_Custo:

    Calcula_Custo = Err
    
    Select Case Err

        Case 55056
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146082)
        
    End Select
    
    Exit Function
    
End Function

Sub Converte_NaturezaProduto(ByVal iNatureza As Integer, sNatureza As String)
'converte a natureza do produto do valor inteiro para o seu correspondente textual (materia prima/produto intermadiario, ...)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Converte_NaturezaProduto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CONVERTE_NATUREZAPRODUTO", X)
    If lErro <> SUCESSO Then Error 55168
    
    
    Call X.Converte_NaturezaProduto(iNatureza, sNatureza)
    
    Exit Sub
    
Erro_Converte_NaturezaProduto:

    Select Case Err

        Case 55168
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146083)
        
    End Select
    
    Exit Sub
    
End Sub

Function Produto_Lock_Shared(lComando As Long, objProduto As ClassProduto) As Long
'Lê o faz o "lock compartilhado" no Produto passado por parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Lock_Shared
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LOCK_SHARED", X)
    If lErro <> SUCESSO Then Error 55236
    
    
    Produto_Lock_Shared = X.Produto_Lock_Shared(lComando, objProduto)

    Exit Function
    
Erro_Produto_Lock_Shared:
    
    Produto_Lock_Shared = Err
    
    Select Case Err

        Case 55236
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146084)
        
    End Select
    
    Exit Function

End Function

Function OrdemDeProducao_Baixa(sOPCodigo As String) As Long
'Baixa a ordem de producao, seus itens e seus empenhos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OrdemDeProducao_Baixa
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ORDEMDEPRODUCAO_BAIXA", X)
    If lErro <> SUCESSO Then Error 55274
    
    
    OrdemDeProducao_Baixa = X.OrdemDeProducao_Baixa(sOPCodigo)

    Exit Function
    
Erro_OrdemDeProducao_Baixa:
    
    OrdemDeProducao_Baixa = Err
    
    Select Case Err

        Case 55274
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146085)
        
    End Select
    
    Exit Function

End Function

Function Empenhos_Excluir_ItemOP1(objItemOP As ClassItemOP, lComando As Long, lComando1 As Long) As Long
'Exclui os empenhos associados a um item de OP

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Empenhos_Excluir_ItemOP1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_EMPENHOS_EXCLUIR_ITEMOP1", X)
    If lErro <> SUCESSO Then Error 55275
    
    
    Empenhos_Excluir_ItemOP1 = X.Empenhos_Excluir_ItemOP1(objItemOP, lComando, lComando1)

    Exit Function
    
Erro_Empenhos_Excluir_ItemOP1:
    
    Empenhos_Excluir_ItemOP1 = Err
    
    Select Case Err

        Case 55275
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146086)
        
    End Select
    
    Exit Function

End Function

Function OrdemDeProducao_Baixa1(alComando() As Long, sOPCodigo As String) As Long
'Verifica se todos os itens da OP estão baixados, se tiverem ==> Baixa a ordem de producao, seus itens e seus empenhos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OrdemDeProducao_Baixa1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ORDEMDEPRODUCAO_BAIXA1", X)
    If lErro <> SUCESSO Then Error 55297
    
    
    OrdemDeProducao_Baixa1 = X.OrdemDeProducao_Baixa1(alComando, sOPCodigo)

    Exit Function
    
Erro_OrdemDeProducao_Baixa1:
    
    OrdemDeProducao_Baixa1 = Err
    
    Select Case Err

        Case 55297
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146087)
        
    End Select
    
    Exit Function

End Function

Function EmpenhoQtdRequisitada_Grava(lComando As Long, lComando1 As Long, objMovEstoque As ClassMovEstoque) As Long
'atualiza o valor da quantidade requisitada no registro de empenho

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EmpenhoQtdRequisitada_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_EMPENHOQTDREQUISITADA_GRAVA", X)
    If lErro <> SUCESSO Then Error 55304
    
    
    EmpenhoQtdRequisitada_Grava = X.EmpenhoQtdRequisitada_Grava(lComando, lComando1, objMovEstoque)

    Exit Function
    
Erro_EmpenhoQtdRequisitada_Grava:
    
    EmpenhoQtdRequisitada_Grava = Err
    
    Select Case Err

        Case 55304
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146088)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_ReqProducao_Grava1(lComando As Long, lComando1 As Long, lComando2 As Long, objItemMovEstoque As ClassItemMovEstoque) As Long
'Grava o Movimento de Estoque de Requisição de Produção
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_ReqProducao_Grava1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_REQPRODUCAO_GRAVA1", X)
    If lErro <> SUCESSO Then Error 55310
    
    
    MovEstoque_ReqProducao_Grava1 = X.MovEstoque_ReqProducao_Grava1(lComando, lComando1, lComando2, objItemMovEstoque)

    Exit Function
    
Erro_MovEstoque_ReqProducao_Grava1:
    
    MovEstoque_ReqProducao_Grava1 = Err
    
    Select Case Err

        Case 55310
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146089)
        
    End Select
    
    Exit Function

End Function

Function InventarioPendente_Critica_Lote(objInvLote As ClassInvLote, iNumIguais As Integer) As Long
'verifica se o numero de lotes em InventarioLotePedente é igual ao de InvLotePendente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InventarioPendente_Critica_Lote
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVENTARIOPENDENTE_CRITICA_LOTE", X)
    If lErro <> SUCESSO Then Error 54707
    
    
    InventarioPendente_Critica_Lote = X.InventarioPendente_Critica_Lote(objInvLote, iNumIguais)

    Exit Function
    
Erro_InventarioPendente_Critica_Lote:
    
    InventarioPendente_Critica_Lote = Err
    
    Select Case Err

        Case 54707
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146090)
        
    End Select
    
    Exit Function

End Function

Function InvLotePendente_Atualiza1(objInvLotePendente As ClassInvLote) As Long
'Atualiza o campo de InvLotePendente que indica o numero de lotes "NumItensAtual"

Dim X As Object, lErro As Long
    
On Error GoTo Erro_InvLotePendente_Atualiza1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_INVLOTEPENDENTE_ATUALIZA1", X)
    If lErro <> SUCESSO Then Error 54708
    
    
    InvLotePendente_Atualiza1 = X.InvLotePendente_Atualiza1(objInvLotePendente)

    Exit Function
    
Erro_InvLotePendente_Atualiza1:
    
    InvLotePendente_Atualiza1 = Err
    
    Select Case Err

        Case 54708
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146091)
        
    End Select
    
    Exit Function

End Function

Function Kit_Le_Produto(objKit As ClassKit, colKits As Collection) As Long
'Função que le todos os kits, produtos iguais ao codigo passado e diferente da versao passada
'retornado na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Kit_Le_Produto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_KIT_LE_PRODUTO", X)
    If lErro <> SUCESSO Then Error 54709
    
    
    Kit_Le_Produto = X.Kit_Le_Produto(objKit, colKits)

    Exit Function
    
Erro_Kit_Le_Produto:
    
    Kit_Le_Produto = Err
    
    Select Case Err

        Case 54709
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146092)
        
    End Select
    
    Exit Function

End Function

Function EstoqueProduto_Le_Todos_Almoxarifados(objEstoqueProduto As ClassEstoqueProduto) As Long
'Lê a soma das Quantidades de todos Almoxarifados que pertence ao Produto passado em objEstoqueProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_Le_Todos_Almoxarifados
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_LE_TODOS_ALMOXARIFADO", X)
    If lErro <> SUCESSO Then Error 54710
    
    
    EstoqueProduto_Le_Todos_Almoxarifados = X.EstoqueProduto_Le_Todos_Almoxarifados(objEstoqueProduto)

    Exit Function
    
Erro_EstoqueProduto_Le_Todos_Almoxarifados:
    
    EstoqueProduto_Le_Todos_Almoxarifados = Err
    
    Select Case Err

        Case 54710
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146093)
        
    End Select
    
    Exit Function

End Function

Public Function MovEstoqueItens_Le1(objMovEstoque As ClassMovEstoque, ByVal iTipoMov As Integer) As Long
'Le os Ítens do MovimentodeEstoque relacionados ao Tipo de Movimento passado como parametro e que possuem o codigo especificado em objMovEstoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoqueItens_Le1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUEITENS_LE1", X)
    If lErro <> SUCESSO Then Error 55382
    
    
    MovEstoqueItens_Le1 = X.MovEstoqueItens_Le1(objMovEstoque, iTipoMov)

    Exit Function
    
Erro_MovEstoqueItens_Le1:
    
    MovEstoqueItens_Le1 = Err
    
    Select Case Err

        Case 55382
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146094)
        
    End Select
    
    Exit Function

End Function

Public Function MovEstoqueItem_Le_Primeiro(objItemMovEstoque As ClassItemMovEstoque) As Long
'Lê o primeiro item correspondentes a partir da FilialEmpresa, Código e retorna no colItens do objeto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoqueItem_Le_Primeiro
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUEITEM_LE_PRIMEIRO", X)
    If lErro <> SUCESSO Then Error 55406
    
    
    MovEstoqueItem_Le_Primeiro = X.MovEstoqueItem_Le_Primeiro(objItemMovEstoque)

    Exit Function
    
Erro_MovEstoqueItem_Le_Primeiro:
    
    MovEstoqueItem_Le_Primeiro = Err
    
    Select Case Err

        Case 55406
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146095)
        
    End Select
    
    Exit Function

End Function

Public Function MovEstoque_Le_NomeTela(objMovEstoque As ClassMovEstoque, sTela As String) As Long
'Esta função retorna o nome da tela
'Relacionada com o Movimento Estoque passado no parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Le_NomeTela
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_LE_NOMETELA", X)
    If lErro <> SUCESSO Then Error 55407
    
    
    MovEstoque_Le_NomeTela = X.MovEstoque_Le_NomeTela(objMovEstoque, sTela)

    Exit Function
    
Erro_MovEstoque_Le_NomeTela:
    
    MovEstoque_Le_NomeTela = Err
    
    Select Case Err

        Case 55407
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146096)
        
    End Select
    
    Exit Function

End Function

Public Function MovEstoque_Le_NomeTela_Inventario(objMovEstoque As ClassMovEstoque, sTela As String, sCodigoInv As String) As Long
'Esta função retorna o nome da tela e Codigo do Inventario
'Relacionada com o Movimento Estoque passado no parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Le_NomeTela_Inventario
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVESTOQUE_LE_NOMETELA_INVENTARIO", X)
    If lErro <> SUCESSO Then Error 55408
    
    
    MovEstoque_Le_NomeTela_Inventario = X.MovEstoque_Le_NomeTela_Inventario(objMovEstoque, sTela, sCodigoInv)

    Exit Function
    
Erro_MovEstoque_Le_NomeTela_Inventario:
    
    MovEstoque_Le_NomeTela_Inventario = Err
    
    Select Case Err

        Case 55408
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146097)
        
    End Select
    
    Exit Function

End Function

Public Function Calcula_Custo_Atual(objProduto As ClassProduto, dCusto As Double, sTipo As String) As Long
'calcula o custo atual do produto passado como parametro e devolve-o de dCusto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Calcula_Custo_Atual
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CALCULA_CUSTO_ATUAL", X)
    If lErro <> SUCESSO Then Error 55429
    
    
    Calcula_Custo_Atual = X.Calcula_Custo_Atual(objProduto, dCusto, sTipo)
    
    Exit Function
    
Erro_Calcula_Custo_Atual:

    Calcula_Custo_Atual = Err
    
    Select Case Err

        Case 55429
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146098)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueProduto_Le_Almoxarifados1(sProduto As String, colAlmoxarifados As Collection) As Long
'Lê todos os Almoxarifados do Produto para todas as filiais e seus nomes reduzidos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_Le_Almoxarifados1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEPRODUTO_LE_ALMOXARIFADOS1", X)
    If lErro <> SUCESSO Then Error 55491
    
    
    EstoqueProduto_Le_Almoxarifados1 = X.EstoqueProduto_Le_Almoxarifados1(sProduto, colAlmoxarifados)
    
    Exit Function
    
Erro_EstoqueProduto_Le_Almoxarifados1:

    EstoqueProduto_Le_Almoxarifados1 = Err
    
    Select Case Err

        Case 55491
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146099)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_ObterRelMesMedia12Meses(sProduto As String, iFilialEmpresa As Integer, dConsumoMedioQtde As Double, dVendaMediaQtde As Double, alComando() As Long) As Long
'IMPORTANTE: especifica para classe ClassRelMesEst de apoio a execucao de relatorios
'obtem consumo medio de um produto em uma filial considerando os 12 meses anteriores ao atual
'Os meses anteriores a implantacao do produto na filial sao desprezados
'ATENCAO: se o alComando() vier zerado, será aberto nesta funcao e ficará aberto.
'poderia retornar -1 p/indicar que nao houve consumo calculado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_ObterRelMesMedia12Meses
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_OBTERRELMESMEDIA12MESES", X)
    If lErro <> SUCESSO Then Error 54817
    
    
    Produto_ObterRelMesMedia12Meses = X.Produto_ObterRelMesMedia12Meses(sProduto, iFilialEmpresa, dConsumoMedioQtde, dVendaMediaQtde, alComando())
    
    Exit Function
    
Erro_Produto_ObterRelMesMedia12Meses:

    Produto_ObterRelMesMedia12Meses = Err
    
    Select Case Err

        Case 54817
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146100)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_ObterRelMesEstMedio(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dQuantidadeMedia As Double, dValorMedio As Double, alComando() As Long) As Long
'IMPORTANTE: especifica para classe ClassRelMesEst de apoio a execucao de relatorios
'Obtém a quantidade e o valores medios de um Produto num periodo para uma filial
'ATENCAO: se o alComando() vier zerado, será aberto nesta funcao e ficará aberto.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_ObterRelMesEstMedio
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_OBTERRELMESESTMEDIO", X)
    If lErro <> SUCESSO Then Error 54818
    
    
    Produto_ObterRelMesEstMedio = X.Produto_ObterRelMesEstMedio(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidadeMedia, dValorMedio, alComando())
    
    Exit Function
    
Erro_Produto_ObterRelMesEstMedio:

    Produto_ObterRelMesEstMedio = Err
    
    Select Case Err

        Case 54818
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146101)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_ObterRelMesEst(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dQuantidadeVendida As Double, dCustoTotal As Double, dQuantidadeEntrada As Double, dQuantidadeSaida As Double, dValorEntrada As Double, dValorSaida As Double, lComando As Long) As Long
'IMPORTANTE: especifica para classe ClassRelMesEst de apoio a execucao de relatorios
'Obtém a quantidade vendida, o custo total (custo do material vendido) e outros dados do Produto correspondentes a um periodo
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_ObterRelMesEst
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_OBTERRELMESEST", X)
    If lErro <> SUCESSO Then Error 54819
    
    
    Produto_ObterRelMesEst = X.Produto_ObterRelMesEst(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dQuantidadeVendida, dCustoTotal, dQuantidadeEntrada, dQuantidadeSaida, dValorEntrada, dValorSaida, lComando)
    
    Exit Function
    
Erro_Produto_ObterRelMesEst:

    Produto_ObterRelMesEst = Err
    
    Select Case Err

        Case 54819
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146102)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_ObterRelMesFat(sProduto As String, dtDataInicial As Date, dtDataFinal As Date, iFilialEmpresa As Integer, dValorFaturado As Double, dDesconto As Double, dQuantidadePedida As Double, dQuantidadeFaturada As Double, dQuantidadeDevolvida As Double, lComando As Long) As Long
'IMPORTANTE: especifica para classe ClassRelMesFat de apoio a execucao de relatorios
'Obtém as quantidades pedida e faturada, o desconto e valor total faturado por Produto entre datas
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_ObterRelMesFat
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_OBTERRELMESFAT", X)
    If lErro <> SUCESSO Then Error 54820
    
    
    Produto_ObterRelMesFat = X.Produto_ObterRelMesFat(sProduto, dtDataInicial, dtDataFinal, iFilialEmpresa, dValorFaturado, dDesconto, dQuantidadePedida, dQuantidadeFaturada, dQuantidadeDevolvida, lComando)
    
    Exit Function
    
Erro_Produto_ObterRelMesFat:

    Produto_ObterRelMesFat = Err
    
    Select Case Err

        Case 54820
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146103)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_ObterRelQtdSldIni(sProduto As String, iAlmoxarifado As Integer, dtDataInicial As Date, dtDataFinal As Date, dQuantInicial As Double, dValorInicial As Double, alComando() As Long, ByVal iApropriacaoCusto As Integer, ByVal iAnoEstApurado As Integer, ByVal iMesEstApurado As Integer) As Long
'IMPORTANTE: especifica para classe ClassRelMesEstAlm de apoio a execucao de relatorios
'Obtém a quantidade inicial e o valor inicial do Produto/Almoxarifado correspondentes até a data anterior a data inicial
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_ObterRelQtdSldIni
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_OBTERRELQTDSLDINI", X)
    If lErro <> SUCESSO Then Error 54821
    
    
    Produto_ObterRelQtdSldIni = X.Produto_ObterRelQtdSldIni(sProduto, iAlmoxarifado, dtDataInicial, dtDataFinal, dQuantInicial, dValorInicial, alComando(), iApropriacaoCusto, iAnoEstApurado, iMesEstApurado)
    
    Exit Function
    
Erro_Produto_ObterRelQtdSldIni:

    Produto_ObterRelQtdSldIni = Err
    
    Select Case Err

        Case 54821
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146104)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_ObterDataUltEmpenho(sProduto As String, dtDataUltEmpenho As Date, lComando As Long) As Long
'obtem a data do ultimo empenho de um produto, se nao houver nenhum empenho retornar DATA_NULA
'ATENCAO: se o lComando vier zerado, será aberto nesta funcao e ficará aberto.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_ObterDataUltEmpenho
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_OBTERDATAULTEMPENHO", X)
    If lErro <> SUCESSO Then Error 54822
    
    
    Produto_ObterDataUltEmpenho = X.Produto_ObterDataUltEmpenho(sProduto, dtDataUltEmpenho, lComando)
    
    Exit Function
    
Erro_Produto_ObterDataUltEmpenho:

    Produto_ObterDataUltEmpenho = Err
    
    Select Case Err

        Case 54822
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146105)
        
    End Select
    
    Exit Function
    
End Function

Function TabelaDePrecoItens_Le_DatasVigentes(objTabelaPrecoItem As ClassTabelaPrecoItem, colDatasVigentes As Collection) As Long
'Lê todas as Datas De Vigência colocando em ordem na Colecao
'Parametros: Produto, Tabela, FilialEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaDePrecoItens_Le_DatasVigentes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TABELADEPRECOITENS_LE_DATASVIGENTES", X)
    If lErro <> SUCESSO Then Error 54917
    
    
    TabelaDePrecoItens_Le_DatasVigentes = X.TabelaDePrecoItens_Le_DatasVigentes(objTabelaPrecoItem, colDatasVigentes)
    
    Exit Function
    
Erro_TabelaDePrecoItens_Le_DatasVigentes:

    TabelaDePrecoItens_Le_DatasVigentes = Err
    
    Select Case Err

        Case 54917
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146106)
        
    End Select
    
    Exit Function
    
End Function

Function TabelaPrecoItem_Le_DataVigente(objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
'Lê a tabela de preço a partir da chave fornecida em objTabelaPreco. Retorna os dados em objTabelaPreco

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPrecoItem_Le_DataVigente
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_TABELAPRECOITEM_LE_DATAVIGENTE", X)
    If lErro <> SUCESSO Then Error 54918
    
    
    TabelaPrecoItem_Le_DataVigente = X.TabelaPrecoItem_Le_DataVigente(objTabelaPrecoItem)
    
    Exit Function
    
Erro_TabelaPrecoItem_Le_DataVigente:

    TabelaPrecoItem_Le_DataVigente = Err
    
    Select Case Err

        Case 54918
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146107)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Le_SituacaoTrib(sSituacaoTrib As String, sProduto As String, iTipoTribICMS As Integer, lComando As Long) As Long
'Le a Situacao Tributaria
'Concatena a OrigemMercadoria da Tabela de Produtos com TipoTribCst da Tabela de TiposTribICMS
'IMPORTANTE : O lComando somente é aberto se vier aberto e permanecerá fechado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Le_SituacaoTrib
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_LE_SITUACAOTRIB", X)
    If lErro <> SUCESSO Then Error 54928
    
    Produto_Le_SituacaoTrib = X.Produto_Le_SituacaoTrib(sSituacaoTrib, sProduto, iTipoTribICMS, lComando)
    
    Exit Function
    
Erro_Produto_Le_SituacaoTrib:

    Produto_Le_SituacaoTrib = Err
    
    Select Case Err

        Case 54928
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146108)
        
    End Select
    
    Exit Function
    
End Function

Function SldMesEstAlm_Atualiza_Ano(objTela As Object, objSldMesEst As ClassSldMesEst, iCancela As Integer) As Long
'Atualiza Ano, abrindo um novo ano(Exercicio) para os produtos/almoxarifado da Filial em questao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEstAlm_Atualiza_Ano
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESESTALM_ATUALIZA_ANO", X)
    If lErro <> SUCESSO Then Error 60525
    
    SldMesEstAlm_Atualiza_Ano = X.SldMesEstAlm_Atualiza_Ano(objTela, objSldMesEst, iCancela)
    
    Exit Function
    
Erro_SldMesEstAlm_Atualiza_Ano:

    SldMesEstAlm_Atualiza_Ano = Err
    
    Select Case Err

        Case 60525
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146109)
        
    End Select
    
    Exit Function
    
End Function

Function SldMesEstAlm_Atualiza_Mes(objTela As Object, iMes As Integer, objSldMesEst As ClassSldMesEst, iCancela As Integer, objEstoque As ClassEstoqueMes) As Long
'Atualiza somente os mes, durante o decorrer do ano

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEstAlm_Atualiza_Mes
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESESTALM_ATUALIZA_MES", X)
    If lErro <> SUCESSO Then Error 60526
    
    SldMesEstAlm_Atualiza_Mes = X.SldMesEstAlm_Atualiza_Mes(objTela, iMes, objSldMesEst, iCancela, objEstoque)
    
    Exit Function
    
Erro_SldMesEstAlm_Atualiza_Mes:

    SldMesEstAlm_Atualiza_Mes = Err
    
    Select Case Err

        Case 60526
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146110)
        
    End Select
    
    Exit Function
    
End Function

Function Estoque_Grava_Trans(objProdutoFilial As ClassProdutoFilial, iControleEstoque As Integer) As Long
'Grava o Estoque do Produto da Filial
'ATENÇÃO: Tem que ser chamado dentro de Transação

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Estoque_Grava_Trans
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUE_GRAVA_TRANS", X)
    If lErro <> SUCESSO Then Error 59858
    
    Estoque_Grava_Trans = X.Estoque_Grava_Trans(objProdutoFilial, iControleEstoque)
    
    Exit Function
    
Erro_Estoque_Grava_Trans:

    Estoque_Grava_Trans = Err
    
    Select Case Err

        Case 59858
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146111)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Grava_Trans(objProduto As ClassProduto, colTabelaPrecoItem As Collection) As Long
'Grava ou Alterar o Produto passado
'ATENÇÃO: Tem que ser chamado dentro de Transação

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Grava_Trans
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_GRAVA_TRANS", X)
    If lErro <> SUCESSO Then Error 59859
    
    Produto_Grava_Trans = X.Produto_Grava_Trans(objProduto, colTabelaPrecoItem)
    
    Exit Function
    
Erro_Produto_Grava_Trans:

    Produto_Grava_Trans = Err
    
    Select Case Err

        Case 59859
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146112)
        
    End Select
    
    Exit Function
    
End Function

Function ProdutoFiliais_Le(sProduto As String, colProdutoFilial As Collection) As Long
'Passa o Produto e obtem uma coleção do Produto para todas as Filiais

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutoFiliais_Le
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTOFILIAIS_LE", X)
    If lErro <> SUCESSO Then Error 59860
    
    ProdutoFiliais_Le = X.ProdutoFiliais_Le(sProduto, colProdutoFilial)
    
    Exit Function
    
Erro_ProdutoFiliais_Le:

    ProdutoFiliais_Le = Err
    
    Select Case Err

        Case 59860
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146113)
        
    End Select
    
    Exit Function
    
End Function

Function Produto_Grava_Light(objProduto As ClassProduto, colTabelaPrecoItem As Collection, objProdutoFilial As ClassProdutoFilial, iControleEstoque As Integer) As Long
'Grava ou Alterar o Produto da versão light passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Produto_Grava_Light
    
    lErro = Rotina_ObterObjetoFuncao("MAT_PRODUTO_GRAVA_LIGHT", X)
    If lErro <> SUCESSO Then Error 59861
    
    Produto_Grava_Light = X.Produto_Grava_Light(objProduto, colTabelaPrecoItem, objProdutoFilial, iControleEstoque)
    
    Exit Function
    
Erro_Produto_Grava_Light:

    Produto_Grava_Light = Err
    
    Select Case Err

        Case 59861
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146114)
        
    End Select
    
    Exit Function
    
End Function

Function ClasseUM_Le_Total(iTotalClasseUM As Integer) As Long
'Le a quantidade de ClasseUM já gravados na Tabela de ClasseUM

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ClasseUM_Le_Total
    
    lErro = Rotina_ObterObjetoFuncao("MAT_CLASSEUM_LE_TOTAL", X)
    If lErro <> SUCESSO Then Error 59862
    
    ClasseUM_Le_Total = X.ClasseUM_Le_Total(iTotalClasseUM)
    
    Exit Function
    
Erro_ClasseUM_Le_Total:

    ClasseUM_Le_Total = Err
    
    Select Case Err

        Case 59862
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146115)
        
    End Select
    
    Exit Function
    
End Function

Function Almoxarifado_Le_Total(iTotalAlmox As Integer) As Long
'Le a quantidade de Almoxarifados já gravados na Tabela de Almoxarifados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Almoxarifado_Le_Total
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ALMOXARIFADO_LE_TOTAL", X)
    If lErro <> SUCESSO Then Error 59863
    
    Almoxarifado_Le_Total = X.Almoxarifado_Le_Total(iTotalAlmox)
    
    Exit Function
    
Erro_Almoxarifado_Le_Total:

    Almoxarifado_Le_Total = Err
    
    Select Case Err

        Case 59863
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146116)
        
    End Select
    
    Exit Function
    
End Function

Function FornecedorProduto_Le_Total(objFornecedorProduto As ClassFornecedorProduto, iTotalFornProduto As Integer) As Long
'Le a quantidade de Fornecedores já gravados na Tabela de FornecedorProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FornecedorProduto_Le_Total
    
    lErro = Rotina_ObterObjetoFuncao("MAT_FORNECEDORPRODUTO_LE_TOTAL", X)
    If lErro <> SUCESSO Then Error 59864
    
    FornecedorProduto_Le_Total = X.FornecedorProduto_Le_Total(objFornecedorProduto, iTotalFornProduto)
    
    Exit Function
    
Erro_FornecedorProduto_Le_Total:

    FornecedorProduto_Le_Total = Err
    
    Select Case Err

        Case 59864
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146117)
        
    End Select
    
    Exit Function
    
End Function

Function EstoqueMes_Le_Mes1(objEstoqueMes As ClassEstoqueMes) As Long
'Le o Ano/Mes que está aberto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueMes_Le_Mes1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTOQUEMES_LE_MES1", X)
    If lErro <> SUCESSO Then Error 60862
    
    
    EstoqueMes_Le_Mes1 = X.EstoqueMes_Le_Mes1(objEstoqueMes)

    Exit Function
    
Erro_EstoqueMes_Le_Mes1:
    
    EstoqueMes_Le_Mes1 = Err
    
    Select Case Err

        Case 60862
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146118)
        
    End Select
    
    Exit Function

End Function

Function Rotina_FechamentoMes(objEstoque As ClassEstoqueMes, iCancela As Integer, objTela As Object) As Long
'Rotina que Chamara as funcoes do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_FechamentoMes
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ROTINA_FECHAMENTOMES", X)
    If lErro <> SUCESSO Then Error 60862
    
    
    Rotina_FechamentoMes = X.Rotina_FechamentoMes(objEstoque, iCancela, objTela)

    Exit Function
    
Erro_Rotina_FechamentoMes:
    
    Rotina_FechamentoMes = Err
    
    Select Case Err

        Case 60862
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146119)
        
    End Select
    
    Exit Function

End Function


Function EstoqueProduto_Le_Todos_Almoxarifados_Filial(objEstoqueProduto As ClassEstoqueProduto) As Long
'Lê a soma das Quantidades de todos Almoxarifados da Filial que pertence ao Produto passado em objEstoqueProduto

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EstoqueProduto_Le_Todos_Almoxarifados_Filial
    
    lErro = Rotina_ObterObjetoFuncao("MAT_ESTPROD_LE_TODOS_ALMOXARIFADO_FILIAL", X)
    If lErro <> SUCESSO Then Error 65087
    
    EstoqueProduto_Le_Todos_Almoxarifados_Filial = X.EstoqueProduto_Le_Todos_Almoxarifados_Filial(objEstoqueProduto)

    Exit Function
    
Erro_EstoqueProduto_Le_Todos_Almoxarifados_Filial:
    
    EstoqueProduto_Le_Todos_Almoxarifados_Filial = Err
    
    Select Case Err

        Case 65087
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146120)
        
    End Select
    
    Exit Function

End Function

Function MovimentoEstoque_Le_Ccl(sCcl As String) As Long
'Verifica se existe algum centro de custo com o centro de custo passado como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovimentoEstoque_Le_Ccl
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("MAT_MOVIMENTOESTOQUE_LE_CCL", X)
    If lErro <> SUCESSO Then Error 60869
    
    
    MovimentoEstoque_Le_Ccl = X.MovimentoEstoque_Le_Ccl(sCcl)

    Exit Function
    
Erro_MovimentoEstoque_Le_Ccl:
    
    MovimentoEstoque_Le_Ccl = Err
    
    Select Case Err

        Case 60869
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146121)
        
    End Select
    
    Exit Function

End Function

Function UMEstoque_Conversao(sProduto As String, sUM As String, dQuantidade As Double, dQuantidadeConvertida As Double) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UMEstoque_Conversao
    
    lErro = Rotina_ObterObjetoFuncao("MAT_UMESTOQUE_CONVERSAO", X)
    If lErro <> SUCESSO Then Error 65270
    
    UMEstoque_Conversao = X.UMEstoque_Conversao(sProduto, sUM, dQuantidade, dQuantidadeConvertida)

    Exit Function
    
Erro_UMEstoque_Conversao:
    
    UMEstoque_Conversao = Err
    
    Select Case Err

        Case 65270
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 146122)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_CustoMedioConsig3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst1_Le_CustoMedioConsig3
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST1_LE_CUSTOMEDIOCONSIG3", X)
    If lErro <> SUCESSO Then gError 71236

    SldMesEst1_Le_CustoMedioConsig3 = X.SldMesEst1_Le_CustoMedioConsig3(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)

    Exit Function
    
Erro_SldMesEst1_Le_CustoMedioConsig3:
    
    SldMesEst1_Le_CustoMedioConsig3 = gErr
    
    Select Case gErr

        Case 71236
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146123)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_CustoMedioConserto(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'calcula o custo médio de material em conserto nosso em poder de terceiros.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst2_Le_CustoMedioConserto
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_CUSTOMEDIOCONSERTO", X)
    If lErro <> SUCESSO Then gError 71237

    SldMesEst2_Le_CustoMedioConserto = X.SldMesEst2_Le_CustoMedioConserto(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)

    Exit Function
    
Erro_SldMesEst2_Le_CustoMedioConserto:
    
    SldMesEst2_Le_CustoMedioConserto = gErr
    
    Select Case gErr

        Case 71237
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146124)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_CustoMedioDemo(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'calcula o custo médio de material em Demonstração nosso em poder de terceiros.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst2_Le_CustoMedioDemo
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_CUSTOMEDIODEMO", X)
    If lErro <> SUCESSO Then gError 71238

    SldMesEst2_Le_CustoMedioDemo = X.SldMesEst2_Le_CustoMedioDemo(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)

    Exit Function
    
Erro_SldMesEst2_Le_CustoMedioDemo:
    
    SldMesEst2_Le_CustoMedioDemo = gErr
    
    Select Case gErr

        Case 71238
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146125)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_CustoMedioOutros(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'calcula o custo médio de Outros materiais nossos em poder de terceiros.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst2_Le_CustoMedioOutros
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_CUSTOMEDIOOUTROS", X)
    If lErro <> SUCESSO Then gError 71239

    SldMesEst2_Le_CustoMedioOutros = X.SldMesEst2_Le_CustoMedioOutros(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)

    Exit Function
    
Erro_SldMesEst2_Le_CustoMedioOutros:
    
    SldMesEst2_Le_CustoMedioOutros = gErr
    
    Select Case gErr

        Case 71239
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146126)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_CustoMedioBenef(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'calcula o custo médio de materiais em beneficiamento nossos em poder de terceiros.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst2_Le_CustoMedioBenef
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_CUSTOMEDIOBENEF", X)
    If lErro <> SUCESSO Then gError 71240

    SldMesEst2_Le_CustoMedioBenef = X.SldMesEst2_Le_CustoMedioBenef(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)

    Exit Function
    
Erro_SldMesEst2_Le_CustoMedioBenef:
    
    SldMesEst2_Le_CustoMedioBenef = gErr
    
    Select Case gErr

        Case 71240
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146127)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_CustoMedioConsig(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long
'calcula o custo médio de materiais em consignação nossos em poder de terceiros.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst2_Le_CustoMedioConsig
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_CUSTOMEDIOCONSIG", X)
    If lErro <> SUCESSO Then gError 71241

    SldMesEst2_Le_CustoMedioConsig = X.SldMesEst2_Le_CustoMedioConsig(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)

    Exit Function
    
Erro_SldMesEst2_Le_CustoMedioConsig:
    
    SldMesEst2_Le_CustoMedioConsig = gErr
    
    Select Case gErr

        Case 71241
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146128)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_CustoMedioConserto3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst1_Le_CustoMedioConserto3
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST1_LE_CUSTOMEDIOCONSERTO3", X)
    If lErro <> SUCESSO Then gError 71242

    SldMesEst1_Le_CustoMedioConserto3 = X.SldMesEst1_Le_CustoMedioConserto3(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)

    Exit Function
    
Erro_SldMesEst1_Le_CustoMedioConserto3:
    
    SldMesEst1_Le_CustoMedioConserto3 = gErr
    
    Select Case gErr

        Case 71242
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146129)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_CustoMedioDemo3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst1_Le_CustoMedioDemo3
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST1_LE_CUSTOMEDIODEMO3", X)
    If lErro <> SUCESSO Then gError 71243

    SldMesEst1_Le_CustoMedioDemo3 = X.SldMesEst1_Le_CustoMedioDemo3(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)

    Exit Function
    
Erro_SldMesEst1_Le_CustoMedioDemo3:
    
    SldMesEst1_Le_CustoMedioDemo3 = gErr
    
    Select Case gErr

        Case 71243
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146130)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_CustoMedioOutros3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst1_Le_CustoMedioOutros3
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST1_LE_CUSTOMEDIOOUTROS3", X)
    If lErro <> SUCESSO Then gError 71244

    SldMesEst1_Le_CustoMedioOutros3 = X.SldMesEst1_Le_CustoMedioOutros3(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)

    Exit Function
    
Erro_SldMesEst1_Le_CustoMedioOutros3:
    
    SldMesEst1_Le_CustoMedioOutros3 = gErr
    
    Select Case gErr

        Case 71244
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146131)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Le_CustoMedioBenef3(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dCustoMedio As Double) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst1_Le_CustoMedioBenef3
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST1_LE_CUSTOMEDIOBENEF3", X)
    If lErro <> SUCESSO Then gError 71245

    SldMesEst1_Le_CustoMedioBenef3 = X.SldMesEst1_Le_CustoMedioBenef3(lComando, iFilialEmpresa, iAno, sProduto, iMes, dCustoMedio)

    Exit Function
    
Erro_SldMesEst1_Le_CustoMedioBenef3:
    
    SldMesEst1_Le_CustoMedioBenef3 = gErr
    
    Select Case gErr

        Case 71245
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146132)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst1_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, objSldDiaEst As ClassSldDiaEst, objTipoMovEstoque As ClassTipoMovEst) As Long
'Atualiza a tabela de SldMesEst1 utilizando os valores passados atraves de objSldDiaEst.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst1_Atualiza
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST1_ATUALIZA", X)
    If lErro <> SUCESSO Then gError 71246

    SldMesEst1_Atualiza = X.SldMesEst1_Atualiza(lComando1, lComando2, objSldDiaEst, objTipoMovEstoque)

    Exit Function
    
Erro_SldMesEst1_Atualiza:
    
    SldMesEst1_Atualiza = gErr
    
    Select Case gErr

        Case 71246
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146133)
        
    End Select
    
    Exit Function

End Function

Function SldMesEstAlm1_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, objSldDiaEst As ClassSldDiaEst, ByVal iAlmoxarifado As Integer, objTipoMovEstoque As ClassTipoMovEst) As Long
'Atualiza a tabela de SldMesEstAlm1 utilizando os valores passados atraves de objSldDiaEst.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEstAlm1_Atualiza
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESESTALM1_ATUALIZA", X)
    If lErro <> SUCESSO Then gError 71247

    SldMesEstAlm1_Atualiza = X.SldMesEstAlm1_Atualiza(lComando1, lComando2, objSldDiaEst, iAlmoxarifado, objTipoMovEstoque)

    Exit Function
    
Erro_SldMesEstAlm1_Atualiza:
    
    SldMesEstAlm1_Atualiza = gErr
    
    Select Case gErr

        Case 71247
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146134)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, objSldDiaEst As ClassSldDiaEst, objTipoMovEstoque As ClassTipoMovEst) As Long
'Atualiza a tabela de SldMesEst1 utilizando os valores passados atraves de objSldDiaEst.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst2_Atualiza
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_ATUALIZA", X)
    If lErro <> SUCESSO Then gError 71248

    SldMesEst2_Atualiza = X.SldMesEst2_Atualiza(lComando1, lComando2, objSldDiaEst, objTipoMovEstoque)

    Exit Function
    
Erro_SldMesEst2_Atualiza:
    
    SldMesEst2_Atualiza = gErr
    
    Select Case gErr

        Case 71248
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146135)
        
    End Select
    
    Exit Function

End Function

Function SldMesEstAlm2_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, objSldDiaEst As ClassSldDiaEst, ByVal iAlmoxarifado As Integer, objTipoMovEstoque As ClassTipoMovEst) As Long
'Atualiza a tabela de SldMesEstAlm1 utilizando os valores passados atraves de objSldDiaEst.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEstAlm2_Atualiza
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESESTALM2_ATUALIZA", X)
    If lErro <> SUCESSO Then gError 71249

    SldMesEstAlm2_Atualiza = X.SldMesEstAlm2_Atualiza(lComando1, lComando2, objSldDiaEst, iAlmoxarifado, objTipoMovEstoque)

    Exit Function
    
Erro_SldMesEstAlm2_Atualiza:
    
    SldMesEstAlm2_Atualiza = gErr
    
    Select Case gErr

        Case 71249
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146136)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantConserto(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dQuantConsertoAtual As Double) As Long
'calcula a quantidade de material em conserto nosso em poder de terceiros.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst2_Le_QuantConserto
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_QUANTCONSERTO", X)
    If lErro <> SUCESSO Then gError 71251

    SldMesEst2_Le_QuantConserto = X.SldMesEst2_Le_QuantConserto(lComando, iFilialEmpresa, iAno, sProduto, iMes, dQuantConsertoAtual)

    Exit Function
    
Erro_SldMesEst2_Le_QuantConserto:
    
    SldMesEst2_Le_QuantConserto = gErr
    
    Select Case gErr

        Case 71251
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146137)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantDemo(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dQuantDemoAtual As Double) As Long
'calcula a quantidade de material em Demonstração nosso em poder de terceiros.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst2_Le_QuantDemo
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_QUANTDEMO", X)
    If lErro <> SUCESSO Then gError 71252

    SldMesEst2_Le_QuantDemo = X.SldMesEst2_Le_QuantDemo(lComando, iFilialEmpresa, iAno, sProduto, iMes, dQuantDemoAtual)

    Exit Function
    
Erro_SldMesEst2_Le_QuantDemo:
    
    SldMesEst2_Le_QuantDemo = gErr
    
    Select Case gErr

        Case 71252
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146138)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantOutros(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dQuantOutrosAtual As Double) As Long
'calcula a quantidade de Outros materiais nosso em poder de terceiros.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst2_Le_QuantOutros
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_QUANTOUTROS", X)
    If lErro <> SUCESSO Then gError 71253

    SldMesEst2_Le_QuantOutros = X.SldMesEst2_Le_QuantOutros(lComando, iFilialEmpresa, iAno, sProduto, iMes, dQuantOutrosAtual)

    Exit Function
    
Erro_SldMesEst2_Le_QuantOutros:
    
    SldMesEst2_Le_QuantOutros = gErr
    
    Select Case gErr

        Case 71253
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146139)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantBenef(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dQuantBenefAtual As Double) As Long
'calcula a quantidade de material em Beneficiamento nosso em poder de terceiros.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst2_Le_QuantBenef
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_QUANTBENEF", X)
    If lErro <> SUCESSO Then gError 71254

    SldMesEst2_Le_QuantBenef = X.SldMesEst2_Le_QuantBenef(lComando, iFilialEmpresa, iAno, sProduto, iMes, dQuantBenefAtual)

    Exit Function
    
Erro_SldMesEst2_Le_QuantBenef:
    
    SldMesEst2_Le_QuantBenef = gErr
    
    Select Case gErr

        Case 71254
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146140)
        
    End Select
    
    Exit Function

End Function

Function SldMesEst2_Le_QuantConsig(ByVal lComando As Long, ByVal iFilialEmpresa As Integer, ByVal iAno As Integer, sProduto As String, ByVal iMes As Integer, dQuantConsigAtual As Double) As Long
'calcula a quantidade de material em Consignação nosso em poder de terceiros.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesEst2_Le_QuantConsig
    
    lErro = Rotina_ObterObjetoFuncao("MAT_SLDMESEST2_LE_QUANTCONSIG", X)
    If lErro <> SUCESSO Then gError 71255

    SldMesEst2_Le_QuantConsig = X.SldMesEst2_Le_QuantConsig(lComando, iFilialEmpresa, iAno, sProduto, iMes, dQuantConsigAtual)

    Exit Function
    
Erro_SldMesEst2_Le_QuantConsig:
    
    SldMesEst2_Le_QuantConsig = gErr
    
    Select Case gErr

        Case 71255
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 146141)
        
    End Select
    
    Exit Function

End Function

