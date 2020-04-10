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
Dim sCategoria As String
Dim sProduto1 As String
Dim sItem As String
Dim sProdutoPai As String
Dim objProduto1 As New ClassProduto

On Error GoTo Erro_Produto_Le_Filhos_Grade

    sCategoria = CATEGORIAPRODUTO_GENERICO

    sProduto1 = objProduto.sCodigo

    'retorna o item do produto/categoria
    lErro = CF("ProdutoCategoria_Le", sProduto1, sCategoria, sItem)
    If lErro <> SUCESSO And lErro <> 92405 Then gError 117692

    'se o produto tem categoria GENERICO e item SIM ==> em sProduto vai estar o produto que foi feita a op
    If lErro = SUCESSO And UCase(sItem) = "SIM" Then
        
        lErro = CF("Produto_Le_ProdutoPai", sProduto1, objProduto1)
        If lErro <> SUCESSO Then gError 117695
        
    Else
    
        Set objProduto1 = objProduto
    
    End If


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
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Produtos WHERE Codigo > ? AND Nivel = ? ORDER BY Codigo", sProduto, objProduto1.sCodigo, objProduto1.iNivel)
    If lErro <> AD_SQL_SUCESSO Then gError 86300

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86301

    'Se não existe nenhum produto do mesmo nível
    If lErro <> AD_SQL_SUCESSO Then
        'BUsca todos os Filhos desse produto nessa filial que possuem os itens para a grade informada
        lErro = Comando_Executar(lComando, "SELECT ProdutoCategoria.Produto, Produtos.Descricao, ProdutoCategoria.Categoria, ProdutoCategoria.Item, Produtos.SiglaUMEstoque FROM ProdutosFilial, Produtos,ProdutoCategoria,GradeCategoria, CategoriaProdutoItem WHERE ProdutoCategoria.Produto IN (SELECT Produto FROM ProdutoCategoria WHERE Categoria IN (SELECT Categoria FROM GradeCategoria WHERE Grade = ? ) GROUP BY Produto HAVING COUNT(*) = (SELECT COUNT(*) FROM GradeCategoria WHERE Grade = ?) ) " _
                                 & " AND ProdutosFilial.Produto = Produtos.Codigo AND GradeCategoria.Grade = ? AND GradeCategoria.Categoria = ProdutoCategoria.Categoria AND Nivel = ? AND ProdutoCategoria.Produto > ? AND ProdutoCategoria.Produto = Produtos.Codigo AND ProdutosFilial.FilialEmpresa = ? AND Produtos.ClasseUM = ? AND CategoriaProdutoItem.Categoria = ProdutoCategoria.Categoria AND CategoriaProdutoItem.Item = ProdutoCategoria.Item ORDER BY Posicao, CategoriaProdutoItem.Ordem", tItemRomaneioGrade.sProduto, tItemRomaneioGrade.sDescricao, tCategProdItem.sCategoria, tCategProdItem.sItem, tItemRomaneioGrade.sSiglaUMEstoque, objProduto1.sGrade, objProduto1.sGrade, objProduto1.sGrade, objProduto1.iNivel + 1, objProduto1.sCodigo, giFilialEmpresa, objProduto1.iClasseUM)

    Else
        'BUsca todos os Filhos desse produto nessa filial que possuem os itens para a grade informada
        lErro = Comando_Executar(lComando, "SELECT ProdutoCategoria.Produto, Produtos.Descricao, ProdutoCategoria.Categoria, ProdutoCategoria.Item, Produtos.SiglaUMEstoque FROM ProdutosFilial, Produtos,ProdutoCategoria,GradeCategoria, CategoriaProdutoItem WHERE ProdutoCategoria.Produto IN (SELECT Produto FROM ProdutoCategoria WHERE Categoria IN (SELECT Categoria FROM GradeCategoria WHERE Grade = ? ) GROUP BY Produto HAVING COUNT(*) = (SELECT COUNT(*) FROM GradeCategoria WHERE Grade = ?) ) " _
                                 & " AND ProdutosFilial.Produto = Produtos.Codigo AND GradeCategoria.Grade = ? AND GradeCategoria.Categoria = ProdutoCategoria.Categoria AND Nivel = ? AND ProdutoCategoria.Produto BETWEEN ? AND ? AND ProdutoCategoria.Produto = Produtos.Codigo AND ProdutosFilial.FilialEmpresa = ? AND Produtos.ClasseUM = ? AND CategoriaProdutoItem.Categoria = ProdutoCategoria.Categoria AND CategoriaProdutoItem.Item = ProdutoCategoria.Item ORDER BY Posicao, CategoriaProdutoItem.Ordem", tItemRomaneioGrade.sProduto, tItemRomaneioGrade.sDescricao, tCategProdItem.sCategoria, tCategProdItem.sItem, tItemRomaneioGrade.sSiglaUMEstoque, objProduto1.sGrade, objProduto1.sGrade, objProduto1.sGrade, objProduto1.iNivel + 1, objProduto1.sCodigo, sProduto, giFilialEmpresa, objProduto1.iClasseUM)

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
        
        Case 86304, 117695, 117696
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179217)
    
    End Select
    
    Call Comando_Fechar(lComando)

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

