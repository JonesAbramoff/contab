Attribute VB_Name = "FriendMAT"
Option Explicit

Public Function TipoDeProduto_Grava_Tabelas_Lock(objTipoDeProduto As ClassTipoDeProduto) As Long
'Faz o lock das tabelas relacionadas a tabela TipoDeProduto para assegurar a gravação de um novo Tipo de Produto
'Chamada DENTRO DE TRANSAÇÃO por TipoDeProduto_Grava

Dim objTipoDeProdutoCategoria As ClassTipoDeProdutoCategoria
Dim sDescricao As String
Dim lComando(2) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim iClasse As Integer

On Error GoTo Erro_TipoDeProduto_Grava_Tabelas_Lock

    For iIndice = 0 To 1
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 31270
    Next

    'Tabela de CategoriaProdutoItem
    For Each objTipoDeProdutoCategoria In objTipoDeProduto.colCategoriaItem

        sDescricao = String(STRING_CATEGORIAPRODUTO_DESCRICAO, 0)

        lErro = Comando_ExecutarLockado(lComando(0), "SELECT Descricao FROM CategoriaProdutoItem WHERE Categoria=? AND Item=?", sDescricao, objTipoDeProdutoCategoria.sCategoria, objTipoDeProdutoCategoria.sItem)
        If lErro <> AD_SQL_SUCESSO Then Error 22593

        lErro = Comando_BuscarPrimeiro(lComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22597
        If lErro = AD_SQL_SEM_DADOS Then Error 22591

        lErro = Comando_LockShared(lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 22594

    Next

    If (Len(objTipoDeProduto.sSiglaUMCompra) > 0 Or Len(objTipoDeProduto.sSiglaUMVenda) > 0 Or Len(objTipoDeProduto.sSiglaUMEstoque) > 0) Then
    'Tabela de Unidades de Medida

        lErro = Comando_ExecutarLockado(lComando(1), "SELECT Classe FROM UnidadesDeMedida WHERE Classe = ? AND (Sigla = ? OR Sigla = ? OR Sigla = ?)", iClasse, objTipoDeProduto.iClasseUM, objTipoDeProduto.sSiglaUMCompra, objTipoDeProduto.sSiglaUMVenda, objTipoDeProduto.sSiglaUMEstoque)
        If lErro <> AD_SQL_SUCESSO Then Error 22595

        lErro = Comando_BuscarPrimeiro(lComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22598
        If lErro = AD_SQL_SEM_DADOS Then Error 22592

        Do While lErro <> AD_SQL_SEM_DADOS

            lErro = Comando_LockShared(lComando(1))
            If lErro <> AD_SQL_SUCESSO Then Error 22596

            lErro = Comando_BuscarProximo(lComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22599

        Loop

    End If

    For iIndice = 0 To 1
        Call Comando_Fechar(lComando(iIndice))
    Next

    TipoDeProduto_Grava_Tabelas_Lock = SUCESSO

    Exit Function

Erro_TipoDeProduto_Grava_Tabelas_Lock:

    TipoDeProduto_Grava_Tabelas_Lock = Err

    Select Case Err

        Case 31270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 22591
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_CADASTRADA", Err, objTipoDeProdutoCategoria.sItem, objTipoDeProdutoCategoria.sCategoria)

        Case 22592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNIDADES_MEDIDAS_NAO_CADASTRADAS", Err)

        Case 22593, 22597
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTOITEM2", Err, objTipoDeProdutoCategoria.sCategoria, objTipoDeProdutoCategoria.sItem)

        Case 22594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIAPRODUTOITEM", Err)

        Case 22595, 22598, 22599
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_UNIDADESDEMEDIDA1", Err, objTipoDeProduto.iClasseUM)

        Case 22596
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_UNIDADESDEMEDIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160708)

    End Select

    For iIndice = 0 To 2
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function


Public Function TipoDeProduto_Grava_CategoriaNaColecao(colCategoriaItemCopia As Collection, sCategoria As String) As Long
'Retorna o Indice da Categoria na coleção ou Zero se não acha-la (na coleção)

Dim iIndice As Integer
Dim objTipoDeProdutoCategoria As New ClassTipoDeProdutoCategoria

    'Pesquisa a Categoria na coleção
    For iIndice = 1 To colCategoriaItemCopia.Count

        Set objTipoDeProdutoCategoria = colCategoriaItemCopia.Item(iIndice)

        'Se achou a Sigla na coleção
        If objTipoDeProdutoCategoria.sCategoria = sCategoria Then

            TipoDeProduto_Grava_CategoriaNaColecao = iIndice
            Exit Function

        End If

    Next

    TipoDeProduto_Grava_CategoriaNaColecao = 0

End Function

Public Function TipoDeProduto_Grava_NovasCategorias(objTipoDeProduto As ClassTipoDeProduto, colCategoriaItem As Collection) As Long
'Percorre as Categorias na coleção incluindo-as no BD

Dim lErro As Long
Dim lComando As Long
Dim objTipoDeProdutoCategoria As New ClassTipoDeProdutoCategoria

On Error GoTo Erro_TipoDeProduto_Grava_NovasCategorias

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 22588

    'Percorre a coleção
    For Each objTipoDeProdutoCategoria In colCategoriaItem

        lErro = Comando_Executar(lComando, "INSERT INTO TipoDeProdutoCategorias (TipoDeProduto, Categoria, Item) VALUES(?,?,?)", objTipoDeProdutoCategoria.iTipoDeProduto, objTipoDeProdutoCategoria.sCategoria, objTipoDeProdutoCategoria.sItem)
        If lErro <> AD_SQL_SUCESSO Then Error 22589

    Next

    Call Comando_Fechar(lComando)

    TipoDeProduto_Grava_NovasCategorias = SUCESSO

    Exit Function

Erro_TipoDeProduto_Grava_NovasCategorias:

    TipoDeProduto_Grava_NovasCategorias = Err

    Select Case Err

        Case 22588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOPRODUTOCATEGORIA", Err, objTipoDeProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 160709)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

