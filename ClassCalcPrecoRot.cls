VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCalcPrecoRot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Rotina_PrecosDeVenda_Calcula(ByVal sNomeArqParam As String, ByVal objCalcPrecoVenda As ClassCalcPrecoVenda) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_PrecosDeVenda_Calcula

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 106642

    lErro = PrecosDeVenda_Calcula(objCalcPrecoVenda)
    If lErro <> SUCESSO Then gError 106643

    Rotina_PrecosDeVenda_Calcula = SUCESSO

    Exit Function

Erro_Rotina_PrecosDeVenda_Calcula:

    Rotina_PrecosDeVenda_Calcula = gErr

    Select Case gErr

        Case 106642, 106643

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144663)

    End Select

    Exit Function

End Function

Private Function PrecosDeVenda_Calcula(ByVal objCalcPrecoVenda As ClassCalcPrecoVenda) As Long

Dim lErro As Long, iIndice As Integer
Dim lTransacao As Long, alComando(1 To 11) As Long

On Error GoTo Erro_PrecosDeVenda_Calcula

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 106761

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If lErro <> SUCESSO Then gError 106762
    Next

    lErro = PrecosDeVenda_Calcula1(objCalcPrecoVenda, alComando)
    If lErro <> SUCESSO Then gError 106762

    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 106763

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    PrecosDeVenda_Calcula = SUCESSO

    Exit Function

Erro_PrecosDeVenda_Calcula:

    PrecosDeVenda_Calcula = gErr

    Select Case gErr

        Case 106761
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 106762
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 106763
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 106771 To 106774

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144664)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function PrecosDeVenda_Calcula1(ByVal objCalcPrecoVenda As ClassCalcPrecoVenda, alComando() As Long) As Long

Dim lErro As Long, iIndice As Integer
Dim objFormacaoPreco As New ClassFormacaoPreco
Dim objFormacaoPrecoTemp As ClassFormacaoPreco
Dim colFormacaoPreco As Collection, sProduto As String
Dim dValor As Double, sProdInicial As String, sProdFinal As String
Dim colTabelas As New Collection
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim objContexto As New ClassContextoPlan

On Error GoTo Erro_PrecosDeVenda_Calcula1

    lErro = PrecosDeVenda_Calcula3(objCalcPrecoVenda, objContexto, alComando)
    If lErro <> SUCESSO Then gError 106825

    sProdInicial = objCalcPrecoVenda.sProdutoDe
    sProdFinal = objCalcPrecoVenda.sProdutoAte

    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELAPRECO_DESCRICAO, colCodigoNome)
    If lErro <> SUCESSO Then gError 106811

    'Preenche a ListBox ListaTabelas com os objetos da coleção
    For Each objCodigoNome In colCodigoNome

        colTabelas.Add objCodigoNome.iCodigo

    Next
    
    objFormacaoPreco.iFilialEmpresa = objCalcPrecoVenda.iFilialEmpresa
    
    For iIndice = 1 To colTabelas.Count
        
        sProduto = String(STRING_PRODUTO, 0)

        If sProdInicial <> "" And sProdFinal <> "" Then

            'Pesquisa a tabela de preço os produtos que estão entre o produto inicial e o produto final
            lErro = Comando_Executar(alComando(1), "SELECT DISTINCT TabelasDePrecoItens.CodProduto FROM TabelasDePrecoItens, Produtos WHERE TabelasDePrecoItens.CodProduto = Produtos.Codigo AND Produtos.Ativo = 0 AND CodTabela = ? AND FilialEmpresa = ? AND CodProduto >= ? AND CodProduto <= ?", sProduto, colTabelas.Item(iIndice), objCalcPrecoVenda.iFilialEmpresa, sProdInicial, sProdFinal)
            If lErro <> AD_SQL_SUCESSO Then gError 106813

        ElseIf sProdInicial <> "" And sProdFinal = "" Then

            'Pesquisa a tabela de preço os produtos maiores ou iguais ao produto inicial
            lErro = Comando_Executar(alComando(1), "SELECT DISTINCT TabelasDePrecoItens.CodProduto FROM TabelasDePrecoItens, Produtos WHERE TabelasDePrecoItens.CodProduto = Produtos.Codigo AND Produtos.Ativo = 0 AND CodTabela = ? AND FilialEmpresa = ? AND CodProduto >= ?", sProduto, colTabelas.Item(iIndice), objCalcPrecoVenda.iFilialEmpresa, sProdInicial)
            If lErro <> AD_SQL_SUCESSO Then gError 106814

        ElseIf sProdInicial = "" And sProdFinal <> "" Then

            'Pesquisa a tabela de preço os produtos menores ou iguais ao produto final
            lErro = Comando_Executar(alComando(1), "SELECT DISTINCT TabelasDePrecoItens.CodProduto FROM TabelasDePrecoItens, Produtos WHERE TabelasDePrecoItens.CodProduto = Produtos.Codigo AND Produtos.Ativo = 0 AND CodTabela = ? AND FilialEmpresa = ? AND CodProduto <= ?", sProduto, colTabelas.Item(iIndice), objCalcPrecoVenda.iFilialEmpresa, sProdFinal)
            If lErro <> AD_SQL_SUCESSO Then gError 106815

        Else

            'Pesquisa todos os produtos da tabela de preço
            lErro = Comando_Executar(alComando(1), "SELECT DISTINCT TabelasDePrecoItens.CodProduto FROM TabelasDePrecoItens, Produtos WHERE TabelasDePrecoItens.CodProduto = Produtos.Codigo AND Produtos.Ativo = 0 AND CodTabela = ? AND FilialEmpresa = ?", sProduto, colTabelas.Item(iIndice), objCalcPrecoVenda.iFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 106816

        End If

        'Tenta ler os produtos da tabela de preco
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106817

        Do While lErro <> AD_SQL_SEM_DADOS

            'carregar as expressoes (planilha)
            objFormacaoPreco.iTabelaPreco = colTabelas.Item(iIndice)
            objFormacaoPreco.sProduto = sProduto
            objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO
    
            Set colFormacaoPreco = New Collection
    
            'Lê a planilha de Formacao de Preço da Tabela/Produto
            lErro = CF("FormacaoPreco_Le1", objFormacaoPreco, colFormacaoPreco)
            If lErro <> SUCESSO And lErro <> 92434 And lErro <> 92432 Then gError 106812
    
            For Each objFormacaoPrecoTemp In colFormacaoPreco
            
                With objFormacaoPrecoTemp
                    .iFilialEmpresa = objCalcPrecoVenda.iFilialEmpresa
                    .iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO
                    .iTabelaPreco = colTabelas.Item(iIndice)
                    .sProduto = sProduto
                End With
            
            Next
            
            objContexto.sProduto = sProduto
            objContexto.iTabelaPreco = colTabelas.Item(iIndice)
            objContexto.iRotinaOrigem = FORMACAO_PRECO_ROTINA_CALCPRECO
        
            'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colFormacaoPreco.Item(?).dValor
            lErro = CF("Avalia_Expressao_FPreco1", colFormacaoPreco, dValor, sProduto, objContexto)
            If lErro <> SUCESSO Then gError 106818

            dValor = dValor '??? antes aqui dividia por 1000, a qtde base p/calculo de custo de mp e emb
        
            'atualizar preco calculado
            lErro = PrecosDeVenda_Calcula2(objCalcPrecoVenda, sProduto, objContexto.iTabelaPreco, dValor, alComando)
            If lErro <> SUCESSO Then gError 106819

            'salvar memoria de calculo
            lErro = PrecosDeVenda_Calcula4(objCalcPrecoVenda, sProduto, objContexto.iTabelaPreco, colFormacaoPreco, dValor, alComando)
            If lErro <> SUCESSO Then gError 106936
            
            'Lê o próximo Produto na tabela
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106820

        Loop

    Next

    PrecosDeVenda_Calcula1 = SUCESSO

    Exit Function

Erro_PrecosDeVenda_Calcula1:

    PrecosDeVenda_Calcula1 = gErr

    Select Case gErr

        Case 106813 To 106817, 106820
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS1", gErr, colTabelas.Item(iIndice))

        Case 106811, 106812, 106818, 106819, 106825, 106936

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144665)

    End Select

    Exit Function

End Function

Private Function PrecosDeVenda_Calcula2(ByVal objCalcPrecoVenda As ClassCalcPrecoVenda, ByVal sProduto As String, ByVal iTabela As Integer, ByVal dPreco As Double, alComando() As Long) As Long
'atualiza o preco calculado na tabela PrecoCalculado

Dim lErro As Long, dPrecoCalculado As Double

On Error GoTo Erro_PrecosDeVenda_Calcula2

    lErro = Comando_ExecutarPos(alComando(2), "SELECT PrecoCalculado FROM PrecoCalculado WHERE FilialEmpresa = ? AND CodProduto = ? AND CodTabela = ? AND DataReferencia = ?", 0, dPrecoCalculado, objCalcPrecoVenda.iFilialEmpresa, sProduto, iTabela, objCalcPrecoVenda.dtDataReferencia)
    If lErro <> AD_SQL_SUCESSO Then gError 106821
    
    lErro = Comando_BuscarProximo(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106822
    
    If lErro = AD_SQL_SUCESSO Then
    
        If Abs(dPreco - dPrecoCalculado) > DELTA_VALORMONETARIO2 Then
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE PrecoCalculado SET PrecoCalculado = ?", alComando(2), dPreco)
            If lErro <> AD_SQL_SUCESSO Then gError 106823
        End If
        
    Else
    
        lErro = Comando_Executar(alComando(4), "INSERT INTO PrecoCalculado (FilialEmpresa,CodProduto,CodTabela,DataReferencia,PrecoCalculado) VALUES (?,?,?,?,?)", objCalcPrecoVenda.iFilialEmpresa, sProduto, iTabela, objCalcPrecoVenda.dtDataReferencia, dPreco)
        If lErro <> AD_SQL_SUCESSO Then gError 106824
        
    End If

    PrecosDeVenda_Calcula2 = SUCESSO

    Exit Function

Erro_PrecosDeVenda_Calcula2:

    PrecosDeVenda_Calcula2 = gErr

    Select Case gErr

        Case 106821, 106822
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRECOCALCULADO", gErr)
            
        Case 106823
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_PRECOCALCULADO", gErr)
            
        Case 106824
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PRECOCALCULADO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144666)

    End Select

    Exit Function

End Function

Private Function PrecosDeVenda_Calcula3(ByVal objCalcPrecoVenda As ClassCalcPrecoVenda, objContexto As ClassContextoPlan, alComando() As Long) As Long

Dim sMnemonico As String, sExpressao As String
Dim lErro As Long, dtDataCF As Date

On Error GoTo Erro_PrecosDeVenda_Calcula3

    objContexto.iFilialFaturamento = objCalcPrecoVenda.iFilialEmpresa
    objContexto.dQuantidade = FORMACAO_PRECO_QTDECALCPRECO
    objContexto.iAno = Year(objCalcPrecoVenda.dtDataReferencia)
    
    lErro = CF("FilialEmpresa_ObtemDataCustoFixo", objContexto.iFilialFaturamento, dtDataCF)
    If lErro <> SUCESSO Then gError 106859
    
    objContexto.dtDataCustoFixo = dtDataCF
    
    lErro = CF("ContextoPlan_InicializaCusto", objContexto)
    If lErro <> SUCESSO Then gError 106829

    PrecosDeVenda_Calcula3 = SUCESSO
     
    Exit Function
    
Erro_PrecosDeVenda_Calcula3:

    PrecosDeVenda_Calcula3 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144667)
     
    End Select
     
    Exit Function

End Function
    
Private Function PrecosDeVenda_Calcula4(ByVal objCalcPrecoVenda As ClassCalcPrecoVenda, ByVal sProduto As String, ByVal iTabelaPreco As Integer, ByVal colFormacaoPreco As Collection, ByVal dValor As Double, alComando() As Long) As Long
'armazena a "memoria de calculo" do produto-tabela de preco

Dim lErro As Long, dValorAux As Double, iLinha As Integer
Dim objFormacaoPreco As ClassFormacaoPreco

On Error GoTo Erro_PrecosDeVenda_Calcula4

    'exclui o calculo anterior para a data, se houver
    lErro = Comando_ExecutarPos(alComando(6), "SELECT Valor FROM FormacaoPrecoCalc WHERE FilialEmpresa = ? AND Data = ? AND TabelaPreco = ? AND Produto = ?", 0, _
        dValorAux, objCalcPrecoVenda.iFilialEmpresa, objCalcPrecoVenda.dtDataReferencia, iTabelaPreco, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 106927
    
    lErro = Comando_BuscarProximo(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106928
    
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_ExecutarPos(alComando(7), "DELETE FROM FormacaoPrecoCalc", alComando(6))
        If lErro <> AD_SQL_SUCESSO Then gError 106929
    
        'exclui as linhas do calculo anterior para a data
        lErro = Comando_ExecutarPos(alComando(8), "SELECT Linha FROM FormacaoPrecoCalcLin WHERE FilialEmpresa = ? AND Data = ? AND TabelaPreco = ? AND Produto = ?", 0, _
            iLinha, objCalcPrecoVenda.iFilialEmpresa, objCalcPrecoVenda.dtDataReferencia, iTabelaPreco, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 106930
        
        lErro = Comando_BuscarProximo(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106931
        
        Do While lErro = AD_SQL_SUCESSO
        
            lErro = Comando_ExecutarPos(alComando(9), "DELETE FROM FormacaoPrecoCalcLin", alComando(8))
            If lErro <> AD_SQL_SUCESSO Then gError 106932
            
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106933
        
        Loop
    
    End If
    
    'inclui os dados do calculo atual
    For iLinha = 1 To colFormacaoPreco.Count
    
        Set objFormacaoPreco = colFormacaoPreco(iLinha)
        
        If iLinha = 1 Then
            
            'inclui os dados gerais do produto
            lErro = Comando_Executar(alComando(10), "INSERT INTO FormacaoPrecoCalc (FilialEmpresa, Data, TabelaPreco, Produto, Valor, Escopo, ItemCategoria) VALUES (?,?,?,?,?,?,?)", _
                objCalcPrecoVenda.iFilialEmpresa, objCalcPrecoVenda.dtDataReferencia, iTabelaPreco, sProduto, dValor, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria)
            If lErro <> AD_SQL_SUCESSO Then gError 106934
        
        End If
        
        'inclui os dados da linha
        lErro = Comando_Executar(alComando(11), "INSERT INTO FormacaoPrecoCalcLin (FilialEmpresa, Data, TabelaPreco, Produto, Linha, Valor, Expressao, Titulo) VALUES (?,?,?,?,?,?,?,?)", _
            objCalcPrecoVenda.iFilialEmpresa, objCalcPrecoVenda.dtDataReferencia, iTabelaPreco, sProduto, iLinha, objFormacaoPreco.dValor, objFormacaoPreco.sExpressao, objFormacaoPreco.sTitulo)
        If lErro <> AD_SQL_SUCESSO Then gError 106935
        
    Next
    
    PrecosDeVenda_Calcula4 = SUCESSO
     
    Exit Function
    
Erro_PrecosDeVenda_Calcula4:

    PrecosDeVenda_Calcula4 = gErr
     
    Select Case gErr
          
        Case 106927, 106928
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORMACAOPRECOCALC", gErr)
        
        Case 106929
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FORMACAOPRECOCALC", gErr)
        
        Case 106930, 106931, 106933
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORMACAOPRECOCALCLIN", gErr)
        
        Case 106932
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FORMACAOPRECOCALCLIN", gErr)
        
        Case 106934
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_FORMACAOPRECOCALC", gErr)
        
        Case 106935
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_FORMACAOPRECOCALCLIN", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144668)
     
    End Select
     
    Exit Function

End Function
