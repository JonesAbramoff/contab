VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Customiza_MovEst_Prod_Grava1(sProduto As String, ByVal iIndice As Integer, colItemMovEst As ColItensMovEstoque) As Long

Dim sCategoria As String
Dim sItem As String
Dim lErro As Long
Dim objItemMovEst As ClassItemMovEstoque
Dim sProduto1 As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Customiza_MovEst_Prod_Grava1

    If iIndice > 0 And iIndice <= colItemMovEst.Count Then

        Set objItemMovEst = colItemMovEst.Item(iIndice)

        sProduto = objItemMovEst.sProduto

    End If

    Customiza_MovEst_Prod_Grava1 = SUCESSO
        
    Exit Function
    
Erro_Customiza_MovEst_Prod_Grava1:

    Customiza_MovEst_Prod_Grava1 = gErr
    
    Select Case gErr
    
        Case 117687, 117688, 117690
        
        Case 117689
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179199)
    
    End Select

    Exit Function

End Function

Function Customiza_RomaneioGrade_TrataParam(objProduto As ClassProduto) As Long
'se o produto for da categoria "GENERICO" vai trocar o produto pelo seu "pai" que deve ser um pai de grade

Dim sCategoria As String
Dim sItem As String
Dim lErro As Long
Dim sProduto1 As String
Dim sProdutoPai As String

On Error GoTo Erro_Customiza_RomaneioGrade_TrataParam

    sCategoria = CATEGORIAPRODUTO_GENERICO

    sProduto1 = objProduto.sCodigo

    'retorna o item do produto/categoria
    lErro = CF("ProdutoCategoria_Le", sProduto1, sCategoria, sItem)
    If lErro <> SUCESSO And lErro <> 92405 Then gError 117692

    'se o produto tem categoria GENERICO e item SIM ==> em sProduto vai estar o produto que foi feita a op
    If lErro = SUCESSO And UCase(sItem) = "SIM" Then
        
        lErro = CF("Produto_Le_ProdutoPai", sProduto1, objProduto)
        If lErro <> SUCESSO Then gError 117697
        
    End If

    Customiza_RomaneioGrade_TrataParam = SUCESSO
        
    Exit Function
    
Erro_Customiza_RomaneioGrade_TrataParam:

    Customiza_RomaneioGrade_TrataParam = gErr
    
    Select Case gErr
    
        Case 117692, 117697
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179200)
    
    End Select

    Exit Function

End Function

Function Customiza_ProducaoEnt_SaidaCel(objProduto As ClassProduto) As Long
'se o produto for da categoria "GENERICO" vai trocar o produto pelo seu "pai" que deve ser um pai de grade

Dim sCategoria As String
Dim sItem As String
Dim lErro As Long
Dim sProduto1 As String
Dim sProdutoPai As String

On Error GoTo Erro_Customiza_ProducaoEnt_SaidaCel

    sCategoria = CATEGORIAPRODUTO_GENERICO

    sProduto1 = objProduto.sCodigo

    'retorna o item do produto/categoria
    lErro = CF("ProdutoCategoria_Le", sProduto1, sCategoria, sItem)
    If lErro <> SUCESSO And lErro <> 92405 Then gError 117693

    'se o produto tem categoria GENERICO e item SIM ==> em sProduto vai estar o produto que foi feita a op
    If lErro = SUCESSO And UCase(sItem) = "SIM" Then
        
        objProduto.sGrade = sCategoria
        
    End If

    Customiza_ProducaoEnt_SaidaCel = SUCESSO
        
    Exit Function
    
Erro_Customiza_ProducaoEnt_SaidaCel:

    Customiza_ProducaoEnt_SaidaCel = gErr
    
    Select Case gErr
    
        Case 117693
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179201)
    
    End Select

    Exit Function

End Function

Function Cust_RomaneioGrade_PreencheAlmox(sProduto As String, objRomaneioGrade As ClassRomaneioGrade) As Long
'troca o produto de grade pelo produto produzido

Dim sCategoria As String
Dim sItem As String
Dim lErro As Long
Dim sProduto1 As String

On Error GoTo Erro_Cust_RomaneioGrade_PreencheAlmox

    sCategoria = CATEGORIAPRODUTO_GENERICO

    sProduto1 = objRomaneioGrade.objObjetoTela.sProduto

    'retorna o item do produto/categoria
    lErro = CF("ProdutoCategoria_Le", sProduto1, sCategoria, sItem)
    If lErro <> SUCESSO And lErro <> 92405 Then gError 117692

    'se o produto tem categoria GENERICO e item SIM ==> em sProduto vai estar o produto que foi feita a op
    If lErro = SUCESSO And UCase(sItem) = "SIM" Then
        
        sProduto = objRomaneioGrade.objObjetoTela.sProduto
        
    End If

    Cust_RomaneioGrade_PreencheAlmox = SUCESSO
        
    Exit Function
    
Erro_Cust_RomaneioGrade_PreencheAlmox:

    Cust_RomaneioGrade_PreencheAlmox = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179202)
    
    End Select

    Exit Function

End Function

Function RelEstMPDisp_Prepara(lNumIntRel As Long, ByVal iFilialEmpresa As Integer) As Long
'gera registros na tabela RelESTMPDisp para a execucao do relatorio de disponibilidade de materias primas

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(1 To 4) As Long
Dim sProduto As String
Dim dQuantDispMP As Double
Dim dQuantKit2Est As Double
Dim objProdutoMP As New ClassProduto

On Error GoTo Erro_RelEstMPDisp_Prepara

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 133069
    Next
    
    'Inicia a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 133070

    'obtem o nº automatico
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_NUMINTREL", lNumIntRel)
    If lErro <> SUCESSO Then gError 133071
    
    sProduto = String(STRING_PRODUTO, 0)
                        
    'Lê o Produto
    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos, ProdutoCategoria WHERE Codigo = Produto AND UPPER(Categoria) = 'MP' And UPPER(Item) = 'SCAMPER' ORDER BY Codigo", sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 133072

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133073

    Do While lErro = AD_SQL_SUCESSO
    
        objProdutoMP.sCodigo = sProduto
    
        'Lê o produto no BD para obter UM de estoque
        lErro = CF("Produto_Le", objProdutoMP)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 133074

        If lErro = 28030 Then gError 133075
    
        dQuantKit2Est = 0
        dQuantDispMP = 0
        
        lErro = RelEstMPDisp_Busca_Arvore(iFilialEmpresa, sProduto, dQuantKit2Est, dQuantDispMP, objProdutoMP, lNumIntRel)
        If lErro <> SUCESSO Then gError 133076
            
        lErro = RelEstMPDisp_Grava(iFilialEmpresa, sProduto, 1, dQuantDispMP, objProdutoMP, lNumIntRel)
        If lErro <> SUCESSO Then gError 133058
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133077

    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 133078
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    RelEstMPDisp_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelEstMPDisp_Prepara:

    RelEstMPDisp_Prepara = gErr
     
    Select Case gErr
          
        Case 133069
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 133070
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
          
        Case 133071, 133074, 133076
          
        Case 133072, 133073, 133077
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
          
        Case 133075
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProdutoMP.sCodigo)
          
        Case 133078
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179203)
     
    End Select
     
    Call Transacao_Rollback
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    Exit Function

End Function

Function RelEstMPDisp_Busca_Arvore(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal dQuantKit3Est As Double, dQuantDispMP As Double, ByVal objProdutoMP As ClassProduto, ByVal lNumIntRel As Long) As Long

Dim lErro As Long
Dim sProdutoRaiz As String
Dim sUM As String
Dim dQuantKit As Double
Dim alComando(1 To 2) As Long
Dim dFator As Double
Dim iIndice As Integer
Dim dQuantKitEst As Double
Dim dQuantKit1 As Double
Dim dQuantKit1Est As Double
Dim dQuantKit2Est As Double
Dim objProduto As New ClassProduto
Dim objProduto1 As New ClassProduto
Dim sProduto1 As String

On Error GoTo Erro_RelEstMPDisp_Busca_Arvore

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 133055
    Next

    sProdutoRaiz = String(STRING_PRODUTO, 0)
    sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)

    lErro = Comando_Executar(alComando(1), "SELECT Kit.ProdutoRaiz, ProdutoKit.UnidadeMed, Quantidade FROM Kit, ProdutoKit WHERE Produto = ? AND Kit.ProdutoRaiz = ProdutoKit.ProdutoRaiz AND Kit.Situacao = 1 AND Nivel = 1", _
    sProdutoRaiz, sUM, dQuantKit, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 133056
                
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133057
        
    If lErro = AD_SQL_SEM_DADOS Then
    
        If dQuantKit3Est <> 0 Then
    
            lErro = RelEstMPDisp_Grava(iFilialEmpresa, sProduto, dQuantKit3Est, dQuantDispMP, objProdutoMP, lNumIntRel)
            If lErro <> SUCESSO Then gError 133058
    
        End If
    
    Else
        
        objProduto.sCodigo = sProduto
    
        'Lê o produto no BD para obter UM de estoque
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 133059

        If lErro = 28030 Then gError 133060
        
        Do While lErro = AD_SQL_SUCESSO
        
            'Obtém fator de conversão para unidade de medida de ItemPV
            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 133061
                
            dQuantKitEst = dQuantKit * dFator
            
            If dQuantKit3Est <> 0 Then dQuantKitEst = dQuantKitEst * dQuantKit3Est
            
            sProduto1 = sProdutoRaiz
            
            sProdutoRaiz = String(STRING_PRODUTO, 0)
            sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
            
            lErro = Comando_Executar(alComando(2), "SELECT ProdutoKit.ProdutoRaiz, ProdutoKit.UnidadeMed, ProdutoKit.Quantidade FROM Kit, ProdutoKit WHERE Produto = ? AND Kit.ProdutoRaiz = ProdutoKit.ProdutoRaiz AND Situacao = 1 And Nivel = 0", _
            sProdutoRaiz, sUM, dQuantKit1, sProduto1)
            If lErro <> AD_SQL_SUCESSO Then gError 133062
                        
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133063
            
            If lErro = AD_SQL_SUCESSO Then
            
                objProduto1.sCodigo = sProduto1
            
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto1)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 133064
        
                If lErro = 28030 Then gError 133065
            
                'Obtém fator de conversão para unidade de medida de ItemPV
                lErro = CF("UM_Conversao", objProduto1.iClasseUM, sUM, objProduto1.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 133066
            
                dQuantKit1Est = dQuantKit1 * dFator
                
                dQuantKit2Est = dQuantKitEst / dQuantKit1Est
                
                sProduto1 = sProdutoRaiz
            
                lErro = RelEstMPDisp_Busca_Arvore(iFilialEmpresa, sProduto1, dQuantKit2Est, dQuantDispMP, objProdutoMP, lNumIntRel)
                If lErro <> SUCESSO Then gError 133067
        
            End If
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133068
        
        Loop
    
    End If
    
    RelEstMPDisp_Busca_Arvore = SUCESSO

    Exit Function

Erro_RelEstMPDisp_Busca_Arvore:

    RelEstMPDisp_Busca_Arvore = gErr
     
    Select Case gErr
          
        Case 133055
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 133056, 133057, 133068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)

        Case 133058, 133059, 133061, 133062, 133063, 133064, 133066, 133067
        
        Case 133060, 133065
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179204)
     
    End Select
     
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    Exit Function

End Function

Function RelEstMPDisp_Grava(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal dQuantKit2Est As Double, dQuantDispMP As Double, ByVal objProdutoMP As ClassProduto, ByVal lNumIntRel As Long) As Long
'dQuantKit2Est tem a quantidade necessaria da materia prima por unidade do produto vendido
'vai pesquisar o estoque disponivel e a quantidade em pedido do produto vendido
'vai pesquisar o quanto de mat prima foi requisitada e ainda nao produzido
            
Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 7) As Long
Dim dQuantPedida As Double
Dim dQuantOP As Double
Dim dQuantProd As Double
Dim lNumIntDocItemOP As Long
Dim dQuantidade As Double
Dim sUM As String
Dim dFator As Double
Dim dQuantReq As Double
Dim dQuantReqNaoUtilizada As Double
Dim dQuantDisp As Double
            
On Error GoTo Erro_RelEstMPDisp_Grava

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 133042
    Next

    lErro = Comando_Executar(alComando(1), "SELECT SUM(QuantDispNossa + QuantReservada) FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND FilialEmpresa = ? AND Produto = ?", _
    dQuantDisp, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 133043
                
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133044

    If dQuantKit2Est <> 0 Then

        dQuantDisp = dQuantDisp * dQuantKit2Est

        lErro = Comando_Executar(alComando(2), "SELECT QuantPedida FROM ProdutosFilial WHERE FilialEmpresa = ? AND Produto = ?", _
        dQuantPedida, iFilialEmpresa, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133045
                    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133046

        dQuantPedida = dQuantPedida * dQuantKit2Est

        lErro = Comando_Executar(alComando(3), "SELECT Quantidade, QuantidadeProd, NumIntDoc FROM ItensOrdemProducao WHERE FilialEmpresa = ? AND Produto = ? AND (Situacao = 0 OR Situacao = 2)", _
        dQuantOP, dQuantProd, lNumIntDocItemOP, iFilialEmpresa, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133047
                    
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133048

        Do While lErro = AD_SQL_SUCESSO
        
            If dQuantOP > dQuantProd Then
            
                sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
            
                lErro = Comando_Executar(alComando(4), "SELECT Quantidade, SiglaUM FROM MovimentoEstoque WHERE FilialEmpresa = ? AND TipoMov = ? AND Produto = ? AND TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? AND NumIntDocEst = 0", _
                dQuantidade, sUM, iFilialEmpresa, MOV_EST_REQ_PRODUCAO, objProdutoMP.sCodigo, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, lNumIntDocItemOP)
                If lErro <> AD_SQL_SUCESSO Then gError 133049
            
                lErro = Comando_BuscarPrimeiro(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133050
            
                Do While lErro = AD_SQL_SUCESSO
            
                    'Obtém fator de conversão para unidade de medida de ItemPV
                    lErro = CF("UM_Conversao", objProdutoMP.iClasseUM, sUM, objProdutoMP.sSiglaUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 133051
        
                    dQuantReq = dQuantReq + dQuantidade * dFator
                    
                    lErro = Comando_BuscarProximo(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133052
                    
                Loop
                
                If (dQuantProd * dQuantKit2Est) < dQuantReq Then
                
                    dQuantReqNaoUtilizada = dQuantReqNaoUtilizada + dQuantReq - (dQuantProd * dQuantKit2Est)
                    
                End If
                                
            End If
            
            dQuantReq = 0
                        
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133053
                                
        Loop

    End If

    dQuantDispMP = dQuantDispMP + dQuantDisp + dQuantReqNaoUtilizada - dQuantPedida

    lErro = Comando_Executar(alComando(5), "INSERT INTO RelEstMPDisp (NumIntRel, Produto, ProdutoMP, QuantEst, QuantPedida, QuantProd, QuantDisp) VALUES (?,?,?,?,?,?,?) ", _
    lNumIntRel, sProduto, objProdutoMP.sCodigo, dQuantDisp, dQuantPedida, dQuantReqNaoUtilizada, dQuantDispMP)
    If lErro <> AD_SQL_SUCESSO Then gError 133054
    
    
    'se estiver gravando a materia prima
    If objProdutoMP.sCodigo = sProduto Then
    
        lErro = Comando_Executar(alComando(6), "SELECT SUM(QuantEst), Sum(QuantPedida), Sum(QuantProd), Sum(QuantDisp) FROM RelEstMPDisp WHERE ProdutoMP = ? And NumIntRel = ?", _
        dQuantDisp, dQuantPedida, dQuantReqNaoUtilizada, dQuantDispMP, sProduto, lNumIntRel)
        If lErro <> AD_SQL_SUCESSO Then gError 126161
    
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO Then gError 126162
        
        If dQuantDisp = 0 And dQuantPedida = 0 And dQuantReqNaoUtilizada = 0 And dQuantDispMP = 0 Then
            
            lErro = Comando_Executar(alComando(7), "DELETE FROM RelEstMPDisp WHERE NumIntRel = ? And ProdutoMP = ?", lNumIntRel, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 126163
    
        End If
    
    End If
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next


    RelEstMPDisp_Grava = SUCESSO

    Exit Function

Erro_RelEstMPDisp_Grava:

    RelEstMPDisp_Grava = gErr
     
    Select Case gErr
          
        Case 126161, 126162
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELESTMPDISP", gErr, lNumIntRel, objProdutoMP.sCodigo)
          
        Case 126163
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELESTMPDISP", gErr, lNumIntRel, objProdutoMP.sCodigo)
          
        Case 133042
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 133043, 133044
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO2", gErr, sProduto)
          
        Case 133045, 133046
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", gErr, iFilialEmpresa, sProduto)
          
        Case 133047, 133048, 133053
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO4", gErr, iFilialEmpresa, sProduto)
          
        Case 133049, 133050, 133052
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE8", gErr, iFilialEmpresa, MOV_EST_REQ_PRODUCAO, objProdutoMP.sCodigo, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, lNumIntDocItemOP)
          
        Case 133051
        
        Case 133054
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELESTMPDISP", gErr, lNumIntRel, sProduto)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179205)
     
    End Select
     
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    Exit Function

End Function


Function LogMovEstoque_Inserir(ByVal objItemMovEst As ClassItemMovEstoque, ByVal iOperacao As Integer) As Long
'insere um movimento de estoque na tabela LogMovEstoque
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim lComando As Long
Dim sComputer As String

On Error GoTo Erro_LogMovEstoque_Inserir

    sComputer = String(128, 0)

    Call GetComputerName(sComputer, 128)

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 126164

    lErro = Comando_Executar(lComando, "INSERT INTO LogMovEstoque (Operacao, DataLog, HoraLog , Usuario, FilialEmpresa, Codigo, NumIntDoc, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao, DataRegistro, Hora, NumIntDocGrade, Observacao, Requisitante, ItemNF, QuantInsumos, Estacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        iOperacao, CDate(Date), CDbl(Time), gsUsuario, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo, objItemMovEst.lNumIntDoc, objItemMovEst.dCusto, objItemMovEst.iApropriacao, objItemMovEst.sProduto, objItemMovEst.sSiglaUM, objItemMovEst.dQuantidade, objItemMovEst.iAlmoxarifado, objItemMovEst.iTipoMov, objItemMovEst.lNumIntDocOrigem, objItemMovEst.iTipoNumIntDocOrigem, objItemMovEst.dtData, objItemMovEst.sCcl, objItemMovEst.lNumIntDocEst, objItemMovEst.lCliente, objItemMovEst.lFornecedor, objItemMovEst.sOPCodigo, objItemMovEst.sDocOrigem, objItemMovEst.sContaContabilEst, objItemMovEst.sContaContabilAplic, _
        objItemMovEst.lHorasMaquina, objItemMovEst.dtDataInicioProducao, objItemMovEst.dtDataRegistro, CDbl(objItemMovEst.dtHora), objItemMovEst.lNumIntDocGrade, objItemMovEst.sObservacao, objItemMovEst.lRequisitante, 0, 0, sComputer)
    If lErro <> AD_SQL_SUCESSO Then gError 126165

    Call Comando_Fechar(lComando)

    LogMovEstoque_Inserir = SUCESSO

    Exit Function

Erro_LogMovEstoque_Inserir:

    LogMovEstoque_Inserir = gErr

    Select Case gErr

        Case 126164
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126165
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOGMOVESTOQUE", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179206)

    End Select

    Exit Function

End Function

'Function RelContainer_Prepara(lNumIntRel As Long, ByVal iFilialEmpresa As Integer) As Long
''gera registros na tabela RelContainer para a execucao do relatorio de disponibilidade de materias primas
'
'Dim lErro As Long
'Dim lTransacao As Long
'Dim iIndice As Integer
'Dim alComando(1 To 4) As Long
'Dim sProduto As String
'Dim dQuantKit2Est As Double
'Dim objProdutoVend As New ClassProduto
'Dim iAchou As Integer
'
'On Error GoTo Erro_RelContainer_Prepara
'
'    'abrir comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 133069
'    Next
'
'    'Inicia a transacao
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 133070
'
'    'obtem o nº automatico
'    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_NUMINTREL", lNumIntRel)
'    If lErro <> SUCESSO Then gError 133071
'
'    sProduto = String(STRING_PRODUTO, 0)
'
'    'Lê os Produtos que podem ser vendidos
'    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos WHERE Faturamento = 1 ORDER BY Codigo", sProduto)
'    If lErro <> AD_SQL_SUCESSO Then gError 133072
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133073
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        objProdutoVend.sCodigo = sProduto
'
'        'Lê o produto no BD para obter UM de estoque
'        lErro = CF("Produto_Le", objProdutoVend)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError 133074
'
'        If lErro = 28030 Then gError 133075
'
'        dQuantKit2Est = 0
'        dQuantDispMP = 0
'
'        iAchou = 0
'
'        lErro = RelContainer_Busca_Arvore(iFilialEmpresa, sProduto, dQuantKit2Est, dQuantDispMP, objProdutoVend, lNumIntRel, sVersao, iAchou)
'        If lErro <> SUCESSO Then gError 133076
'
'        lErro = Comando_BuscarProximo(alComando(1))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133077
'
'    Loop
'
'    'Confirma a transação
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then gError 133078
'
'   'fechar comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    RelContainer_Prepara = SUCESSO
'
'    Exit Function
'
'Erro_RelContainer_Prepara:
'
'    RelContainer_Prepara = gErr
'
'    Select Case gErr
'
'        Case 133069
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 133070
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'        Case 133071, 133074, 133076
'
'        Case 133072, 133073, 133077
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", Err)
'
'        Case 133075
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProdutoVend.sCodigo)
'
'        Case 133078
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179207)
'
'    End Select
'
'    Call Transacao_Rollback
'
'   'fechar comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Function RelContainer_Busca_Arvore(ByVal iFilialEmpresa As Integer, ByVal sProdutoRaiz As String, ByVal dQuantKit3Est As Double, dQuantDispMP As Double, ByVal objProdutoVend As ClassProduto, ByVal lNumIntRel As Long, iAchou As Integer) As Long
'
'Dim lErro As Long
'Dim sUM As String
'Dim dQuantKit As Double
'Dim alComando(1 To 2) As Long
'Dim dFator As Double
'Dim iIndice As Integer
'Dim dQuantKitEst As Double
'Dim dQuantKit1 As Double
'Dim dQuantKit1Est As Double
'Dim dQuantKit2Est As Double
'Dim objProduto As New ClassProduto
'Dim objProduto1 As New ClassProduto
'Dim sProduto1 As String
'
'On Error GoTo Erro_RelContainer_Busca_Arvore
'
'    'abrir comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 133055
'    Next
'
'    sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
'
'    lErro = Comando_Executar(alComando(1), "SELECT ProdutoKit.UnidadeMed, Quantidade FROM Kit, ProdutoKit WHERE Produto = ? AND Kit.ProdutoRaiz = ProdutoKit.ProdutoRaiz AND Kit.Situacao = 1 AND Nivel = 0", _
'    sUM, dQuantKit, sProdutoRaiz)
'    If lErro <> AD_SQL_SUCESSO Then gError 133056
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133057
'
'    If lErro = AD_SQL_SEM_DADOS Then
'
'        If dQuantKit3Est <> 0 Then
'
'            'Lê o Produto
'            lErro = Comando_Executar(alComando(3), "SELECT Codigo FROM Produtos, ProdutoCategoria WHERE Codigo = Produto AND UPPER(Categoria) = 'MP' And UPPER(Item) = 'SCAMPER' ORDER BY Codigo", sProdutoRaiz)
'            If lErro <> AD_SQL_SUCESSO Then gError 133072
'
'            lErro = Comando_BuscarPrimeiro(alComando(3))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133057
'
'            If lErro = AD_SQL_SUCESSO Then
'
'                iAchou = 1
'
'                lErro = RelContainer_Grava(iFilialEmpresa, sProdutoRaiz, dQuantKit3Est, dQuantDispMP, objProdutoVend, lNumIntRel)
'                If lErro <> SUCESSO Then gError 133058
'
'            End If
'
'        End If
'
'    Else
'
'        objProduto.sCodigo = sProdutoRaiz
'
'        'Lê o produto no BD para obter UM de estoque
'        lErro = CF("Produto_Le", objProduto)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError 133059
'
'        If lErro = 28030 Then gError 133060
'
'        'Obtém fator de conversão para unidade de medida de ItemPV
'        lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
'        If lErro <> SUCESSO Then gError 133061
'
'        dQuantKitEst = dQuantKit * dFator
'
'        If dQuantKit3Est <> 0 Then dQuantKitEst = dQuantKitEst * dQuantKit3Est
'
'        sProduto1 = String(STRING_PRODUTO, 0)
'        sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
'
'        lErro = Comando_Executar(alComando(2), "SELECT ProdutoKit.ProdutoRaiz, ProdutoKit.UnidadeMed, ProdutoKit.Quantidade FROM Kit, ProdutoKit WHERE ProdutoKit.ProdutoRaiz = ? AND Kit.ProdutoRaiz = ProdutoKit.ProdutoRaiz AND Situacao = 1 And Nivel = 1", _
'        sProduto1, sUM, dQuantKit1, sProdutoRaiz)
'        If lErro <> AD_SQL_SUCESSO Then gError 133062
'
'        lErro = Comando_BuscarPrimeiro(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133063
'
'        Do While lErro = AD_SQL_SUCESSO
'
'            objProduto1.sCodigo = sProduto1
'
'            'Lê o produto no BD para obter UM de estoque
'            lErro = CF("Produto_Le", objProduto1)
'            If lErro <> SUCESSO And lErro <> 28030 Then gError 133064
'
'            If lErro = 28030 Then gError 133065
'
'            'Obtém fator de conversão para unidade de medida de ItemPV
'            lErro = CF("UM_Conversao", objProduto1.iClasseUM, sUM, objProduto1.sSiglaUMEstoque, dFator)
'            If lErro <> SUCESSO Then gError 133066
'
'            dQuantKit1Est = dQuantKit1 * dFator
'
'            dQuantKit2Est = dQuantKitEst / dQuantKit1Est
'
'            lErro = RelEstMPDisp_Busca_Arvore(iFilialEmpresa, sProduto1, dQuantKit2Est, dQuantDispMP, objProdutoVend, lNumIntRel)
'            If lErro <> SUCESSO Then gError 133067
'
'            If iAchou = 1 Then Exit Do
'
'            lErro = Comando_BuscarProximo(alComando(1))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133068
'
'        Loop
'
'    End If
'
'    RelContainer_Busca_Arvore = SUCESSO
'
'    Exit Function
'
'Erro_RelContainer_Busca_Arvore:
'
'    RelContainer_Busca_Arvore = gErr
'
'    Select Case gErr
'
'        Case 133055
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 133056, 133057, 133068
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)
'
'        Case 133058, 133059, 133061, 133062, 133063, 133064, 133066, 133067
'
'        Case 133060, 133065
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179208)
'
'    End Select
'
'   'fechar comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function
'
'Function RelContainer_Grava(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal dQuantKit2Est As Double, dQuantDispMP As Double, ByVal objProdutoVend As ClassProduto, ByVal lNumIntRel As Long) As Long
''dQuantKit2Est tem a quantidade necessaria da materia prima por unidade do produto vendido
''vai pesquisar o estoque disponivel e a quantidade em pedido do produto vendido
''vai pesquisar o quanto de mat prima foi requisitada e ainda nao produzido
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim alComando(1 To 7) As Long
'Dim dQuantPedida As Double
'Dim dQuantidade As Double
'Dim sUM As String
'Dim dFator As Double
'Dim dQuantDisp As Double
'Dim dQuantSaldo As Double
'Dim dtEntrada As Date
'Dim dQuantResto As Double
'
'On Error GoTo Erro_RelContainer_Grava
'
'    'abrir comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 133042
'    Next
'
'    lErro = Comando_Executar(alComando(1), "SELECT SUM(QuantDispNossa + QuantReservada) FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND FilialEmpresa = ? AND Produto = ?", _
'    dQuantDisp, iFilialEmpresa, objProdutoVend.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 133043
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133044
'
'    dQuantDisp = dQuantDisp
'
'    lErro = Comando_Executar(alComando(2), "SELECT QuantPedida FROM ProdutosFilial WHERE FilialEmpresa = ? AND Produto = ?", _
'    dQuantPedida, iFilialEmpresa, objProdutoVend.sCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 133045
'
'    lErro = Comando_BuscarPrimeiro(alComando(2))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133046
'
'    dQuantSaldo = dQuantDisp - dQuantPedida
'
'    dQuantResto = dQuantSaldo
'
'    sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
'
'    'Busca o Ítem
'    lErro = Comando_Executar(alComando(3), "SELECT NFiscal.DataEntrada, ItensNFiscal.UnidadeMed, ItensNFiscal.Quantidade,  ItensNFiscal.DataEntrega FROM NFiscal, ItensNFiscal WHERE NFiscal.Serie = 'DI' AND ItensNFiscal.Produto = ? AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND NFiscal.TipoNFiscal = TiposDocInfo.Codigo AND NFiscal.FilialEmpresa = ? AND NFiscal.Status <> ? AND ItensNFiscal.Status <> ? AND (TiposDocInfo.Tipo = ? OR TiposDocInfo.Tipo =?) ORDER BY DataEntrada DESC", _
'    dtEntrada, sUM, dQuantidade, iFilialEmpresa, STATUS_EXCLUIDO, STATUS_EXCLUIDO, DOCINFO_NF_INT_ENTRADA, DOCINFO_NF_EXTERNA)
'    If lErro <> AD_SQL_SUCESSO Then Error 39321
'
'    lErro = Comando_BuscarPrimeiro(alComando(3))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 39322
'
'    If lErro = AD_SQL_SEM_DADOS Then
'
'
'    Else
'
'        Do While lErro = AD_SQL_SUCESSO
'
'
'        lErro = Comando_Executar(alComando(5), "INSERT INTO RelIdadeEstoque (NumIntRel, Produto, ProdutoMP, QuantEst, QuantPedida, QuantProd, QuantDisp) VALUES (?,?,?,?,?,?,?) ", _
'        lNumIntRel, sProduto, objProdutoRaiz.sCodigo, dQuantDisp, dQuantPedida, dQuantReqNaoUtilizada, dQuantDispMP)
'        If lErro <> AD_SQL_SUCESSO Then gError 133054
'
'
'            lErro = Comando_BuscarProximo(alComando(3))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133053
'
'        Loop
'
'    End If
'
'    dQuantDispMP = dQuantDispMP + dQuantDisp + dQuantReqNaoUtilizada - dQuantPedida
'
'    lErro = Comando_Executar(alComando(5), "INSERT INTO RelEstMPDisp (NumIntRel, Produto, ProdutoMP, QuantEst, QuantPedida, QuantProd, QuantDisp) VALUES (?,?,?,?,?,?,?) ", _
'    lNumIntRel, sProduto, objProdutoRaiz.sCodigo, dQuantDisp, dQuantPedida, dQuantReqNaoUtilizada, dQuantDispMP)
'    If lErro <> AD_SQL_SUCESSO Then gError 133054
'
'
'    'se estiver gravando a materia prima
'    If objProdutoRaiz.sCodigo = sProduto Then
'
'        lErro = Comando_Executar(alComando(6), "SELECT SUM(QuantEst), Sum(QuantPedida), Sum(QuantProd), Sum(QuantDisp) FROM RelEstMPDisp WHERE ProdutoMP = ? And NumIntRel = ?", _
'        dQuantDisp, dQuantPedida, dQuantReqNaoUtilizada, dQuantDispMP, sProduto, lNumIntRel)
'        If lErro <> AD_SQL_SUCESSO Then gError 126161
'
'        lErro = Comando_BuscarPrimeiro(alComando(6))
'        If lErro <> AD_SQL_SUCESSO Then gError 126162
'
'        If dQuantDisp = 0 And dQuantPedida = 0 And dQuantReqNaoUtilizada = 0 And dQuantDispMP = 0 Then
'
'            lErro = Comando_Executar(alComando(7), "DELETE FROM RelEstMPDisp WHERE NumIntRel = ? And ProdutoMP = ?", lNumIntRel, sProduto)
'            If lErro <> AD_SQL_SUCESSO Then gError 126163
'
'        End If
'
'    End If
'
'   'fechar comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'
'    RelContainer_Grava = SUCESSO
'
'    Exit Function
'
'Erro_RelContainer_Grava:
'
'    RelContainer_Grava = gErr
'
'    Select Case gErr
'
'        Case 126161, 126162
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELESTMPDISP", gErr, lNumIntRel, objProdutoRaiz.sCodigo)
'
'        Case 126163
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELESTMPDISP", gErr, lNumIntRel, objProdutoRaiz.sCodigo)
'
'        Case 133042
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 133043, 133044
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO2", gErr, sProduto)
'
'        Case 133045, 133046
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", gErr, iFilialEmpresa, sProduto)
'
'        Case 133047, 133048, 133053
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO4", gErr, iFilialEmpresa, sProduto)
'
'        Case 133049, 133050, 133052
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE8", gErr, iFilialEmpresa, MOV_EST_REQ_PRODUCAO, objProdutoRaiz.sCodigo, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, lNumIntDocItemOP)
'
'        Case 133051
'
'        Case 133054
'            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELESTMPDISP", gErr, lNumIntRel, sProduto)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179209)
'
'    End Select
'
'   'fechar comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Function RelContainer_Prepara(lNumIntRel As Long, ByVal iFilialEmpresa As Integer, sProdIni As String, sProdFim As String) As Long
'gera registros na tabela RelContainer para a execucao do relatorio de saldo de produtos vs containers

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(1 To 4) As Long
Dim sProduto As String
Dim dQuantDispMP As Double
Dim dQuantKit2Est As Double
Dim objProdutoMP As New ClassProduto

On Error GoTo Erro_RelContainer_Prepara

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126520
    Next
    
    'Inicia a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 126521

    'obtem o nº automatico
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_NUMINTREL", lNumIntRel)
    If lErro <> SUCESSO Then gError 126522
    
    sProduto = String(STRING_PRODUTO, 0)
                        
    If Len(Trim(sProdIni)) = 0 And Len(Trim(sProdFim)) = 0 Then
        
        lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos, ProdutoCategoria WHERE Codigo = Produto AND UPPER(Categoria) = 'MP' And UPPER(Item) = 'SCAMPER' ORDER BY Codigo", sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 126523
    
    ElseIf Len(Trim(sProdIni)) > 0 And Len(Trim(sProdFim)) > 0 Then

        lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos, ProdutoCategoria WHERE Codigo = Produto AND UPPER(Categoria) = 'MP' And UPPER(Item) = 'SCAMPER' AND Codigo >= ? AND Codigo <= ? ORDER BY Codigo", sProduto, sProdIni, sProdFim)
        If lErro <> AD_SQL_SUCESSO Then gError 126524

    ElseIf Len(Trim(sProdIni)) > 0 Then

        lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos, ProdutoCategoria WHERE Codigo = Produto AND UPPER(Categoria) = 'MP' And UPPER(Item) = 'SCAMPER' AND Codigo >= ? ORDER BY Codigo", sProduto, sProdIni)
        If lErro <> AD_SQL_SUCESSO Then gError 126525

    Else

        lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos, ProdutoCategoria WHERE Codigo = Produto AND UPPER(Categoria) = 'MP' And UPPER(Item) = 'SCAMPER' AND Codigo <= ? ORDER BY Codigo", sProduto, sProdFim)
        If lErro <> AD_SQL_SUCESSO Then gError 126526

    End If

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126527

    Do While lErro = AD_SQL_SUCESSO
    
        objProdutoMP.sCodigo = sProduto
    
        'Lê o produto no BD para obter UM de estoque
        lErro = CF("Produto_Le", objProdutoMP)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 126528

        If lErro = 28030 Then gError 126529
    
        dQuantKit2Est = 0
        dQuantDispMP = 0
        
        lErro = RelContainer_Busca_Arvore(iFilialEmpresa, sProduto, dQuantKit2Est, dQuantDispMP, objProdutoMP, lNumIntRel)
        If lErro <> SUCESSO Then gError 126530
            
        lErro = RelContainer_Grava(iFilialEmpresa, sProduto, 1, dQuantDispMP, objProdutoMP, lNumIntRel)
        If lErro <> SUCESSO Then gError 126531
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126532

    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 126533
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    RelContainer_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelContainer_Prepara:

    RelContainer_Prepara = gErr
     
    Select Case gErr
          
        Case 126520
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 126521
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
          
        Case 126522, 126528, 126530, 126531
          
        Case 126523 To 126527, 126532
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
          
        Case 126529
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProdutoMP.sCodigo)
          
        Case 126533
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179210)
     
    End Select
     
    Call Transacao_Rollback
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    Exit Function

End Function

Function RelContainer_Busca_Arvore(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal dQuantKit3Est As Double, dQuantDispMP As Double, ByVal objProdutoMP As ClassProduto, ByVal lNumIntRel As Long) As Long

Dim lErro As Long
Dim sProdutoRaiz As String
Dim sUM As String
Dim dQuantKit As Double
Dim alComando(1 To 2) As Long
Dim dFator As Double
Dim iIndice As Integer
Dim dQuantKitEst As Double
Dim dQuantKit1 As Double
Dim dQuantKit1Est As Double
Dim dQuantKit2Est As Double
Dim objProduto As New ClassProduto
Dim objProduto1 As New ClassProduto
Dim sProduto1 As String

On Error GoTo Erro_RelContainer_Busca_Arvore

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 133141
    Next

    sProdutoRaiz = String(STRING_PRODUTO, 0)
    sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)

    lErro = Comando_Executar(alComando(1), "SELECT Kit.ProdutoRaiz, ProdutoKit.UnidadeMed, Quantidade FROM Kit, ProdutoKit WHERE Produto = ? AND Kit.ProdutoRaiz = ProdutoKit.ProdutoRaiz AND Kit.Situacao = 1 AND Nivel = 1", _
    sProdutoRaiz, sUM, dQuantKit, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 133142
                
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133143
        
    If lErro = AD_SQL_SEM_DADOS Then
    
        If dQuantKit3Est <> 0 Then
    
            lErro = RelContainer_Acumula(iFilialEmpresa, sProduto, dQuantKit3Est, dQuantDispMP, objProdutoMP, lNumIntRel)
            If lErro <> SUCESSO Then gError 133144
    
        End If
    
    Else
        
        objProduto.sCodigo = sProduto
    
        'Lê o produto no BD para obter UM de estoque
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 133145

        If lErro = 28030 Then gError 133146
        
        Do While lErro = AD_SQL_SUCESSO
        
            'Obtém fator de conversão para unidade de medida de ItemPV
            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 133147
                
            dQuantKitEst = dQuantKit * dFator
            
            If dQuantKit3Est <> 0 Then dQuantKitEst = dQuantKitEst * dQuantKit3Est
            
            sProduto1 = sProdutoRaiz
            
            sProdutoRaiz = String(STRING_PRODUTO, 0)
            sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
            
            lErro = Comando_Executar(alComando(2), "SELECT ProdutoKit.ProdutoRaiz, ProdutoKit.UnidadeMed, ProdutoKit.Quantidade FROM Kit, ProdutoKit WHERE Produto = ? AND Kit.ProdutoRaiz = ProdutoKit.ProdutoRaiz AND Situacao = 1 And Nivel = 0", _
            sProdutoRaiz, sUM, dQuantKit1, sProduto1)
            If lErro <> AD_SQL_SUCESSO Then gError 133148
                        
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133149
            
            If lErro = AD_SQL_SUCESSO Then
            
                objProduto1.sCodigo = sProduto1
            
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto1)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 133150
        
                If lErro = 28030 Then gError 133151
            
                'Obtém fator de conversão para unidade de medida de ItemPV
                lErro = CF("UM_Conversao", objProduto1.iClasseUM, sUM, objProduto1.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 133152
            
                dQuantKit1Est = dQuantKit1 * dFator
                
                dQuantKit2Est = dQuantKitEst / dQuantKit1Est
                
                sProduto1 = sProdutoRaiz
            
                lErro = RelContainer_Busca_Arvore(iFilialEmpresa, sProduto1, dQuantKit2Est, dQuantDispMP, objProdutoMP, lNumIntRel)
                If lErro <> SUCESSO Then gError 133153
        
            End If
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133154
        
        Loop
    
    End If
    
    RelContainer_Busca_Arvore = SUCESSO

    Exit Function

Erro_RelContainer_Busca_Arvore:

    RelContainer_Busca_Arvore = gErr
     
    Select Case gErr
          
        Case 133141
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 133142, 133143, 133154, 133148, 133149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)

        Case 133144, 133145, 133147, 133150, 133152, 133153
        
        Case 133146, 133151
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179211)
     
    End Select
     
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    Exit Function

End Function

Function RelContainer_Grava(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal dQuantKit2Est As Double, dQuantDispMP As Double, ByVal objProdutoMP As ClassProduto, ByVal lNumIntRel As Long) As Long
'dQuantKit2Est tem a quantidade necessaria da materia prima por unidade do produto vendido
'vai pesquisar o estoque disponivel e a quantidade em pedido do produto vendido
'vai pesquisar o quanto de mat prima foi requisitada e ainda nao produzido
            
Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 3) As Long
Dim dQuantNF As Double
Dim sUM As String
Dim dFator As Double
Dim iSeq As Integer
Dim dtDataEntrada As Date
Dim dQuantidade As Double

            
On Error GoTo Erro_RelContainer_Grava

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 133155
    Next

    lErro = Comando_Executar(alComando(3), "SELECT SUM(QuantDispNossa + QuantReservada) FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND FilialEmpresa = ? AND Produto = ?", _
    dQuantidade, iFilialEmpresa, objProdutoMP.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 133156
                
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133157
    iSeq = 1
    
    dQuantDispMP = dQuantDispMP + dQuantidade

    lErro = Comando_Executar(alComando(1), "INSERT INTO RelContainer (NumIntRel, Produto, Seq, QuantNF, QuantDisp, DataEmissao) VALUES (?,?,?,?,?,?) ", _
    lNumIntRel, objProdutoMP.sCodigo, iSeq, 0, dQuantDispMP, DATA_NULA)
    If lErro <> AD_SQL_SUCESSO Then gError 133123

    If dQuantDispMP > 0 Then

        sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)

        lErro = Comando_Executar(alComando(2), "SELECT DataEntrada, UnidadeMed, Quantidade FROM NFiscal, ItensNFiscal WHERE ItensNFiscal.NumIntNF = NFIscal.NumIntDoc AND FilialEmpresa = ? AND Serie = ? AND Produto = ? ORDER BY DataEmissao DESC", _
        dtDataEntrada, sUM, dQuantNF, iFilialEmpresa, "DI", objProdutoMP.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 133124
                
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133125

        Do While lErro = AD_SQL_SUCESSO And dQuantDispMP > 0
                
            'Obtém fator de conversão para unidade de medida de ItemPV
            lErro = CF("UM_Conversao", objProdutoMP.iClasseUM, sUM, objProdutoMP.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 133126
            
            dQuantNF = dQuantNF * dFator
            iSeq = iSeq + 1
            
            If dQuantDispMP > dQuantNF Then
            
                dQuantidade = dQuantNF
                
            Else
            
                dQuantidade = dQuantDispMP
                
            End If
            
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelContainer (NumIntRel, Produto, Seq, QuantNF, QuantDisp, DataEmissao) VALUES (?,?,?,?,?,?) ", _
            lNumIntRel, objProdutoMP.sCodigo, iSeq, dQuantidade, 0, dtDataEntrada)
            If lErro <> AD_SQL_SUCESSO Then gError 133127
            
            dQuantDispMP = dQuantDispMP - dQuantNF
            
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133128
                        
        Loop
                    
    End If
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelContainer_Grava = SUCESSO

    Exit Function

Erro_RelContainer_Grava:

    RelContainer_Grava = gErr
     
    Select Case gErr
          
        Case 133122
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 133123, 133127
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELESTMPDISP", gErr, lNumIntRel, objProdutoMP.sCodigo, iSeq)
          
        Case 133124, 133125, 133128
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", gErr)
          
        Case 133126
        
        Case 133156, 133157
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO2", gErr, sProduto)
          
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179212)
     
    End Select
     
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    Exit Function

End Function

Function RelContainer_Acumula(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal dQuantKit2Est As Double, dQuantDispMP As Double, ByVal objProdutoMP As ClassProduto, ByVal lNumIntRel As Long) As Long
'dQuantKit2Est tem a quantidade necessaria da materia prima por unidade do produto vendido
'vai pesquisar o estoque disponivel e a quantidade em pedido do produto vendido
'vai pesquisar o quanto de mat prima foi requisitada e ainda nao produzido
            
Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 7) As Long
Dim dQuantPedida As Double
Dim dQuantOP As Double
Dim dQuantProd As Double
Dim lNumIntDocItemOP As Long
Dim dQuantidade As Double
Dim sUM As String
Dim dFator As Double
Dim dQuantReq As Double
Dim dQuantReqNaoUtilizada As Double
Dim dQuantDisp As Double
            
On Error GoTo Erro_RelContainer_Acumula

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 133129
    Next

    lErro = Comando_Executar(alComando(1), "SELECT SUM(QuantDispNossa + QuantReservada) FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND FilialEmpresa = ? AND Produto = ?", _
    dQuantDisp, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 133130
                
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133131

    dQuantDisp = dQuantDisp * dQuantKit2Est

    If dQuantKit2Est <> 0 Then

        lErro = Comando_Executar(alComando(2), "SELECT QuantPedida FROM ProdutosFilial WHERE FilialEmpresa = ? AND Produto = ?", _
        dQuantPedida, iFilialEmpresa, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133132
                    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133133

        dQuantPedida = dQuantPedida * dQuantKit2Est

        lErro = Comando_Executar(alComando(3), "SELECT Quantidade, QuantidadeProd, NumIntDoc FROM ItensOrdemProducao WHERE FilialEmpresa = ? AND Produto = ? AND (Situacao = 0 OR Situacao = 2)", _
        dQuantOP, dQuantProd, lNumIntDocItemOP, iFilialEmpresa, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133134
                    
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133135

        Do While lErro = AD_SQL_SUCESSO
        
            If dQuantOP > dQuantProd Then
            
                sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
            
                lErro = Comando_Executar(alComando(4), "SELECT Quantidade, SiglaUM FROM MovimentoEstoque WHERE FilialEmpresa = ? AND TipoMov = ? AND Produto = ? AND TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? AND NumIntDocEst = 0", _
                dQuantidade, sUM, iFilialEmpresa, MOV_EST_REQ_PRODUCAO, objProdutoMP.sCodigo, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, lNumIntDocItemOP)
                If lErro <> AD_SQL_SUCESSO Then gError 133136
            
                lErro = Comando_BuscarPrimeiro(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133137
            
                Do While lErro = AD_SQL_SUCESSO
            
                    'Obtém fator de conversão para unidade de medida de ItemPV
                    lErro = CF("UM_Conversao", objProdutoMP.iClasseUM, sUM, objProdutoMP.sSiglaUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 133138
        
                    dQuantReq = dQuantReq + dQuantidade * dFator
                    
                    lErro = Comando_BuscarProximo(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133139
                    
                Loop
                
                If (dQuantProd * dQuantKit2Est) < dQuantReq Then
                
                    dQuantReqNaoUtilizada = dQuantReqNaoUtilizada + dQuantReq - (dQuantProd * dQuantKit2Est)
                    
                End If
                                
            End If
            
            dQuantReq = 0
                        
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133140
                                
        Loop

    End If

    dQuantDispMP = dQuantDispMP + dQuantDisp + dQuantReqNaoUtilizada - dQuantPedida
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelContainer_Acumula = SUCESSO

    Exit Function

Erro_RelContainer_Acumula:

    RelContainer_Acumula = gErr
     
    Select Case gErr
          
        Case 133129
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 133130, 133131
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO2", gErr, sProduto)
          
        Case 133132, 133133
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", gErr, iFilialEmpresa, sProduto)
          
        Case 133134, 133135, 133140
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO4", gErr, iFilialEmpresa, sProduto)
          
        Case 133136, 133137, 133139
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE8", gErr, iFilialEmpresa, MOV_EST_REQ_PRODUCAO, objProdutoMP.sCodigo, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, lNumIntDocItemOP)
          
        Case 133138
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179213)
     
    End Select
     
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    Exit Function

End Function

Function RelMapaFat_Prepara(lNumIntRel As Long, ByVal iFilialEmpresa As Integer, sProdIni As String, sProdFim As String) As Long
'gera registros na tabela RelMapaFat para a execucao do mapa de faturamento

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim alComando(1 To 4) As Long
Dim sProduto As String
Dim dQuantDispMP As Double
Dim dQuantKit2Est As Double
Dim objProdutoMP As New ClassProduto

On Error GoTo Erro_RelMapaFat_Prepara

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 133243
    Next
    
    'Inicia a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 133230

    'obtem o nº automatico
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_NUMINTREL", lNumIntRel)
    If lErro <> SUCESSO Then gError 133231
    
    sProduto = String(STRING_PRODUTO, 0)
                        
    If Len(Trim(sProdIni)) = 0 And Len(Trim(sProdFim)) = 0 Then
        
        lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos, ProdutoCategoria WHERE Codigo = Produto AND UPPER(Categoria) = 'MP' And UPPER(Item) = 'SCAMPER' ORDER BY Codigo", sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133232
    
    ElseIf Len(Trim(sProdIni)) > 0 And Len(Trim(sProdFim)) > 0 Then

        lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos, ProdutoCategoria WHERE Codigo = Produto AND UPPER(Categoria) = 'MP' And UPPER(Item) = 'SCAMPER' AND Codigo >= ? AND Codigo <= ? ORDER BY Codigo", sProduto, sProdIni, sProdFim)
        If lErro <> AD_SQL_SUCESSO Then gError 133233

    ElseIf Len(Trim(sProdIni)) > 0 Then

        lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos, ProdutoCategoria WHERE Codigo = Produto AND UPPER(Categoria) = 'MP' And UPPER(Item) = 'SCAMPER' AND Codigo >= ? ORDER BY Codigo", sProduto, sProdIni)
        If lErro <> AD_SQL_SUCESSO Then gError 133234

    Else

        lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM Produtos, ProdutoCategoria WHERE Codigo = Produto AND UPPER(Categoria) = 'MP' And UPPER(Item) = 'SCAMPER' AND Codigo <= ? ORDER BY Codigo", sProduto, sProdFim)
        If lErro <> AD_SQL_SUCESSO Then gError 133235

    End If

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133236

    Do While lErro = AD_SQL_SUCESSO
    
        objProdutoMP.sCodigo = sProduto
    
        'Lê o produto no BD para obter UM de estoque
        lErro = CF("Produto_Le", objProdutoMP)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 133237

        If lErro = 28030 Then gError 133238
    
        dQuantKit2Est = 0
        dQuantDispMP = 0
        
        lErro = RelMapaFat_Busca_Arvore(iFilialEmpresa, sProduto, dQuantKit2Est, dQuantDispMP, objProdutoMP, lNumIntRel)
        If lErro <> SUCESSO Then gError 133239
            
        lErro = RelMapaFat_Grava(iFilialEmpresa, sProduto, 1, dQuantDispMP, objProdutoMP, lNumIntRel)
        If lErro <> SUCESSO Then gError 133240
            
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133241

    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 133242
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    RelMapaFat_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelMapaFat_Prepara:

    RelMapaFat_Prepara = gErr
     
    Select Case gErr
          
        Case 133230
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
          
        Case 133231, 133237, 133239, 133240
          
        Case 133232, 133233, 133234, 133235, 133236, 133241
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
          
        Case 133238
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProdutoMP.sCodigo)
          
        Case 133242
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
          
        Case 133243
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179214)
     
    End Select
     
    Call Transacao_Rollback
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    Exit Function

End Function

Private Function RelMapaFat_Busca_Arvore(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal dQuantKit3Est As Double, dQuantDispMP As Double, ByVal objProdutoMP As ClassProduto, ByVal lNumIntRel As Long) As Long

Dim lErro As Long
Dim sProdutoRaiz As String
Dim sUM As String
Dim dQuantKit As Double
Dim alComando(1 To 2) As Long
Dim dFator As Double
Dim iIndice As Integer
Dim dQuantKitEst As Double
Dim dQuantKit1 As Double
Dim dQuantKit1Est As Double
Dim dQuantKit2Est As Double
Dim objProduto As New ClassProduto
Dim objProduto1 As New ClassProduto
Dim sProduto1 As String

On Error GoTo Erro_RelMapaFat_Busca_Arvore

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 133244
    Next

    sProdutoRaiz = String(STRING_PRODUTO, 0)
    sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)

    lErro = Comando_Executar(alComando(1), "SELECT Kit.ProdutoRaiz, ProdutoKit.UnidadeMed, Quantidade FROM Kit, ProdutoKit WHERE Produto = ? AND Kit.ProdutoRaiz = ProdutoKit.ProdutoRaiz AND Kit.Situacao = 1 AND Nivel = 1", _
    sProdutoRaiz, sUM, dQuantKit, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 133245
                
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133246
        
    If lErro = AD_SQL_SEM_DADOS Then
    
        If dQuantKit3Est <> 0 Then
    
            lErro = RelMapaFat_Grava(iFilialEmpresa, sProduto, dQuantKit3Est, dQuantDispMP, objProdutoMP, lNumIntRel)
            If lErro <> SUCESSO Then gError 133247
    
        End If
    
    Else
        
        objProduto.sCodigo = sProduto
    
        'Lê o produto no BD para obter UM de estoque
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 133248

        If lErro = 28030 Then gError 133249
        
        Do While lErro = AD_SQL_SUCESSO
        
            'Obtém fator de conversão para unidade de medida de ItemPV
            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 133250
                
            dQuantKitEst = dQuantKit * dFator
            
            If dQuantKit3Est <> 0 Then dQuantKitEst = dQuantKitEst * dQuantKit3Est
            
            sProduto1 = sProdutoRaiz
            
            sProdutoRaiz = String(STRING_PRODUTO, 0)
            sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
            
            lErro = Comando_Executar(alComando(2), "SELECT ProdutoKit.ProdutoRaiz, ProdutoKit.UnidadeMed, ProdutoKit.Quantidade FROM Kit, ProdutoKit WHERE Produto = ? AND Kit.ProdutoRaiz = ProdutoKit.ProdutoRaiz AND Situacao = 1 And Nivel = 0", _
            sProdutoRaiz, sUM, dQuantKit1, sProduto1)
            If lErro <> AD_SQL_SUCESSO Then gError 133251
                        
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133252
            
            If lErro = AD_SQL_SUCESSO Then
            
                objProduto1.sCodigo = sProduto1
            
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto1)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 133253
        
                If lErro = 28030 Then gError 133254
            
                'Obtém fator de conversão para unidade de medida de ItemPV
                lErro = CF("UM_Conversao", objProduto1.iClasseUM, sUM, objProduto1.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 133255
            
                dQuantKit1Est = dQuantKit1 * dFator
                
                dQuantKit2Est = dQuantKitEst / dQuantKit1Est
                
                sProduto1 = sProdutoRaiz
            
                lErro = RelMapaFat_Busca_Arvore(iFilialEmpresa, sProduto1, dQuantKit2Est, dQuantDispMP, objProdutoMP, lNumIntRel)
                If lErro <> SUCESSO Then gError 133256
        
            End If
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133257
        
        Loop
    
    End If
    
    RelMapaFat_Busca_Arvore = SUCESSO

    Exit Function

Erro_RelMapaFat_Busca_Arvore:

    RelMapaFat_Busca_Arvore = gErr
     
    Select Case gErr
          
        Case 133244
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 133245, 133246, 133251, 133252, 133257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOKIT", gErr)

        Case 133247, 133248, 133250, 133253, 133255, 133256
        
        Case 133249, 133254
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179215)
     
    End Select
     
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    Exit Function

End Function

Private Function RelMapaFat_Grava(ByVal iFilialEmpresa As Integer, ByVal sProduto As String, ByVal dQuantKit2Est As Double, dQuantDispMP As Double, ByVal objProdutoMP As ClassProduto, ByVal lNumIntRel As Long) As Long
'dQuantKit2Est tem a quantidade necessaria da materia prima por unidade do produto vendido
'vai pesquisar o estoque disponivel e a quantidade em pedido do produto vendido
'vai pesquisar o quanto de mat prima foi requisitada e ainda nao produzido
            
Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 8) As Long
Dim dQuantPedida As Double
Dim dQuantOP As Double
Dim dQuantProd As Double
Dim lNumIntDocItemOP As Long
Dim dQuantidade As Double
Dim sUM As String
Dim dFator As Double
Dim dQuantReq As Double
Dim dQuantReqNaoUtilizada As Double
Dim dQuantDisp As Double
Dim dValorFaturado1 As Double
Dim dValorFaturado2 As Double
Dim dValorFaturado3 As Double
Dim dQuantFaturada1 As Double
Dim dQuantFaturada2 As Double
Dim dQuantFaturada3 As Double
Dim iMesAtual As Integer
Dim iAnoAtual As Integer

            
On Error GoTo Erro_RelMapaFat_Grava

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 133258
    Next

    lErro = Comando_Executar(alComando(1), "SELECT SUM(QuantDispNossa + QuantReservada) FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND FilialEmpresa = ? AND Produto = ?", _
    dQuantDisp, iFilialEmpresa, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 133259
                
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133260

    If dQuantKit2Est <> 0 Then

        dQuantDisp = dQuantDisp * dQuantKit2Est

        lErro = Comando_Executar(alComando(2), "SELECT QuantPedida FROM ProdutosFilial WHERE FilialEmpresa = ? AND Produto = ?", _
        dQuantPedida, iFilialEmpresa, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133261
                    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133262

        dQuantPedida = dQuantPedida * dQuantKit2Est

        lErro = Comando_Executar(alComando(3), "SELECT Quantidade, QuantidadeProd, NumIntDoc FROM ItensOrdemProducao WHERE FilialEmpresa = ? AND Produto = ? AND (Situacao = 0 OR Situacao = 2)", _
        dQuantOP, dQuantProd, lNumIntDocItemOP, iFilialEmpresa, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133263
                    
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133264

        Do While lErro = AD_SQL_SUCESSO
        
            If dQuantOP > dQuantProd Then
            
                sUM = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)
            
                lErro = Comando_Executar(alComando(4), "SELECT Quantidade, SiglaUM FROM MovimentoEstoque WHERE FilialEmpresa = ? AND TipoMov = ? AND Produto = ? AND TipoNumIntDocOrigem = ? AND NumIntDocOrigem = ? AND NumIntDocEst = 0", _
                dQuantidade, sUM, iFilialEmpresa, MOV_EST_REQ_PRODUCAO, objProdutoMP.sCodigo, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, lNumIntDocItemOP)
                If lErro <> AD_SQL_SUCESSO Then gError 133265
            
                lErro = Comando_BuscarPrimeiro(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133266
            
                Do While lErro = AD_SQL_SUCESSO
            
                    'Obtém fator de conversão para unidade de medida de ItemPV
                    lErro = CF("UM_Conversao", objProdutoMP.iClasseUM, sUM, objProdutoMP.sSiglaUMEstoque, dFator)
                    If lErro <> SUCESSO Then gError 133267
        
                    dQuantReq = dQuantReq + dQuantidade * dFator
                    
                    lErro = Comando_BuscarProximo(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133268
                    
                Loop
                
                If (dQuantProd * dQuantKit2Est) < dQuantReq Then
                
                    dQuantReqNaoUtilizada = dQuantReqNaoUtilizada + dQuantReq - (dQuantProd * dQuantKit2Est)
                    
                End If
                                
            End If
            
            dQuantReq = 0
                        
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133269
                                
        Loop

    End If

    dQuantDispMP = dQuantDispMP + dQuantDisp + dQuantReqNaoUtilizada - dQuantPedida


    iMesAtual = Month(gdtDataHoje)
    iAnoAtual = Year(gdtDataHoje)
    
    If iMesAtual > 2 Then
    
        lErro = Comando_Executar(alComando(8), "SELECT QuantFaturada" & CStr(iMesAtual - 2) & ",  ValorFaturado" & CStr(iMesAtual - 2) & ", QuantFaturada" & CStr(iMesAtual - 1) & ", ValorFaturado" & CStr(iMesAtual - 1) & ", QuantFaturada" & CStr(iMesAtual) & ", ValorFaturado" & CStr(iMesAtual) & " FROM SldMesFAT WHERE FilialEmpresa = ? AND Ano=? AND Produto = ?", _
        dQuantFaturada1, dValorFaturado1, dQuantFaturada2, dValorFaturado2, dQuantFaturada3, dValorFaturado3, iFilialEmpresa, iAnoAtual, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133270
    
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133271
    
    ElseIf iMesAtual = 2 Then

        lErro = Comando_Executar(alComando(8), "SELECT QuantFaturada" & CStr(iMesAtual - 1) & ", ValorFaturado" & CStr(iMesAtual - 1) & ", QuantFaturada" & CStr(iMesAtual) & ", ValorFaturado" & CStr(iMesAtual) & " FROM SldMesFAT WHERE FilialEmpresa = ? AND Ano=? AND Produto = ?", _
        dQuantFaturada2, dValorFaturado2, dQuantFaturada3, dValorFaturado3, iFilialEmpresa, iAnoAtual, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133272

        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133273

        lErro = Comando_Executar(alComando(8), "SELECT QuantFaturada12, ValorFaturado12 FROM SldMesFAT WHERE FilialEmpresa = ? AND Ano=? AND Produto = ?", _
        dQuantFaturada1, dValorFaturado1, iFilialEmpresa, iAnoAtual - 1, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133274

        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133275

    Else

        lErro = Comando_Executar(alComando(8), "SELECT QuantFaturada" & CStr(iMesAtual) & ", ValorFaturado" & CStr(iMesAtual) & " FROM SldMesFAT WHERE FilialEmpresa = ? AND Ano=? AND Produto = ?", _
        dQuantFaturada3, dValorFaturado3, iFilialEmpresa, iAnoAtual, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133276

        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133277

        lErro = Comando_Executar(alComando(8), "SELECT QuantFaturada11, ValorFaturado11, QuantFaturada12, ValorFaturado12 FROM SldMesFAT WHERE FilialEmpresa = ? AND Ano=? AND Produto = ?", _
        dQuantFaturada1, dValorFaturado1, dQuantFaturada2, dValorFaturado2, iFilialEmpresa, iAnoAtual - 1, sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 133278

        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 133279

    End If

    lErro = Comando_Executar(alComando(5), "INSERT INTO RelMapaFat (NumIntRel, Produto, ProdutoMP, ValorFaturado1, ValorFaturado2, ValorFaturado3, QuantDisp, QuantFaturada1, QuantFaturada2, QuantFaturada3, QuantDispMP) VALUES (?,?,?,?,?,?,?,?,?,?,?) ", _
    lNumIntRel, sProduto, objProdutoMP.sCodigo, dValorFaturado1, dValorFaturado2, dValorFaturado3, dQuantDisp, dQuantFaturada1, dQuantFaturada2, dQuantFaturada3, dQuantDispMP)
    If lErro <> AD_SQL_SUCESSO Then gError 133280
    
    'se estiver gravando a materia prima
    If objProdutoMP.sCodigo = sProduto Then
    
        lErro = Comando_Executar(alComando(6), "SELECT SUM(ValorFaturado1), Sum(ValorFaturado2), Sum(ValorFaturado3), Sum(Abs(QuantDisp)), Sum(QuantFaturada1), Sum(QuantFaturada2), Sum(QuantFaturada3) FROM RelMapaFat WHERE ProdutoMP = ? And NumIntRel = ?", _
        dValorFaturado1, dValorFaturado2, dValorFaturado3, dQuantDisp, dQuantFaturada1, dQuantFaturada2, dQuantFaturada3, sProduto, lNumIntRel)
        If lErro <> AD_SQL_SUCESSO Then gError 133289
    
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO Then gError 133290
        
        If dQuantDisp = 0 And dValorFaturado1 = 0 And dValorFaturado2 = 0 And dValorFaturado3 = 0 And dQuantFaturada1 = 0 And dQuantFaturada2 = 0 And dQuantFaturada3 = 0 Then
            
            lErro = Comando_Executar(alComando(7), "DELETE FROM RelMapaFat WHERE NumIntRel = ? And ProdutoMP = ?", lNumIntRel, sProduto)
            If lErro <> AD_SQL_SUCESSO Then gError 133291
    
        End If
    
    End If
    
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelMapaFat_Grava = SUCESSO

    Exit Function

Erro_RelMapaFat_Grava:

    RelMapaFat_Grava = gErr
     
    Select Case gErr
          
        Case 133258
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
          
        Case 133259, 133260
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTOQUEPRODUTO2", gErr, sProduto)
          
        Case 133261, 133262
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOSFILIAL", gErr, iFilialEmpresa, sProduto)
          
        Case 133263, 133264, 133269
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSORDEMPRODUCAO4", gErr, iFilialEmpresa, sProduto)
          
        Case 133265, 133266, 133268
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE8", gErr, iFilialEmpresa, MOV_EST_REQ_PRODUCAO, objProdutoMP.sCodigo, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP, lNumIntDocItemOP)
          
        Case 133267
        
        Case 133270 To 133273, 133276, 133277
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFAT", gErr, iAnoAtual, iFilialEmpresa, sProduto)

        Case 133274, 133275, 133278, 133279
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFAT", gErr, iAnoAtual - 1, iFilialEmpresa, sProduto)

        Case 133280
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELMAPAFAT", gErr, lNumIntRel, sProduto, objProdutoMP.sCodigo)
          
        Case 133289, 133290
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELMAPAFAT", gErr, lNumIntRel, objProdutoMP.sCodigo)
          
        Case 133291
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELMAPAFAT", gErr, lNumIntRel, objProdutoMP.sCodigo)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179216)
     
    End Select
     
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    Exit Function

End Function


