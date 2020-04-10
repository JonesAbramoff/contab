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

Function Customizacoes_MovEstoque(ByVal objMovEst As ClassMovEstoque) As Long
'Lê Almoxarifado Padrão do Produto da iFilialEmpresa

Dim lErro As Long
Dim iAlmoxarifado As Integer
Dim lComando As Long
Dim lComando2 As Long

On Error GoTo Erro_Customizacoes_MovEstoque

    If objMovEst.iFilialEmpresa = 52 Then objMovEst.iFilialEmpresa = 51
    
    If objMovEst.iFilialEmpresa = 2 Then objMovEst.iFilialEmpresa = 1

    Customizacoes_MovEstoque = SUCESSO

    Exit Function

Erro_Customizacoes_MovEstoque:

    Customizacoes_MovEstoque = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 175935)

    End Select

    Exit Function

End Function

Function TabelaPrecoItem_Le1(ByVal objTabelaPrecoItem As ClassTabelaPrecoItem, Optional dtDataLeitura As Date = DATA_NULA) As Long
'Lê a tabela de preço a partir da chave fornecida em objTabelaPreco. Retorna os dados em objTabelaPreco

'*** ALTERADO POR LUIZ G.F.NOGUEIRA EM 02/04/2002 ***
'dtDataLeitura indica que a leitura deve ser feita utilizando a data passada na variável. Seu valor padrão é DATA_NULA. Isso significa
'que se o programador não passar essa variável, a leitura será feita utilizando a data atual (Date)
'****************************************************

Dim lErro As Long
Dim lComando As Long
Dim tTabelaPrecoItem As typeTabelaPrecoItem

On Error GoTo Erro_TabelaPrecoItem_Le1

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 126781

    'Pesquisa no BD Ítens de Tabela de Preço
    lErro = Comando_Executar(lComando, "SELECT Preco, DataVigencia FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto = ? ORDER BY DataVigencia DESC", tTabelaPrecoItem.dPreco, tTabelaPrecoItem.dtDataVigencia, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 126782

    'Verifica se Ítem existe
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 126783

    'Não encontrou o Ítem
    If lErro = AD_SQL_SEM_DADOS Then gError 126784

    'Preenche objTabelaPrecoItem
    objTabelaPrecoItem.dPreco = tTabelaPrecoItem.dPreco
    objTabelaPrecoItem.dtDataVigencia = tTabelaPrecoItem.dtDataVigencia

    'Fechamento comando
    Call Comando_Fechar(lComando)

    TabelaPrecoItem_Le1 = SUCESSO

    Exit Function

Erro_TabelaPrecoItem_Le1:

    TabelaPrecoItem_Le1 = gErr

    Select Case gErr

        Case 126781
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 126782, 126783
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS1", gErr, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.sCodProduto)

        Case 126784

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175940)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Customizacao_FilialEmpresa(iFilialEmpresa As Integer) As Long
'Altera a filialempresa passada como parametro

Dim lErro As Long
Dim iAlmoxarifado As Integer
Dim lComando As Long
Dim lComando2 As Long

On Error GoTo Erro_Customizacao_FilialEmpresa

    If iFilialEmpresa = 52 Then iFilialEmpresa = 51
    
    If iFilialEmpresa = 2 Then iFilialEmpresa = 1

    Customizacao_FilialEmpresa = SUCESSO

    Exit Function

Erro_Customizacao_FilialEmpresa:

    Customizacao_FilialEmpresa = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 175941)

    End Select

    Exit Function

End Function

Function Produto_Testa_Quant_Requerida(objItemOP As ClassItemOP, sProdMaiorReq As String) As Long
'verifica se  a quantidade dos insumos (kit) que serao necessarios para produzir a quantidade passada é maior que a quantidade requerida.

Dim colItensMovEstoque As New colItensMovEstoque
Dim colItensMovEstoque1 As New colItensMovEstoque
Dim colItensMovEstoque2 As New colItensMovEstoque
Dim lErro As Long
Dim dFator As Double
Dim objItemMovEst As ClassItemMovEstoque
Dim objProduto As ClassProduto
Dim objItemMovEst1 As ClassItemMovEstoque
Dim objItemOP1 As New ClassItemOP
Dim sProdutoMascarado As String
Dim sProdutoMascarado1 As String
Dim dQuantProdAgora As Double
Dim iIndice As Integer
Dim objItemMovEstAgora As ClassItemMovEstoque
Dim objItemMovEstProd As ClassItemMovEstoque

On Error GoTo Erro_Produto_Testa_Quant_Requerida

    objItemOP1.iFilialEmpresa = objItemOP.iFilialEmpresa
    objItemOP1.sCodigo = objItemOP.sCodigo
    objItemOP1.sProduto = objItemOP.sProduto

    lErro = CF("ItemOP_Le", objItemOP1)
    If lErro <> SUCESSO And lErro <> 34711 Then gError 126600

    If lErro = 34711 Then gError 126601
    
    'quantidade que esta sendo produzida agora
    dQuantProdAgora = objItemOP.dQuantidade
    
    'quantidade que esta produzindo agora + quantidade ja produzida
    objItemOP.dQuantidade = dQuantProdAgora + objItemOP1.dQuantidadeProd
    objItemOP.sVersao = objItemOP1.sVersao

    lErro = CF("ProdutoKit_Le_QuantInsumos", objItemOP, colItensMovEstoque)
    If lErro <> SUCESSO Then gError 126609


    objItemOP.dQuantidade = dQuantProdAgora

    lErro = CF("ProdutoKit_Le_QuantInsumos", objItemOP, colItensMovEstoque1)
    If lErro <> SUCESSO Then gError 126609

    objItemOP.dQuantidade = objItemOP1.dQuantidadeProd

    lErro = CF("ProdutoKit_Le_QuantInsumos", objItemOP, colItensMovEstoque2)
    If lErro <> SUCESSO Then gError 126609

    For Each objItemMovEst In colItensMovEstoque

        iIndice = iIndice + 1
        
        Set objItemMovEstAgora = colItensMovEstoque1.Item(iIndice)
        Set objItemMovEstProd = colItensMovEstoque2.Item(iIndice)

        Set objProduto = New ClassProduto
            
        objProduto.sCodigo = objItemMovEst.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 126610
        
        If lErro <> SUCESSO Then gError 126611

        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemMovEst.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 126605

        objItemMovEst.dQuantidadeEst = objItemMovEst.dQuantidade * dFator

        Set objItemMovEst1 = New ClassItemMovEstoque

        objItemMovEst1.sProduto = objItemMovEst.sProduto
        objItemMovEst1.sOPCodigo = objItemOP.sCodigo
        objItemMovEst1.iFilialEmpresa = objItemOP.iFilialEmpresa
        objItemMovEst1.sProdutoOP = objItemOP.sProduto

        lErro = CF("ReqProducao_Total_Produto_OP", objProduto, objItemMovEst1)
        If lErro <> SUCESSO Then gError 126612

        'se a quatindade de insumo requerida for menor que a que deveria ser usada de acordo com o kit
        If objItemMovEst1.dQuantidadeEst < objItemMovEst.dQuantidadeEst - QTDE_ESTOQUE_DELTA2 Then
            
            If Len(sProdMaiorReq) = 0 Then sProdMaiorReq = Chr(10)
            
            sProdutoMascarado = String(STRING_PRODUTO, 0)
    
            lErro = Mascara_RetornaProdutoTela(objItemOP.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 126615
            
            sProdutoMascarado1 = String(STRING_PRODUTO, 0)
    
            lErro = Mascara_RetornaProdutoTela(objItemMovEst.sProduto, sProdutoMascarado1)
            If lErro <> SUCESSO Then gError 126616
            
            objItemMovEstAgora.dQuantidadeEst = objItemMovEstAgora.dQuantidade * dFator
            objItemMovEstProd.dQuantidadeEst = objItemMovEstProd.dQuantidade * dFator
            
            
            sProdMaiorReq = sProdMaiorReq & "Produto: " & sProdutoMascarado & " Insumo: " & sProdutoMascarado1 & " Quant. Req.: " & objItemMovEst1.dQuantidadeEst - objItemMovEstProd.dQuantidadeEst & " Quant. Kit.:" & objItemMovEstAgora.dQuantidadeEst & Chr(10) & _
            "Quant.Req.Outras Produções: " & objItemMovEstProd.dQuantidadeEst & Chr(10)
            
        End If

    Next

    Produto_Testa_Quant_Requerida = SUCESSO

    Exit Function

Erro_Produto_Testa_Quant_Requerida:

    Produto_Testa_Quant_Requerida = gErr

    Select Case gErr

        Case 126600, 126609, 126610, 126612, 126615, 126616
        
        Case 126601
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP1.sProduto, objItemOP1.sCodigo)

        Case 126611
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 126608)

    End Select

    Exit Function

End Function

Function Produto_Quant_Requerida_Superior(objMovEstoque As ClassMovEstoque, sReqMaiorOP As String) As Long
'verifica se  a quantidade requerida dos insumos supera o que esta especificado nos kits

Dim colItensMovEstoque As colItensMovEstoque
Dim lErro As Long
Dim dFator As Double
Dim objItemMovEst As ClassItemMovEstoque
Dim objItemMovEst1 As ClassItemMovEstoque
Dim objItemMovEstInsumos As ClassItemMovEstoque
Dim objItemMovEstReq As ClassItemMovEstoque
Dim objProduto As ClassProduto
Dim objItemOP As ClassItemOP
Dim sProdutoMascarado As String
Dim sProdutoMascarado1 As String
Dim colItens As New Collection
Dim iAchou As Integer

On Error GoTo Erro_Produto_Quant_Requerida_Superior


    'pesquisa na colecao colItens se o mesmo produto, produtoOP, sOPCodigo ja esta cadastrado
    'se ja estiver, converte para unidade de estoque e armazena a quantidade
    'senao cria um novo item em colItens
    For Each objItemMovEst In objMovEstoque.colItens
    
        iAchou = 0
        
        For Each objItemMovEst1 In colItens
            
            If objItemMovEst1.sProduto = objItemMovEst.sProduto And _
            objItemMovEst1.sOPCodigo = objItemMovEst.sOPCodigo And _
            objItemMovEst1.sProdutoOP = objItemMovEst.sProdutoOP Then
                            
                Set objProduto = New ClassProduto

                objProduto.sCodigo = objItemMovEst.sProduto
        
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 178253
                
                If lErro <> SUCESSO Then gError 178251
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemMovEst.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 178254
        
                objItemMovEst1.dQuantidadeEst = objItemMovEst1.dQuantidadeEst + (objItemMovEst.dQuantidade * dFator)
                        
                iAchou = 1
                
                Exit For
                
            End If
                
        Next
            
        If iAchou = 0 Then
        
            Set objProduto = New ClassProduto

            objProduto.sCodigo = objItemMovEst.sProduto
    
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 178255
            
            If lErro <> SUCESSO Then gError 178252
    
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemMovEst.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 178256
        
            objItemMovEst.dQuantidadeEst = objItemMovEst.dQuantidade * dFator
        
            colItens.Add objItemMovEst
            
        End If

    Next
    
    For Each objItemMovEst In colItens
    
        Set objItemOP = New ClassItemOP
    
        objItemOP.iFilialEmpresa = objMovEstoque.iFilialEmpresa
        objItemOP.sCodigo = objItemMovEst.sOPCodigo
        objItemOP.sProduto = objItemMovEst.sProdutoOP
    
        lErro = CF("ItemOP_Le", objItemOP)
        If lErro <> SUCESSO And lErro <> 34711 Then gError 126618
    
        If lErro = 34711 Then gError 126619
        
        Set colItensMovEstoque = New colItensMovEstoque
        
        lErro = CF("ProdutoKit_Le_QuantInsumos", objItemOP, colItensMovEstoque)
        If lErro <> SUCESSO Then gError 126620

        For Each objItemMovEstInsumos In colItensMovEstoque

            If objItemMovEstInsumos.sProduto = objItemMovEst.sProduto Then

                Set objProduto = New ClassProduto

                objProduto.sCodigo = objItemMovEst.sProduto
        
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 126621
                
                If lErro <> SUCESSO Then gError 126622
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemMovEstInsumos.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 126624
        
                objItemMovEstInsumos.dQuantidadeEst = objItemMovEstInsumos.dQuantidade * dFator
        
                Set objItemMovEstReq = New ClassItemMovEstoque
        
                objItemMovEstReq.sProduto = objItemMovEst.sProduto
                objItemMovEstReq.sOPCodigo = objItemOP.sCodigo
                objItemMovEstReq.iFilialEmpresa = objItemOP.iFilialEmpresa
                objItemMovEstReq.sProdutoOP = objItemOP.sProduto
        
                lErro = CF("ReqProducao_Total_Produto_OP", objProduto, objItemMovEstReq)
                If lErro <> SUCESSO Then gError 126625
                
                objItemMovEst.dQuantidadeEst = StrParaDbl(Formata_Estoque(objItemMovEst.dQuantidadeEst))
                objItemMovEstReq.dQuantidadeEst = StrParaDbl(Formata_Estoque(objItemMovEstReq.dQuantidadeEst))
                objItemMovEstInsumos.dQuantidadeEst = StrParaDbl(Formata_Estoque(objItemMovEstInsumos.dQuantidadeEst))
        
                'se a quatindade requerida for superior a quantidade que seria necessaria
                If (objItemMovEst.dQuantidadeEst + objItemMovEstReq.dQuantidadeEst) > objItemMovEstInsumos.dQuantidadeEst + QTDE_ESTOQUE_DELTA2 Then
                
                    'If Len(sReqMaiorOP) = 0 Then sReqMaiorOP = Chr(10)
                    If Len(sReqMaiorOP) <> 0 Then sReqMaiorOP = sReqMaiorOP & vbNewLine
                    
                    sProdutoMascarado = String(STRING_PRODUTO, 0)
            
                    lErro = Mascara_RetornaProdutoTela(objItemMovEst.sProduto, sProdutoMascarado)
                    If lErro <> SUCESSO Then gError 126626
                    
                    sReqMaiorOP = sReqMaiorOP & "Prod:" & Trim(sProdutoMascarado) & " OP:" & objItemOP.sCodigo & " Qtd.Req:" & objItemMovEst.dQuantidadeEst + objItemMovEstReq.dQuantidadeEst & " Qtd.Kit:" & objItemMovEstInsumos.dQuantidadeEst
        
                End If
    
            End If
    
        Next

    Next

    Produto_Quant_Requerida_Superior = SUCESSO

    Exit Function

Erro_Produto_Quant_Requerida_Superior:

    Produto_Quant_Requerida_Superior = gErr

    Select Case gErr

        Case 126618, 126620, 126621, 126623 To 126626, 178253 To 278256
        
        Case 126619
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PARTICIPA_OP", gErr, objItemOP.sProduto, objItemOP.sCodigo)

        Case 126622, 178251, 178252
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 126608)

    End Select

    Exit Function

End Function


