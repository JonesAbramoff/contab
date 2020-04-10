VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Distribuicao_ProcessaAux(ByVal objCT As Object) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dQuantidade As Double
Dim sUM As String
Dim dFator As Double
Dim objItemNF As ClassItemNF
Dim dFatorAux As Double
Dim objNF As ClassNFiscal
Dim sProdutoMascarado As String
Dim iLinha As Integer

On Error GoTo Erro_Distribuicao_ProcessaAux

    iLinha = objCT.objTela.objGridItens.objGrid.Row

    'Remove as alocações do grid de alocações
    Call objCT.Remove_Distribuicao(iLinha)
    
    Set objNF = objCT.objTela.gobjNFiscal
    Set objItemNF = objNF.ColItensNF.Item(iLinha)
    
    If Not (objNF.ColItensNF.Item(iLinha).objInfoUsu Is Nothing) Then
        dQuantidade = StrParaDbl(objCT.objTela.objGridItens.objGrid.TextMatrix(iLinha, objCT.objTela.iGrid_Quantidade_Col))
        dFatorAux = dQuantidade / objItemNF.dQuantidade
    End If
    
    lErro = CF("MovEstoque_ItemNF_Customizado", objItemNF, objNF)
    If lErro <> SUCESSO Then gError 197727
    
    If Not objItemNF Is Nothing Then

        'Formata o Produto
        If objNF.ColItensNF.Item(iLinha).objInfoUsu Is Nothing Then
            sProdutoMascarado = objCT.objTela.objGridItens.objGrid.TextMatrix(iLinha, objCT.objTela.iGrid_Produto_Col)
            lErro = CF("Produto_Formata", sProdutoMascarado, sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 89502
        Else
            sProduto = objItemNF.sProduto
            lErro = Mascara_RetornaProdutoTela(sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 89502
        End If
    
        objProduto.sCodigo = sProduto
    
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 89503
        
        'se o produto não está cadastrado ==> erro
        If lErro <> SUCESSO Then gError 89504
    
        'recolhe a quantidade do grid de itens
        If objNF.ColItensNF.Item(iLinha).objInfoUsu Is Nothing Then
            dQuantidade = StrParaDbl(objCT.objTela.objGridItens.objGrid.TextMatrix(iLinha, objCT.objTela.iGrid_Quantidade_Col))
        Else
            dQuantidade = objItemNF.dQuantidade * dFatorAux
            objItemNF.dQuantidade = dQuantidade
            objNF.ColItensNF.Item(iLinha).dQuantidade = objNF.ColItensNF.Item(iLinha).dQuantidade * dFatorAux
        End If
        'Verifica se o produto tem o controle de estoque <> PRODUTO_CONTROLE_SEM_ESTOQUE
        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 Then
        
            'recolhe a UM do ItemNF
            If objNF.ColItensNF.Item(iLinha).objInfoUsu Is Nothing Then
                sUM = objCT.objTela.objGridItens.objGrid.TextMatrix(iLinha, objCT.objTela.iGrid_UnidadeMed_Col)
            Else
                sUM = objItemNF.sUnidadeMed
            End If
    
            'Faz a conversão da UM da tela para a UM de estoque
            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 89506
        
            'Converte a quantidade para a UM de estoque
            dQuantidade = dQuantidade * dFator
    
            'Tenta Alocar o produto no Almoxarifado padrão
            lErro = objCT.AlocaAlmoxarifadoPradrao(iLinha, dQuantidade, objProduto)
            If lErro <> SUCESSO Then gError 89505
            
            If Not (objNF.ColItensNF.Item(iLinha).objInfoUsu Is Nothing) Then
                objCT.objGridDist.objGrid.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_ProdutoAlmoxDist_Col) = sProdutoMascarado
            End If
        
        End If
        
    End If

    Distribuicao_ProcessaAux = SUCESSO

    Exit Function

Erro_Distribuicao_ProcessaAux:

    Distribuicao_ProcessaAux = gErr

    Select Case gErr

        Case 89502, 89503, 89505

        Case 89504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151572)

    End Select

    Exit Function

End Function

Public Function Preenche_GridDistribuicaoAux(ByVal objCT As Object, objNFiscal As Object) As Long
'Preenche o Grid com as Alocações da Nota Fiscal

Dim objItemAloc As Object
Dim iIndice As Integer
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantidade As Double
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim iContador As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoMascarado As String

On Error GoTo Erro_Preenche_GridDistribuicaoAux
    
    'Limpa o grid de alocações
    Call Grid_Limpa(objCT.objGridDist)

    'Para cada item da NF
    For iIndice = 1 To objNFiscal.ColItensNF.Count
    
        Set objItemNF = objNFiscal.ColItensNF.Item(iIndice)
                
        lErro = CF("MovEstoque_ItemNF_Customizado", objItemNF, objNFiscal)
        If lErro <> SUCESSO Then gError 197727
        
        If Not objItemNF Is Nothing Then
                    
            'Para cada alocação do Item de NF
            For Each objItemAloc In objItemNF.ColAlocacoes
    
                objCT.objGridDist.iLinhasExistentes = objCT.objGridDist.iLinhasExistentes + 1
            
                objProduto.sCodigo = objItemNF.sProduto
                
                lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
                
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 51081
                If lErro <> SUCESSO Then gError 51082
            
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 51083
                
                'Coloca os dados da alocação na tela
                objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_ItemNFDist_Col) = objNFiscal.ColItensNF(iIndice).iItem
                objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_ProdutoAlmoxDist_Col) = sProdutoMascarado 'objTela.objGridItens.objGrid.TextMatrix(objNFiscal.ColItensNF(iIndice).iItem, objTela.iGrid_Produto_Col)
                objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_AlmoxDist_Col) = objItemAloc.sAlmoxarifado
                objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_UMDist_Col) = objProduto.sSiglaUMEstoque
                objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_QuantDist_Col) = Formata_Estoque(objItemAloc.dQuantidade)
                dQuantidade = StrParaDbl(objCT.objTela.objGridItens.objGrid.TextMatrix(iIndice, objCT.objTela.iGrid_Quantidade_Col)) * dFator
                
                objCT.GridDist.TextMatrix(objCT.objGridDist.iLinhasExistentes, objCT.iGrid_QuantItemNFDist_Col) = Formata_Estoque(dQuantidade)
            Next
    
            
            If objNFiscal.ColItensNF(iIndice).iPossuiGrade Then Call objCT.Atualiza_Grid_Distribuicao(objNFiscal.ColItensNF(iIndice))

        End If

    Next

    Preenche_GridDistribuicaoAux = SUCESSO

    Exit Function

Erro_Preenche_GridDistribuicaoAux:

    Preenche_GridDistribuicaoAux = gErr
    
    Select Case gErr
    
        Case 51081, 51083, 197727
        
        Case 51082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objCT.ProdutoAlmoxDist.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151568)
            
    End Select
    
    Exit Function

End Function

Public Function Distribuicao_Obtem_Produto_Cust(ByVal objNF As ClassNFiscal, ByVal objItemNF As ClassItemNF, sProdutoMascarado As String) As Long

Dim lErro As Long

On Error GoTo Erro_Distribuicao_Obtem_Produto_Cust:

        lErro = CF("MovEstoque_ItemNF_Customizado", objItemNF, objNF)
        If lErro <> SUCESSO Then gError 197727
        
        If Not objItemNF Is Nothing Then
        
            lErro = Mascara_RetornaProdutoTela(objItemNF.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 197727
        
        End If
    
    Distribuicao_Obtem_Produto_Cust = SUCESSO

    Exit Function

Erro_Distribuicao_Obtem_Produto_Cust:

    Distribuicao_Obtem_Produto_Cust = gErr
    
    Select Case gErr
    
        Case 197727
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151568)
            
    End Select
    
    Exit Function

End Function
