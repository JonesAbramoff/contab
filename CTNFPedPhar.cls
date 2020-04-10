VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFPedPhar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function NFP_GridItens_InicializaColunas(ByVal objCT As Object) As Long

    'Colunas da Grid
    objCT.iGrid_Produto_Col = 1
    objCT.iGrid_DescProduto_Col = 2
    objCT.iGrid_UnidadeMed_Col = 3
    objCT.iGrid_Quantidade_Col = 4
    objCT.gobjInfoUsu.iGrid_ProdutoKit_Col = 5
    objCT.gobjInfoUsu.iGrid_UMKit_Col = 6
'    objCT.gobjInfoUsu.iGrid_VersaoKit_Col = 6
    objCT.gobjInfoUsu.iGrid_QuantKit_Col = 7
    objCT.iGrid_ValorUnitario_Col = 8
    objCT.iGrid_PercDesc_Col = 9
    objCT.iGrid_Desconto_Col = 10
    objCT.iGrid_PrecoTotal_Col = 11
    objCT.iGrid_PrecoTotalB_Col = 12
    objCT.iGrid_DataEntrega_Col = 13
    
    NFP_GridItens_InicializaColunas = SUCESSO
    
End Function

Public Function NFP_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Pack")
    objGridInt.colColuna.Add ("UM Kit")
'    objGridInt.colColuna.Add ("Versao")
    objGridInt.colColuna.Add ("Quant.Pack")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Data Entrega")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoItem.Name)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)
    objGridInt.colCampo.Add (ProdutoKit(objCT).Name)
    objGridInt.colCampo.Add (UMKit(objCT).Name)
'    objGridInt.colCampo.Add (VersaoKit(objCT).Name)
    objGridInt.colCampo.Add (QuantKit(objCT).Name)
    objGridInt.colCampo.Add (objCT.PrecoUnitario.Name)
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (objCT.Desconto.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotal.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotalB.Name)
    objGridInt.colCampo.Add (objCT.DataEntrega.Name)

    Call NFP_GridItens_InicializaColunas(objCT)

    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    objCT.GridItens.ColWidth(0) = 500

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFP_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Function NFP_Preenche_GridItens(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal iLinha As Integer) As Long

Dim objItemNFPhar As ClassItemNF
Dim sProdutoEnxuto As String
Dim dValorDesconto As Double
Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_NFP_Preenche_GridItens

    'permite que a tributacao reflita a inclusao de uma linha no grid
    lErro = objCT.gobjTribTab.Inclusao_Item_Grid(iLinha, objItemNF.sProduto)
    If lErro <> SUCESSO Then gError 197713

    If Not (objItemNF.objInfoUsu Is Nothing) Then
    
        Set objItemNFPhar = objItemNF.objInfoUsu

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNFPhar.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 197693

        objCT.Produto.PromptInclude = False
        objCT.Produto.Text = sProdutoEnxuto
        objCT.Produto.PromptInclude = True

        objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_ProdutoKit_Col) = objCT.Produto.Text
        objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_UMKit_Col) = objItemNFPhar.sUnidadeMed
 '       objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_VersaoKit_Col) = objItemNFPhar.sVersaoKit
        objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_QuantKit_Col) = Formata_Estoque(objItemNFPhar.dQuantidade)
    
    End If

    dValorDesconto = objItemNF.dValorDesconto

    dValorTotal = (objItemNF.dPrecoUnitario * objItemNF.dQuantidade) - dValorDesconto

    objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Desconto_Col) = Format(dValorDesconto, "Standard")
    objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_PrecoTotal_Col) = Format(dValorTotal, "Standard")

''*** incluidos p/tratamento de tributacao *******************************
    lErro = objCT.gobjTribTab.Alteracao_Item_Grid(iLinha)
    If lErro <> SUCESSO Then gError 197712
'*** fim tributacao

    NFP_Preenche_GridItens = SUCESSO
   
    Exit Function

Erro_NFP_Preenche_GridItens:

    NFP_Preenche_GridItens = gErr
    
    Select Case gErr

        Case 197693, 197712, 197713
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197694)
    
    End Select
    
    Exit Function
    
End Function

Public Function NFP_Transfere_Dados_Pedido_NFiscal(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal objItemPV As ClassItemPedido, ByVal dQuantidade As Double, ByVal objNFiscal As ClassNFiscal) As Long
    
Dim objItemNFPhar As New ClassItemNF
Dim objKit As New ClassKit
Dim objProdutoKit As ClassProdutoKit
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim lErro As Long
Dim objItemNF1 As ClassItemNF
Dim dValorTotalKit As Double
Dim dValorItemPedido As Double
Dim dValorItensKit As Double
Dim dPrecoUnitario As Double
Dim dDifAcerto As Double

On Error GoTo Erro_NFP_Transfere_Dados_Pedido_NFiscal

    objKit.sProdutoRaiz = objItemPV.sProduto

    'Le a Versao Padrao para o Kit
    lErro = CF("Kit_Le_Padrao", objKit)
    If lErro <> SUCESSO And lErro <> 106304 Then gError 197695

    'Se encontrou
    If lErro = SUCESSO Then

        objNFiscal.ColItensNF.Remove (objNFiscal.ColItensNF.Count)

        lErro = CF("Kit_Le_Componentes", objKit)
        If lErro <> SUCESSO And lErro <> 21831 Then gError 197696
        
        iIndice = 0
        
        '####################################
        'Incluído por Wagner
        For Each objProdutoKit In objKit.colComponentes
        
            If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
                dValorTotalKit = dValorTotalKit + objProdutoKit.dCustoStandard
            End If
        
        Next
        
        dValorItensKit = 0
        dValorItemPedido = objItemNF.dPrecoUnitario * objItemNF.dQuantidade
        '####################################
        
        For Each objProdutoKit In objKit.colComponentes
        
            If objProdutoKit.iNivel <> KIT_NIVEL_RAIZ Then
            
                iIndice = iIndice + 1
            
                'Guarda o código formatado do produto em objProduto
                objProduto.sCodigo = objProdutoKit.sProduto
                
                'Lê o produto a partir do código passado em objProduto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 197697
            
                'Se o produto não foi encontrado => erro
                If lErro = 28030 Then gError 197698
                
                dPrecoUnitario = Arredonda_Moeda(((objProdutoKit.dCustoStandard / dValorTotalKit) * dValorItemPedido) / (objProdutoKit.dQuantidade * objItemNF.dQuantidade))
                dValorItensKit = dValorItensKit + dPrecoUnitario * (objProdutoKit.dQuantidade * objItemNF.dQuantidade)
            
                If iIndice = 1 Then
            
                    Set objItemNF1 = objNFiscal.ColItensNF.Add(0, 0, objProdutoKit.sProduto, objProdutoKit.sUnidadeMed, objProdutoKit.dQuantidade * objItemNF.dQuantidade, dPrecoUnitario, objItemNF.dPercDesc, (objProdutoKit.dQuantidade * objItemNF.dQuantidade) * (dPrecoUnitario) * objItemNF.dPercDesc, objItemNF.dtDataEntrega, objProduto.sDescricao, 0, objItemPV.lNumIntDoc, 0, 0, objItemPV.lNumIntDoc, objItemNF.ColAlocacoes, 0, "", "", objItemPV.iStatus, 0, "", 0, 0, objProduto.iControleEstoque, objProduto.sSiglaUMEstoque, objProduto.iClasseUM, 0)
                    
                Else
                
                    Set objItemNF1 = objNFiscal.ColItensNF.Add(0, 0, objProdutoKit.sProduto, objProdutoKit.sUnidadeMed, objProdutoKit.dQuantidade * objItemNF.dQuantidade, dPrecoUnitario, objItemNF.dPercDesc, (objProdutoKit.dQuantidade * objItemNF.dQuantidade) * (dPrecoUnitario) * objItemNF.dPercDesc, objItemNF.dtDataEntrega, objProduto.sDescricao, 0, objItemPV.lNumIntDoc, 0, 0, objItemPV.lNumIntDoc, Nothing, 0, "", "", objItemPV.iStatus, 0, "", 0, 0, objProduto.iControleEstoque, objProduto.sSiglaUMEstoque, objProduto.iClasseUM, 0)
                    objCT.gcolItemPedido.Add objItemPV.lNumIntDoc
                    objCT.gcolItemQuant.Add objItemPV.dQuantidade
                    objCT.gcolItemUM.Add objItemPV.sUnidadeMed 'Inserido por Wagner
                    
                End If
                                            
                                
                                
                Set objItemNF1.objInfoUsu = objItemNF
                
            End If
        
        Next

        '####################################
        'Incluído por Wagner
        'Faz o ajuste no valor para bater exatamente
        If Abs(dValorItensKit - dValorItemPedido) > DELTA_VALORMONETARIO Then
        
            dDifAcerto = Arredonda_Moeda(dValorItemPedido - dValorItensKit)
        
            If dDifAcerto < 0 Then
                For Each objItemNF1 In objNFiscal.ColItensNF
                    objItemNF1.dValorDesconto = objItemNF1.dValorDesconto + Abs(dDifAcerto)
                    Exit For
                Next
            Else
                For Each objItemNF1 In objNFiscal.ColItensNF
                    If dDifAcerto / objItemNF1.dQuantidade >= 0.01 - DELTA_VALORMONETARIO Then
                        objItemNF1.dPrecoUnitario = objItemNF1.dPrecoUnitario + Formata_Estoque(dDifAcerto / objItemNF1.dQuantidade)
                    Else
                        objItemNF1.dPrecoUnitario = objItemNF1.dPrecoUnitario + 0.01
                        objItemNF1.dValorDesconto = objItemNF1.dValorDesconto + (0.01 * objItemNF1.dQuantidade) - dDifAcerto
                    End If
                    Exit For
                Next
            End If
        
        End If
        '####################################
        
    End If
    
    NFP_Transfere_Dados_Pedido_NFiscal = SUCESSO

    Exit Function

Erro_NFP_Transfere_Dados_Pedido_NFiscal:

    NFP_Transfere_Dados_Pedido_NFiscal = gErr

    Select Case gErr

        Case 197695 To 197697

        Case 197698
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197699)

    End Select

    Exit Function
    
End Function

Public Property Get ProdutoKit(ByVal objCT As Object) As Object
     Set ProdutoKit = objCT.objUserControl.Controls("ProdutoKit")
End Property

Public Property Get UMKit(ByVal objCT As Object) As Object
     Set UMKit = objCT.objUserControl.Controls("UMKit")
End Property

Public Property Get QuantKit(ByVal objCT As Object) As Object
     Set QuantKit = objCT.objUserControl.Controls("QuantKit")
End Property

'Public Property Get VersaoKit(ByVal objCT As Object) As Object
'     Set VersaoKit = objCT.objUserControl.Controls("VersaoKit")
'End Property

Public Sub ProdutoKit_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ProdutoKit_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)
End Sub

Public Sub ProdutoKit_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)
End Sub

Public Sub ProdutoKit_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
    Set objCT.objGridItens.objControle = ProdutoKit(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub QuantKit_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub QuantKit_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)
End Sub

Public Sub QuantKit_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)
End Sub

Public Sub QuantKit_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
    Set objCT.objGridItens.objControle = QuantKit(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Function NFP_Move_GridItem_Memoria(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal iLinha As Integer) As Long

Dim objItemNFPhar As New ClassItemNF
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_NFP_Move_GridItem_Memoria

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_ProdutoKit_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 197714

    'Armazena produto
    objItemNFPhar.sProduto = sProdutoFormatado
    objItemNFPhar.dQuantidade = StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_QuantKit_Col))
    objItemNFPhar.sUnidadeMed = objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_UMKit_Col)

    Set objItemNF.objInfoUsu = objItemNFPhar
    
    NFP_Move_GridItem_Memoria = SUCESSO

    Exit Function

Erro_NFP_Move_GridItem_Memoria:

    NFP_Move_GridItem_Memoria = gErr

    Select Case gErr

        Case 197714

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177091)

    End Select

    Exit Function
    
End Function

Function NFP_Alocacao_Processa(ByVal objCT As Object) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dQuantidade As Double
Dim sUM As String
Dim iIndice As Integer

On Error GoTo Erro_NFP_Alocacao_Processa

    If Not objCT.gobjInfoUsu Is Nothing Then

        If Len(Trim(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_ProdutoKit_Col))) <> 0 Then
            
            For iIndice = 1 To objCT.GridItens.Row
                If objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ProdutoKit_Col) = objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_ProdutoKit_Col) Then Exit For
            Next
    
            If iIndice = objCT.GridItens.Row Then
    
                'Remove as alocações do grid de alocações
                Call objCT.Remove_Alocacoes(iIndice)
            
                'Formata o Produto
                lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ProdutoKit_Col), sProduto, iPreenchido)
                If lErro <> SUCESSO Then gError 197697
        
                objProduto.sCodigo = sProduto
            
                'Lê o produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 197698
                If lErro = 28030 Then gError 197699
            
                'recolhe a quantidade do grid de itens
                dQuantidade = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_QuantKit_Col))
            
                'Verifica se o produto tem o controle de estoque <> PRODUTO_CONTROLE_SEM_ESTOQUE
                If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA Then
                    
                    'recolhe a UM do ItemNF
                    sUM = objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_UMKit_Col)
                    
                    'Tenta Alocar o produto no Almoxarifado padrão
                    lErro = CF2(objCT, "NFP_AlocaAlmoxarifadoPadrao", dQuantidade, objProduto, sUM)
                    If lErro <> SUCESSO And lErro <> 23796 Then gError 197700 '??? Jones 09/02/05 passou a tratar 23796
                    
                    If lErro = 23796 Then gError 197701 '??? Jones 09/02/05 passou a tratar 23796
                
                End If
    
            End If
    
        Else
    
            lErro = objCT.NFP_Alocacao_Processa(objCT)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 197702
            
            If lErro = 23796 Then gError 197703
    
        End If

    Else

        lErro = objCT.NFP_Alocacao_Processa(objCT)
        If lErro <> SUCESSO And lErro <> 23796 Then gError 197702
        
        If lErro = 23796 Then gError 197703

    End If
    
    NFP_Alocacao_Processa = SUCESSO

    Exit Function

Erro_NFP_Alocacao_Processa:

    NFP_Alocacao_Processa = gErr

    Select Case gErr

        Case 197697, 197698, 197700, 197701, 197702, 197703

        Case 197699
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197704)

    End Select

    Exit Function

End Function

Function NFP_AlocaAlmoxarifadoPadrao(ByVal objCT As Object, dQuantidade As Double, objProduto As ClassProduto, sUM As String) As Long
'Tenta fazer a alocação do produto no almoxarifado padrão. Caso não consiga chama a tela de Alocação de produto.

Dim lErro As Long
Dim dQuantAlocar As Double
Dim dFator As Double
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dQuantDisponivel As Double
Dim objItemNF As ClassItemNF
Dim colOutrosProdutos As New Collection
Dim sProduto As String
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim objItemNFAloc As ClassItemNFAlocacao
Dim sProdutoEnxuto As String
Dim iLinhas As Integer
Dim dAcrescimo As Double
Dim iNumCasasDec As Integer
Dim objItemPV As New ClassItemPedido
Dim colReservaBD As New colReservaItem
Dim dTotal As Double
Dim sProdutoAnt As String

On Error GoTo Erro_NFP_AlocaAlmoxarifadoPadrao

    'Faz a conversão da UM da tela para a UM de estoque
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 197714

    'Converte a quantidade para a UM de estoque
    dQuantAlocar = Arredonda_Estoque(dQuantidade * dFator) 'Alterado por Wagner

    'Busca o Almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 197715
    
    If lErro = 23796 Then gError 197716
    
    'Se encontrou
    If iAlmoxarifado > 0 Then

        objAlmoxarifado.iCodigo = iAlmoxarifado
        'Lê o Aloxarifado
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> 25056 And lErro <> SUCESSO Then gError 197717
        
        If lErro = 25056 Then gError 197718
        
        objEstoqueProduto.sProduto = objProduto.sCodigo
        objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
        'Le os estoques desse produto nesse almoxarifado
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 197719
        
        If lErro = 21306 Then gError 197720 'Não encontrou

        objItemPV.lNumIntDoc = objCT.gcolItemPedido(objCT.GridItens.Row)
        objItemPV.sProduto = objProduto.sCodigo

        lErro = CF("ReservasItemPV_Le_NumIntOrigem", objItemPV, colReservaBD)
        If lErro <> SUCESSO And lErro <> 51601 Then gError 197721
        
        For iIndice = 1 To colReservaBD.Count
            If objEstoqueProduto.iAlmoxarifado = colReservaBD(iIndice).iAlmoxarifado Then
                objEstoqueProduto.dQuantDispNossa = objEstoqueProduto.dQuantDispNossa + colReservaBD(iIndice).dQuantidade
                Exit For
            End If
        Next
        
        dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
        'Remove as alocações do produto do grid de alocações
        Call objCT.Remove_Alocacoes(objCT.GridItens.Row)

        iLinhas = objCT.objGridAlocacoes.iLinhasExistentes
        
        dQuantAlocar = Arredonda_Estoque(dQuantAlocar)

        'Verifica se a Quantidade disponível é maior que a quantidade a alocar
        If (dQuantAlocar - dQuantDisponivel) < QTDE_ESTOQUE_DELTA Then

            iLinhas = iLinhas + 1
            'Coloca no Grid alocações a alocação para o almoxarifado padrão
            objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_Item_Col) = objCT.GridItens.Row
            objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_ProdutoAloc_Col) = objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_ProdutoKit_Col)
            objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_AlmoxAloc_Col) = objAlmoxarifado.sNomeReduzido
            objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_QuantAloc_Col) = Formata_Estoque(dQuantAlocar)
            objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_QuantVend_Col) = Formata_Estoque(dQuantAlocar)
            objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_UMAloc_Col) = objProduto.sSiglaUMEstoque
        
        'Se não for
        Else
            'Se a função não estiver sendo usada pelo carregamento da tela
            'se prepara para chamar a tela de localização
            If Not objCT.gbCarregandoTela Then
                
                Set objItemNF = New ClassItemNF
                'Recolhe os dados do item
                objItemNF.iItem = objCT.GridItens.Row
                objItemNF.sProduto = objProduto.sCodigo
                objItemNF.sDescricaoItem = objProduto.sDescricao
                objItemNF.dQuantidade = dQuantidade
                objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
                objItemNF.lNumIntItemPedVenda = objCT.gcolItemPedido(objCT.GridItens.Row)
    
                'Recolhe todos os produtos dos outros itens
                For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
                    If iIndice <> objCT.GridItens.Row Then
                        lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Produto_Col), sProduto, iPreenchido)
                        If lErro <> SUCESSO Then gError 197722
                        'Adiciona na coleção de produtos
                        If sProdutoAnt <> sProduto Then colOutrosProdutos.Add sProduto
                        sProdutoAnt = sProduto
                    End If
                Next
    
                'Chama a tela de Localização de Produto
                Call Chama_Tela_Modal("LocalizacaoProduto1", objItemNF, colOutrosProdutos, dQuantAlocar, DOCINFO_NFISFVPV)
                If giRetornoTela = vbCancel Then gError 197723 'Se nada foi feito lá
                If giRetornoTela = vbOK Then
    
                    'Se o produto foi substituido
                    If objProduto.sCodigo <> objItemNF.sProduto Then gError 197724
                                    
                    sProdutoEnxuto = String(STRING_PRODUTO, 0)
    
                    lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
                    If lErro Then gError 197725
    
                    objCT.Produto.PromptInclude = False
                    objCT.Produto.Text = sProdutoEnxuto
                    objCT.Produto.PromptInclude = True
                    
                    'Calcula o Total Alocado
                    For Each objItemNFAloc In objItemNF.ColAlocacoes
                        dTotal = dTotal + objItemNFAloc.dQuantidade
                    Next
                    
                    'Para cada alocação feita para o item
                    For Each objItemNFAloc In objItemNF.ColAlocacoes
    
                        iLinhas = iLinhas + 1
                        'coloca no grid de alocações os dados das alocações
                        objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_Item_Col) = objItemNF.iItem
                        objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_ProdutoAloc_Col) = objCT.Produto.Text
                        objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_AlmoxAloc_Col) = objItemNFAloc.sAlmoxarifado
                        objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_QuantAloc_Col) = Formata_Estoque(objItemNFAloc.dQuantidade)
                        objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_QuantVend_Col) = Formata_Estoque(dTotal)
                        objCT.GridAlocacao.TextMatrix(iLinhas, objCT.iGrid_UMAloc_Col) = objItemNF.sUMEstoque
    
                    Next
                
                    'Quantidade.Text = Formata_Estoque(dTotal)
                    If Abs(dQuantAlocar - dTotal) > QTDE_ESTOQUE_DELTA Then
                        objCT.Quantidade.Text = Formata_Estoque(dTotal / dFator)
                    Else
                        objCT.Quantidade.Text = Formata_Estoque(StrParaDbl(objCT.Quantidade.Text))
                    End If
    
                End If
            End If
        End If
        
        'atualiza o numero de linhas existentes no grid de alocações
        objCT.objGridAlocacoes.iLinhasExistentes = iLinhas
    End If

    NFP_AlocaAlmoxarifadoPadrao = SUCESSO

    Exit Function

Erro_NFP_AlocaAlmoxarifadoPadrao:

    NFP_AlocaAlmoxarifadoPadrao = gErr

    Select Case gErr

        Case 197714 To 197717, 197719, 197721, 197722, 197725

        Case 197718
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case 197720
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 197723
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_LOCALIZACAO", gErr, objProduto.sCodigo)

        Case 197724
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_SUBSTITUIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197726)

    End Select

    Exit Function

End Function

Function NFP_Preenche_GridAlocacoes(ByVal objCT As Object, objNFiscal As ClassNFiscal, Optional bGeraAlocacao As Boolean = False) As Long
'Preenche o Grid com as Alocações da Nota Fiscal

Dim objItemAloc As ClassItemNFAlocacao
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim iContador As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemNF As ClassItemNF
Dim sProdutoEnxuto As String
Dim objItemNFPhar As New ClassItemNF
Dim iAchou As Integer

On Error GoTo Erro_NFP_Preenche_GridAlocacoes
    
    'Limpa o grid de alocações
    Call Grid_Limpa(objCT.objGridAlocacoes)

    objCT.objGridAlocacoes.iLinhasExistentes = 0
    'Para cada item da NF
    For iIndice = 1 To objNFiscal.ColItensNF.Count
                
        If objNFiscal.ColItensNF.Item(iIndice).colItensRomaneioGrade.Count = 0 Then
        
            Call objCT.AlocacoesNF_Agrupa(objNFiscal.ColItensNF.Item(iIndice).ColAlocacoes)
            
            iContador = 0
            'Para cada alocação do Item de NF
            For Each objItemAloc In objNFiscal.ColItensNF.Item(iIndice).ColAlocacoes
    
                iContador = iContador + 1
                objCT.objGridAlocacoes.iLinhasExistentes = objCT.objGridAlocacoes.iLinhasExistentes + 1
                iIndice1 = objCT.objGridAlocacoes.iLinhasExistentes
                
                objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 197727
                If lErro <> SUCESSO Then gError 197728
                
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objNFiscal.ColItensNF(iIndice).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 197729
                
                sProdutoEnxuto = String(STRING_PRODUTO, 0)

                lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 197730
                
                Set objItemNF = objNFiscal.ColItensNF(iIndice)
                
                If Not objItemNF.objInfoUsu Is Nothing Then
                    
                    Set objItemNFPhar = objItemNF.objInfoUsu
                    
                    If Len(Trim(objItemNFPhar.sProduto)) > 0 Then
                
                        sProdutoEnxuto = String(STRING_PRODUTO, 0)
        
                        lErro = Mascara_RetornaProdutoEnxuto(objItemNFPhar.sProduto, sProdutoEnxuto)
                        If lErro <> SUCESSO Then gError 197731
                
                        objProduto.sCodigo = objItemNFPhar.sProduto
                        
                        'Lê o Produto
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 28030 Then gError 197732
                        If lErro <> SUCESSO Then gError 197733
                        
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNFPhar.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
                        If lErro <> SUCESSO Then gError 197734
                
                    End If
                    
                End If

                objCT.Produto.PromptInclude = False
                objCT.Produto.Text = sProdutoEnxuto
                objCT.Produto.PromptInclude = True
                
                'Coloca os dados da alocação na tela
                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_Item_Col) = objNFiscal.ColItensNF(iIndice).iItem
                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_ProdutoAloc_Col) = objCT.Produto.Text
                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_AlmoxAloc_Col) = objItemAloc.sAlmoxarifado
                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_UMAloc_Col) = objProduto.sSiglaUMEstoque
                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_QuantAloc_Col) = Formata_Estoque(objItemAloc.dQuantidade)
                
                If Len(Trim(objItemNFPhar.sProduto)) > 0 Then
                
                    dQuantidade = objItemNFPhar.dQuantidade
                
                Else
                
                
                    dQuantidade = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col))
                
                End If
                    
                dQuantidade = Arredonda_Estoque(dQuantidade * dFator)
                
'                If StrParaDbl(Formata_Estoque(dQuantidade)) < dQuantidade Then
    
                    'Calcula o número de casas decimais do Formato de Estoque
'                    iNumCasasDec = Len(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
                    
'                    dAcrescimo = 0
'                    If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
                    
                    'dQuantidade = dQuantidade + dAcrescimo
                
                'End If
                
                
                
                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_QuantVend_Col) = Formata_Estoque(dQuantidade)
    
            Next
            
            'Quando os dados que estão sendo carregados são originados por um pedido
            'e um item não possui localização tenta gerar automaticamente a alocação
            If iContador = 0 And bGeraAlocacao Then
            
                objCT.GridItens.Row = iIndice
                
                lErro = CF2(objCT, "NFP_Alocacao_Processa")
                If lErro <> SUCESSO And lErro <> 23796 And lErro <> 197701 And lErro <> 197703 Then gError 197735
                
            End If
        Else
            
            Call objCT.Atualiza_Grid_Alocacao(objNFiscal.ColItensNF.Item(iIndice))
            
            objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 197736
            If lErro <> SUCESSO Then gError 197737

            If objProduto.iKitVendaComp <> MARCADO Then
                Call objCT.Atualiza_Grid_Itens(objNFiscal.ColItensNF.Item(iIndice))
            End If
        
        End If

    Next
    
    If objCT.gbCarregandoTela Then
    
        For Each objItemNF In objNFiscal.ColItensNF
                
            If objItemNF.objTributacaoItemNF Is Nothing Then
                iAchou = 1
                Exit For
            End If
        Next
        
        If iAchou <> 1 Then
            lErro = objCT.gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
            If lErro <> SUCESSO Then gError 197738
        Else
            lErro = objCT.gobjTribTab.ExtraiNF_Pack(objNFiscal)
            If lErro <> SUCESSO Then gError 197738
        End If
    
    End If
    
    NFP_Preenche_GridAlocacoes = SUCESSO

    Exit Function

Erro_NFP_Preenche_GridAlocacoes:

    NFP_Preenche_GridAlocacoes = gErr
    
    Select Case gErr
    
        Case 197727, 197729 To 197732, 197734 To 197736, 197738
        
        Case 197728, 197733, 197737
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197739)
            
    End Select

End Function

Public Function NFP_Traz_NFiscal_Tela2(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal, ByVal lNumIntPedido As Long) As Long

Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim lErro As Long

On Error GoTo Erro_NFP_Traz_NFiscal_Tela2

    objFilialCliente.iCodFilial = Codigo_Extrai(objCT.Filial.Caption)
    
    'Lê a FilialCliente
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", objCT.Cliente.Caption, objFilialCliente)
    If lErro <> SUCESSO Then gError 197708

    objCliente.sNomeReduzido = objCT.Cliente.Caption
    
    'Lê o Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 197709
        
    'Se não achou o Cliente --> erro
    If lErro = 12348 Then gError 197710

    Call objCT.gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, objFilialCliente.iCodFilial)

    NFP_Traz_NFiscal_Tela2 = SUCESSO
    
    Exit Function
    
Erro_NFP_Traz_NFiscal_Tela2:

    NFP_Traz_NFiscal_Tela2 = gErr

    Select Case gErr
    
        Case 197708, 197709
    
        Case 197710
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197711)
            
    End Select

End Function

