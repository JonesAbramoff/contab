VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFRemPedPhar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Public Property Get ProdutoKit(ByVal objCT As Object) As Object
'     Set ProdutoKit = objCT.objUserControl.Controls("ProdutoKit")
'End Property
'
'Public Property Get UMKit(ByVal objCT As Object) As Object
'     Set UMKit = objCT.objUserControl.Controls("UMKit")
'End Property
'
'Public Property Get QuantKit(ByVal objCT As Object) As Object
'     Set QuantKit = objCT.objUserControl.Controls("QuantKit")
'End Property
'
'Public Function NFRP_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
''Inicializa o Grid de Itens
'
'    Set objGridInt.objForm = objCT
'
'    'Títulos das colunas
'    objGridInt.colColuna.Add ("Item")
'    objGridInt.colColuna.Add ("Produto")
'    objGridInt.colColuna.Add ("Descrição")
'    objGridInt.colColuna.Add ("U.M.")
'    objGridInt.colColuna.Add ("Quantidade")
'    objGridInt.colColuna.Add ("Pack")
'    objGridInt.colColuna.Add ("UM Kit")
''    objGridInt.colColuna.Add ("Versao")
'    objGridInt.colColuna.Add ("Quant.Pack")
'    objGridInt.colColuna.Add ("Preço Unitário")
'    objGridInt.colColuna.Add ("% Desconto")
'    objGridInt.colColuna.Add ("Desconto")
'    objGridInt.colColuna.Add ("Preço Total")
'    objGridInt.colColuna.Add ("Data Entrega")
'
'    'Controles que participam do Grid
'    objGridInt.colCampo.Add (objCT.Produto.Name)
'    objGridInt.colCampo.Add (objCT.DescricaoItem.Name)
'    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
'    objGridInt.colCampo.Add (objCT.Quantidade.Name)
'    objGridInt.colCampo.Add (ProdutoKit(objCT).Name)
'    objGridInt.colCampo.Add (UMKit(objCT).Name)
''    objGridInt.colCampo.Add (VersaoKit(objCT).Name)
'    objGridInt.colCampo.Add (QuantKit(objCT).Name)
'    objGridInt.colCampo.Add (objCT.PrecoUnitario.Name)
'    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
'    objGridInt.colCampo.Add (objCT.Desconto.Name)
'    objGridInt.colCampo.Add (objCT.PrecoTotal.Name)
'    objGridInt.colCampo.Add (objCT.DataEntrega.Name)
'
'    Call CF("NFP_GridItens_InicializaColunas", objCT.objCTNFiscalPedido)
'
'    'Grid do GridInterno
'    objGridInt.objGrid = objCT.GridItens
'
'    'Todas as linhas do grid
'    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1
'
'    'Linhas visíveis do grid
'    objGridInt.iLinhasVisiveis = 5
'
'    'Largura da primeira coluna
'    objCT.GridItens.ColWidth(0) = 500
'
'    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'
'    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
'    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
'
'    'Chama função que inicializa o Grid
'    Call Grid_Inicializa(objGridInt)
'
'    NFRP_Inicializa_Grid_Itens = SUCESSO
'
'    Exit Function
'
'End Function

'Function NFRP_Alocacao_Processa(ByVal objCT As Object) As Long
'
'Dim lErro As Long
'Dim sProduto As String
'Dim iPreenchido As Integer
'Dim objProduto As New ClassProduto
'Dim dQuantidade As Double
'Dim sUM As String
'Dim iIndice As Integer
'
'On Error GoTo Erro_NFRP_Alocacao_Processa
'
'    If Not objCT.gobjInfoUsu Is Nothing Then
'
'        If Len(Trim(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_ProdutoKit_Col))) <> 0 Then
'
'            For iIndice = 1 To objCT.GridItens.Row
'                If objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ProdutoKit_Col) = objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_ProdutoKit_Col) Then Exit For
'            Next
'
'            If iIndice = objCT.GridItens.Row Then
'
'                'Remove as alocações do grid de alocações
'                Call objCT.Remove_Alocacoes(iIndice)
'
'                'Formata o Produto
'                lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ProdutoKit_Col), sProduto, iPreenchido)
'                If lErro <> SUCESSO Then gError 197697
'
'                objProduto.sCodigo = sProduto
'
'                'Lê o produto
'                lErro = CF("Produto_Le", objProduto)
'                If lErro <> SUCESSO And lErro <> 28030 Then gError 197698
'                If lErro = 28030 Then gError 197699
'
'                'recolhe a quantidade do grid de itens
'                dQuantidade = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_QuantKit_Col))
'
'                'Verifica se o produto tem o controle de estoque <> PRODUTO_CONTROLE_SEM_ESTOQUE
'                If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA Then
'
'                    'recolhe a UM do ItemNF
'                    sUM = objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_UMKit_Col)
'
'                    'Tenta Alocar o produto no Almoxarifado padrão
'                    lErro = CF2(objCT, "NFP_AlocaAlmoxarifadoPadrao", dQuantidade, objProduto, sUM)
'                    If lErro <> SUCESSO And lErro <> 23796 Then gError 197700 '??? Jones 09/02/05 passou a tratar 23796
'
'                    If lErro = 23796 Then gError 197701 '??? Jones 09/02/05 passou a tratar 23796
'
'                End If
'
'            End If
'
'        Else
'
'            lErro = objCT.NFRP_Alocacao_Processa(objCT)
'            If lErro <> SUCESSO And lErro <> 23796 Then gError 197702
'
'            If lErro = 23796 Then gError 197703
'
'        End If
'
'    Else
'
'        lErro = objCT.NFRP_Alocacao_Processa(objCT)
'        If lErro <> SUCESSO And lErro <> 23796 Then gError 197702
'
'        If lErro = 23796 Then gError 197703
'
'    End If
'
'    NFRP_Alocacao_Processa = SUCESSO
'
'    Exit Function
'
'Erro_NFRP_Alocacao_Processa:
'
'    NFRP_Alocacao_Processa = gErr
'
'    Select Case gErr
'
'        Case 197697, 197698, 197700, 197701, 197702, 197703
'
'        Case 197699
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197704)
'
'    End Select
'
'    Exit Function
'
'End Function

'Function NFRP_Preenche_GridAlocacoes(ByVal objCT As Object, objNFiscal As ClassNFiscal, Optional bGeraAlocacao As Boolean = False) As Long
''Preenche o Grid com as Alocações da Nota Fiscal
'
'Dim objItemAloc As ClassItemNFAlocacao
'Dim iIndice As Integer
'Dim iIndice1 As Integer
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim dFator As Double
'Dim dQuantidade As Double
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim iNumCasasDec As Integer
'Dim dAcrescimo As Double
'Dim iContador As Integer
'Dim objItemRomaneio As ClassItemRomaneioGrade
'Dim objItemNF As ClassItemNF
'Dim sProdutoEnxuto As String
'Dim objItemNFPhar As New ClassItemNF
'Dim iAchou As Integer
'
'On Error GoTo Erro_NFRP_Preenche_GridAlocacoes
'
'    'Limpa o grid de alocações
'    Call Grid_Limpa(objCT.objGridAlocacoes)
'
'    objCT.objGridAlocacoes.iLinhasExistentes = 0
'    'Para cada item da NF
'    For iIndice = 1 To objNFiscal.ColItensNF.Count
'
'        If objNFiscal.ColItensNF.Item(iIndice).colItensRomaneioGrade.Count = 0 Then
'
'            Call objCT.AlocacoesNF_Agrupa(objNFiscal.ColItensNF.Item(iIndice).ColAlocacoes)
'
'            iContador = 0
'            'Para cada alocação do Item de NF
'            For Each objItemAloc In objNFiscal.ColItensNF.Item(iIndice).ColAlocacoes
'
'                iContador = iContador + 1
'                objCT.objGridAlocacoes.iLinhasExistentes = objCT.objGridAlocacoes.iLinhasExistentes + 1
'                iIndice1 = objCT.objGridAlocacoes.iLinhasExistentes
'
'                objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
'                'Lê o Produto
'                lErro = CF("Produto_Le", objProduto)
'                If lErro <> SUCESSO And lErro <> 28030 Then gError 197727
'                If lErro <> SUCESSO Then gError 197728
'
'                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objNFiscal.ColItensNF(iIndice).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
'                If lErro <> SUCESSO Then gError 197729
'
'                sProdutoEnxuto = String(STRING_PRODUTO, 0)
'
'                lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
'                If lErro <> SUCESSO Then gError 197730
'
'                Set objItemNF = objNFiscal.ColItensNF(iIndice)
'
'                If Not objItemNF.objInfoUsu Is Nothing Then
'
'                    Set objItemNFPhar = objItemNF.objInfoUsu
'
'                    If Len(Trim(objItemNFPhar.sProduto)) > 0 Then
'
'                        sProdutoEnxuto = String(STRING_PRODUTO, 0)
'
'                        lErro = Mascara_RetornaProdutoEnxuto(objItemNFPhar.sProduto, sProdutoEnxuto)
'                        If lErro <> SUCESSO Then gError 197731
'
'                        objProduto.sCodigo = objItemNFPhar.sProduto
'
'                        'Lê o Produto
'                        lErro = CF("Produto_Le", objProduto)
'                        If lErro <> SUCESSO And lErro <> 28030 Then gError 197732
'                        If lErro <> SUCESSO Then gError 197733
'
'                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNFPhar.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
'                        If lErro <> SUCESSO Then gError 197734
'
'                    End If
'
'                End If
'
'                objCT.Produto.PromptInclude = False
'                objCT.Produto.Text = sProdutoEnxuto
'                objCT.Produto.PromptInclude = True
'
'                'Coloca os dados da alocação na tela
'                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_Item_Col) = objNFiscal.ColItensNF(iIndice).iItem
'                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_ProdutoAloc_Col) = objCT.Produto.Text
'                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_AlmoxAloc_Col) = objItemAloc.sAlmoxarifado
'                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_UMAloc_Col) = objProduto.sSiglaUMEstoque
'                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_QuantAloc_Col) = Formata_Estoque(objItemAloc.dQuantidade)
'
'                If Len(Trim(objItemNFPhar.sProduto)) > 0 Then
'
'                    dQuantidade = objItemNFPhar.dQuantidade
'
'                Else
'
'
'                    dQuantidade = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col))
'
'                End If
'
'                dQuantidade = Arredonda_Estoque(dQuantidade * dFator)
'
''                If StrParaDbl(Formata_Estoque(dQuantidade)) < dQuantidade Then
'
'                    'Calcula o número de casas decimais do Formato de Estoque
''                    iNumCasasDec = Len(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
''                    dAcrescimo = 0
''                    If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                    'dQuantidade = dQuantidade + dAcrescimo
'
'                'End If
'
'
'
'                objCT.GridAlocacao.TextMatrix(iIndice1, objCT.iGrid_QuantVend_Col) = Formata_Estoque(dQuantidade)
'
'            Next
'
'            'Quando os dados que estão sendo carregados são originados por um pedido
'            'e um item não possui localização tenta gerar automaticamente a alocação
'            If iContador = 0 And bGeraAlocacao Then
'
'                objCT.GridItens.Row = iIndice
'
'                lErro = CF2(objCT, "NFP_Alocacao_Processa")
'                If lErro <> SUCESSO And lErro <> 23796 And lErro <> 197701 And lErro <> 197703 Then gError 197735
'
'            End If
'        Else
'
'            Call objCT.Atualiza_Grid_Alocacao(objNFiscal.ColItensNF.Item(iIndice))
'
'            objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
'            'Lê o Produto
'            lErro = CF("Produto_Le", objProduto)
'            If lErro <> SUCESSO And lErro <> 28030 Then gError 197736
'            If lErro <> SUCESSO Then gError 197737
'
'            If objProduto.iKitVendaComp <> MARCADO Then
'                Call objCT.Atualiza_Grid_Itens(objNFiscal.ColItensNF.Item(iIndice))
'            End If
'
'        End If
'
'    Next
'
'    If objCT.gbCarregandoTela Then
'
'        For Each objItemNF In objNFiscal.ColItensNF
'
'            If objItemNF.objTributacaoItemNF Is Nothing Then
'                iAchou = 1
'                Exit For
'            End If
'        Next
'
'        If iAchou <> 1 Then
'            lErro = objCT.gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
'            If lErro <> SUCESSO Then gError 197738
'        Else
'            lErro = objCT.gobjTribTab.ExtraiNF_Pack(objNFiscal)
'            If lErro <> SUCESSO Then gError 197738
'        End If
'
'    End If
'
'    NFRP_Preenche_GridAlocacoes = SUCESSO
'
'    Exit Function
'
'Erro_NFRP_Preenche_GridAlocacoes:
'
'    NFRP_Preenche_GridAlocacoes = gErr
'
'    Select Case gErr
'
'        Case 197727, 197729 To 197732, 197734 To 197736, 197738
'
'        Case 197728, 197733, 197737
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197739)
'
'    End Select
'
'End Function
'

Public Function NFRP_Limpa_Tela(ByVal objCT As Object) As Long
    objCT.FornEntTerc.Text = "403"
    Call objCT.FornEntTerc_validate(bSGECancelDummy)
    NFRP_Limpa_Tela = SUCESSO
End Function

Function NFRP_CargaPosFormLoad(ByVal objCT As Object) As Long
    objCT.FornEntTerc.Text = "403"
    Call objCT.FornEntTerc_validate(bSGECancelDummy)
    NFRP_CargaPosFormLoad = SUCESSO
End Function
