VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaPan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Function PV_Rotina_Grid_Enable1(ByVal objCT As Object, iLinha As Integer, objControl As Object, iCaminho As Integer) As Long
    
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim iTipo As Integer
Dim sUnidadeMed As String
Dim iIndice As Integer

On Error GoTo Erro_PV_Rotina_Grid_Enable1
    
    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 31389

    Select Case objControl.Name
        
        Case objCT.Produto.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                objCT.Produto.Enabled = False
            Else
                objCT.Produto.Enabled = True
            End If
        
        Case objCT.UnidadeMed.Name
            'guarda a um go grid nessa coluna
            sUM = objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_UnidadeMed_Col)
            'Se já houve um faturamento (parcial ou não ) desse item desabilita a UM
            If Len(Trim(objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_QuantFaturada_Col))) > 0 Then
                If CDbl(objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_QuantFaturada_Col)) > 0 Then
                    objCT.UnidadeMed.Enabled = False
                Else
                    If iCaminho = ROTINA_GRID_CLICK Or iCaminho = ROTINA_GRID_ENTRADA_CELULA Or iCaminho = ROTINA_GRID_ABANDONA_CELULA Then
                        objCT.UnidadeMed.Enabled = True
                    Else
                        objCT.UnidadeMed.Enabled = False
                    End If
                End If
            Else
                    If iCaminho = ROTINA_GRID_CLICK Or iCaminho = ROTINA_GRID_ENTRADA_CELULA Or iCaminho = ROTINA_GRID_ABANDONA_CELULA Then
                        objCT.UnidadeMed.Enabled = True
                    Else
                        objCT.UnidadeMed.Enabled = False
                    End If
            End If
            
            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = objCT.UnidadeMed.Text
            
            objCT.UnidadeMed.Clear

            If iProdutoPreenchido <> PRODUTO_VAZIO Then

                objProduto.sCodigo = sProdutoFormatado
                'Lê o produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 26644

                If lErro = 28030 Then gError 26645

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lê as UMs do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 26646
                'Carrega a combo de UMs
                For Each objUM In colSiglas
                    objCT.UnidadeMed.AddItem objUM.sSigla
                Next
                
                'Tento selecionar na Combo a Unidade anterior
                If objCT.UnidadeMed.ListCount <> 0 Then
    
                    For iIndice = 0 To objCT.UnidadeMed.ListCount - 1
    
                        If objCT.UnidadeMed.List(iIndice) = sUnidadeMed Then
                            objCT.UnidadeMed.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If
            
            Else
                objCT.UnidadeMed.Enabled = False
            End If

        Case objCT.TabPrecoItemPV.Name, objCT.ComissaoItemPV.Name, objCT.PercentDesc.Name, objCT.Desconto.Name, objCT.DataEntrega.Name, objCT.PercentMaisReceb.Name, objCT.PercentMenosReceb.Name, objCT.RecebForaFaixa.Name, objCT.Prioridade.Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                If iCaminho = ROTINA_GRID_CLICK Or iCaminho = ROTINA_GRID_ENTRADA_CELULA Or iCaminho = ROTINA_GRID_ABANDONA_CELULA Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If
            
        Case objCT.PrecoUnitario.Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or gobjFAT.iBloqEditPrecoPV = MARCADO Then
                objControl.Enabled = False
            Else
                If iCaminho = ROTINA_GRID_CLICK Or iCaminho = ROTINA_GRID_ENTRADA_CELULA Or iCaminho = ROTINA_GRID_ABANDONA_CELULA Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If
        
        Case objCT.DescricaoProduto.Name
            If iCaminho = ROTINA_GRID_CLICK Or iCaminho = ROTINA_GRID_ENTRADA_CELULA Or iCaminho = ROTINA_GRID_ABANDONA_CELULA Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
                
        '******** O TRATAMENTO DE QUANTIDADE E QUANTCANCELADA FOI DESTACADO PARA TRATAR GRADE
        Case objCT.Quantidade.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or left(objCT.GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case objCT.QuantCancelada.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or left(objCT.GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                If iCaminho = ROTINA_GRID_CLICK Or iCaminho = ROTINA_GRID_ENTRADA_CELULA Or iCaminho = ROTINA_GRID_ABANDONA_CELULA Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If


        Case objCT.TipoBloqueio.Name, objCT.ResponsavelBL.Name
            'Se for algum Bloqueio previamente cadastrado, desabilita o controle
            If iLinha <= objCT.giGridBloqueioLinhasBloqueadas Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case objCT.Desconto1Ate.Name, objCT.Desconto1Valor.Name, objCT.Desconto1Percentual.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(objCT.GridParcelas.TextMatrix(iLinha, objCT.iGrid_Desc1Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                iTipo = Codigo_Extrai(objCT.GridParcelas.TextMatrix(iLinha, objCT.iGrid_Desc1Codigo_Col))
                If objControl.Name = objCT.Desconto1Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = objCT.Desconto1Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    objCT.Desconto1Valor.Enabled = True
                ElseIf objControl.Name = objCT.Desconto1Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = objCT.Percentual) Then
                    objCT.Desconto1Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If
        
        Case objCT.Desconto2Ate.Name, objCT.Desconto2Valor.Name, objCT.Desconto2Percentual.Name
            'Habilita os campos de desconto em sequencia
            iTipo = Codigo_Extrai(objCT.GridParcelas.TextMatrix(iLinha, objCT.iGrid_Desc2Codigo_Col))
            If Len(Trim(objCT.GridParcelas.TextMatrix(iLinha, objCT.iGrid_Desc2Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                If objControl.Name = objCT.Desconto2Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = objCT.Desconto2Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    objCT.Desconto2Valor.Enabled = True
                ElseIf objControl.Name = Desconto2Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = objCT.Percentual) Then
                    objCT.Desconto2Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If
        
        Case objCT.Desconto3Ate.Name, objCT.Desconto3Valor.Name, objCT.Desconto3Percentual.Name
            'Habilita os campos de desconto em sequencia
            iTipo = Codigo_Extrai(objCT.GridParcelas.TextMatrix(iLinha, objCT.iGrid_Desc3Codigo_Col))
            If Len(Trim(objCT.GridParcelas.TextMatrix(iLinha, objCT.iGrid_Desc3Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                If objControl.Name = objCT.Desconto3Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = objCT.Desconto3Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    objCT.Desconto3Valor.Enabled = True
                ElseIf objControl.Name = Desconto3Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = objCT.Percentual) Then
                    objCT.Desconto3Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If
        
        
        Case objCT.ValorParcela.Name
            'Se o vencimento estiver preenchido, habilita o controle
            If Len(Trim(objCT.GridParcelas.TextMatrix(objCT.GridParcelas.Row, objCT.iGrid_Vencimento_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case objCT.TipoDesconto2.Name, objCT.TipoDesconto3.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(objCT.GridParcelas.TextMatrix(iLinha, objCT.GridParcelas.Col - 4))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
        Case objCT.QuantFatAMais.Name
            objControl.Enabled = False
        
        Case Else
            lErro = CF2(objCT, "PV_Rotina_Grid_Enable", iLinha, objControl, iCaminho, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 124178
    
    End Select

    'embalagem
    lErro = objCT.gobjEmbalagem.Rotina_Grid_Enable_Emb(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 96127
    
    lErro = CF2(objCT, "PV_Rotina_Grid_Enable2", iLinha, objControl, iCaminho, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 124178
       
    PV_Rotina_Grid_Enable1 = SUCESSO
    
    Exit Function

Erro_PV_Rotina_Grid_Enable1:

    PV_Rotina_Grid_Enable1 = gErr

    Select Case gErr

        Case 26644, 26646, 31389, 96127, 124178

        Case 26645
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157983)

    End Select

    Exit Function

End Function

