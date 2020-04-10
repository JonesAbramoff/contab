VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get PrecoTabela(ByVal objCT As Object) As Object
     Set PrecoTabela = objCT.objUserControl.Controls("PrecoTabela")
End Property

Public Property Get DescTabela(ByVal objCT As Object) As Object
     Set DescTabela = objCT.objUserControl.Controls("DescTabela")
End Property

Public Sub PrecoTabela_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PrecoTabela_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.objgriditens)
End Sub

Public Sub PrecoTabela_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objgriditens)
End Sub

Public Sub PrecoTabela_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
    Set objCT.objgriditens.objControle = PrecoTabela(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objgriditens)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub DescTabela_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DescTabela_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.objgriditens)
End Sub

Public Sub DescTabela_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objgriditens)
End Sub

Public Sub DescTabela_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
    Set objCT.objgriditens.objControle = DescTabela(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objgriditens)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Function PV_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim iColunasEscondidas As Integer

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. Pedida")
    
    If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then
        objGridInt.colColuna.Add ("Quant. Cancelada")
    End If
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        objGridInt.colColuna.Add ("Tab. Preço")
    End If
    
    '??? webasto
    objGridInt.colColuna.Add ("Preço Tabela")
    objGridInt.colColuna.Add ("Desc. Tabela")
    
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Quant Reservada")
    objGridInt.colColuna.Add ("Quant Faturada")
    objGridInt.colColuna.Add ("Prioridade")
    
    If gobjFAT.iExibePercFatPV = MARCADO Then
        objGridInt.colColuna.Add ("% a Menos Fat")
        objGridInt.colColuna.Add ("% a Mais Fat")
        objGridInt.colColuna.Add ("Ação Receb Fora Faixa")
        objGridInt.colColuna.Add ("Quant Fat. A Mais")
    End If
    
    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        objGridInt.colColuna.Add ("Comissão")
    End If

    'Controles que participam do Grid
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoProduto.Name)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)

    If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then
        objGridInt.colCampo.Add (objCT.QuantCancelada.Name)
    Else
        objCT.QuantCancelada.left = POSICAO_FORA_TELA
    End If
    
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        objGridInt.colCampo.Add (objCT.TabPrecoItemPV.Name)
    Else
        objCT.TabPrecoItemPV.left = POSICAO_FORA_TELA
    End If
    
    '??? webasto
    objGridInt.colCampo.Add (PrecoTabela(objCT).Name)
    objGridInt.colCampo.Add (DescTabela(objCT).Name)
    
    objGridInt.colCampo.Add (objCT.PrecoUnitario.Name)
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (objCT.Desconto.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotal.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotalB.Name)
    objGridInt.colCampo.Add (objCT.DataEntrega.Name)
    objGridInt.colCampo.Add (objCT.QuantReservadaPV.Name)
    objGridInt.colCampo.Add (objCT.QuantFaturada.Name)
    objGridInt.colCampo.Add (objCT.Prioridade.Name)
    
    If gobjFAT.iExibePercFatPV = MARCADO Then
        objGridInt.colCampo.Add (objCT.PercentMenosReceb.Name)
        objGridInt.colCampo.Add (objCT.PercentMaisReceb.Name)
        objGridInt.colCampo.Add (objCT.RecebForaFaixa.Name)
        objGridInt.colCampo.Add (objCT.QuantFatAMais.Name)
    Else
        objCT.PercentMenosReceb.left = POSICAO_FORA_TELA
        objCT.PercentMaisReceb.left = POSICAO_FORA_TELA
        objCT.RecebForaFaixa.left = POSICAO_FORA_TELA
        objCT.QuantFatAMais.left = POSICAO_FORA_TELA
    End If
    
    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        objGridInt.colCampo.Add (objCT.ComissaoItemPV.Name)
    Else
        objCT.ComissaoItemPV.left = POSICAO_FORA_TELA
    End If

    'Colunas do Grid
    objCT.iGrid_ItemProduto_Col = 0
    objCT.iGrid_Produto_Col = 1
    objCT.iGrid_DescProduto_Col = 2
    objCT.iGrid_UnidadeMed_Col = 3
    objCT.iGrid_QuantPedida_Col = 4

    If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then
        objCT.iGrid_QuantCancel_Col = 5
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If
    
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        objCT.iGrid_TabPrecoItem_Col = 6 - iColunasEscondidas
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If
    
    '??? webasto
    objCT.gobjInfoUsu.iGrid_PrecoTabela_Col = 7 - iColunasEscondidas
    objCT.gobjInfoUsu.iGrid_DescTabela_Col = 8 - iColunasEscondidas
    
    '??? webasto: alterados
    objCT.iGrid_ValorUnitario_Col = 9 - iColunasEscondidas
    objCT.iGrid_PercDesc_Col = 10 - iColunasEscondidas
    objCT.iGrid_Desconto_Col = 11 - iColunasEscondidas
    objCT.iGrid_PrecoTotal_Col = 12 - iColunasEscondidas
    objCT.iGrid_PrecoTotalB_Col = 13 - iColunasEscondidas
    objCT.iGrid_DataEntrega_Col = 14 - iColunasEscondidas
    objCT.iGrid_QuantReservada_Col = 15 - iColunasEscondidas
    objCT.iGrid_QuantFaturada_Col = 16 - iColunasEscondidas
    objCT.iGrid_Prioridade_Col = 17 - iColunasEscondidas
    
    If gobjFAT.iExibePercFatPV = MARCADO Then

        objCT.iGrid_PercentMenosReceb_Col = 18 - iColunasEscondidas
        objCT.iGrid_PercentMaisReceb_Col = 19 - iColunasEscondidas
        objCT.iGrid_RecebForaFaixa_Col = 20 - iColunasEscondidas
        objCT.iGrid_QuantFatAMais_Col = 21 - iColunasEscondidas
    Else
        iColunasEscondidas = iColunasEscondidas + 4
    End If
    
    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        objCT.iGrid_ComissaoItem_Col = 20 - iColunasEscondidas
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If

    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    objCT.GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    PV_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Function PV_Carrega_Grid_Itens_Linha(ByVal objCT As Object, ByVal objPV As ClassPedidoDeVenda, ByVal iLinha As Integer) As Long

Dim objItemPVWB As ClassItemPedidoWB
Dim objItemPV As ClassItemPedido

On Error GoTo Erro_PV_Carrega_Grid_Itens_Linha

    Set objItemPV = objPV.colItensPedido.Item(iLinha)

    If Not (objItemPV.objInfoUsu Is Nothing) Then
    
        Set objItemPVWB = objItemPV.objInfoUsu

        objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoTabela_Col) = Format(objItemPVWB.dPrecoTabela, gobjFAT.sFormatoPrecoUnitario)
        objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_DescTabela_Col) = Format(objItemPVWB.dDescTabela, "Percent")

    End If

    PV_Carrega_Grid_Itens_Linha = SUCESSO
    
    Exit Function
    
Erro_PV_Carrega_Grid_Itens_Linha:

    PV_Carrega_Grid_Itens_Linha = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177116)
    
    End Select
    
    Exit Function

End Function

Public Function PV_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Saida_Celula_GridItens

    'Verifica qual a coluna do Grid em questão
    Select Case objGridInt.objGrid.Col
    
        Case objCT.iGrid_PercDesc_Col
            lErro = objCT.Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 177745
    
        Case objCT.gobjInfoUsu.iGrid_PrecoTabela_Col
            lErro = Saida_Celula_PrecoTabela(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 177745
    
        Case objCT.gobjInfoUsu.iGrid_DescTabela_Col
            lErro = Saida_Celula_DescTabela(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 177745
    
    End Select

    PV_Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_PV_Saida_Celula_GridItens:

    PV_Saida_Celula_GridItens = gErr

    Select Case gErr
    
        Case 177745

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177124)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoTabela(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PrecoTabela

    Set objGridInt.objControle = PrecoTabela(objCT)

    If Len(Trim(PrecoTabela(objCT).Text)) > 0 Then
        
        lErro = Valor_Positivo_Critica(PrecoTabela(objCT).Text)
        If lErro <> SUCESSO Then gError 177131
        
        PrecoTabela(objCT).Text = Format(PrecoTabela(objCT).Text, gobjFAT.sFormatoPrecoUnitario)
        
        '??? calcular novo preco unitario
        
    End If
       
    'Faz a saída de célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 177132

    lErro = PrecoUnitario_Calcula(objCT, objCT.GridItens.Row)
    If lErro <> SUCESSO Then gError 177132
    
    'Recalcula o preço total do produto
    Call CF2(objCT, "PV_PrecoTotal_Calcula", objCT.GridItens.Row)
    
    'Recalcula o valor total da nota
    Call objCT.ValorTotal_Calcula
    
    Saida_Celula_PrecoTabela = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoTabela:

    Saida_Celula_PrecoTabela = gErr

    Select Case gErr

        Case 177131 To 177132
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177133)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescTabela(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long, dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_DescTabela

    Set objGridInt.objControle = DescTabela(objCT)

    If Len(Trim(DescTabela(objCT).Text)) > 0 Then
        
        lErro = Porcentagem_Critica(DescTabela(objCT).Text)
        If lErro <> SUCESSO Then gError 26694

        dPercentDesc = CDbl(DescTabela(objCT).Text)
        
        If Format(dPercentDesc, "#0.#0\%") <> objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_DescTabela_Col) Then
            
            'se for igual a 100% -> erro
            If dPercentDesc = 100 Then gError 26695

            DescTabela(objCT).Text = Format(DescTabela(objCT), "Fixed")
        
            '??? calcular novo Desc unitario
            
        End If
        
    End If
       
    'Faz a saída de célula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 177132

    lErro = PrecoUnitario_Calcula(objCT, objCT.GridItens.Row)
    If lErro <> SUCESSO Then gError 177132
    
    'Recalcula o preço total do produto
    Call CF2(objCT, "PV_PrecoTotal_Calcula", objCT.GridItens.Row)
    
    'Recalcula o valor total da nota
    Call objCT.ValorTotal_Calcula
    
    Saida_Celula_DescTabela = SUCESSO

    Exit Function

Erro_Saida_Celula_DescTabela:

    Saida_Celula_DescTabela = gErr

    Select Case gErr

        Case 177132, 26694
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 26695
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177133)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Function PV_PrecoUnitario_Calcula(ByVal objCT As Object, ByVal iLinha As Integer, ByVal sUM As String, ByVal objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lErro As Long
Dim dPercAcresFin As Double
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_PV_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 26638

    If lErro = 28030 Then gError 26639
    'Converte a quantidade para a UM de Venda
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 26640

    dPrecoUnitario = objTabelaPrecoItem.dPreco * dFator

    'Recolhe o percentual de acréscimo financeiro
    dPercAcresFin = StrParaDbl(objCT.PercAcrescFin.Text) / 100

    'Calcula o Preço unitário
    If dPercAcresFin <> 0 Then
        dPrecoUnitario = dPrecoUnitario * (1 + dPercAcresFin)
    End If

    'webasto: este preco calculado é o preço de tabela
    objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoTabela_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
    
    lErro = PrecoUnitario_Calcula(objCT, objCT.GridItens.Row)
    If lErro <> SUCESSO Then gError 177132
    
    PV_PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PV_PrecoUnitario_Calcula:

    PV_PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 26638, 26640

        Case 26639
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157981)

    End Select

    Exit Function

End Function

Public Function PV_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Rotina_Grid_Enable

    Select Case objControl.Name

       Case PrecoTabela(objCT).Name, DescTabela(objCT).Name
        
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
    End Select
    
    PV_Rotina_Grid_Enable = SUCESSO

    Exit Function

Erro_PV_Rotina_Grid_Enable:

    PV_Rotina_Grid_Enable = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177150)

    End Select

    Exit Function

End Function

Public Function PV_Move_GridItem_Memoria(ByVal objCT As Object, ByVal objPV As ClassPedidoDeVenda, ByVal objItemPV As ClassItemPedido, ByVal iLinha As Integer, ByVal sProduto As String) As Long

Dim objItemPVWB As New ClassItemPedidoWB

On Error GoTo Erro_PV_Move_GridItem_Memoria

    objItemPVWB.dPrecoTabela = StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoTabela_Col))
    objItemPVWB.dDescTabela = PercentParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_DescTabela_Col))

    Set objItemPV.objInfoUsu = objItemPVWB
    
    PV_Move_GridItem_Memoria = SUCESSO

    Exit Function

Erro_PV_Move_GridItem_Memoria:

    PV_Move_GridItem_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177153)

    End Select

    Exit Function
    
End Function

Public Function PV_Valida_Grid_Itens(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    PV_Valida_Grid_Itens = SUCESSO
End Function

Private Function PrecoUnitario_Calcula(ByVal objCT As Object, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim dPrecoTabela As Double, dDescTabela As Double
Dim dPrecoUnitarioNovo As Double, dPrecoUnitarioAntigo As Double
Dim iPreenchido As Integer, sProduto As String, sProduto1 As String
Dim objProduto As New ClassProduto, dIPI As Double

On Error GoTo Erro_PrecoUnitario_Calcula

    sProduto1 = objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col)
    
    'Verifica se o Produto está preenchido
    If Len(Trim(sProduto1)) > 0 Then
        
        'Formata o produto
        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 27682
    
        If iPreenchido = PRODUTO_PREENCHIDO Then
        
            objProduto.sCodigo = sProduto
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 177926
            
            If lErro = SUCESSO Then dIPI = objProduto.dIPIAliquota
    
        End If
    
    End If
    
    dPrecoTabela = StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoTabela_Col))
    dDescTabela = PercentParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_DescTabela_Col))

    dPrecoUnitarioNovo = Arredonda_Moeda(dPrecoTabela * (1 - dDescTabela) / (1 + dIPI))
    
    dPrecoUnitarioAntigo = StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_ValorUnitario_Col))
    
    If dPrecoUnitarioNovo <> 0 And Abs(dPrecoUnitarioAntigo - dPrecoUnitarioNovo) > DELTA_VALORMONETARIO2 Then
        objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_ValorUnitario_Col) = Format(dPrecoUnitarioNovo, gobjFAT.sFormatoPrecoUnitario)
    End If

    PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PrecoUnitario_Calcula:

    PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 27682, 177926
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177153)

    End Select

    Exit Function
    
End Function
