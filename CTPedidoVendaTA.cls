VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get PrecoBase(ByVal objCT As Object) As Object
     Set PrecoBase = objCT.objUserControl.Controls("PrecoBase")
End Property

Public Property Get PercAcresc(ByVal objCT As Object) As Object
     Set PercAcresc = objCT.objUserControl.Controls("PercAcresc")
End Property

Public Property Get Acrescimo(ByVal objCT As Object) As Object
     Set Acrescimo = objCT.objUserControl.Controls("Acrescimo")
End Property

Public Property Get PercAcrescimoItens(ByVal objCT As Object) As Object
     Set PercAcrescimoItens = objCT.objUserControl.Controls("PercAcrescimoItens")
End Property

Public Property Get ValorAcrescimoItens(ByVal objCT As Object) As Object
     Set ValorAcrescimoItens = objCT.objUserControl.Controls("ValorAcrescimoItens")
End Property

Public Property Get PercVariacaoItens(ByVal objCT As Object) As Object
     Set PercVariacaoItens = objCT.objUserControl.Controls("PercVariacaoItens")
End Property

Public Property Get ValorVariacaoItens(ByVal objCT As Object) As Object
     Set ValorVariacaoItens = objCT.objUserControl.Controls("ValorVariacaoItens")
End Property

Public Function PV_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

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
    objGridInt.colColuna.Add ("Preço Base")
    objGridInt.colColuna.Add ("% Acréscimo")
    objGridInt.colColuna.Add ("Acréscimo")
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
    
    objGridInt.colCampo.Add (PrecoBase(objCT).Name)
    objGridInt.colCampo.Add (PercAcresc(objCT).Name)
    objGridInt.colCampo.Add (Acrescimo(objCT).Name)
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
    
    objCT.gobjInfoUsu.iGrid_PrecoBase_Col = 7 - iColunasEscondidas
    objCT.gobjInfoUsu.iGrid_PercAcresc_Col = 8 - iColunasEscondidas
    objCT.gobjInfoUsu.iGrid_Acrescimo_Col = 9 - iColunasEscondidas
    iColunasEscondidas = iColunasEscondidas - 3
    
    objCT.iGrid_ValorUnitario_Col = 7 - iColunasEscondidas
    objCT.iGrid_PercDesc_Col = 8 - iColunasEscondidas
    objCT.iGrid_Desconto_Col = 9 - iColunasEscondidas
    objCT.iGrid_PrecoTotal_Col = 10 - iColunasEscondidas
    objCT.iGrid_PrecoTotalB_Col = 11 - iColunasEscondidas
    objCT.iGrid_DataEntrega_Col = 12 - iColunasEscondidas
    objCT.iGrid_QuantReservada_Col = 13 - iColunasEscondidas
    objCT.iGrid_QuantFaturada_Col = 14 - iColunasEscondidas
    objCT.iGrid_Prioridade_Col = 15 - iColunasEscondidas
    
    If gobjFAT.iExibePercFatPV = MARCADO Then

        objCT.iGrid_PercentMenosReceb_Col = 16 - iColunasEscondidas
        objCT.iGrid_PercentMaisReceb_Col = 17 - iColunasEscondidas
        objCT.iGrid_RecebForaFaixa_Col = 18 - iColunasEscondidas
        objCT.iGrid_QuantFatAMais_Col = 19 - iColunasEscondidas
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

Function PV_Traz_Pedido_Tela2(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItemPV As ClassItemPedido
Dim objItemOV As ClassItemOV
Dim objItemOVTA As ClassItemOVTA
Dim iItem As Integer, bAchou As Boolean

On Error GoTo Erro_PV_Traz_Pedido_Tela2

    If Not (objOrcamentoVenda Is Nothing) Then
    
        If objOrcamentoVenda.colItens.Count = 0 Then
            lErro = CF("OrcamentoVenda_Le_Completo", objOrcamentoVenda)
            If lErro <> SUCESSO And lErro <> 101232 Then gError ERRO_SEM_MENSAGEM
        End If

        For iIndice = 1 To objPedidoVenda.colItensPedido.Count
        
            Set objItemPV = objPedidoVenda.colItensPedido.Item(iIndice)
            
            bAchou = False
            For Each objItemOV In objOrcamentoVenda.colItens
                If objItemOV.sProduto = objItemPV.sProduto Then
                    bAchou = True
                    Exit For
                End If
            Next

            If bAchou Then
            
                If Not (objItemOV.objInfoUsu Is Nothing) Then
                
                    Set objItemOVTA = objItemOV.objInfoUsu
                    
                    If objItemOVTA.dAcrescimo = 0 And objItemOVTA.dPercAcresc <> 0 Then
                        objItemOVTA.dAcrescimo = objItemOVTA.dPrecoBase * objItemOV.dQuantidade * objItemOVTA.dPercAcresc
                    End If
                           
                    objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoBase_Col) = Format(objItemOVTA.dPrecoBase, gobjFAT.sFormatoPrecoUnitario)
                    objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercAcresc_Col) = Format(objItemOVTA.dPercAcresc, "Percent")
                    objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Acrescimo_Col) = Format(objItemOVTA.dAcrescimo, "STANDARD")
                
                End If
            
            End If
            
        Next

    End If
    
    Call ValorAcrescimoItens_Calcula(objCT, objOrcamentoVenda)
    
    Exit Function

Erro_PV_Traz_Pedido_Tela2:

    PV_Traz_Pedido_Tela2 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157924)

    End Select

    Exit Function
    
End Function

Function ValorAcrescimoItens_Calcula(ByVal objCT As Object, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long, iIndice As Integer
Dim dAcrescimo As Double, dPercDesc As Double
Dim dTotal As Double, dDesconto As Double
Dim objItemOV As ClassItemOV
Dim objItemOVTA As ClassItemOVTA

On Error GoTo Erro_ValorAcrescimoItens_Calcula

    dAcrescimo = 0
    dPercDesc = 0
    dDesconto = 0
    If Not (objOrcamentoVenda Is Nothing) Then
        If objOrcamentoVenda.colItens.Count > 0 Then
            dTotal = 0
            For Each objItemOV In objOrcamentoVenda.colItens
                If Not (objItemOV.objInfoUsu Is Nothing) Then
                    Set objItemOVTA = objItemOV.objInfoUsu
                Else
                    Set objItemOVTA = New ClassItemOVTA
                End If
                If objItemOV.objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                    dAcrescimo = dAcrescimo + objItemOVTA.dAcrescimo
                    dTotal = dTotal + objItemOVTA.dPrecoBase * objItemOV.dQuantidade
                    dDesconto = dDesconto + objItemOV.dValorDesconto
                End If
             Next
            If dTotal > 0 Then dPercDesc = Arredonda_Moeda(dAcrescimo / dTotal, 4)
            ValorAcrescimoItens(objCT).Text = Format(dAcrescimo, "Standard")
            PercAcrescimoItens(objCT).Text = Format(dPercDesc * 100, "FIXED")
        Else
            ValorAcrescimoItens(objCT).Text = Format(0, "Standard")
            PercAcrescimoItens(objCT).Text = Format(0, "FIXED")
        End If
        
    End If

    ValorVariacaoItens(objCT).Caption = Format(dAcrescimo - dDesconto, "Standard")
    If dTotal > 0 Then
        PercVariacaoItens(objCT).Caption = Format((dAcrescimo - dDesconto) / dTotal, "PERCENT")
    Else
        PercVariacaoItens(objCT).Caption = Format(0, "FIXED")
    End If
    
    ValorAcrescimoItens_Calcula = SUCESSO

    Exit Function

Erro_ValorAcrescimoItens_Calcula:

    ValorAcrescimoItens_Calcula = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208382)

    End Select

    Exit Function

End Function
