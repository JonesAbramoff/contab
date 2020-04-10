VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTOrcamentoVendaTA"
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

Public Property Get PercParticVend2(ByVal objCT As Object) As Object
     Set PercParticVend2 = objCT.objUserControl.Controls("PercParticVend2")
End Property

Public Property Get NomeExibicao(ByVal objCT As Object) As Object
     Set NomeExibicao = objCT.objUserControl.Controls("NomeExibicao")
End Property

Public Function OV_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim iIncremento As Integer
Dim objUserControl As Object

    Set objGridInt.objForm = objCT
    Set objUserControl = objCT

    'Títulos das colunas
    If gobjFAT.iOVExibeOrdItens = MARCADO Then
        objGridInt.colColuna.Add (" ")
    Else
        objGridInt.colColuna.Add ("Item")
    End If
    If gobjFAT.iOVExibeTS = MARCADO Then
        objGridInt.colColuna.Add ("S")
    End If
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colColuna.Add ("Versão do Kit")
        objGridInt.colColuna.Add ("Versao do Kit Base")
    End If
    
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens4", objGridInt)
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Base")
    objGridInt.colColuna.Add ("% Acréscimo")
    objGridInt.colColuna.Add ("Acréscimo")
    objGridInt.colColuna.Add ("Preço Unitário")
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens1", objGridInt)
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Prazo Entrega")
    objGridInt.colColuna.Add ("Data Entrega")
    
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Motivo Perda")
    objGridInt.colColuna.Add ("Observação")

    'Controles que participam do Grid
    If gobjFAT.iOVExibeTS = MARCADO Then
        objGridInt.colCampo.Add (objCT.SomaItem.Name)
    End If
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoProduto.Name)
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colCampo.Add (objCT.VersaoKit.Name)
        objGridInt.colCampo.Add (objCT.VersaoKitBase.Name)
        
    Else
    
        objCT.VersaoKit.left = POSICAO_FORA_TELA
        objCT.VersaoKit.TabStop = False
    
        objCT.VersaoKitBase.left = POSICAO_FORA_TELA
        objCT.VersaoKitBase.TabStop = False

    End If

    Call CF("Orcamento_Inicializa_Grid_Itens5", objGridInt, objUserControl)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.quantidade.Name)
    objGridInt.colCampo.Add (PrecoBase(objCT).Name)
    objGridInt.colCampo.Add (PercAcresc(objCT).Name)
    objGridInt.colCampo.Add (Acrescimo(objCT).Name)
    objGridInt.colCampo.Add (objCT.PrecoUnitario.Name)
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens2", objGridInt, objUserControl)
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (objCT.Desconto.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotal.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotalB.Name)
    objGridInt.colCampo.Add (objCT.PrazoEntregaItem.Name)
    objGridInt.colCampo.Add (objCT.dataEntrega.Name)
    
    objGridInt.colCampo.Add (objCT.StatusItem.Name)
    objGridInt.colCampo.Add (objCT.MotivoPerdaItem.Name)
    objGridInt.colCampo.Add (objCT.Observacao.Name)
    
    'Colunas do Grid
    objCT.iGrid_ItemProduto_Col = 0
    If gobjFAT.iOVExibeTS = MARCADO Then
        objCT.iGrid_SomaItem_Col = 1
        iIncremento = iIncremento + 1
    End If
    objCT.iGrid_Produto_Col = 1 + iIncremento
    objCT.iGrid_DescProduto_Col = 2 + iIncremento
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objCT.iGrid_VersaoKit_Col = 3 + iIncremento
        objCT.iGrid_VersaoKitBase_Col = 4 + iIncremento
        iIncremento = iIncremento + 2
        
    End If
    
    Call CF("Orcamento_Inicializa_Grid_Itens6", iIncremento)
    objCT.iGrid_UnidadeMed_Col = 3 + iIncremento
    objCT.iGrid_Quantidade_Col = 4 + iIncremento
    
    objCT.gobjInfoUsu.iGrid_PrecoBase_Col = 5 + iIncremento
    objCT.gobjInfoUsu.iGrid_PercAcresc_Col = 6 + iIncremento
    objCT.gobjInfoUsu.iGrid_Acrescimo_Col = 7 + iIncremento
    iIncremento = iIncremento + 3
    
    objCT.iGrid_PrecoUnitario_Col = 5 + iIncremento
    Call CF("Orcamento_Inicializa_Grid_Itens3", iIncremento)
    objCT.iGrid_PercDesc_Col = 6 + iIncremento
    objCT.iGrid_Desconto_Col = 7 + iIncremento
    objCT.iGrid_PrecoTotal_Col = 8 + iIncremento
    objCT.iGrid_PrecoTotalB_Col = 9 + iIncremento
    objCT.iGrid_PrazoEntrega_Col = 10 + iIncremento
    objCT.iGrid_DataEntrega_Col = 11 + iIncremento
    
    objCT.iGrid_Status_Col = 12 + iIncremento
    objCT.iGrid_MotivoPerda_Col = 13 + iIncremento
    objCT.iGrid_Observacao_Col = 14 + iIncremento
    
    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    If gobjFAT.iOVExibeTS = MARCADO Then
        objGridInt.iLinhasVisiveis = 7
    Else
        objGridInt.iLinhasVisiveis = 8
    End If

    'Largura da primeira coluna
    objCT.GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    OV_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Function OV_Move_GridItens_Memoria(ByVal objCT As Object, ByVal objItemOV As ClassItemOV, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objItemOVTA As New ClassItemOVTA

On Error GoTo Erro_OV_Move_GridItens_Memoria

    objItemOVTA.dPrecoBase = StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoBase_Col))
    objItemOVTA.dPercAcresc = PercentParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PercAcresc_Col))
    objItemOVTA.dAcrescimo = PercentParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Acrescimo_Col))
    
    Set objItemOV.objInfoUsu = objItemOVTA

    OV_Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_OV_Move_GridItens_Memoria:

    OV_Move_GridItens_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function OV_Traz_GridItens_Tela(ByVal objCT As Object, ByVal objItemOV As ClassItemOV, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objItemOVTA As ClassItemOVTA

On Error GoTo Erro_OV_Traz_GridItens_Tela

    Set objItemOVTA = objItemOV.objInfoUsu
    
    'O campo de acréscimo foi incluido depois e por isso pode estar zerado
    If objItemOVTA.dAcrescimo = 0 And objItemOVTA.dPercAcresc <> 0 Then
        objItemOVTA.dAcrescimo = objItemOVTA.dPrecoBase * objItemOV.dQuantidade * objItemOVTA.dPercAcresc
    End If

    objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoBase_Col) = Format(objItemOVTA.dPrecoBase, gobjFAT.sFormatoPrecoUnitario)
    objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PercAcresc_Col) = Format(objItemOVTA.dPercAcresc, "Percent")
    objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Acrescimo_Col) = Format(objItemOVTA.dAcrescimo, "STANDARD")
    
    OV_Traz_GridItens_Tela = SUCESSO

    Exit Function

Erro_OV_Traz_GridItens_Tela:

    OV_Traz_GridItens_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function OV_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case objCT.gobjInfoUsu.iGrid_PrecoBase_Col
            lErro = Saida_Celula_PrecoBase(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 210459

        Case objCT.gobjInfoUsu.iGrid_PercAcresc_Col
            lErro = Saida_Celula_PercAcresc(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 210460

        Case objCT.gobjInfoUsu.iGrid_Acrescimo_Col
            lErro = Saida_Celula_Acrescimo(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 210460


    End Select

    OV_Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_OV_Saida_Celula_GridItens:

    OV_Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 210459, 210460

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210461)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoBase(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Base que está deixando de ser a corrente

Dim lErro As Long
Dim bPrecoBaseIgual As Boolean

On Error GoTo Erro_Saida_Celula_PrecoBase

    bPrecoBaseIgual = False

    Set objGridInt.objControle = PrecoBase(objCT)

    If Len(Trim(PrecoBase(objCT).Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoBase(objCT).Text)
        If lErro <> SUCESSO Then gError 210462
        
        PrecoBase(objCT).Text = Format(PrecoBase(objCT).Text, gobjFAT.sFormatoPrecoUnitario)

    End If

    'Comparação com Preço Base anterior
    If StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) = StrParaDbl(PrecoBase(objCT).Text) Then bPrecoBaseIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 210463

    'Se mudou o preço Base ou o Preço unitário da linha atual não está preenchido
    If Not bPrecoBaseIgual Then
                        
        objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_Acrescimo_Col) = Format(StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) * StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Quantidade_Col)) * (PercentParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PercAcresc_Col))), gobjFAT.sFormatoPrecoUnitario)
                        
        'Reclacula o preço unitário
        lErro = Recalcula_PrecoUnitario(objCT, objCT.GridItens.Row)
        If lErro <> SUCESSO Then gError 210464
        
    End If

    Saida_Celula_PrecoBase = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoBase:

    Saida_Celula_PrecoBase = gErr

    Select Case gErr

        Case 210462 To 210464
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210465)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercAcresc(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual de Acrescimo do valor base que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentAcresc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dQuantidade As Double
Dim sValorPercAnterior As String

On Error GoTo Erro_Saida_Celula_PercAcresc

    Set objGridInt.objControle = PercAcresc(objCT)

    If Len(PercAcresc(objCT).Text) > 0 Then
        
        'Critica a porcentagem
        lErro = Porcentagem_Critica_Negativa(PercAcresc(objCT).Text)
        If lErro <> SUCESSO Then gError 210466

        dPercentAcresc = StrParaDbl(PercAcresc(objCT).Text)

        If Format(dPercentAcresc, "#0.#0\%") <> objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PercAcresc_Col) Then

            PercAcresc(objCT).Text = Format(dPercentAcresc, "Fixed")

        End If

    End If

    sValorPercAnterior = objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PercAcresc_Col)
    If sValorPercAnterior = "" Then sValorPercAnterior = "0,00%"
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 210467
    
    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_Acrescimo_Col) = Format(StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) * StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Quantidade_Col)) * (PercentParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PercAcresc_Col))), gobjFAT.sFormatoPrecoUnitario)
    
    'Se foi alterada
    If Format(dPercentAcresc, "#0.#0\%") <> sValorPercAnterior Then

        'Reclacula o preço unitário
        lErro = Recalcula_PrecoUnitario(objCT, objCT.GridItens.Row)
        If lErro <> SUCESSO Then gError 210468

    End If

    Saida_Celula_PercAcresc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercAcresc:

    Saida_Celula_PercAcresc = gErr

    Select Case gErr

        Case 210466 To 210467
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 210468

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210469)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Acrescimo(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual de Acrescimo do valor base que está deixando de ser a corrente

Dim lErro As Long
Dim dAcresc As Double
Dim dAcrescAnt As Double

On Error GoTo Erro_Saida_Celula_Acrescimo

    Set objGridInt.objControle = Acrescimo(objCT)

    If Len(Acrescimo(objCT).Text) > 0 Then
        
        'Critica a porcentagem
        lErro = Valor_NaoNegativo_Critica(Acrescimo(objCT).Text)
        If lErro <> SUCESSO Then gError 210466

        dAcresc = StrParaDbl(Acrescimo(objCT).Text)

        Acrescimo(objCT).Text = Format(dAcresc, "Fixed")

    End If
    
    dAcrescAnt = StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_Acrescimo_Col))
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 210467
    
    If StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Quantidade_Col)) > 0 Then
        objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PercAcresc_Col) = Format(StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_Acrescimo_Col)) / ((StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PrecoBase_Col))) * StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Quantidade_Col))), "PERCENT")
    End If
    
    'Se foi alterada
    If Abs(dAcresc - dAcrescAnt) > DELTA_VALORMONETARIO Then

        'Reclacula o preço unitário
        lErro = Recalcula_PrecoUnitario(objCT, objCT.GridItens.Row)
        If lErro <> SUCESSO Then gError 210468
        
    End If

    Saida_Celula_Acrescimo = SUCESSO

    Exit Function

Erro_Saida_Celula_Acrescimo:

    Saida_Celula_Acrescimo = gErr

    Select Case gErr

        Case 210466 To 210467
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 210468

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210469)

    End Select

    Exit Function

End Function

Private Function Recalcula_PrecoUnitario(ByVal objCT As Object, ByVal iLinhaGrid As Integer, Optional bPulaRecalculo As Boolean = False) As Long

Dim dPrecoBase As Double
Dim dPrecoUnitario As Double
Dim dDias As Double
Dim dTaxaCliente As Double
Dim dTaxaFinDia As Double
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim lErro As Long

On Error GoTo Erro_Recalcula_PrecoUnitario

    'precobase * ( 1 + ((1 + ((1 + taxamensal/100) ^ (1 / 30)) - 1) ^ dias) - 1)
    
    'Se o acréscimo financeiro é igual a zero
'    If StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_Acrescimo_Col)) = 0 Then
'
'        'Preço unitário = preço base
'        objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_PrecoUnitario_Col) = objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)
'
'    'Se é maior que zero
'    Else
    
        'If StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_Quantidade_Col)) > 0 Then
        dPrecoUnitario = StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) * (1 + PercentParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PercAcresc_Col)))
'        Else
'            dPrecoUnitario = StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) + StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_Acrescimo_Col))
'        End If
        objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_PrecoUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        
    objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_Acrescimo_Col) = Format((StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_PrecoUnitario_Col)) - StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.gobjInfoUsu.iGrid_PrecoBase_Col))) * StrParaDbl(objCT.GridItens.TextMatrix(iLinhaGrid, objCT.iGrid_Quantidade_Col)), "STANDARD")
        
'    End If

    If Not bPulaRecalculo Then
        
        If iLinhaGrid <= objCT.objGridItens.iLinhasExistentes Then
            
            'Recalcula preço total
            Call objCT.PrecoTotal_Calcula(iLinhaGrid)
        
            lErro = objCT.gobjTribTab.Alteracao_Item_Grid(iLinhaGrid)
            If lErro <> SUCESSO Then gError 210470
        
        End If
    
        lErro = objCT.ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 210471
    
        lErro = objCT.FormPreco_Calcula(iLinhaGrid)
        If lErro <> SUCESSO Then gError 210471
    
    End If
    
    Recalcula_PrecoUnitario = SUCESSO
    
    Exit Function
    
Erro_Recalcula_PrecoUnitario:

    Recalcula_PrecoUnitario = gErr
    
    Select Case gErr
    
        Case 210470, 210471
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210472)
    
    End Select
    
    Resume Next
    
    Exit Function
    
End Function


Public Sub PrecoBase_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoBase_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)

End Sub

Public Sub PrecoBase_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)

End Sub

Public Sub PrecoBase_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = PrecoBase(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercAcresc_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercAcresc_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)

End Sub

Public Sub PercAcresc_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)

End Sub

Public Sub PercAcresc_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = PercAcresc(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Acrescimo_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Acrescimo_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)

End Sub

Public Sub Acrescimo_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)

End Sub

Public Sub Acrescimo_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = Acrescimo(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function OV_Rotina_Grid_Enable(ByVal objCT As Object, ByVal sNomeControle As String, ByVal iProdutoPreenchido As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Rotina_Grid_Enable

    Select Case sNomeControle
        
        Case PrecoBase(objCT).Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                PrecoBase(objCT).Enabled = False
            Else
                PrecoBase(objCT).Enabled = True
            End If
    
        Case PercAcresc(objCT).Name
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                PercAcresc(objCT).Enabled = False
            Else
                PercAcresc(objCT).Enabled = True
            End If
            
        Case Acrescimo(objCT).Name
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                Acrescimo(objCT).Enabled = False
            Else
                Acrescimo(objCT).Enabled = True
            End If
    
        Case objCT.PrecoUnitario.Name
            objCT.PrecoUnitario.Enabled = False
    
    
    End Select
    
    OV_Rotina_Grid_Enable = SUCESSO
     
    Exit Function
    
Erro_OV_Rotina_Grid_Enable:

    OV_Rotina_Grid_Enable = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210473)
     
    End Select
     
    Exit Function

End Function

Public Function OV_Move_PrecoUnitario_Tela(ByVal objCT As Object, ByVal dPrecoUnitario As Double, ByVal iLinha As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Move_PrecoUnitario_Tela

    If dPrecoUnitario > 0 Then
        objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoBase_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
    Else
        objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PrecoBase_Col) = ""
    End If

    'Reclacula o preço unitário
    lErro = Recalcula_PrecoUnitario(objCT, iLinha)
    If lErro <> SUCESSO Then gError 210458

    OV_Move_PrecoUnitario_Tela = SUCESSO
     
    Exit Function
    
Erro_OV_Move_PrecoUnitario_Tela:

    OV_Move_PrecoUnitario_Tela = gErr
     
    Select Case gErr
          
        Case 210458
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210459)
     
    End Select
     
    Exit Function


End Function

Public Function OV_Form_Load(ByVal objCT As Object) As Long

    PrecoBase(objCT).Format = gobjFAT.sFormatoPrecoUnitario

    OV_Form_Load = SUCESSO
    
End Function


Public Sub ValorAcrescimoItens_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercAcrescimoItens_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorAcrescimoItens_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long, dTotal As Double
Dim dValorAcrescimo As Double
Dim vbMsg As VbMsgBoxResult, iIndice As Integer
Dim dAcrescimoItens As Double, dFator As Double

On Error GoTo Erro_ValorAcrescimoItens_Validate

    dValorAcrescimo = 0

    'Verifica se o Valor está preenchido
    If Len(Trim(ValorAcrescimoItens(objCT).Text)) > 0 Then
    
        'Faz a Crítica do Valor digitado
        lErro = Valor_NaoNegativo_Critica(ValorAcrescimoItens(objCT).Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dValorAcrescimo = CDbl(ValorAcrescimoItens(objCT).Text)
        
        'Coloca o Valor formatado na tela
        ValorAcrescimoItens(objCT).Text = Format(dValorAcrescimo, "Standard")
        
    End If
    
    'Se houve alguma alteração
    If Abs(objCT.gobjInfoUsu.dValorAcrescimoItensAnt - dValorAcrescimo) > DELTA_VALORMONETARIO Then
        
        'Se o Acrescimo foi alterado nos itens pegunta se quer que o sistema recalcule
        If objCT.gobjInfoUsu.iAcrescimoAlterado = REGISTRO_ALTERADO Then
        
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ACRESCIMO_ITENS_ALTERADO")
            If vbMsg = vbNo Then
                dValorAcrescimo = objCT.gobjInfoUsu.dValorAcrescimoItensAnt
                ValorAcrescimoItens(objCT).Text = Format(dValorAcrescimo, "Standard")
                gError ERRO_SEM_MENSAGEM
            End If
            objCT.gobjInfoUsu.iAcrescimoAlterado = 0
               
        End If
        
        lErro = ValorAcrescimoItens_Aplica(objCT)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objCT.gobjInfoUsu.dValorAcrescimoItensAnt = dValorAcrescimo
    
    End If
    
    dTotal = 0
    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
        If StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotalB_Col)) > 0 And objCT.gobjOrcamentoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
            dTotal = dTotal + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) * StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col))
        End If
    Next

    If dTotal > 0 Then
        objCT.gobjInfoUsu.dPercAcrescimoItensAnt = Arredonda_Moeda(dValorAcrescimo / dTotal, 4)
    Else
        objCT.gobjInfoUsu.dPercAcrescimoItensAnt = 0
    End If
    PercAcrescimoItens(objCT).Text = Format(objCT.gobjInfoUsu.dPercAcrescimoItensAnt * 100, "FIXED")
        
    Exit Sub

Erro_ValorAcrescimoItens_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

Public Sub PercAcrescimoItens_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim dPercAcrescimo As Double, dTotal As Double
Dim vbMsg As VbMsgBoxResult, iIndice As Integer
Dim dAcrescimoItens As Double, dFator As Double

On Error GoTo Erro_PercAcrescimoItens_Validate

    dPercAcrescimo = 0

    'Verifica se o Valor está preenchido
    If Len(Trim(PercAcrescimoItens(objCT).Text)) > 0 Then
    
        'Faz a Crítica do Valor digitado
        lErro = Porcentagem_Critica(PercAcrescimoItens(objCT).Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dPercAcrescimo = StrParaDbl(PercAcrescimoItens(objCT).Text) / 100

    End If
    
    'Se houve alguma alteração
    If Abs(objCT.gobjInfoUsu.dPercAcrescimoItensAnt - dPercAcrescimo) > DELTA_VALORMONETARIO2 Then
        
        'Se o Acrescimo foi alterado nos itens pegunta se quer que o sistema recalcule
        If objCT.gobjInfoUsu.iAcrescimoAlterado = REGISTRO_ALTERADO Then
        
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ACRESCIMO_ITENS_ALTERADO")
            If vbMsg = vbNo Then
                dPercAcrescimo = objCT.gobjInfoUsu.dPercAcrescimoItensAnt
                PercAcrescimoItens(objCT).Text = Format(dPercAcrescimo * 100, "FIXED")
                gError ERRO_SEM_MENSAGEM
            End If
            objCT.gobjInfoUsu.iAcrescimoAlterado = 0
                
        End If
        
        dTotal = 0
        For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
            If StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotalB_Col)) > 0 And objCT.gobjOrcamentoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                dTotal = dTotal + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) * StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col))
            End If
        Next
        
        If dTotal > 0 Then
        
            ValorAcrescimoItens(objCT).Text = Format(Arredonda_Moeda(dPercAcrescimo * dTotal), "Standard")
            Call ValorAcrescimoItens_Validate(objCT, bSGECancelDummy)
            
        Else
            For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
                objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercAcresc_Col) = Format(dPercAcrescimo, "Percent")
            Next
        End If
    
    End If
        
    Exit Sub

Erro_PercAcrescimoItens_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

Function ValorAcrescimoItens_Aplica(ByVal objCT As Object) As Long

Dim lErro As Long, dTotal As Double, dValorItem As Double, dAcrescimoItem As Double
Dim dAcrescimo As Double, dFator As Double, dAcrescimoAplicado As Double, dDiferenca As Double
Dim dPercDesc As Double, iIndice As Integer, dValorTotal As Double
Dim iLinha As Integer

On Error GoTo Erro_ValorAcrescimoItens_Aplica

    If objCT.objGridItens.iLinhasExistentes > 0 Then
    
        dTotal = 0
        For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
            If StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotalB_Col)) > 0 And objCT.gobjOrcamentoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                dTotal = dTotal + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) * StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col))
            End If
        Next
        
        'dTotal = StrParaDbl(objCT.ValorProdutos2.Caption)
        dAcrescimo = StrParaDbl(ValorAcrescimoItens(objCT).Text)
        dFator = dAcrescimo / (dTotal)
    
        For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
            If StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotalB_Col)) > 0 And objCT.gobjOrcamentoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                If iLinha = 0 Then iLinha = iIndice
                dValorItem = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) * StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col))
                dAcrescimoItem = Arredonda_Moeda(dValorItem * dFator)
                dAcrescimoAplicado = dAcrescimoAplicado + dAcrescimoItem
                objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Acrescimo_Col) = Format(dAcrescimoItem, "STANDARD")
            End If
        Next
        If Abs(dAcrescimo - dAcrescimoAplicado) > DELTA_VALORMONETARIO And iLinha > 0 Then
            objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Acrescimo_Col) = Format(StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Acrescimo_Col)) + dAcrescimoAplicado - dAcrescimo, "STANDARD")
        End If
        
        For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
            If StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotalB_Col)) > 0 And objCT.gobjOrcamentoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                dValorItem = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) * StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col))
                dPercDesc = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Acrescimo_Col)) / dValorItem
                objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercAcresc_Col) = Format(dPercDesc, "Percent")
                
                'Reclacula o preço unitário
                lErro = Recalcula_PrecoUnitario(objCT, iIndice)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                dValorTotal = dValorTotal + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))
            End If
        
        Next
    
    End If
        
    'Coloca valor total dos produtos na tela
    objCT.ValorProdutos.Caption = Format(dValorTotal, "Standard")

    'Calcula o valor total da nota
    lErro = objCT.ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    ValorAcrescimoItens_Aplica = SUCESSO

    Exit Function

Erro_ValorAcrescimoItens_Aplica:

    ValorAcrescimoItens_Aplica = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208382)

    End Select

    Exit Function

End Function

Function ValorAcrescimoItens_Calcula(ByVal objCT As Object) As Long

Dim lErro As Long, iIndice As Integer
Dim dAcrescimo As Double, dPercDesc As Double
Dim dTotal As Double, dDesconto As Double

On Error GoTo Erro_ValorAcrescimoItens_Calcula

    dAcrescimo = 0
    dPercDesc = 0
    dDesconto = 0
    If Not (objCT.objGridItens Is Nothing) Then
        If objCT.objGridItens.iLinhasExistentes > 0 Then
            dTotal = 0
            For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
                If objCT.gobjOrcamentoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                    dAcrescimo = dAcrescimo + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Acrescimo_Col))
                    dTotal = dTotal + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PrecoBase_Col)) * StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col))
                    dDesconto = dDesconto + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Desconto_Col))
                End If
             Next
            If dTotal > 0 Then dPercDesc = Arredonda_Moeda(dAcrescimo / dTotal, 4)
            ValorAcrescimoItens(objCT).Text = Format(dAcrescimo, "Standard")
            PercAcrescimoItens(objCT).Text = Format(dPercDesc * 100, "FIXED")
        Else
            ValorAcrescimoItens(objCT).Text = Format(0, "Standard")
            PercAcrescimoItens(objCT).Text = Format(0, "FIXED")
        End If
        
        objCT.gobjInfoUsu.dValorAcrescimoItensAnt = dAcrescimo
        objCT.gobjInfoUsu.dPercAcrescimoItensAnt = dPercDesc
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

Function OV_ValorTotal_Calcula(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_OV_ValorTotal_Calcula

    Call ValorAcrescimoItens_Calcula(objCT)

    OV_ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_OV_ValorTotal_Calcula:

    OV_ValorTotal_Calcula = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157776)

    End Select

    Exit Function

End Function

Function OV_Limpa_OrcamentoVenda2(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Limpa_OrcamentoVenda2

    objCT.gobjInfoUsu.iAcrescimoAlterado = 0
    objCT.gobjInfoUsu.dValorAcrescimoItensAnt = 0
    objCT.gobjInfoUsu.dPercAcrescimoItensAnt = 0
    
    ValorVariacaoItens(objCT).Caption = Format(0, "Standard")
    PercVariacaoItens(objCT).Caption = Format(0, "PERCENT")
    
    OV_Limpa_OrcamentoVenda2 = SUCESSO

    Exit Function

Erro_OV_Limpa_OrcamentoVenda2:

    OV_Limpa_OrcamentoVenda2 = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157776)

    End Select

    Exit Function

End Function

Public Function OV_Saida_Celula_Quantidade(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Saida_Celula_Quantidade

    'Reclacula o preço unitário
    lErro = Recalcula_PrecoUnitario(objCT, objCT.GridItens.Row, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    OV_Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_OV_Saida_Celula_Quantidade:

    OV_Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210469)

    End Select

    Exit Function
    
End Function

Public Function OV_Saida_Celula_PrecoUnitario(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Saida_Celula_PrecoUnitario

    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_Acrescimo_Col) = Format((StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_PrecoUnitario_Col)) - StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PrecoBase_Col))) * StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Quantidade_Col)), "STANDARD")
    
    If StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Quantidade_Col)) > 0 Then
        objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PercAcresc_Col) = Format((StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_Acrescimo_Col)) / StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Quantidade_Col)) / StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.gobjInfoUsu.iGrid_PrecoBase_Col))), "PERCENT")
    End If

    'Reclacula o preço unitário
    lErro = Recalcula_PrecoUnitario(objCT, objCT.GridItens.Row, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    OV_Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_OV_Saida_Celula_PrecoUnitario:

    OV_Saida_Celula_PrecoUnitario = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210469)

    End Select

    Exit Function
    
End Function

Function OV_Traz_OrcamentoVenda_Tela(ByVal objCT As Object, objOrcamentoVenda As ClassOrcamentoVenda, ByVal iBrowse As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Traz_OrcamentoVenda_Tela

    Call ValorAcrescimoItens_Calcula(objCT)
    
    If objOrcamentoVenda.iVendedor2 <> 0 Then
        PercParticVend2(objCT).Text = Format(objOrcamentoVenda.dPercParticVend2 * 100, "FIXED")
    Else
        PercParticVend2(objCT).Text = ""
    End If

    NomeExibicao(objCT).Text = objOrcamentoVenda.sNomeCli
    
    OV_Traz_OrcamentoVenda_Tela = SUCESSO

    Exit Function

Erro_OV_Traz_OrcamentoVenda_Tela:

    OV_Traz_OrcamentoVenda_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210469)

    End Select

    Exit Function
    
End Function

Public Function OV_AnaliseVenda_Preenche(ByVal objCT As Object, ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal objAnaliseVendas As ClassAnaliseVendaInfo) As Long

Dim lErro As Long, iIndice As Integer
Dim objItemOVTA As New ClassItemOVTA
Dim objItemOV As New ClassItemOV
Dim objItemAnalise As ClassItemAnaliseVenda

On Error GoTo Erro_OV_AnaliseVenda_Preenche

    iIndice = 0
    For Each objItemOV In objOrcamentoVenda.colItens
        iIndice = iIndice + 1
        Set objItemAnalise = objAnaliseVendas.colItens(iIndice)
        Set objItemOVTA = objItemOV.objInfoUsu
        objItemAnalise.dPrecoBase = objItemOVTA.dPrecoBase
    Next

    OV_AnaliseVenda_Preenche = SUCESSO

    Exit Function

Erro_OV_AnaliseVenda_Preenche:

    OV_AnaliseVenda_Preenche = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 210469)

    End Select

    Exit Function

End Function

Public Function OV_Valida_Grid_Itens(ByVal objCT As Object, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim sProduto1 As String
Dim iPreenchido As Integer

On Error GoTo Erro_OV_Valida_Grid_Itens

    sProduto1 = objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 210942

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 210939
    If lErro = 28030 Then gError 210940

    If objProduto.iAtivo = PRODUTO_INATIVO Then gError 210941
    
    OV_Valida_Grid_Itens = SUCESSO

    Exit Function

Erro_OV_Valida_Grid_Itens:

    OV_Valida_Grid_Itens = gErr

    Select Case gErr

        Case 210939, 210942

        Case 210940
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 210941
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INATIVO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function OV_Move_OrcamentoVenda_Memoria(ByVal objCT As Object, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long

On Error GoTo Erro_OV_Move_OrcamentoVenda_Memoria

    If objOrcamentoVenda.iVendedor2 <> 0 And Len(Trim(PercParticVend2(objCT).Text)) = 0 Then gError 210941
    
    objOrcamentoVenda.dPercParticVend2 = StrParaDbl(PercParticVend2(objCT).Text) / 100

    objOrcamentoVenda.sNomeCli = NomeExibicao(objCT).Text
    
    Exit Function

Erro_OV_Move_OrcamentoVenda_Memoria:

    OV_Move_OrcamentoVenda_Memoria = gErr

    Select Case gErr

        Case 210941
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCPARTICVEND_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function
    
End Function

Public Sub PercParticVend2_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercParticVend2_Validate

    'Verifica se o Valor está preenchido
    If Len(Trim(PercParticVend2(objCT).Text)) > 0 Then
    
        'Faz a Crítica do Valor digitado
        lErro = Porcentagem_Critica(PercParticVend2(objCT).Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
        
    Exit Sub

Erro_PercParticVend2_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

Public Function OV_Cliente_Validate(ByVal objCT As Object, ByVal objCli As ClassCliente) As Long
    NomeExibicao(objCT).Text = objCli.sNomeReduzido
    OV_Cliente_Validate = SUCESSO
End Function
