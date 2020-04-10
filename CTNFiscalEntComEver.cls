VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFEntCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get TotalMoeda(ByVal objCT As Object) As Object
     Set TotalMoeda = objCT.objUserControl.Controls("TotalMoeda")
End Property

Public Property Get ValorUnitarioMoeda(ByVal objCT As Object) As Object
     Set ValorUnitarioMoeda = objCT.objUserControl.Controls("ValorUnitarioMoeda")
End Property

Function NFEntCom_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Preço Unitário")
    
    'sport
    objGridInt.colColuna.Add ("Valor Moeda")
    
    objGridInt.colColuna.Add ("% Desc.")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoItem.Name)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)
    objGridInt.colCampo.Add (objCT.Ccl.Name)
    objGridInt.colCampo.Add (objCT.ValorUnitario.Name)
    
    'sport
    objGridInt.colCampo.Add (ValorUnitarioMoeda(objCT).Name)
    
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (objCT.Desconto.Name)
    objGridInt.colCampo.Add (objCT.ValorTotal.Name)
    objGridInt.colCampo.Add (objCT.precoTotalb.Name)
    
    'Se é permitido que o valor unitário do ItemNF é diferente do valor unitário do ItemPC
    If gobjCOM.iNFDiferentePC = NFISCAL_NAO_ACEITA_DIFERENCA_PC Then
        objCT.ValorUnitario.Enabled = False
    Else
        objCT.ValorUnitario.Enabled = True
    End If
    
    'Colunas da Grid
    objCT.iGrid_Produto_Col = 1
    objCT.iGrid_DescProduto_Col = 2
    objCT.iGrid_UnidadeMed_Col = 3
    objCT.iGrid_Quantidade_Col = 4
'distribuicao
'    iGrid_Almoxarifado_Col = 5
    objCT.iGrid_Ccl_Col = 5
    objCT.iGrid_ValorUnitario_Col = 6
    
    'sport
    objCT.iGrid_ValorUnitarioMoeda_Col = 7
    
    objCT.iGrid_PercDesc_Col = 8
    objCT.iGrid_Desconto_Col = 9
    objCT.iGrid_PrecoTotal_Col = 10
    objCT.iGrid_PrecoTotalb_Col = 11

    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_NF + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    objCT.GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFEntCom_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Sub ValorUnitarioMoeda_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorUnitarioMoeda_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)

End Sub

Public Sub ValorUnitarioMoeda_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)

End Sub

Public Sub ValorUnitarioMoeda_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = ValorUnitarioMoeda(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ValorUnitarioMoeda_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Function NFEntCom_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_NFEntCom_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case objCT.iGrid_ValorUnitarioMoeda_Col
            lErro = Saida_Celula_ValorUnitarioMoeda(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 124164

    End Select

    NFEntCom_Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_NFEntCom_Saida_Celula_GridItens:

    NFEntCom_Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 124164 To 124168

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179127)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorUnitarioMoeda(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorUnitarioMoeda As Double
Dim dValorUnitario  As Double

On Error GoTo Erro_Saida_Celula_ValorUnitarioMoeda

    Set objGridInt.objControle = ValorUnitarioMoeda(objCT)

    'Se estiver preenchido
    If Len(Trim(ValorUnitarioMoeda(objCT).ClipText)) > 0 Then

        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorUnitarioMoeda(objCT).Text)
        If lErro <> SUCESSO Then gError 66340

        dValorUnitario = StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_ValorUnitario_Col))
       
        If dValorUnitario - StrParaDbl(ValorUnitarioMoeda(objCT).Text) < -DELTA_VALORMONETARIO2 Then gError 140324

        ValorUnitarioMoeda(objCT).Text = Format(ValorUnitarioMoeda(objCT).Text, gobjFAT.sFormatoPrecoUnitario)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 66341

''*** incluidos p/tratamento de tributacao *******************************
    lErro = objCT.gobjTribTab.Alteracao_Item_Grid(objCT.GridItens.Row)
    If lErro <> SUCESSO Then gError 65764
'*** fim tributacao
    
    Call objCT.ValorTotal_Calcula
    
    Saida_Celula_ValorUnitarioMoeda = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorUnitarioMoeda:

    Saida_Celula_ValorUnitarioMoeda = gErr

    Select Case gErr

        Case 66340, 66341
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 140324
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_MOEDA_MAIOR", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179128)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Function NFEntCom_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iCaminho As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_NFEntCom_Rotina_Grid_Enable

    Select Case objControl.Name
        
        Case ValorUnitarioMoeda(objCT).Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
    
    End Select
    
    NFEntCom_Rotina_Grid_Enable = SUCESSO
     
    Exit Function
    
Erro_NFEntCom_Rotina_Grid_Enable:

    NFEntCom_Rotina_Grid_Enable = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179129)
     
    End Select
     
    Exit Function

End Function

Function NFEntCom_Move_GridItem_Memoria(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal iIndice As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_NFEntCom_Move_GridItem_Memoria

    objItemNF.dPrecoUnitarioMoeda = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_ValorUnitarioMoeda_Col))
    
    NFEntCom_Move_GridItem_Memoria = SUCESSO
     
    Exit Function
    
Erro_NFEntCom_Move_GridItem_Memoria:

    NFEntCom_Move_GridItem_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179130)
     
    End Select
     
    Exit Function

End Function

Function NFEntCom_Preenche_GridItens(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long, iIndice As Integer, objItemNF As ClassItemNF

On Error GoTo Erro_NFEntCom_Preenche_GridItens

    For iIndice = 1 To objNFiscal.ColItensNF.Count
       
        Set objItemNF = objNFiscal.ColItensNF(iIndice)
        
        objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_ValorUnitarioMoeda_Col) = Format(objItemNF.dPrecoUnitarioMoeda, FORMATO_PRECO_UNITARIO_EXTERNO)
    
    Next
    
    NFEntCom_Preenche_GridItens = SUCESSO
     
    Exit Function
    
Erro_NFEntCom_Preenche_GridItens:

    NFEntCom_Preenche_GridItens = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179131)
     
    End Select
     
    Exit Function

End Function

'####################################################################
'Inserido por Wagner
Function NFEntCom_Traz_NFiscal_Tela(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim dTotalMoeda As Double

On Error GoTo Erro_NFEntCom_Traz_NFiscal_Tela

    lErro = Calcula_Total_Moeda(objCT, dTotalMoeda)
    If lErro <> SUCESSO Then gError 140323

    TotalMoeda(objCT).Text = Format(dTotalMoeda, "STANDARD")
    
    NFEntCom_Traz_NFiscal_Tela = SUCESSO
     
    Exit Function
    
Erro_NFEntCom_Traz_NFiscal_Tela:

    NFEntCom_Traz_NFiscal_Tela = gErr
     
    Select Case gErr
    
        Case 140323
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179132)
     
    End Select
     
    Exit Function
    
End Function

Function NFEntCom_Valida(ByVal objCT As Object) As Long

Dim lErro As Long
Dim dTotalMoeda As Double

On Error GoTo Erro_NFEntCom_Valida

    lErro = Calcula_Total_Moeda(objCT, dTotalMoeda)
    If lErro <> SUCESSO Then gError 140326

    If Abs(StrParaDbl(TotalMoeda(objCT).Text) - dTotalMoeda) > DELTA_VALORMONETARIO Then gError 140327
    
    NFEntCom_Valida = SUCESSO
     
    Exit Function
    
Erro_NFEntCom_Valida:

    NFEntCom_Valida = gErr
     
    Select Case gErr
    
        Case 140326
        
        Case 140327
            Call Rotina_Erro(vbOKOnly, "ERRO_TOTALMOEDA_DIFERENTE", gErr, Format(StrParaDbl(TotalMoeda(objCT).Text), "STANDARD"), Format(dTotalMoeda, "STANDARD"))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179133)
     
    End Select
     
    Exit Function
     
End Function

Public Sub TotalMoeda_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TotalMoeda_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TotalMoeda_Validate

    'Verifica se o Valor Frete foi preenchido
    If Len(Trim(TotalMoeda(objCT).Text)) > 0 Then
    
        'faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(TotalMoeda(objCT).Text)
        If lErro <> SUCESSO Then gError 140321
              
        'Coloca o Valor Formatado na tela
        TotalMoeda(objCT).Text = Format(TotalMoeda(objCT).Text, "Standard")

    End If
    
    Exit Sub

Erro_TotalMoeda_Validate:

    Cancel = True

    Select Case gErr
    
        Case 140321
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179134)

    End Select

    Exit Sub

End Sub

Function Calcula_Total_Moeda(ByVal objCT As Object, dTotalMoeda As Double) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dValorMoeda As Double
Dim dTotalMoedaItem As Double
Dim dQuantidade As Double
Dim dPercDesconto As Double
Dim dDesconto As Double
Dim dTotalProdutoMoeda As Double

On Error GoTo Erro_Calcula_Total_Moeda

    dTotalMoeda = 0

    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes

        If Len(Trim(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))) <> 0 Then

            dValorMoeda = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_ValorUnitarioMoeda_Col))
            dQuantidade = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Quantidade_Col))

            If Len(Trim(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PercDesc_Col))) > 0 Then
                dPercDesconto = PercentParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PercDesc_Col))
                dDesconto = StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Desconto_Col))
            End If
            
            'Calcula o Valor Real
            Call objCT.ValorReal_Calcula(dQuantidade, dValorMoeda, dPercDesconto, dDesconto, dTotalMoedaItem)
            
            If dValorMoeda <> 0 Then
                dTotalProdutoMoeda = dTotalProdutoMoeda + dTotalMoedaItem
            Else
                dTotalProdutoMoeda = dTotalProdutoMoeda + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))
            End If

        End If

    Next
    
    dTotalMoeda = StrParaDbl(objCT.total.Text) - StrParaDbl(objCT.subtotal.Caption) + dTotalProdutoMoeda
    
    Calcula_Total_Moeda = SUCESSO
     
    Exit Function
    
Erro_Calcula_Total_Moeda:

    Calcula_Total_Moeda = gErr
     
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179135)
     
    End Select
     
    Exit Function
    
End Function
'####################################################################



