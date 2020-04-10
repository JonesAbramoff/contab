VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- inicio dos properties get dos controles da tela
'###################################################################
'Inserido por wagner
'Public Property Get PercDesconto(ByVal objCT As Object) As Object
'     Set PercDesconto = objCT.objUserControl.Controls("PercDesconto")
'End Property
'###################################################################

Public Property Get Moeda(ByVal objCT As Object) As Object
     Set Moeda = objCT.objUserControl.Controls("Moeda")
End Property

Public Property Get TipoDesc1(ByVal objCT As Object) As Object
     Set TipoDesc1 = objCT.objUserControl.Controls("TipoDesc1")
End Property

Public Property Get TipoDesc2(ByVal objCT As Object) As Object
     Set TipoDesc2 = objCT.objUserControl.Controls("TipoDesc2")
End Property

Public Property Get TipoDesc3(ByVal objCT As Object) As Object
     Set TipoDesc3 = objCT.objUserControl.Controls("TipoDesc3")
End Property

Public Property Get PercentDesc2(ByVal objCT As Object) As Object
     Set PercentDesc2 = objCT.objUserControl.Controls("PercentDesc2")
End Property

Public Property Get PercentDesc3(ByVal objCT As Object) As Object
     Set PercentDesc3 = objCT.objUserControl.Controls("PercentDesc3")
End Property

Public Property Get ValorFreteAux(ByVal objCT As Object) As Object
     Set ValorFreteAux = objCT.objUserControl.Controls("ValorFreteAux")
End Property

Public Property Get ValorSeguroAux(ByVal objCT As Object) As Object
     Set ValorSeguroAux = objCT.objUserControl.Controls("ValorSeguroAux")
End Property

Public Property Get ValorDespesasAux(ByVal objCT As Object) As Object
     Set ValorDespesasAux = objCT.objUserControl.Controls("ValorDespesasAux")
End Property

Public Property Get SomaFrete(ByVal objCT As Object) As Object
     Set SomaFrete = objCT.objUserControl.Controls("SomaFrete")
End Property

Public Property Get ValorFreteAuxInd(ByVal objCT As Object, ByVal Index As Integer) As Object
Dim objAux As Object
    Set objAux = objCT.objUserControl.Controls("ValorFreteAux")
    Set ValorFreteAuxInd = objAux(Index)
End Property

Public Property Get ValorSeguroAuxInd(ByVal objCT As Object, ByVal Index As Integer) As Object
Dim objAux As Object
    Set objAux = objCT.objUserControl.Controls("ValorSeguroAux")
    Set ValorSeguroAuxInd = objAux(Index)
End Property

Public Property Get ValorDespesasAuxInd(ByVal objCT As Object, ByVal Index As Integer) As Object
Dim objAux As Object
    Set objAux = objCT.objUserControl.Controls("ValorDespesasAux")
    Set ValorDespesasAuxInd = objAux(Index)
End Property

Public Property Get SomaFreteInd(ByVal objCT As Object, ByVal Index As Integer) As Object
Dim objAux As Object
    Set objAux = objCT.objUserControl.Controls("SomaFrete")
    Set SomaFreteInd = objAux(Index)
End Property

'--- fim dos properties get dos controles da tela

Public Sub Moeda_Change(ByVal objCT As Object)
    
    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Moeda_Click(ByVal objCT As Object)

End Sub

Public Function PV_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. Pedida")
    objGridInt.colColuna.Add ("Quant. Cancelada")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desc 1")
    objGridInt.colColuna.Add ("Tipo Desc 1")
    objGridInt.colColuna.Add ("% Desc 2")
    objGridInt.colColuna.Add ("Tipo Desc 2")
    objGridInt.colColuna.Add ("% Desc 3")
    objGridInt.colColuna.Add ("Tipo Desc 3")
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

    'Controles que participam do Grid
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoProduto.Name)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)
    objGridInt.colCampo.Add (objCT.QuantCancelada.Name)
    objGridInt.colCampo.Add (objCT.PrecoUnitario.Name)
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (TipoDesc1(objCT).Name)
    objGridInt.colCampo.Add (PercentDesc2(objCT).Name)
    objGridInt.colCampo.Add (TipoDesc2(objCT).Name)
    objGridInt.colCampo.Add (PercentDesc3(objCT).Name)
    objGridInt.colCampo.Add (TipoDesc3(objCT).Name)
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
    
    'Colunas do Grid
    objCT.iGrid_ItemProduto_Col = 0
    objCT.iGrid_Produto_Col = 1
    objCT.iGrid_DescProduto_Col = 2
    objCT.iGrid_UnidadeMed_Col = 3
    objCT.iGrid_QuantPedida_Col = 4
    objCT.iGrid_QuantCancel_Col = 5
    objCT.iGrid_ValorUnitario_Col = 6
    objCT.iGrid_PercDesc_Col = 7
    
    objCT.gobjInfoUsu.iGrid_TipoDesc1_Col = 8
    objCT.gobjInfoUsu.iGrid_PercDesc2_Col = 9
    objCT.gobjInfoUsu.iGrid_TipoDesc2_Col = 10
    objCT.gobjInfoUsu.iGrid_PercDesc3_Col = 11
    objCT.gobjInfoUsu.iGrid_TipoDesc3_Col = 12
    
    objCT.iGrid_Desconto_Col = 13
    objCT.iGrid_PrecoTotal_Col = 14
    objCT.iGrid_PrecoTotalB_Col = 15
    objCT.iGrid_DataEntrega_Col = 16
    objCT.iGrid_QuantReservada_Col = 17
    objCT.iGrid_QuantFaturada_Col = 18
    objCT.iGrid_Prioridade_Col = 19
    
    If gobjFAT.iExibePercFatPV = MARCADO Then

        objCT.iGrid_PercentMenosReceb_Col = 20
        objCT.iGrid_PercentMaisReceb_Col = 21
        objCT.iGrid_RecebForaFaixa_Col = 22
        objCT.iGrid_QuantFatAMais_Col = 23

    End If

    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS_HICARE + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

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

Public Function PV_CargaPosFormLoad(ByVal objCT As Object, bTrazendoDoc As Boolean) As Long
    
Dim lErro As Long

On Error GoTo Erro_PV_CargaPosFormLoad

    'carregar moedas
    lErro = Carrega_Moeda(objCT)
    If lErro <> SUCESSO Then gError 124170
    
    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPO_DESCONTO, TipoDesc1(objCT), False, False)
    If lErro <> SUCESSO Then gError 124175
    TipoDesc1(objCT).AddItem "", 0
    TipoDesc1(objCT).ItemData(0) = 0
    
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPO_DESCONTO, TipoDesc2(objCT), False, False)
    If lErro <> SUCESSO Then gError 124176
    TipoDesc2(objCT).AddItem "", 0
    TipoDesc2(objCT).ItemData(0) = 0
    
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPO_DESCONTO, TipoDesc3(objCT), False, False)
    If lErro <> SUCESSO Then gError 124177
    TipoDesc3(objCT).AddItem "", 0
    TipoDesc3(objCT).ItemData(0) = 0
    
    PV_CargaPosFormLoad = SUCESSO
         
    Exit Function
    
Erro_PV_CargaPosFormLoad:

    PV_CargaPosFormLoad = gErr
     
    Select Case gErr
          
        Case 124170
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179317)
     
    End Select
     
    Exit Function

End Function

Public Function PV_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long

    'liberar objetos globais referentes a customizacao
    
End Function

Public Sub ValorFreteAux_Change(ByVal objCT As Object, Index As Integer)
    objCT.iAlterado = REGISTRO_ALTERADO
    objCT.giValorFreteAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorFreteAux_Validate(ByVal objCT As Object, Index As Integer, Cancel As Boolean)
Dim dValorFrete As Double
Dim lErro As Long

On Error GoTo Erro_ValorFreteAux_Validate

    If objCT.giValorFreteAlterado = 0 Then Exit Sub

    dValorFrete = 0

    If Len(Trim(ValorFreteAuxInd(objCT, Index).Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(ValorFreteAuxInd(objCT, Index).Text)
        If lErro <> SUCESSO Then gError 26650

        dValorFrete = CDbl(ValorFreteAuxInd(objCT, Index).Text)

        ValorFreteAuxInd(objCT, Index).Text = Format(dValorFrete, "Standard")

    End If

    lErro = TotalFrete_Atualiza(objCT)
    If lErro <> SUCESSO Then gError 124171
    
    objCT.giValorFreteAlterado = 0

    Exit Sub

Erro_ValorFreteAux_Validate:

    Cancel = True

    Select Case gErr

        Case 26650, 124171

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179318)

    End Select

    Exit Sub

End Sub

Public Sub ValorSeguroAux_Change(ByVal objCT As Object, Index As Integer)
    objCT.iAlterado = REGISTRO_ALTERADO
    objCT.giValorSeguroAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorSeguroAux_Validate(ByVal objCT As Object, Index As Integer, Cancel As Boolean)

Dim dValorSeguro As Double
Dim lErro As Long

On Error GoTo Erro_ValorSeguroAux_Validate

    If objCT.giValorSeguroAlterado = 0 Then Exit Sub

    dValorSeguro = 0

    If Len(Trim(ValorSeguroAuxInd(objCT, Index).Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(ValorSeguroAuxInd(objCT, Index).Text)
        If lErro <> SUCESSO Then gError 26650

        dValorSeguro = CDbl(ValorSeguroAuxInd(objCT, Index).Text)

        ValorSeguroAuxInd(objCT, Index).Text = Format(dValorSeguro, "Standard")

    End If

    lErro = TotalSeguro_Atualiza(objCT)
    If lErro <> SUCESSO Then gError 124171
    
    objCT.giValorSeguroAlterado = 0

    Exit Sub

Erro_ValorSeguroAux_Validate:

    Cancel = True

    Select Case gErr

        Case 26650, 124171

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179319)

    End Select

    Exit Sub

End Sub

Public Sub ValorDespesasAux_Change(ByVal objCT As Object, Index As Integer)
    objCT.iAlterado = REGISTRO_ALTERADO
    objCT.giValorDespesasAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorDespesasAux_Validate(ByVal objCT As Object, Index As Integer, Cancel As Boolean)

Dim dValorDespesas As Double
Dim lErro As Long

On Error GoTo Erro_ValorDespesasAux_Validate

    If objCT.giValorDespesasAlterado = 0 Then Exit Sub

    dValorDespesas = 0

    If Len(Trim(ValorDespesasAuxInd(objCT, Index).Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(ValorDespesasAuxInd(objCT, Index).Text)
        If lErro <> SUCESSO Then gError 26650

        dValorDespesas = CDbl(ValorDespesasAuxInd(objCT, Index).Text)

        ValorDespesasAuxInd(objCT, Index).Text = Format(dValorDespesas, "Standard")

    End If

    lErro = TotalDespesas_Atualiza(objCT)
    If lErro <> SUCESSO Then gError 124171
    
    objCT.giValorDespesasAlterado = 0

    Exit Sub

Erro_ValorDespesasAux_Validate:

    Cancel = True

    Select Case gErr

        Case 26650, 124171

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179320)

    End Select

    Exit Sub

End Sub

Private Function TotalFrete_Atualiza(ByVal objCT As Object) As Long

Dim lErro As Long
Dim dValorFrete As Double

On Error GoTo Erro_TotalFrete_Atualiza

    dValorFrete = IIf(SomaFreteInd(objCT, 1).Value = vbChecked, StrParaDbl(ValorFreteAuxInd(objCT, 1).Text), 0) + IIf(SomaFreteInd(objCT, 2).Value = vbChecked, StrParaDbl(ValorFreteAuxInd(objCT, 2).Text), 0)
    
    'Para tributação
    objCT.gobjPedidoVenda.dValorFrete = dValorFrete
        
    If dValorFrete = 0 Then
        objCT.ValorFrete.Text = ""
    Else
        objCT.ValorFrete.Text = Format(dValorFrete, "Standard")
    End If
    
    'Recalcula valor total
    lErro = objCT.ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 51040

    TotalFrete_Atualiza = SUCESSO
     
    Exit Function
    
Erro_TotalFrete_Atualiza:

    TotalFrete_Atualiza = gErr
     
    Select Case gErr
          
        Case 51040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179321)
     
    End Select
     
    Exit Function

End Function

Public Sub TipoDesc1_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDesc1_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)

End Sub

Public Sub TipoDesc1_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)

End Sub

Public Sub TipoDesc1_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = TipoDesc1(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub TipoDesc1_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Function PV_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case objCT.gobjInfoUsu.iGrid_TipoDesc1_Col
            lErro = Saida_Celula_Padrao(objGridInt, TipoDesc1(objCT))
            If lErro <> SUCESSO Then gError 124164

        Case objCT.gobjInfoUsu.iGrid_TipoDesc2_Col
            lErro = Saida_Celula_Padrao(objGridInt, TipoDesc2(objCT))
            If lErro <> SUCESSO Then gError 124165

        Case objCT.gobjInfoUsu.iGrid_TipoDesc3_Col
            lErro = Saida_Celula_Padrao(objGridInt, TipoDesc3(objCT))
            If lErro <> SUCESSO Then gError 124166

        Case objCT.iGrid_PercDesc_Col
            lErro = Saida_Celula_PercDesc(objCT, objGridInt, objCT.PercentDesc, objGridInt.objGrid.Col)
            If lErro <> SUCESSO Then gError 124167
        
        Case objCT.gobjInfoUsu.iGrid_PercDesc2_Col
            lErro = Saida_Celula_PercDesc(objCT, objGridInt, PercentDesc2(objCT), objGridInt.objGrid.Col)
            If lErro <> SUCESSO Then gError 124167

        Case objCT.gobjInfoUsu.iGrid_PercDesc3_Col
            lErro = Saida_Celula_PercDesc(objCT, objGridInt, PercentDesc3(objCT), objGridInt.objGrid.Col)
            If lErro <> SUCESSO Then gError 124168

    End Select

    PV_Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_PV_Saida_Celula_GridItens:

    PV_Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 124164 To 124168

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179322)

    End Select

    Exit Function

End Function

'??? mover p/admlib ou outro
Function Saida_Celula_Padrao(ByVal objGridInt As AdmGrid, objControle As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 124163
    
    Saida_Celula_Padrao = SUCESSO
     
    Exit Function
    
Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr
     
    Select Case gErr
          
        Case 124163
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179323)
     
    End Select
     
    Exit Function

End Function

Public Sub TipoDesc2_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDesc2_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)

End Sub

Public Sub TipoDesc2_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)

End Sub

Public Sub TipoDesc2_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = TipoDesc2(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub TipoDesc2_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDesc3_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDesc3_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)

End Sub

Public Sub TipoDesc3_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)

End Sub

Public Sub TipoDesc3_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = TipoDesc3(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub TipoDesc3_Click(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentDesc2_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentDesc2_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)

End Sub

Public Sub PercentDesc2_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)

End Sub

Public Sub PercentDesc2_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = PercentDesc2(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub PercentDesc3_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentDesc3_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.objGridItens)

End Sub

Public Sub PercentDesc3_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGridItens)

End Sub

Public Sub PercentDesc3_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGridItens.objControle = PercentDesc3(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Function PV_Move_Pedido_Memoria(ByVal objCT As Object, objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Move_Pedido_Memoria

    objPedidoVenda.iMoeda = Codigo_Extrai(Moeda(objCT).Text)
    
    Set objPedidoVenda.objInfoUsu = New ClassPedidoDeVendaHic
    
    With objPedidoVenda.objInfoUsu

        .dValorFrete1 = StrParaDbl(ValorFreteAuxInd(objCT, 1).Text)
        .dValorSeguro1 = StrParaDbl(ValorSeguroAuxInd(objCT, 1).Text)
        .dValorOutrasDespesas1 = StrParaDbl(ValorDespesasAuxInd(objCT, 1).Text)
        .dValorFrete2 = StrParaDbl(ValorFreteAuxInd(objCT, 2).Text)
        .dValorSeguro2 = StrParaDbl(ValorSeguroAuxInd(objCT, 2).Text)
        .dValorOutrasDespesas2 = StrParaDbl(ValorDespesasAuxInd(objCT, 2).Text)
        .iFlagCompl1 = SomaFreteInd(objCT, 1).Value
        .iFlagCompl2 = SomaFreteInd(objCT, 2).Value
    
    End With
    
    PV_Move_Pedido_Memoria = SUCESSO
     
    Exit Function
    
Erro_PV_Move_Pedido_Memoria:

    PV_Move_Pedido_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179324)
     
    End Select
     
    Exit Function

End Function

Private Function Carrega_Moeda(ByVal objCT As Object) As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 103371
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        Moeda(objCT).AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        Moeda(objCT).ItemData(Moeda(objCT).NewIndex) = objMoeda.iCodigo
    
    Next

    Moeda(objCT).ListIndex = 0
    
    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 103371
        
        Case 103372
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179325)
    
    End Select

End Function

Public Sub SomaFrete_Click(ByVal objCT As Object, Index As Integer)

Dim lErro As Long

On Error GoTo Erro_SomaFrete_Click

    objCT.iAlterado = REGISTRO_ALTERADO
    
    If Not objCT.gbCarregandoTela Then
    
        lErro = TotalFrete_Atualiza(objCT)
        If lErro <> SUCESSO Then gError 124172
    
        lErro = TotalSeguro_Atualiza(objCT)
        If lErro <> SUCESSO Then gError 124173
    
        lErro = TotalDespesas_Atualiza(objCT)
        If lErro <> SUCESSO Then gError 124174
    
    End If
    
    Exit Sub
     
Erro_SomaFrete_Click:

    Select Case gErr
          
        Case 124172 To 124174
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179326)
     
    End Select
     
    Exit Sub

End Sub

Private Function TotalSeguro_Atualiza(ByVal objCT As Object) As Long

Dim lErro As Long
Dim dValorSeguro As Double

On Error GoTo Erro_TotalSeguro_Atualiza

    dValorSeguro = IIf(SomaFreteInd(objCT, 1).Value = vbChecked, StrParaDbl(ValorSeguroAuxInd(objCT, 1).Text), 0) + IIf(SomaFreteInd(objCT, 2).Value = vbChecked, StrParaDbl(ValorSeguroAuxInd(objCT, 2).Text), 0)
    
    'Para tributação
    objCT.gobjPedidoVenda.dValorSeguro = dValorSeguro
        
    If dValorSeguro = 0 Then
        objCT.ValorSeguro.Text = ""
    Else
        objCT.ValorSeguro.Text = Format(dValorSeguro, "Standard")
    End If
    
    'Recalcula valor total
    lErro = objCT.ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 51040

    TotalSeguro_Atualiza = SUCESSO
     
    Exit Function
    
Erro_TotalSeguro_Atualiza:

    TotalSeguro_Atualiza = gErr
     
    Select Case gErr
          
        Case 51040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179327)
     
    End Select
     
    Exit Function

End Function

Private Function TotalDespesas_Atualiza(ByVal objCT As Object) As Long

Dim lErro As Long
Dim dValorDespesas As Double

On Error GoTo Erro_TotalDespesas_Atualiza

    dValorDespesas = IIf(SomaFreteInd(objCT, 1).Value = vbChecked, StrParaDbl(ValorDespesasAuxInd(objCT, 1).Text), 0) + IIf(SomaFreteInd(objCT, 2).Value = vbChecked, StrParaDbl(ValorDespesasAuxInd(objCT, 2).Text), 0)
    
    'Para tributação
    objCT.gobjPedidoVenda.dValorOutrasDespesas = dValorDespesas
        
    If dValorDespesas = 0 Then
        objCT.ValorDespesas.Text = ""
    Else
        objCT.ValorDespesas.Text = Format(dValorDespesas, "Standard")
    End If
    
    'Recalcula valor total
    lErro = objCT.ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 51040

    TotalDespesas_Atualiza = SUCESSO
     
    Exit Function
    
Erro_TotalDespesas_Atualiza:

    TotalDespesas_Atualiza = gErr
     
    Select Case gErr
          
        Case 51040
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179328)
     
    End Select
     
    Exit Function

End Function


Public Function PV_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iCaminho As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Rotina_Grid_Enable

    Select Case objControl.Name
        
        Case TipoDesc1(objCT).Name, TipoDesc2(objCT).Name, TipoDesc3(objCT).Name, PercentDesc2(objCT).Name, PercentDesc3(objCT).Name
            'Se o produto estiver preenchido, habilita o controle
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179329)
     
    End Select
     
    Exit Function

End Function

Function PV_Move_GridItem_Memoria(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objItemPedido As ClassItemPedido, ByVal iIndice As Integer, ByVal sProduto1 As String) As Long

Dim lErro As Long, lItemData As Long

On Error GoTo Erro_PV_Move_GridItem_Memoria

    Call Combo_Obtem_ItemData(TipoDesc1(objCT), objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_TipoDesc1_Col), lItemData)
    objItemPedido.iTipoDesc1 = lItemData
    
    Call Combo_Obtem_ItemData(TipoDesc2(objCT), objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_TipoDesc2_Col), lItemData)
    objItemPedido.iTipoDesc2 = lItemData
    
    Call Combo_Obtem_ItemData(TipoDesc3(objCT), objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_TipoDesc3_Col), lItemData)
    objItemPedido.iTipoDesc3 = lItemData
    
    objItemPedido.dPercDesc1 = PercentParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PercDesc_Col))
    objItemPedido.dPercDesc2 = PercentParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercDesc2_Col))
    objItemPedido.dPercDesc3 = PercentParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercDesc3_Col))
    
    PV_Move_GridItem_Memoria = SUCESSO
     
    Exit Function
    
Erro_PV_Move_GridItem_Memoria:

    PV_Move_GridItem_Memoria = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179330)
     
    End Select
     
    Exit Function

End Function

Function PV_Carrega_Grid_Itens(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal bProcessaReserva As Boolean) As Long

Dim lErro As Long, iIndice As Integer, objItemPedido As ClassItemPedido, sTipo As String

On Error GoTo Erro_PV_Carrega_Grid_Itens

    For iIndice = 1 To objPedidoVenda.colItensPedido.Count
       
        Set objItemPedido = objPedidoVenda.colItensPedido(iIndice)
        
        lErro = Combo_Obtem_String(TipoDesc1(objCT), objItemPedido.iTipoDesc1, sTipo)
        If lErro <> SUCESSO Then gError 124182
        objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_TipoDesc1_Col) = sTipo
        
        lErro = Combo_Obtem_String(TipoDesc2(objCT), objItemPedido.iTipoDesc2, sTipo)
        If lErro <> SUCESSO Then gError 124183
        objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_TipoDesc2_Col) = sTipo
        
        lErro = Combo_Obtem_String(TipoDesc3(objCT), objItemPedido.iTipoDesc3, sTipo)
        If lErro <> SUCESSO Then gError 124184
        objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_TipoDesc3_Col) = sTipo
        
        objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PercDesc_Col) = Format(objItemPedido.dPercDesc1, "Percent")
        objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercDesc2_Col) = Format(objItemPedido.dPercDesc2, "Percent")
        objCT.GridItens.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercDesc3_Col) = Format(objItemPedido.dPercDesc3, "Percent")
    
    Next
    
    PV_Carrega_Grid_Itens = SUCESSO
     
    Exit Function
    
Erro_PV_Carrega_Grid_Itens:

    PV_Carrega_Grid_Itens = gErr
     
    Select Case gErr
          
        Case 124182 To 124184
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179331)
     
    End Select
     
    Exit Function

End Function

Function PV_Traz_Pedido_Tela(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Traz_Pedido_Tela

    Call Combo_Seleciona_ItemData(Moeda(objCT), objPedidoVenda.iMoeda)
        
    lErro = CF("PedidoDeVenda_Le_InfoUsu", objPedidoVenda)
    If lErro <> SUCESSO Then gError 124189
    
    ValorFreteAuxInd(objCT, 1).Text = Format(objPedidoVenda.objInfoUsu.dValorFrete1, "Standard")
    ValorSeguroAuxInd(objCT, 1).Text = Format(objPedidoVenda.objInfoUsu.dValorSeguro1, "Standard")
    ValorDespesasAuxInd(objCT, 1).Text = Format(objPedidoVenda.objInfoUsu.dValorOutrasDespesas1, "Standard")
    ValorFreteAuxInd(objCT, 2).Text = Format(objPedidoVenda.objInfoUsu.dValorFrete2, "Standard")
    ValorSeguroAuxInd(objCT, 2).Text = Format(objPedidoVenda.objInfoUsu.dValorSeguro2, "Standard")
    ValorDespesasAuxInd(objCT, 2).Text = Format(objPedidoVenda.objInfoUsu.dValorOutrasDespesas2, "Standard")
        
    SomaFreteInd(objCT, 1).Value = objPedidoVenda.objInfoUsu.iFlagCompl1
    SomaFreteInd(objCT, 2).Value = objPedidoVenda.objInfoUsu.iFlagCompl2
       
    PV_Traz_Pedido_Tela = SUCESSO
     
    Exit Function
    
Erro_PV_Traz_Pedido_Tela:

    PV_Traz_Pedido_Tela = gErr
     
    Select Case gErr
          
        Case 124189
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179332)
     
    End Select
     
    Exit Function
    
End Function

Function PV_Cliente_Validate(ByVal objCT As Object, ByVal lCliente As Long) As Long

Dim lErro As Long, dSaldoAtraso As Double

On Error GoTo Erro_PV_Cliente_Validate

'    lErro = CF("Cliente_Obtem_SaldoAtraso", lCliente, dSaldoAtraso)
'    If lErro <> SUCESSO Then gError 124198
'
'    If dSaldoAtraso <> 0 Then gError 124199

    lErro = CF("Libera_BloqueioCliente", lCliente)
    If lErro <> SUCESSO Then gError 132307
        
    PV_Cliente_Validate = SUCESSO
     
    Exit Function
    
Erro_PV_Cliente_Validate:

    PV_Cliente_Validate = gErr
     
    Select Case gErr
          
'        Case 124198
'
'        Case 124199
'            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_COM_ATRASO", gErr)
'
        Case 132307
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179333)
     
    End Select
     
    Exit Function

End Function

Function Saida_Celula_PercDesc(ByVal objCT As Object, objGridInt As AdmGrid, ByVal objControle As Object, ByVal iGrid As Integer) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dQuantidade As Double
Dim dQuantidadeCancelada As Double
Dim sValorPercAnterior As String

On Error GoTo Erro_Saida_Celula_PercDesc

    Set objGridInt.objControle = objControle

    If Len(objControle.Text) > 0 Then
    
        'Critica a porcentagem
        lErro = Porcentagem_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 26694

        dPercentDesc = CDbl(objControle.Text)
        
        If Format(dPercentDesc, "#0.#0\%") <> objCT.GridItens.TextMatrix(objCT.GridItens.Row, iGrid) Then
            
            'se for igual a 100% -> erro
            If dPercentDesc = 100 Then gError 26695

            objControle.Text = Format(dPercentDesc, "Fixed")
        
        End If

    End If

    sValorPercAnterior = objCT.GridItens.TextMatrix(objCT.GridItens.Row, iGrid)
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26696
    
    'Se foi alterada
    If Format(dPercentDesc, "#0.#0\%") <> sValorPercAnterior Then
    
        'Recalcula o preço total
        Call CF2(objCT, "PV_PrecoTotal_Calcula", objCT.GridItens.Row)
        lErro = objCT.ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 51044

    End If

    Saida_Celula_PercDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercDesc:

    Saida_Celula_PercDesc = gErr

    Select Case gErr

        Case 26694, 26696, 51044
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26695
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179334)

    End Select

    Exit Function

End Function

Public Function PV_PrecoTotal_Calcula(ByVal objCT As Object, ByVal iLinha As Integer) As Long

Dim dPrecoTotal As Double
Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dQuantidadeCancelada As Double
Dim dDesconto As Double
Dim dPercDesc1 As Double, dPercDesc2 As Double, dPercDesc3 As Double
Dim lTamanho As Long
Dim dValorTotal As Double
Dim iIndice As Integer
Dim dValorDesconto As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_PV_PrecoTotal_Calcula

    'Quantidades e preço unitário
    dPrecoUnitario = StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_ValorUnitario_Col))
    dQuantidade = StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_QuantPedida_Col))
    dQuantidadeCancelada = StrParaDbl(objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_QuantCancel_Col))

    'Cálculo do desconto
    lTamanho = Len(Trim(objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_PercDesc_Col)))
    If lTamanho > 0 Then
        dPercDesc1 = CDbl(Format(objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_PercDesc_Col), "General Number"))
    Else
        dPercDesc1 = 0
    End If

    lTamanho = Len(Trim(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PercDesc2_Col)))
    If lTamanho > 0 Then
        dPercDesc2 = CDbl(Format(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PercDesc2_Col), "General Number"))
    Else
        dPercDesc2 = 0
    End If

    lTamanho = Len(Trim(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PercDesc3_Col)))
    If lTamanho > 0 Then
        dPercDesc3 = CDbl(Format(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PercDesc3_Col), "General Number"))
    Else
        dPercDesc3 = 0
    End If

    dPrecoTotal = dPrecoUnitario * (dQuantidade - dQuantidadeCancelada)

    'Se percentual for >0 tira o desconto
    If (dPercDesc1 + dPercDesc2 + dPercDesc3) > 0 Then
        dDesconto = dPrecoTotal * (1 - ((1 - dPercDesc1) * (1 - dPercDesc2) * (1 - dPercDesc3)))
    End If
    
    dPrecoTotalReal = dPrecoTotal - dDesconto

    'Coloca valor do desconto no Grid
    If dDesconto > 0 Then
        objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Desconto_Col) = Format(dDesconto, "Standard")
    Else
        objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Desconto_Col) = ""
    End If

    'Coloca preco total do ítem no grid
    objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_PrecoTotal_Col) = Format(dPrecoTotalReal, "Standard")

    'Calcula a soma dos valores de produtos
    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
        If Len(Trim(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))) > 0 Then
            If CDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))
        End If
    Next

    If objCT.gdDesconto > 0 And objCT.giValorDescontoManual = 0 Then 'Alterado por Wagner
        dValorDesconto = objCT.gdDesconto * dValorTotal
    ElseIf Len(Trim(objCT.ValorDesconto.Text)) > 0 And IsNumeric(objCT.ValorDesconto.Text) Then
        dValorDesconto = CDbl(objCT.ValorDesconto.Text)
    End If
    dValorTotal = dValorTotal - dValorDesconto
    
    'Verifica se o valor de desconto é maior que o valor dos produtos
    If dValorTotal < 0 And dValorDesconto > 0 Then

        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", dValorDesconto, dValorTotal + dValorDesconto) 'Alterado por Wagner

        objCT.gdDesconto = 0
        objCT.ValorDesconto.Text = ""
        objCT.giValorDescontoAlterado = 0
        dValorDesconto = 0

        'Para tributação
        objCT.gobjPedidoVenda.dValorDesconto = dValorDesconto

        'Faz a atualização dos valores da tributação
        lErro = objCT.AtualizarTributacao()
        If lErro <> SUCESSO Then gError 56913
    
        'Calcula a soma dos valores de produtos
        dValorTotal = 0
        For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
            If Len(Trim(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))) > 0 Then
                If CDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))
            End If
        Next
    
    End If

    'Coloca valor total dos produtos na tela
    objCT.ValorProdutos.Caption = Format(dValorTotal, "Standard")

    '#####################################################
    'Inserido por Wagner
    lErro = CF2(objCT, "PV_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132008
    '#####################################################

    'Altera ítem na coleção de ítens de gobjPedidoVenda que dá suporte p/ Tributação
    Call objCT.Tributacao_Alteracao_Item_Grid(iLinha)

    PV_PrecoTotal_Calcula = SUCESSO
    
    Exit Function

Erro_PV_PrecoTotal_Calcula:

    PV_PrecoTotal_Calcula = gErr
    
    Select Case gErr

        Case 56913 'tratados nas rotinas chamadas
        
        Case 132008

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179335)

    End Select

    Exit Function

End Function

'###########################################################################
'Inserido por Wagner
Function PV_Botao_Estoque(ByVal objCT As Object, ByVal objEstoqueProd As ClassEstoqueProduto, ByVal colSelecao As Collection, sSelecao As String) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim iAlmoxarifado As Integer
Dim sInterroga As String

On Error GoTo Erro_PV_Botao_Estoque

    
    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
    
        lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 131866
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            If Len(Trim(sInterroga)) > 0 Then
                sInterroga = sInterroga & ", ?"
            Else
                sInterroga = "?"
            End If
    
            '*********** INCLUIDO PARA O TRATAMENTO DE ITEM DE GRADE ****************
            If objCT.gobjPedidoVenda.colItensPedido(iIndice).iPossuiGrade = MARCADO Then
                Call objCT.BotaoReserva_Click
                Exit Function
            End If
    
            objEstoqueProd.sProduto = sProdutoFormatado
    
            colSelecao.Add sProdutoFormatado
    
            lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, sProdutoFormatado, iAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 131867
    
            If lErro = SUCESSO Then objEstoqueProd.iAlmoxarifado = iAlmoxarifado
    
    
        End If
        
    Next
    
    colSelecao.Add "%"
    
    sSelecao = "Produto IN (" & sInterroga & ")"

    PV_Botao_Estoque = SUCESSO
    
    Exit Function

Erro_PV_Botao_Estoque:

    PV_Botao_Estoque = gErr

    Select Case gErr

        Case 131866 To 131867 'Tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179336)

    End Select

    Exit Function
    
End Function
'###########################################################################

'###########################################################################
'Inserido por Wagner 04/01/2005
Function PV_Muda_Tabulacao(ByVal objCT As Object, ByVal iKey As Integer, ByVal iLinha As Integer, ByVal iColuna As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Muda_Tabulacao

    If iKey = vbKeyReturn Then

        'Se Passou do terceiro desconto ou passou do segundo sem o mesmo estar preenchido
        '=> vai para próxima linha
        If (iColuna > objCT.gobjInfoUsu.iGrid_TipoDesc3_Col) Or _
        ((iColuna = objCT.gobjInfoUsu.iGrid_PercDesc3_Col) And _
        (Len(Trim(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_TipoDesc2_Col))) = 0) And _
        (Len(Trim(objCT.GridItens.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_PercDesc2_Col))) = 0)) Then

             SendKeys "{HOME}{DOWN}", True

        End If
                
    End If
    
    PV_Muda_Tabulacao = SUCESSO
    
    Exit Function

Erro_PV_Muda_Tabulacao:

    PV_Muda_Tabulacao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179337)

    End Select

    Exit Function
    
End Function

'Public Sub PercDesconto_Change(ByVal objCT As Object)
'
'    objCT.iAlterado = REGISTRO_ALTERADO
'    objCT.giValorDescontoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PercDesconto_Validate(ByVal objCT As Object, Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValorProdutos As Double
'Dim iIndice As Integer
'Dim dPercDesconto As Double
'
'On Error GoTo Erro_PercDesconto_Validate
'
'    'Verifica se o valor foi alterado
'    If objCT.giValorDescontoAlterado = 0 Then Exit Sub
'
'    'Vale o desconto que foi colocado aqui
'    objCT.giValorDescontoManual = 1
'
'    dPercDesconto = StrParaDbl(PercDesconto(objCT).Text)
'
'    'Calcula a soma dos valores de produtos
'    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
'        If Len(Trim(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))) > 0 Then
'            If StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col)) > 0 Then dValorProdutos = dValorProdutos + StrParaDbl(objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_PrecoTotal_Col))
'        End If
'    Next
'
'    'Verifica se o Valor está preenchido
'    If dPercDesconto > 0 Then
'
'        'Faz a Crítica do Valor digitado
'        lErro = Porcentagem_Critica(PercDesconto(objCT).Text)
'        If lErro <> SUCESSO Then gError 132000
'
'    End If
'
'    objCT.ValorDesconto.Text = Format((StrParaDbl(objCT.ValorProdutos.Caption) + StrParaDbl(objCT.ValorDesconto)) * (dPercDesconto / 100), "STANDARD")
'
'    'Recalcula o Valor do desconto
'    lErro = objCT.ValorDesconto_Validate(bSGECancelDummy)
'    If lErro <> SUCESSO Then gError 132001
'
'    objCT.giValorDescontoAlterado = 0
'
'    Exit Sub
'
'Erro_PercDesconto_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 132000 To 132001
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179338)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Function PV_ValorDesconto_Validate(ByVal objCT As Object) As Long

'Dim lErro As Long
'Dim dValorDesconto As Double
'Dim dValorProdutos As Double
'
'On Error GoTo Erro_PV_ValorDesconto_Validate
'
'    dValorDesconto = StrParaDbl(objCT.ValorDesconto)
'    dValorProdutos = StrParaDbl(objCT.ValorProdutos.Caption)
'
'    If 0 > dValorProdutos Then gError 132009
'
'    If dValorDesconto + dValorProdutos <> 0 Then
'        PercDesconto(objCT).Text = (dValorDesconto / (dValorDesconto + dValorProdutos)) * 100
'    Else
'        PercDesconto(objCT).Text = "0"
'    End If
'
'    PV_ValorDesconto_Validate = SUCESSO
'
'    Exit Function
'
'Erro_PV_ValorDesconto_Validate:
'
'    PV_ValorDesconto_Validate = gErr
'
'    Select Case gErr
'
'        Case 132009
'            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos + dValorDesconto)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179339)
'
'    End Select
'
'    Exit Function
    
End Function
'###########################################################################

'#########################################
'Inserido por Wagner 01/03/2005
Function PV_NF_DataSaida(ByVal objCT As Object, ByVal gobjGeracaoNFiscal As ClassGeracaoNFiscal) As Long
    gobjGeracaoNFiscal.dtDataSaida = DATA_NULA
End Function
'#########################################

