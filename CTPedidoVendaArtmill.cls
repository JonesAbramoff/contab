VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const TAB_Itens = 2

Public Property Get DescricaoItem2(ByVal objct As Object) As Object
     Set DescricaoItem2 = objct.objUserControl.Controls("DescricaoItem2")
End Property

Public Property Get DescricaoItem3(ByVal objct As Object) As Object
     Set DescricaoItem3 = objct.objUserControl.Controls("DescricaoItem3")
End Property

'precodesc
Public Property Get PrecoDesc(ByVal objct As Object) As Object
     Set PrecoDesc = objct.objUserControl.Controls("PrecoDesc")
End Property

Public Function PV_Form_Load(ByVal objct As Object) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Form_Load

'    Call CF2(objct, "NFCF_Form_Load_PrecoDesc", "Frame2", 3)
    
    PV_Form_Load = SUCESSO
    
    Exit Function
    
Erro_PV_Form_Load:

    PV_Form_Load = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184276)

    End Select
    
    Exit Function

End Function

Public Function PV_Inicializa_Grid_Itens(ByVal objct As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = objct

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Cor")
    objGridInt.colColuna.Add ("Detalhe")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Qtde Pedida")
    objGridInt.colColuna.Add ("Qtde Canc.")
    objGridInt.colColuna.Add ("Preço Unitário")
    'precodesc
    objGridInt.colColuna.Add ("Preço Desc.")
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
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (objct.Produto.Name)
    objGridInt.colCampo.Add (objct.DescricaoProduto.Name)
    objGridInt.colCampo.Add (DescricaoItem2(objct).Name)
    objGridInt.colCampo.Add (DescricaoItem3(objct).Name)
    objGridInt.colCampo.Add (objct.UnidadeMed.Name)
    objGridInt.colCampo.Add (objct.quantidade.Name)
    objGridInt.colCampo.Add (objct.QuantCancelada.Name)
    objGridInt.colCampo.Add (objct.PrecoUnitario.Name)
    'precodesc
    objGridInt.colCampo.Add (PrecoDesc(objct).Name)
    objGridInt.colCampo.Add (objct.PercentDesc.Name)
    objGridInt.colCampo.Add (objct.Desconto.Name)
    objGridInt.colCampo.Add (objct.PrecoTotal.Name)
    objGridInt.colCampo.Add (objct.PrecoTotalB.Name)
    objGridInt.colCampo.Add (objct.dataEntrega.Name)
    objGridInt.colCampo.Add (objct.QuantReservadaPV.Name)
    objGridInt.colCampo.Add (objct.QuantFaturada.Name)
    objGridInt.colCampo.Add (objct.Prioridade.Name)

    If gobjFAT.iExibePercFatPV = MARCADO Then
        objGridInt.colCampo.Add (objct.PercentMenosReceb.Name)
        objGridInt.colCampo.Add (objct.PercentMaisReceb.Name)
        objGridInt.colCampo.Add (objct.RecebForaFaixa.Name)
        objGridInt.colCampo.Add (objct.QuantFatAMais.Name)
    Else
        objct.PercentMenosReceb.left = POSICAO_FORA_TELA
        objct.PercentMaisReceb.left = POSICAO_FORA_TELA
        objct.RecebForaFaixa.left = POSICAO_FORA_TELA
        objct.QuantFatAMais.left = POSICAO_FORA_TELA
    End If
    
    With objct
        'Colunas do Grid
        .iGrid_ItemProduto_Col = 0
        .iGrid_Produto_Col = 1
        .iGrid_DescProduto_Col = 2
        .gobjInfoUsu.iGrid_DescProduto2_Col = 3
        .gobjInfoUsu.iGrid_DescProduto3_Col = 4
        .iGrid_UnidadeMed_Col = 5
        .iGrid_QuantPedida_Col = 6
        .iGrid_QuantCancel_Col = 7
        .iGrid_ValorUnitario_Col = 8
        'precodesc
        .gobjInfoUsu.iGrid_PrecoDesc_Col = 9
        .iGrid_PercDesc_Col = 10
        .iGrid_Desconto_Col = 11
        .iGrid_PrecoTotal_Col = 12
        .iGrid_PrecoTotalB_Col = 13
        .iGrid_DataEntrega_Col = 14
        .iGrid_QuantReservada_Col = 15
        .iGrid_QuantFaturada_Col = 16
    End With
    objct.iGrid_Prioridade_Col = 17
    
    If gobjFAT.iExibePercFatPV = MARCADO Then

        objct.iGrid_PercentMenosReceb_Col = 18
        objct.iGrid_PercentMaisReceb_Col = 19
        objct.iGrid_RecebForaFaixa_Col = 20
        objct.iGrid_QuantFatAMais_Col = 21

    End If
    
    objct.ComissaoItemPV.left = POSICAO_FORA_TELA
    objct.TabPrecoItemPV.left = POSICAO_FORA_TELA

    'Grid do GridInterno
    objGridInt.objGrid = objct.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 15

    'Largura da primeira coluna
    objct.GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    PV_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Function PV_Preenche_GridItens(ByVal objct As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objItemPedido As ClassItemPedido, ByVal objProduto As ClassProduto, ByVal iIndice As Integer, ByVal dPercDesc As Double) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Preenche_GridItens

    Call CF2(objct, "Artmill_PreencheCorDetalhe", objItemPedido.sProduto, iIndice)
    
    'precodesc
    objct.GridItens.TextMatrix(iIndice, objct.gobjInfoUsu.iGrid_PrecoDesc_Col) = Format(objItemPedido.dPrecoUnitario * (1 - dPercDesc), gobjFAT.sFormatoPrecoUnitario)
    
    PV_Preenche_GridItens = SUCESSO
    
    Exit Function
    
Erro_PV_Preenche_GridItens:

    PV_Preenche_GridItens = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184280)

    End Select
    
    Exit Function

End Function

Function PV_Rotina_Grid_Enable2(ByVal objct As Object, iLinha As Integer, objControl As Object, iCaminho As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long
    
    If objControl.Name = objct.DescricaoProduto.Name Then objControl.Enabled = False
    
    PV_Rotina_Grid_Enable2 = SUCESSO
    
End Function

Public Function PV_PrecoTotal_Calcula3(ByVal objct As Object, ByVal iLinha As Integer, ByVal dQuantidade As Double, ByVal dPrecoUnitario As Double, ByVal dPercentDesc As Double) As Long

Dim dPrecoUnitDesc As Double, dPrecoTotal As Double, dPrecoTotalReal As Double, dDesconto As Double
    
    'precodesc
    objct.GridItens.TextMatrix(iLinha, objct.gobjInfoUsu.iGrid_PrecoDesc_Col) = Format(dPrecoUnitario * (1 - dPercentDesc), "Standard")
    
    dPrecoUnitDesc = StrParaDbl(objct.GridItens.TextMatrix(iLinha, objct.gobjInfoUsu.iGrid_PrecoDesc_Col))

    dPrecoTotal = dPrecoUnitario * dQuantidade
    dPrecoTotalReal = dPrecoUnitDesc * dQuantidade

    'Se percentual for >0 tira o desconto
    dDesconto = dPrecoTotal - dPrecoTotalReal
    
    objct.GridItens.TextMatrix(iLinha, objct.iGrid_Desconto_Col) = Format(dDesconto, "Standard")
    
    'Coloca preco total do ítem no grid
    objct.GridItens.TextMatrix(iLinha, objct.iGrid_PrecoTotal_Col) = Format(dPrecoTotalReal, "Standard")
    objct.GridItens.TextMatrix(iLinha, objct.iGrid_PrecoTotalB_Col) = Format(dPrecoTotal, "Standard")
    
    PV_PrecoTotal_Calcula3 = SUCESSO
    
End Function

Public Function PV_Produto_Saida_Celula5(ByVal objct As Object, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long, dPercentDesc As Double
Dim dtDataEntrega As Date

On Error GoTo Erro_PV_Produto_Saida_Celula5

    If Len(Trim(objct.Cliente.Text)) > 0 Then

        lErro = CF("Artmill_PreencheDescontoICMS", objct.Cliente.Text, Codigo_Extrai(objct.Filial.Text), dPercentDesc, objct.gobjInfoUsu.gdDesconto1)
        If lErro <> SUCESSO Then gError 126430

        If dPercentDesc <> 0 Then objct.GridItens.TextMatrix(objct.GridItens.Row, objct.iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")

    End If
    
    lErro = CF2(objct, "PV_Traz_Produto_Tela", objProduto)
    If lErro <> SUCESSO Then gError 184291
    
    PV_Produto_Saida_Celula5 = SUCESSO
    
    Exit Function
    
Erro_PV_Produto_Saida_Celula5:

    PV_Produto_Saida_Celula5 = gErr

    Select Case gErr

        Case 126430, 184291
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184290)

    End Select
    
    Exit Function

End Function

Public Function PV_Produto_Saida_Celula(ByVal objct As Object, ByVal bPreenchePreco As Boolean, ByVal dPrecoUnitario As Double) As Long
            
Dim dPercentDesc As Double

On Error GoTo Erro_PV_Produto_Saida_Celula

    If dPrecoUnitario > 0 Then
        dPercentDesc = PercentParaDbl(objct.GridItens.TextMatrix(objct.GridItens.Row, objct.iGrid_PercDesc_Col))
        objct.GridItens.TextMatrix(objct.GridItens.Row, objct.gobjInfoUsu.iGrid_PrecoDesc_Col) = Format(dPrecoUnitario * (1 - dPercentDesc), "Standard")
    Else
        objct.GridItens.TextMatrix(objct.GridItens.Row, objct.gobjInfoUsu.iGrid_PrecoDesc_Col) = ""
    End If

    PV_Produto_Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_PV_Produto_Saida_Celula:

    PV_Produto_Saida_Celula = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184290)

    End Select
    
    Exit Function
    
End Function

Public Function PV_Traz_Produto_Tela(ByVal objct As Object, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim dtDataEntrega As Date

On Error GoTo Erro_PV_Traz_Produto_Tela

    Call CF2(objct, "Artmill_PreencheCorDetalhe", objProduto.sCodigo, objct.GridItens.Row)
    
    'artmill
    dtDataEntrega = MaskedParaDate(objct.DataEntregaPV)
    If dtDataEntrega <> DATA_NULA Then
        objct.GridItens.TextMatrix(objct.GridItens.Row, objct.iGrid_DataEntrega_Col) = Format(dtDataEntrega, "dd/mm/yyyy")
    Else
        objct.GridItens.TextMatrix(objct.GridItens.Row, objct.iGrid_DataEntrega_Col) = ""
    End If
    'fim artmill

    PV_Traz_Produto_Tela = SUCESSO
    
    Exit Function
    
Erro_PV_Traz_Produto_Tela:

    PV_Traz_Produto_Tela = gErr

    Select Case gErr

        Case 126430
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184290)

    End Select
    
    Exit Function

End Function

Public Sub BotaoImprimir_Click(ByVal objct As Object, ByVal iIndice As Integer)

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim dValor As Double
Dim objPedidoVendaAux As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoImprimir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(objct.Codigo.Text)) = 0 Then gError 93215
    If Len(Trim(objct.Cliente.Text)) = 0 Then gError 93216
    If Len(Trim(objct.Filial.Text)) = 0 Then gError 93217
    If Len(Trim(objct.FilialEntrega.Text)) = 0 Then gError 93218
    If Len(Trim(objct.DataEmissao.ClipText)) = 0 Then gError 93219
    'If Len(Trim(objct.NaturezaOp.Text)) = 0 Then gError 93230
        
    'Chama rotina de Gravação
    lErro = objct.Gravar_Registro()
    If lErro <> SUCESSO Then gError 126794
        
    lErro = objct.Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 93228

    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 93214

    If lErro = 26509 Then gError 93231
    
    objPedidoVendaAux.lCodigo = objPedidoVenda.lCodigo
    objPedidoVendaAux.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
    
    'Coloca o Pedido de Venda na tela
    lErro = objct.Traz_Pedido_Tela(objPedidoVendaAux)
    If lErro <> SUCESSO Then gError 126795
        
    If iIndice = 0 Then
    
        lErro = objRelatorio.ExecutarDireto("Pedido de Venda", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 0, "PedVenda", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
        If lErro <> SUCESSO Then gError 93229

    ElseIf iIndice = 1 Then

        lErro = objRelatorio.ExecutarDireto("Documento Interno", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 0, "PedVend5", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
        If lErro <> SUCESSO Then gError 126729

    ElseIf iIndice = 2 Then

        lErro = objRelatorio.ExecutarDireto("Pedido Interno", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 0, "PedVend4", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
        If lErro <> SUCESSO Then gError 126730

    ElseIf iIndice = 3 Then

        lErro = objRelatorio.ExecutarDireto("Ordem de Expedicao", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 0, "OrdExp", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
        If lErro <> SUCESSO Then gError 126731

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImprimir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 93214, 93229, 126729 To 126731, 126794, 126795

        Case 93215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)

        Case 93216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 93217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 93218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
        
        Case 93219
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 93230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 93231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177183)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Function PV_Cliente_Formata(ByVal objct As Object, ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim objTipoCliente As New ClassTipoCliente

On Error GoTo Erro_PV_Cliente_Formata

    'Se o Tipo estiver preenchido
    If objCliente.iTipo > 0 Then
        objTipoCliente.iCodigo = objCliente.iTipo
        'Lê o Tipo de Cliente
        lErro = CF("TipoCliente_Le", objTipoCliente)
        If lErro <> SUCESSO And lErro <> 19062 Then gError 126822
    End If

    objct.gdDesconto = 0
    'Guarda o valor do desconto do cliente
    If objCliente.dDesconto > 0 Then
        objct.gobjInfoUsu.gdDesconto1 = objCliente.dDesconto
    ElseIf objTipoCliente.dDesconto > 0 Then
        objct.gobjInfoUsu.gdDesconto1 = objTipoCliente.dDesconto
    Else
        objct.gobjInfoUsu.gdDesconto1 = 0
    End If

    PV_Cliente_Formata = SUCESSO
    
    Exit Function
    
Erro_PV_Cliente_Formata:

    PV_Cliente_Formata = gErr

    Select Case gErr

        Case 126822

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184292)

    End Select
    
    Exit Function

End Function

Public Function PV_UserControl_KeyDown(ByVal objct As Object, KeyCode As Integer, Shift As Integer) As Long
    If KeyCode = vbKeyReturn Then
        If objct.ActiveControl Is objct.Opcao And objct.Opcao.SelectedItem.Index = TAB_Itens Then
            objct.GridItens.Row = objct.objGridItens.iLinhasExistentes + 1
            objct.GridItens.Col = 1
            objct.GridItens.SetFocus
            SendKeys TECLA_ENTER
        End If
    End If
    PV_UserControl_KeyDown = SUCESSO
End Function

Function PV_DataEntregaPV_Validate(ByVal objct As Object, Cancel As Boolean) As Long
    
    PV_DataEntregaPV_Validate = SUCESSO
    
End Function

