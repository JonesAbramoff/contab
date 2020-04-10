VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPV_ConsultaHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--- inicio dos properties get dos controles da tela

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
Public Function PVC_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long

    'liberar objetos globais referentes a customizacao
    
End Function

Public Function PVC_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
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
    objGridInt.colColuna.Add ("Preço Total")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Quant Reservada")
    objGridInt.colColuna.Add ("Quant Faturada")

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
    objGridInt.colCampo.Add (objCT.DataEntrega.Name)
    objGridInt.colCampo.Add (objCT.QuantReservadaPV.Name)
    objGridInt.colCampo.Add (objCT.QuantFaturada.Name)

    'Colunas do Grid
    objCT.iGrid_ItemProduto_Col = 0
    objCT.iGrid_Produto_Col = 1
    objCT.iGrid_DescProduto_Col = 2
    objCT.iGrid_UnidadeMed_Col = 3
    objCT.iGrid_QuantPedida_Col = 4
    objCT.iGrid_QuantCancel_Col = 5
    objCT.iGrid_PrecoUnitario_Col = 6
    objCT.iGrid_PercDesc_Col = 7
    
    objCT.gobjInfoUsu.iGrid_TipoDesc1_Col = 8
    objCT.gobjInfoUsu.iGrid_PercDesc2_Col = 9
    objCT.gobjInfoUsu.iGrid_TipoDesc2_Col = 10
    objCT.gobjInfoUsu.iGrid_PercDesc3_Col = 11
    objCT.gobjInfoUsu.iGrid_TipoDesc3_Col = 12
    
    objCT.iGrid_Desconto_Col = 13
    objCT.iGrid_PrecoTotal_Col = 14
    objCT.iGrid_DataEntrega_Col = 15
    objCT.iGrid_QuantReservada_Col = 16
    objCT.iGrid_QuantFaturada_Col = 17

    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS_HICARE + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    objCT.GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    PVC_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Function PVC_Carrega_Grid_Itens(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long, iIndice As Integer, objItemPedido As ClassItemPedido, sTipo As String

On Error GoTo Erro_PVC_Carrega_Grid_Itens

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
    
    PVC_Carrega_Grid_Itens = SUCESSO
     
    Exit Function
    
Erro_PVC_Carrega_Grid_Itens:

    PVC_Carrega_Grid_Itens = gErr
     
    Select Case gErr
          
        Case 124182 To 124184
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179340)
     
    End Select
     
    Exit Function

End Function

Function PVC_Traz_Pedido_Tela(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long

On Error GoTo Erro_PVC_Traz_Pedido_Tela

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
          
    PVC_Traz_Pedido_Tela = SUCESSO
     
    Exit Function
    
Erro_PVC_Traz_Pedido_Tela:

    PVC_Traz_Pedido_Tela = gErr
     
    Select Case gErr
          
        Case 124189
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179341)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179342)
    
    End Select

End Function

Public Function PVC_CargaPosFormLoad(ByVal objCT As Object) As Long
    
Dim lErro As Long

On Error GoTo Erro_PVC_CargaPosFormLoad

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
    
    PVC_CargaPosFormLoad = SUCESSO
         
    Exit Function
    
Erro_PVC_CargaPosFormLoad:

    PVC_CargaPosFormLoad = gErr
     
    Select Case gErr
          
        Case 124170
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179343)
     
    End Select
     
    Exit Function

End Function
