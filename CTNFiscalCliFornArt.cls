VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalCliFornArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function NFCF_Cliente_Validate2(ByVal objct As Object, ByVal objCliente As ClassCliente, ByVal objTipoCliente As ClassTipoCliente) As Long

Dim lErro As Long

On Error GoTo Erro_NFCF_Cliente_Validate2

    objct.gdDesconto = 0
    
    'Verifica se o cliente possui um desconto
    If objCliente.dDesconto > 0 Then
        objct.gobjInfoUsu.gdDesconto1 = objCliente.dDesconto
    ElseIf objTipoCliente.dDesconto > 0 Then
        objct.gobjInfoUsu.gdDesconto1 = objTipoCliente.dDesconto
    Else
        objct.gobjInfoUsu.gdDesconto1 = 0
    End If
    
    NFCF_Cliente_Validate2 = SUCESSO
    
    Exit Function
    
Erro_NFCF_Cliente_Validate2:

    NFCF_Cliente_Validate2 = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184282)

    End Select
    
    Exit Function

End Function

Function NFCF_Limpa_Desconto(ByVal objct As Object) As Long
    objct.gobjInfoUsu.gdDesconto1 = 0
    
    NFCF_Limpa_Desconto = SUCESSO
End Function

Function NFCF_Cliente_Validate3(ByVal objct As Object, ByVal objCliente As ClassCliente) As Long
    NFCF_Cliente_Validate3 = SUCESSO
End Function

Public Function NFCF_Cliente_Filial_Validate(ByVal objct As Object, ByVal objFilialCliente As ClassFilialCliente) As Long

Dim lErro As Long

On Error GoTo Erro_NFCF_Cliente_Filial_Validate

    lErro = Desconto_Item_Recalcula(objct)
    If lErro <> SUCESSO Then gError 126728
    
    NFCF_Cliente_Filial_Validate = SUCESSO
    
    Exit Function
    
Erro_NFCF_Cliente_Filial_Validate:

    NFCF_Cliente_Filial_Validate = gErr

    Select Case gErr

        Case 126728
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184281)

    End Select
    
    Exit Function

End Function

Private Function Desconto_Item_Recalcula(ByVal objct As Object) As Long

Dim iIndice As Integer
Dim dPercentDesc As Double
Dim bCalculou As Boolean
Dim lErro As Long
    
On Error GoTo Erro_Desconto_Item_Recalcula
    
    'Para cada linha do Grid de Itens
    For iIndice = 1 To objct.objGridItens.iLinhasExistentes
    
        'Se o percentual de desconto nao estiver preenchido
        If Len(Trim(objct.GridItens.TextMatrix(iIndice, objct.iGrid_PercDesc_Col))) = 0 Then
        
            lErro = CF("Artmill_PreencheDescontoICMS", objct.Cliente.Text, Codigo_Extrai(objct.Filial.Text), dPercentDesc, objct.gobjInfoUsu.gdDesconto1)
            If lErro <> SUCESSO Then gError 126725
    
            If dPercentDesc <> 0 Then objct.GridItens.TextMatrix(iIndice, objct.iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")

            Call objct.PrecoTotal_Calcula(iIndice)
            bCalculou = True

        End If
        
    Next
    
    If bCalculou Then Call objct.ValorTotal_Calcula
    
    Desconto_Item_Recalcula = SUCESSO
    
    Exit Function
    
Erro_Desconto_Item_Recalcula:
    
    Desconto_Item_Recalcula = gErr
    
    Select Case gErr

        Case 126725

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 176180)

    End Select
    
    Exit Function
    
End Function

Public Sub Artmill_PreencheCorDetalhe(ByVal objct As Object, ByVal sProduto As String, ByVal iIndice As Integer)

Dim lErro As Long, objProduto As New ClassProduto

On Error GoTo Erro_Artmill_PreencheCorDetalhe

    objProduto.sCodigo = sProduto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 124137

    'Se o produto não está cadastrado, erro
    If lErro <> SUCESSO Then gError 124139
    
    objct.GridItens.TextMatrix(iIndice, objct.gobjInfoUsu.iGrid_DescProduto2_Col) = objProduto.sCor
    
    lErro = CF("Produto_Le_InfoUsu", objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 124138
    
    'Se o produto não está cadastrado, erro
    If lErro <> SUCESSO Then gError 124140
    
    objct.GridItens.TextMatrix(iIndice, objct.gobjInfoUsu.iGrid_DescProduto3_Col) = objProduto.objInfoUsu.sDetalheCor

    Exit Sub
     
Erro_Artmill_PreencheCorDetalhe:

    Select Case gErr
          
        Case 124137, 124138
        
        Case 124139, 124140
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 176198)
     
    End Select
     
    Exit Sub

End Sub

'Public Function NFCF_Form_Load_PrecoDesc(ByVal objct As Object, ByVal sContainer As String, ByVal iIndice As Integer) As Long
'
'Dim lErro As Long, objAux As Object, objPrecoUnitario As Object, objPrecoDesc As Object
'
'On Error GoTo Erro_NFCF_Form_Load_PrecoDesc
'
'    Set objPrecoUnitario = objct.objUserControl.Controls("PrecoUnitario")
'
'    If sContainer <> "" Then
'        Set objAux = objct.objUserControl.Controls(sContainer)
'        Set objPrecoDesc = objct.objUserControl.Controls.Add("MsMask.MaskEdBox", "PrecoDesc", objAux(iIndice))
'    Else
'        Set objPrecoDesc = objct.objUserControl.Controls.Add("MsMask.MaskEdBox", "PrecoDesc")
'    End If
'
'    objPrecoDesc.Enabled = False
'    objPrecoDesc.Width = objPrecoUnitario.Width
'    objPrecoDesc.Height = objPrecoUnitario.Height
'    objPrecoDesc.Format = objPrecoUnitario.Format
'
'    NFCF_Form_Load_PrecoDesc = SUCESSO
'
'    Exit Function
'
'Erro_NFCF_Form_Load_PrecoDesc:
'
'    NFCF_Form_Load_PrecoDesc = gErr
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184276)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function NFCF_BotaoImprimirPI_Click(ByVal objct As Object) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_NFCF_BotaoImprimirPI_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    objPedidoVenda.lCodigo = objct.glCodPedido
    objPedidoVenda.iFilialEmpresa = objct.giFilialPedido
    
    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 132431

    If lErro = 26509 Then gError 132432
    
    lErro = objRelatorio.ExecutarDireto("Pedido Interno", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 1, "PedVend4", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError 132433

'    'Limpa a Tela
'    Call Limpa_Tela_NFiscalFaturaPedido
    
    GL_objMDIForm.MousePointer = vbDefault

    NFCF_BotaoImprimirPI_Click = SUCESSO
    
    Exit Function

Erro_NFCF_BotaoImprimirPI_Click:

    NFCF_BotaoImprimirPI_Click = gErr
    
    Select Case gErr

        Case 132431, 132433
        
        Case 132432
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 176821)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Function

End Function
'###########################################################

