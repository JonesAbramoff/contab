VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Function BotaoEntrega_Click(ByVal objCT As Object) As Long

Dim lErro As Long
Dim sProdutoTela As String, colDatas As Collection

On Error GoTo Erro_BotaoEntrega_Click

    sProdutoTela = objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col)
    Set colDatas = objCT.gobjInfoUsu.gcolDatas
    Call Chama_Tela_Modal("PVEntrega", sProdutoTela, colDatas)


    BotaoEntrega_Click = SUCESSO

    Exit Function

Erro_BotaoEntrega_Click:

    BotaoEntrega_Click = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180810)

    End Select

    Exit Function
    
End Function

Public Sub PV_Quantidade_GotFocus(ByVal objCT As Object)

Dim lErro As Long


On Error GoTo Erro_PV_Quantidade_GotFocus

    If objCT.objGridItens.lErroSaidaCelula = 0 Then

        objCT.gdQuantAnterior = StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_QuantPedida_Col))

        lErro = Trata_Data_Entrega(objCT)
        If lErro <> SUCESSO Then gError 199833

        Call Grid_Entrada_Celula(objCT.objGridItens, objCT.iAlterado)

        objCT.objGridItens.lErroSaidaCelula = 1

    Else

        objCT.objGridItens.lErroSaidaCelula = 0

        objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_QuantPedida_Col) = objCT.gdQuantAnterior

    End If

    Exit Sub

Erro_PV_Quantidade_GotFocus:

    Select Case gErr

        Case 199833

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199834)

    End Select

    Exit Sub

End Sub

Public Sub PV_Trata_Quantidade_ItemPV(ByVal objCT As Object)

Dim objDataEntrega As ClassDataEntrega
Dim dQuantidade As Double

    For Each objDataEntrega In objCT.gobjPedidoVenda.colItensPedido(objCT.GridItens.Row).colItemPVEntrega
    
        dQuantidade = dQuantidade + objDataEntrega.dQuantidade
        
    Next

    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_QuantPedida_Col) = Formata_Estoque(dQuantidade)

End Sub

Private Function Trata_Data_Entrega(ByVal objCT As Object) As Long

Dim lErro As Long
Dim objDataEntregaParam As New ClassDataEntregaParam

Dim dQuantidade As Double
Dim objDataEntrega As ClassDataEntrega
Dim sProdutoTela As String
Dim colItemDataEntregaPV As Collection

On Error GoTo Erro_Trata_Data_Entrega

    
    Set objCT.objGridItens.objControle = objCT.Quantidade
    
    lErro = Grid_Abandona_Celula(objCT.objGridItens)
    If lErro <> SUCESSO Then gError 199830
            
    objDataEntregaParam.sProduto = objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col)

    If Len(objDataEntregaParam.sProduto) = 0 Then gError 199831
    
    Set objDataEntregaParam.colDataEntrega = objCT.gobjPedidoVenda.colItensPedido(objCT.GridItens.Row).colItemPVEntrega

    Call Chama_Tela_Modal("DataEntrega", objDataEntregaParam, 1, 1)

    Call PV_Trata_Quantidade_ItemPV(objCT)

    objCT.Show
    
    Trata_Data_Entrega = SUCESSO

    Exit Function

Erro_Trata_Data_Entrega:

    Trata_Data_Entrega = gErr

    Select Case gErr

        Case 199830

        Case 199831
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199832)

    End Select

    Exit Function

End Function


