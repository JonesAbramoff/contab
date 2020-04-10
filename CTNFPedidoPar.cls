VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFPedidoPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function NFP_Quantidade_GotFocus(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_NFP_Quantidade_GotFocus

    If objCT.objGridItens.lErroSaidaCelula = 0 Then

        objCT.gdQuantAnterior = StrParaDbl(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Quantidade_Col))
        
        lErro = Trata_Data_Entrega(objCT)
        If lErro <> SUCESSO Then gError 199835

        Call Grid_Entrada_Celula(objCT.objGridItens, objCT.iAlterado)

        objCT.objGridItens.lErroSaidaCelula = 1

    Else

        objCT.objGridItens.lErroSaidaCelula = 0

        objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Quantidade_Col) = objCT.gdQuantAnterior

    End If

    NFP_Quantidade_GotFocus = SUCESSO
    
    Exit Function

Erro_NFP_Quantidade_GotFocus:

    NFP_Quantidade_GotFocus = gErr
    
    Select Case gErr

        Case 199835

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199836)

    End Select

    Exit Function

End Function

Private Function Trata_Data_Entrega(ByVal objCT As Object) As Long

Dim lErro As Long
Dim objDataEntregaParam As New ClassDataEntregaParam

On Error GoTo Erro_Trata_Data_Entrega

    
    Set objCT.objGridItens.objControle = objCT.Quantidade
    
    lErro = Grid_Abandona_Celula(objCT.objGridItens)
    If lErro <> SUCESSO Then gError 199837
            
    objDataEntregaParam.sProduto = objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col)

    If Len(objDataEntregaParam.sProduto) = 0 Then gError 199838

    Set objDataEntregaParam.colDataEntrega = objCT.gobjNFiscal.ColItensNF(objCT.GridItens.Row).colDataEntrega
    
    Set objDataEntregaParam.colDataEntregaPV = objCT.gobjNFiscal.ColItensNF(objCT.GridItens.Row).colDataEntregaPV
    
    Call Chama_Tela_Modal("DataEntrega", objDataEntregaParam)

    Call NFP_Trata_Quantidade_ItemNF(objCT)

    objCT.Show
    
    Trata_Data_Entrega = SUCESSO

    Exit Function

Erro_Trata_Data_Entrega:

    Trata_Data_Entrega = gErr

    Select Case gErr

        Case 199837

        Case 199838
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199839)

    End Select

    Exit Function

End Function

Public Sub NFP_Trata_Quantidade_ItemNF(ByVal objCT As Object)

Dim objDataEntrega As ClassDataEntrega
Dim dQuantidade As Double

    For Each objDataEntrega In objCT.gobjNFiscal.ColItensNF(objCT.GridItens.Row).colDataEntrega
    
        dQuantidade = dQuantidade + objDataEntrega.dQuantidade
        
    Next

    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

End Sub

Public Function NFP_BotaoEntrega_Click(ByVal objCT As Object) As Long

Dim lErro As Long
Dim objDataEntregaParam As New ClassDataEntregaParam

On Error GoTo Erro_NFP_BotaoEntrega_Click

    lErro = Trata_Data_Entrega(objCT)
    If lErro <> SUCESSO Then gError 199835
    
    NFP_BotaoEntrega_Click = SUCESSO

    Exit Function

Erro_NFP_BotaoEntrega_Click:

    NFP_BotaoEntrega_Click = gErr

    Select Case gErr

        Case 199835
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183238)

    End Select

    Exit Function
    
End Function


