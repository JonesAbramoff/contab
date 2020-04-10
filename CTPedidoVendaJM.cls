VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaJM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function PV_Saida_Celula_PrecoUnitario2(ByVal objCT As Object, ByVal dPrecoUnitario As Double, ByVal iLinha As Integer) As Long
    
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim dPrecoUnitarioTab As Double
Dim objPVMotivo As ClassPVMotivo
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem

On Error GoTo Erro_PV_Saida_Celula_PrecoUnitario2

    lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 199681

    objProduto.sCodigo = sProduto

    If Len(Trim(objCT.TabelaPreco.Text)) > 0 Then

        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(objCT.TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo

        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 199682
        
        'preco nao encontrado na tabela
        If lErro <> SUCESSO Then gError 199683
        
        lErro = CF2(objCT, "PV_PrecoUnitario_Calcula", iLinha, objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitarioTab)
        If lErro <> SUCESSO Then gError 199684

        If dPrecoUnitario < dPrecoUnitarioTab Then gError 199685
    
    Else
        'tabela de preco nao preenchida
        gError 199686

    End If

    PV_Saida_Celula_PrecoUnitario2 = SUCESSO
    
    Exit Function

Erro_PV_Saida_Celula_PrecoUnitario2:

    PV_Saida_Celula_PrecoUnitario2 = gErr

    Select Case gErr
    
        Case 199681, 199682, 199684
        
        Case 199683
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECO_PRODUTO_NAO_CADASTRADO", gErr, objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col))

        Case 199685
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECO_MENOR_PRECOTABELA", gErr, dPrecoUnitario, dPrecoUnitarioTab)

        Case 199686
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199687)

    End Select

    Exit Function
    
End Function

