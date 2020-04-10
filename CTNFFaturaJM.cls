VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFFaturaJM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function NFF_Saida_Celula_PrecoUnitario(ByVal objCT As Object, ByVal dPrecoUnitario As Double, ByVal iLinha As Integer) As Long
    
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim dPrecoUnitarioTab As Double
Dim objPVMotivo As ClassPVMotivo
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim sUM As String

On Error GoTo Erro_NFF_Saida_Celula_PrecoUnitario

    lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 199809

    objProduto.sCodigo = sProduto

    If Len(Trim(objCT.TabelaPreco.Text)) > 0 Then

        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(objCT.TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo

        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 199810
        
        'preco nao encontrado na tabela
        If lErro <> SUCESSO Then gError 199811
        
        sUM = objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_UnidadeMed_Col)
        
        lErro = CF2(objCT, "NFF_PrecoUnitario_Calcula", sUM, objTabelaPrecoItem, dPrecoUnitarioTab)
        If lErro <> SUCESSO Then gError 199812

        If dPrecoUnitario < dPrecoUnitarioTab Then gError 199813
    
    Else
        'tabela de preco nao preenchida
        gError 199814

    End If

    NFF_Saida_Celula_PrecoUnitario = SUCESSO
    
    Exit Function

Erro_NFF_Saida_Celula_PrecoUnitario:

    NFF_Saida_Celula_PrecoUnitario = gErr

    Select Case gErr
    
        Case 199809, 199810, 199812
        
        Case 199811
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECO_PRODUTO_NAO_CADASTRADO", gErr, objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col))

        Case 199813
            Call Rotina_Erro(vbOKOnly, "ERRO_PRECO_MENOR_PRECOTABELA", gErr, dPrecoUnitario, dPrecoUnitarioTab)

        Case 199814
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199807)

    End Select

    Exit Function
    
End Function

