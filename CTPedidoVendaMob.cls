VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVendaMob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function PV_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_PV_Form_Load

    Set objCT.gobjInfoUsu = New CTPedidoVendaVGMob
    
    Exit Function

Erro_PV_Form_Load:

    PV_Form_Load = gErr

    Select Case gErr
    
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197734)

    End Select

    Exit Function

End Function

Public Function PV_Produto_Saida_Celula6(ByVal objCT As Object) As Long
    
Dim objPVMotivo As New ClassPVMotivo
    
    objCT.gobjInfoUsu.gcolMotivos.Add objPVMotivo
    
    PV_Produto_Saida_Celula6 = SUCESSO
    
End Function

Public Function PV_Saida_Celula_PrecoUnitario1(ByVal objCT As Object, ByVal dPrecoUnitario As Double, ByVal iLinha As Integer) As Long
    
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim dPrecoUnitarioTab As Double
Dim objPVMotivo As ClassPVMotivo
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem

On Error GoTo Erro_PV_Saida_Celula_PrecoUnitario1

    lErro = CF("Produto_Formata", objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 197735

    objProduto.sCodigo = sProduto

    If Len(Trim(objCT.TabelaPreco.Text)) > 0 Then

        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(objCT.TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo

        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 197736

        lErro = CF2(objCT, "PV_PrecoUnitario_Calcula", iLinha, objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitarioTab)
        If lErro <> SUCESSO Then gError 197737

        If dPrecoUnitario < dPrecoUnitarioTab Then
    
            Set objPVMotivo = objCT.gobjInfoUsu.gcolMotivos(iLinha)
    
            objPVMotivo.sProduto = sProduto
            objPVMotivo.dPrecoTabela = dPrecoUnitarioTab
            objPVMotivo.dPrecoInformado = dPrecoUnitario
            objPVMotivo.sMotivo = objCT.gobjInfoUsu.gcolMotivos(iLinha).sMotivo
    
            Call Chama_Tela_Modal("PVMotivo", objPVMotivo)
            
            If Len(Trim(objPVMotivo.sMotivo)) = 0 Then gError 197738
    
        End If

    End If

    PV_Saida_Celula_PrecoUnitario1 = SUCESSO
    
    Exit Function

Erro_PV_Saida_Celula_PrecoUnitario1:

    PV_Saida_Celula_PrecoUnitario1 = gErr

    Select Case gErr
    
        Case 197735 To 197738

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197739)

    End Select

    Exit Function
    
End Function

Public Function PV_Exclui_ItemPV(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    objCT.gobjInfoUsu.gcolMotivos.Remove (iLinha)
    PV_Exclui_ItemPV = SUCESSO
End Function

Function PV_Move_Pedido_Memoria(ByVal objCT As Object, objPedidoVenda As ClassPedidoDeVenda) As Long
    Set objPedidoVenda.objInfoUsu = objCT.gobjInfoUsu
    PV_Move_Pedido_Memoria = SUCESSO
End Function

Public Function PV_Codigo_Enabled(ByVal objCT As Object, ByVal bFlag As Boolean) As Long
    Set objCT.gobjInfoUsu.gcolMotivos = New Collection
    PV_Codigo_Enabled = SUCESSO
End Function

