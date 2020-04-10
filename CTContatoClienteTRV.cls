VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTContatoCliTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Marca(ByVal objCT As Object) As Object
     Set Marca = objCT.objUserControl.Controls("Marca")
End Property

Public Sub BotaoEmissores_Click(ByVal objCT As Object)

Dim objForn As New ClassFornecedor
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoEmissores_Click

    If objCT.GridClientes.Row = 0 Then gError 200154
        
    colSelecao.Add LCodigo_Extrai(objCT.GridClientes.TextMatrix(objCT.GridClientes.Row, objCT.iGrid_Cliente_Col))
    
    Call Chama_Tela("TRVEmissoresLista", colSelecao, objForn, Nothing, "Cliente = ? ")

    Exit Sub

Erro_BotaoEmissores_Click:

    Select Case gErr
    
        Case 200154
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200155)

    End Select

    Exit Sub

End Sub

Public Sub Marca_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
    objCT.iAlteradoFiltro = REGISTRO_ALTERADO
End Sub

Public Function CONTCLI_Move_Selecao_Memoria(ByVal objCT As Object, ByVal objContatoCliSel As ClassContatoCliSel) As Long

On Error GoTo Erro_CONTCLI_Move_Selecao_Memoria

    objContatoCliSel.iMarca = Marca(objCT).ItemData(Marca(objCT).ListIndex)
    
    CONTCLI_Move_Selecao_Memoria = SUCESSO
    
    Exit Function

Erro_CONTCLI_Move_Selecao_Memoria:

    CONTCLI_Move_Selecao_Memoria = gErr

    Select Case gErr
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200987)

    End Select

    Exit Function
    
End Function
