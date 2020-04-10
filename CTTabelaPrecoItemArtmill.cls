VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTTabelaPrecoItemArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get Cor(ByVal objCT As Object) As Object
     Set Cor = objCT.objUserControl.Controls("Cor")
End Property

Public Property Get Detalhe(ByVal objCT As Object) As Object
     Set Detalhe = objCT.objUserControl.Controls("Detalhe")
End Property

Function TabPreco_Produto_ExibeInfo(ByVal objCT As Object, ByVal sProduto As String, Optional ByVal sCor As String = "") As Long

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_TabPreco_Produto_ExibeInfo
    
    If sProduto = "" Then
        Cor(objCT).Caption = ""
        Detalhe(objCT).Caption = ""
    Else
        Cor(objCT).Caption = sCor
        objProduto.sCodigo = sProduto
        lErro = CF("Produto_Le_InfoUsu", objProduto)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 140380
    
        If lErro = SUCESSO Then
            Detalhe(objCT).Caption = objProduto.objInfoUsu.sDetalheCor
        End If
    
    End If
    
    TabPreco_Produto_ExibeInfo = SUCESSO
    
    Exit Function
    
Erro_TabPreco_Produto_ExibeInfo:

    TabPreco_Produto_ExibeInfo = gErr

    Select Case gErr

        Case 140380

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184267)
    
    End Select
    
    Exit Function

End Function

