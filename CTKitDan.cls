VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTKitDan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'a qtde do produto pai terá que ser sempre 1, para poder ser multiplicada pela qtde de kits da OS.

Public Sub BotaoOS_Click(ByVal objCT As Object)

Dim lErro As Long, sProduto As String
Dim objDan_OS As New ClassDan_OS, sSelecao As String
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoOS_Click

    If objCT.colComponentes.Count > 0 Then
        
        sProduto = objCT.colComponentes.Item(1).sProdutoRaiz
        sSelecao = "EXISTS (SELECT 1 FROM Dan_ItensOS WHERE Produto = ? AND Dan_ItensOS.OS = Dan_OSCli.OS)"
        colSelecao.Add sProduto
    Else
        sProduto = ""
    End If
    
    Call Chama_Tela("Dan_OSCliLista", colSelecao, objDan_OS, Nothing, sSelecao)

    Exit Sub

Erro_BotaoOS_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184831)

    End Select

    Exit Sub

End Sub

Public Function KT_GravarRegistro(ByVal objCT As Object, ByVal objKit As ClassKit) As Long
    
On Error GoTo Erro_KT_GravarRegistro
    
    If objKit.colComponentes.Count <> 0 Then
        If objKit.colComponentes.Item(1).dQuantidade <> 1 Then gError 184848
    End If
    
    KT_GravarRegistro = SUCESSO
    
    Exit Function
    
Erro_KT_GravarRegistro:

    KT_GravarRegistro = gErr

    Select Case gErr

        Case 184848
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDE_KIT_DEFINIDA_NA_OS", gErr, Error$)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184847)

    End Select
    
    Exit Function

End Function


