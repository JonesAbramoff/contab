VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTCobrPorEmailTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get TipoDoc(ByVal objCT As Object) As Object
     Set TipoDoc = objCT.objUserControl.Controls("TipoDoc")
End Property

Function CE_Move_TabSelecao_Memoria(ByVal objCT As Object, ByVal objCobrancaEmailSel As ClassCobrancaPorEmailSel) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objCobrSelTRP As New ClassCobrPorEmailSelTRP

On Error GoTo Erro_CE_Move_TabSelecao_Memoria

    Set objCobrancaEmailSel.objInfoUsu = objCobrSelTRP

    If objCT.giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Or objCT.giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
    
        For iLinha = 0 To TipoDoc(objCT).ListCount - 1
            
            If TipoDoc(objCT).Selected(iLinha) Then
                objCobrSelTRP.colTiposDoc.Add TipoDoc(objCT).ItemData(iLinha)
            End If
        
        Next
    
    End If
   
    CE_Move_TabSelecao_Memoria = SUCESSO

    Exit Function

Erro_CE_Move_TabSelecao_Memoria:

    CE_Move_TabSelecao_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187027)

    End Select

    Exit Function

End Function

Function CE_BotaoEmail_Click(ByVal objCT As Object, iGeraRelac As Integer)

'    If objCT.giTipoTela = TIPOTELA_EMAIL_COBRANCA_FATURA Or objCT.giTipoTela = TIPOTELA_EMAIL_AVISO_PAGTO_CP Then
'        iGeraRelac = MARCADO
'    End If
    CE_BotaoEmail_Click = SUCESSO
End Function

Public Sub TipoDoc_Click(ByVal objCT As Object)
    objCT.iAlteradoFiltro = REGISTRO_ALTERADO
End Sub
