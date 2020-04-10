Attribute VB_Name = "TRPGeraComiInt"
Option Explicit


Public Function Rotina_GerComiInt_AtualizaTelaBatch()
'Atualiza tela de acompanhamento do Batch

Dim lErro As Long
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Rotina_GerComiInt_AtualizaTelaBatch


    'Atualiza tela de acompanhamento do Batch
    lErro = DoEvents()

    TelaAcompanhaBatchFAT1.dValorAtual = TelaAcompanhaBatchFAT1.dValorAtual + 1
    TelaAcompanhaBatchFAT1.TotReg.Caption = CStr(TelaAcompanhaBatchFAT1.dValorAtual)
    TelaAcompanhaBatchFAT1.ProgressBar1.Value = CInt((TelaAcompanhaBatchFAT1.dValorAtual / TelaAcompanhaBatchFAT1.dValorTotal) * 100)

    If TelaAcompanhaBatchFAT1.iCancelaBatch = CANCELA_BATCH Then

        vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_GERCOMIINT")

        If vbMsgBox = vbYes Then gError 197372

        TelaAcompanhaBatchFAT1.iCancelaBatch = 0

    End If
    
    Rotina_GerComiInt_AtualizaTelaBatch = SUCESSO
    
    Exit Function

Erro_Rotina_GerComiInt_AtualizaTelaBatch:

    Rotina_GerComiInt_AtualizaTelaBatch = gErr

    Select Case gErr

        Case 197372

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197373)

    End Select
       
    Exit Function

End Function

