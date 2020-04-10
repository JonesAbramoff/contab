Attribute VB_Name = "ImportCoinfo"
Option Explicit

Public Const CATEGORIA_PRODUTO_VERSAO_KIT = "Versão Kit"

Public Function Rotina_Importar_AtualizaTelaBatch()
'Atualiza tela de acompanhamento do Batch

Dim lErro As Long
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Rotina_Importar_AtualizaTelaBatch

    'Atualiza tela de acompanhamento do Batch
    DoEvents

    TelaAcompanhaBatchFAT.dValorAtualImp = TelaAcompanhaBatchFAT.dValorAtualImp + 1
    TelaAcompanhaBatchFAT.TotArq.Caption = CStr(TelaAcompanhaBatchFAT.dValorAtualImp)
    TelaAcompanhaBatchFAT.ProgressBar1.Value = CInt((TelaAcompanhaBatchFAT.dValorAtualImp / TelaAcompanhaBatchFAT.dValorTotalImp) * 100)

    If (TelaAcompanhaBatchFAT.iCancelaBatch = CANCELA_BATCH) Or (TelaAcompanhaBatchFAT.Cancelar.Enabled = False) Then

        vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_IMPORTAR_COINFO")

        If vbMsgBox = vbYes Then gError 129933

        TelaAcompanhaBatchFAT.Cancelar.Enabled = True
        TelaAcompanhaBatchFAT.iCancelaBatch = 0

    End If
    
    Rotina_Importar_AtualizaTelaBatch = SUCESSO
    
    Exit Function

Erro_Rotina_Importar_AtualizaTelaBatch:

    Rotina_Importar_AtualizaTelaBatch = gErr

    Select Case gErr

        Case 129933

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161553)

    End Select
       
    Exit Function

End Function

Public Function Rotina_Atualizar_AtualizaTelaBatch()
'Atualiza tela de acompanhamento do Batch

Dim lErro As Long
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Rotina_Atualizar_AtualizaTelaBatch

    'Atualiza tela de acompanhamento do Batch
    DoEvents

    TelaAcompanhaBatchFAT.dValorAtualAtu = TelaAcompanhaBatchFAT.dValorAtualAtu + 1
    TelaAcompanhaBatchFAT.TotReg.Caption = CStr(TelaAcompanhaBatchFAT.dValorAtualAtu)
    TelaAcompanhaBatchFAT.ProgressBar2.Value = CInt((TelaAcompanhaBatchFAT.dValorAtualAtu / TelaAcompanhaBatchFAT.dValorTotalAtu) * 100)

    DoEvents
    
    If (TelaAcompanhaBatchFAT.iCancelaBatch = CANCELA_BATCH) Or (TelaAcompanhaBatchFAT.Cancelar.Enabled = False) Then

        vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_IMPORTAR_COINFO")

        If vbMsgBox = vbYes Then gError 129933

        TelaAcompanhaBatchFAT.iCancelaBatch = 0

    End If
    
    Rotina_Atualizar_AtualizaTelaBatch = SUCESSO
    
    Exit Function

Erro_Rotina_Atualizar_AtualizaTelaBatch:

    Rotina_Atualizar_AtualizaTelaBatch = gErr

    Select Case gErr

        Case 129933

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161553)

    End Select
       
    Exit Function

End Function

