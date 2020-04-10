VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTChequesPag3AHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function CHQA_Trata_Parametros(ByVal objCT As Object, ByVal objChequesPagAvulso As ClassChequesPagAvulso) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_CHQA_Trata_Parametros
    
    objCT.Observacao.Text = objChequesPagAvulso.objInfoChequePag.sObservacao

    CHQA_Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_CHQA_Trata_Parametros:

    CHQA_Trata_Parametros = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179246)
        
    End Select

    Exit Function
        
End Function

Public Sub BotaoReter_Click(ByVal objCT As Object)

Dim lErro As Long
Dim objInfoChequePag As ClassInfoChequePag

On Error GoTo Erro_BotaoReter_Click

    'Verifica se o NumCheque está preenchido
    If Len(Trim(objCT.NumCheque.Text)) = 0 Then gError 132886

    '##########################
    'Inserido por Wagner
    objCT.gobjChequesPagAvulso.objInfoChequePag.sFavorecido = objCT.Nominal.Text
    objCT.gobjChequesPagAvulso.objInfoChequePag.sObservacao = objCT.Observacao.Text
    objCT.gobjChequesPagAvulso.lNumCheque = StrParaLong(objCT.NumCheque.Text)
    '##########################

   lErro = CF("ParcelaPag_Grava_ChequeAvulso", objCT.gobjChequesPagAvulso)
   If lErro <> SUCESSO Then gError 132572

   Call Rotina_Aviso(vbOKOnly, "AVISO_CHEQUE_RETIDO")

   Exit Sub

Erro_BotaoReter_Click:

    Select Case gErr
    
        Case 132572
        
        Case 132886
            Call Rotina_Erro(vbOKOnly, "ERRO_PROXCHEQUE_NAO_INFORMADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 179247)
    
    End Select

    Exit Sub
    
End Sub
