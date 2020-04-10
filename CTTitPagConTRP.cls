VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTTitPagConTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub BotaoExcluir_Click(ByVal objCT As Object)

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim vbMsgRes As VbMsgBoxResult
Dim colParcelas As New ColParcelaPagar
Dim iFilialAux As Integer

On Error GoTo Erro_BotaoExcluir_Click

    iFilialAux = giFilialEmpresa
    giFilialEmpresa = objCT.ifilialempresatela

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se campos identificadores estão preenchidos
    If Len(Trim(objCT.Fornecedor.Text)) = 0 Then gError 196410
    If Len(Trim(objCT.Filial.Text)) = 0 Then gError 196411
    If Len(Trim(objCT.NumeroTitulo.Text)) = 0 Then gError 196412
    If Len(Trim(objCT.Tipo.Text)) = 0 Then gError 196413
    
    'Move dados da Tela para objTituloPagar e colParcelas
    lErro = objCT.Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then gError 196419
    
    objTituloPagar.iFilialEmpresa = objCT.ifilialempresatela

    'Verifica se campos identificadores correspondem a Título Baixado
    lErro = CF("TituloPagarBaixado_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18556 Then gError 196414

    'Se corresponder, erro
    If lErro = SUCESSO Then gError 196415

    'Lê Título Pagar a partir dos campos identificadores
    lErro = CF("TituloPagar_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18551 Then gError 196416

    'Se não encontrou ---> ERRO
    If lErro <> SUCESSO Then gError 196417

    'Pede confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_NFFATPAG", objTituloPagar.lNumTitulo)

    'Se não confirmar, sai
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui Nota Fiscal Fatura (incluindo dados contabeis (contabilidade))
    If objTituloPagar.sSiglaDocumento <> TIPODOC_FATURA_A_PAGAR Then
        lErro = CF("NFFatPag_Exclui", objTituloPagar, objCT.objContabil)
    Else
        lErro = CF("FaturaPagar_Exclui", objTituloPagar, objCT.objContabil)
    End If
    If lErro <> SUCESSO Then gError 196418

    'Limpa a Tela
    Call objCT.Limpa_Tela_TituloPagar

    giFilialEmpresa = iFilialAux

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    giFilialEmpresa = iFilialAux

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 196410
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 196411
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 196412
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)

        Case 196413
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)

        Case 196414, 196416, 196418, 196419

        Case 196415
            Call Rotina_Erro(vbOKOnly, "ERRO_NFFATPAG_BAIXADA_EXCLUSAO", gErr, objTituloPagar.lNumTitulo)

        Case 196417
            Call Rotina_Erro(vbOKOnly, "ERRO_NFFATPAG_NAO_CADASTRADA1", gErr, objTituloPagar.lNumTitulo)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196420)

    End Select

    Exit Sub

End Sub
