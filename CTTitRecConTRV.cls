VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTTitRecConTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub BotaoImprimir_Click(ByVal objCT As Object)

Dim lErro As Long
Dim sNomeArquivo As String
Dim sConteudo As String

On Error GoTo Erro_BotaoImprimir_Click

    lErro = CF("TRVConfig_Le", TRVCONFIG_DIRETORIO_FAT_HTML, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 194199
    
    sNomeArquivo = sConteudo & gsEmpresaTRVHTML & String(6 - Len(objCT.NumeroTitulo.Text), "0") & CStr(objCT.NumeroTitulo.Text) & ".html"

    Call Shell("explorer.exe " & sNomeArquivo, vbMaximizedFocus)
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 194199

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194200)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click(ByVal objCT As Object)

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim vbMsgRes As VbMsgBoxResult
Dim iFilialAux As Integer

On Error GoTo Erro_BotaoExcluir_Click

    iFilialAux = giFilialEmpresa
    giFilialEmpresa = objCT.ifilialempresatela

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se campos identificadores estão preenchidos
    If Len(Trim(objCT.Cliente.ClipText)) = 0 Then gError 194410
    If Len(Trim(objCT.Filial.Text)) = 0 Then gError 194411
    If Len(Trim(objCT.NumeroTitulo.Text)) = 0 Then gError 194412
    If Len(Trim(objCT.Tipo.Text)) = 0 Then gError 194413
    
    lErro = objCT.Move_Tela_Memoria(objTituloReceber)
    If lErro <> SUCESSO Then gError 194414
    
    objTituloReceber.iFilialEmpresa = objCT.ifilialempresatela
    
    'Verifica se campos identificadores correspondem a Título Baixado
    lErro = CF("TitulosRec_Le_Todos", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 58997 Then gError 194415
    
    'Se corresponder, erro
'    If lErro = SUCESSO Then gError 194416
    
'    'Lê Título Receber a partir dos campos identificadores
'    lErro = CF("TituloReceber_Le_Numero", objTituloReceber)
'    If lErro <> SUCESSO And lErro <> 26434 Then gError 194417
    
    'Se não encontrou  o Título à Receber --> erro
    If lErro <> SUCESSO Then gError 194418
    
    If objTituloReceber.iStatus = STATUS_BAIXADO Then gError 194416
    
    'Pede confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TITULORECEBER", objTituloReceber.lNumTitulo)
    
    'Se não confirmar, sai
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui o Titulo (inclusive a sua parte contábil)
    lErro = CF("TituloReceber_Exclui", objTituloReceber, objCT.objContabil)
    If lErro <> SUCESSO Then gError 194419
    
    'Limpa a Tela
    Call objCT.BotaoLimpar_Click
        
    giFilialEmpresa = iFilialAux
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    giFilialEmpresa = iFilialAux

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
        
        Case 194410
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 194411
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
    
        Case 194412
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
           
        Case 194413
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)

        Case 194414, 194415, 194417, 194419

        Case 194416
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_BAIXADO_EXCLUSAO", gErr, objTituloReceber.lNumTitulo, objTituloReceber.sSiglaDocumento)
            
        Case 194418
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO1", gErr, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo)
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194420)
            
    End Select
    
    Exit Sub
            
End Sub
