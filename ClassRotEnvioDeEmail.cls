VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotEnviodeEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function Rotina_Envia_Emails(ByVal colEnvioEmail As Collection) As Long

Dim lErro As Long
Dim objEmail As Object
Dim sConteudoHtml As String
Dim sSMTPHost As String
Dim sUsuario As String
Dim sSenha As String
Dim objEnvioEmail As ClassEnvioDeEmail
Dim sArquivoHtml As String
Dim sNomeExibicao As String
Dim iLinha As Integer
Dim sTipoEnvio As String
Dim sArqLic As String
Dim sLogo As String
Dim sImagens As String
Dim sURLImgCB As String
Dim sURLLogo As String
Dim sAgencia As String
Dim sConta As String
Dim objEmailConfig As New ClassEmailConfig
Dim lPorta As Long, iSSL As Integer
Dim sErro As String, sPosEnv As String

On Error GoTo Erro_Rotina_Envia_Emails
    
    objEmailConfig.sUsuario = gsUsuario
    
    lErro = CF("EmailConfig_Le", objEmailConfig)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    If lErro <> SUCESSO Then gError 202830
    
    If objEmail Is Nothing Then
        If objEmailConfig.iPgmEmail = 1 Then
            Set objEmail = CreateObject("ADEnviarEmail.ClassVBEnvioEmail")
        Else
            Set objEmail = CreateObject("SGEUtil.Email")
        End If
    End If
    
    sSMTPHost = objEmailConfig.sSMTP
    sUsuario = objEmailConfig.sSMTPUsu
    sSenha = objEmailConfig.sSMTPSenha
    lPorta = objEmailConfig.lSMTPPorta
    iSSL = objEmailConfig.iSSL
    
    If objEmailConfig.iPgmEmail = 1 Then
        lErro = objEmail.Inicializa(sSMTPHost, sUsuario, sSenha, TelaAcompanhaBatchCPR2, "Trata_Progresso", "Trata_Falha", "Trata_Sucesso", "Trata_Status", lPorta)
    Else
        lErro = objEmail.Inicializa(sSMTPHost, sUsuario, sSenha, lPorta, iSSL)
    End If
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iLinha = 0
    For Each objEnvioEmail In colEnvioEmail
    
        If Len(Trim(objEnvioEmail.sDe)) = 0 Then objEnvioEmail.sDe = sUsuario
    
        iLinha = iLinha + 1
        
        sPosEnv = CStr(iLinha) & " de " & CStr(colEnvioEmail.Count)
                 
        If Len(Trim(objEnvioEmail.sModelo)) <> 0 Then
            
            lErro = Gera_Arquivo_Html(objEnvioEmail.objTela, objEnvioEmail.sModelo, sArquivoHtml, sConteudoHtml, objEnvioEmail.iLinha)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        Else
            sConteudoHtml = objEnvioEmail.sMensagem
        End If
            
        If Len(Trim(objEnvioEmail.sDe)) = 0 Then gError 189422
        
        If Len(Trim(objEnvioEmail.sNomeExibicao)) = 0 Then
            sNomeExibicao = objEnvioEmail.sDe
        Else
            sNomeExibicao = objEnvioEmail.sNomeExibicao
        End If

        'Envia pelo WinSock
        If objEmailConfig.iPgmEmail = 1 Then
            lErro = objEmail.Email_Enviar(objEnvioEmail.sDe, sNomeExibicao, objEnvioEmail.sEmail, objEnvioEmail.sAssunto, sConteudoHtml, objEnvioEmail.sAnexo, objEnvioEmail.sCC, "", objEnvioEmail.sEmailResp, IIf(objEnvioEmail.iConfirmacaoLeitura = MARCADO, True, False))
        Else
            lErro = objEmail.Enviar_Email(objEnvioEmail.sDe, sNomeExibicao, objEnvioEmail.sEmailResp, objEnvioEmail.sEmail, objEnvioEmail.sCC, "", objEnvioEmail.sAssunto, sConteudoHtml, 1, objEnvioEmail.sAnexo, "", objEnvioEmail.iConfirmacaoLeitura, sErro)
            If lErro = SUCESSO Then
                Call TelaAcompanhaBatchCPR2.Trata_Status("SUCESSO NO ENVIO " & sPosEnv & ": " & objEnvioEmail.sAssunto)
            Else
                Call TelaAcompanhaBatchCPR2.Trata_Status("FALHA NO ENVIO " & sPosEnv & ": " & objEnvioEmail.sAssunto & sErro)
            End If
        End If
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = TelaAcompanhaBatchCPR2.ProcessouItem
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = CF("EmailsEnviados_Grava", objEnvioEmail)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objEnvioEmail.iGeraRelac = MARCADO Then
            lErro = Gera_Relacionamento(objEnvioEmail.lClienteRelac, objEnvioEmail.iFilialCliRelac, IIf(objEnvioEmail.iTipoDoc <> EMAIL_TIPODOC_PARCELASREC, 0, objEnvioEmail.lNumIntDocParc), objEnvioEmail.sTextoRelac)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
        
        If Len(Trim(objEnvioEmail.sModelo)) <> 0 Then
            
            lErro = Apaga_Arquivo_Html(sArquivoHtml)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
                   
'        lErro = TelaAcompanhaBatchCPR2.ProcessouItem
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    Next
    
    If Not (objEmail Is Nothing) Then
        Call objEmail.Finaliza
        Set objEmail = Nothing
    End If
    
    Rotina_Envia_Emails = SUCESSO

    Exit Function
    
Erro_Rotina_Envia_Emails:

    Rotina_Envia_Emails = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 189422
            Call Rotina_Erro(vbOKOnly, "ERRO_DE_NAO_PREENCHIDO", gErr, objEnvioEmail.iLinha)
        
        Case 202830
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAILCONFIG_NAO_CADASTRADO", gErr, objEmailConfig.sUsuario)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187040)

    End Select
    
    If Not (objEmail Is Nothing) Then
        Call objEmail.Finaliza
        Set objEmail = Nothing
    End If
    
    Exit Function
    
End Function

Function Gera_Arquivo_Html(ByVal objTela As Object, ByVal sModelo As String, sArquivo As String, sConteudoHtml As String, Optional ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objTS1 As TextStream
Dim objTS2 As TextStream
Dim objFSO As New FileSystemObject
Dim sRegistro As String
Dim bArqAberto1 As Boolean
Dim bArqAberto2 As Boolean
Dim iPos1 As Long
Dim iPos2 As Long
Dim objMnemonicoValor As ClassMnemonicoValor
Dim sValorTexto As String
Dim sDirTsk As String

On Error GoTo Erro_Gera_Arquivo_Html

    sArquivo = ""

    If Len(Trim(sModelo)) > 0 Then
    
        sDirTsk = String(512, 0)
        Call GetPrivateProfileString("Forprint", "DirTsks", "0", sDirTsk, 512, "ADM100.INI")
        sDirTsk = Replace(sDirTsk, Chr(0), "")
        
        sArquivo = sDirTsk & gsUsuario & Format(Now, "YYYYMMDDHHMMSS") & ".html"
    
        bArqAberto1 = False
        Set objTS1 = objFSO.OpenTextFile(sModelo, IOMode.ForReading, False)
        bArqAberto1 = True

        bArqAberto2 = False
        Set objTS2 = objFSO.OpenTextFile(sArquivo, IOMode.ForWriting, True)
        bArqAberto2 = True

        'Até chegar ao fim do arquivo
        Do While Not objTS1.AtEndOfLine
        
            'Busca o próximo registro do arquivo
            sRegistro = objTS1.ReadAll
             
            iPos1 = InStr(1, sRegistro, "ZZZ")
             
            Do While iPos1 <> 0
             
                iPos2 = InStr(iPos1 + 3, sRegistro, "ZZZ")
                 
                Set objMnemonicoValor = New ClassMnemonicoValor
                Set objMnemonicoValor.colValor = New Collection
                 
                objMnemonicoValor.sMnemonico = Mid(sRegistro, iPos1 + 3, iPos2 - iPos1 - 3)
                 
                lErro = objTela.Calcula_Mnemonico(objMnemonicoValor, sValorTexto, iLinha)
                If lErro <> SUCESSO Then gError 189320
             
                sRegistro = Replace(sRegistro, "ZZZ" & objMnemonicoValor.sMnemonico & "ZZZ", sValorTexto)
                
                iPos1 = InStr(1, sRegistro, "ZZZ")
            Loop
            
            objTS2.Write sRegistro
            
            sConteudoHtml = sRegistro

        Loop
        
        objTS1.Close
        bArqAberto1 = False
        objTS2.Close
        bArqAberto2 = False

    End If

    Gera_Arquivo_Html = SUCESSO

    Exit Function

Erro_Gera_Arquivo_Html:

    Gera_Arquivo_Html = gErr

    Select Case gErr
    
        Case 53 'ERRO_ARQUIVO_FTP_NAO_ENCONTRADO
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_FTP_NAO_ENCONTRADO", gErr, sModelo)
            
        Case 189320

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189320)

    End Select
    
    'fechar aquivo texto
    If bArqAberto1 Then
        objTS1.Close
    End If

    'fechar aquivo texto
    If bArqAberto2 Then
        objTS2.Close
        Call Apaga_Arquivo_Html(sArquivo)
    End If

    Exit Function

End Function

Private Function Gera_Relacionamento(ByVal lCliente As Long, ByVal iFilial As Integer, ByVal lNumIntParcRec As Long, ByVal sMsg As String) As Long

Dim lErro As Long
Dim objRelacCli As New ClassRelacClientes

On Error GoTo Erro_Gera_Relacionamento

    objRelacCli.lCliente = lCliente
    objRelacCli.iFilialCliente = iFilial
    objRelacCli.dtData = Date
    objRelacCli.dtDataPrevReceb = DATA_NULA
    objRelacCli.dtDataProxCobr = DATA_NULA
    objRelacCli.iFilialEmpresa = giFilialEmpresa
    If Len(Trim(sMsg)) < STRING_BUFFER_MAX_TEXTO Then
        objRelacCli.sAssunto1 = sMsg
    Else
        objRelacCli.sAssunto1 = left(sMsg, STRING_BUFFER_MAX_TEXTO - 1)
        objRelacCli.sAssunto2 = Mid(sMsg, STRING_BUFFER_MAX_TEXTO)
    End If
    objRelacCli.lNumIntParcRec = lNumIntParcRec
    objRelacCli.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA
    objRelacCli.dtHora = Time
    
    objRelacCli.lTipo = TIPO_RELACIONAMENTO_COBRANCA
    
    'Alterado para gerar o relacionamento com status encerrado, senão fica como se algo estivesse pendente
    objRelacCli.iStatus = RELACIONAMENTOCLIENTES_STATUS_ENCERRADO
    objRelacCli.dtDataFim = Date
    objRelacCli.dtHoraFim = Now
    
    lErro = CF("RelacionamentoClientes_Grava", objRelacCli, True, gsUsuario)
    If lErro <> SUCESSO Then gError 189324
    
    Gera_Relacionamento = SUCESSO

    Exit Function

Erro_Gera_Relacionamento:

    Gera_Relacionamento = gErr

    Select Case gErr
    
        Case 189324

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189321)

    End Select

    Exit Function

End Function

Function Apaga_Arquivo_Html(ByVal sArquivo As String) As Long

Dim lErro As Long
Dim objFSO As New FileSystemObject

On Error GoTo Erro_Apaga_Arquivo_Html

    If Len(Trim(sArquivo)) > 0 Then

        Call objFSO.DeleteFile(sArquivo)
        
    End If

    Apaga_Arquivo_Html = SUCESSO

    Exit Function

Erro_Apaga_Arquivo_Html:

    Apaga_Arquivo_Html = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 189322)

    End Select

    Exit Function

End Function
