VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'API para envio de E-mails
'via VB
'Considera somente a criacao de somente
'um documento...

Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Constante que indica o tipo do anexo
'0 quer dizer que eh um arquivo de dados
Private Const mapData = 0

'Separa os atributos do e-mail
Private Const MAILSEPARADOR = ";"

' Class Enum for host name string validation
Public Enum VALIDATE_HOST_METHOD
    VALIDATE_HOST_NONE = 0
    VALIDATE_HOST_SYNTAX = 1
    VALIDATE_HOST_PING = 2
    VALIDATE_HOST_DNS = 3
End Enum

' Class Enum for email address string validation
Public Enum VALIDATE_METHOD
    VALIDATE_NONE = 0
    VALIDATE_SYNTAX = 1
End Enum

Public Function Email_Enviar(ByVal sTo As String, ByVal sSubject As String, ByVal sTexto As String, ByVal sAnexo As String, ByVal sAliasAnexo As String, ByVal sCc As String, ByVal sOutrosAnexos As String, ByVal sCco As String) As Long
Dim sEmailConfig As String
Dim lRetorno As Long, sErro As String
Dim iPgmEmail As Integer, sSMTP As String, lPorta As Long, sEmail As String, sUsu As String, sSenha As String, iSSL As Integer, sNomeExibicao As String, iConfirmacaoLeitura As Integer
Dim iHtml As Integer, sAnexos As String
On Error GoTo Erro_Email_Enviar
    sEmailConfig = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "EmailConfig", "", sEmailConfig, 255, "ADM100.INI")
    sEmailConfig = Left(sEmailConfig, lRetorno)
    Call Email_Config_Obtem_Dados(sEmailConfig, iPgmEmail, sSMTP, lPorta, sEmail, sNomeExibicao, sUsu, sSenha, iSSL, iConfirmacaoLeitura)
    If iPgmEmail = 1 Then
        Email_Enviar = Email_Enviar1(sTo, sSubject, sTexto, sAnexo, sAliasAnexo, sCc, sOutrosAnexos, sCco)
    Else
        Dim objEmail As Object
        Set objEmail = CreateObject("SGEUtil.Email")
        If InStr(sTexto, "<body>") <> 0 Then iHtml = 1
        sAnexos = sAnexo
        If Len(Trim(sOutrosAnexos)) > 0 Then
            If Right(sAnexo, 1) <> ";" And Len(Trim(sAnexos)) > 0 Then sAnexos = sAnexos & ";"
            sAnexos = sAnexos & sOutrosAnexos
        End If
        If Len(Trim(sAliasAnexo)) <> 0 Then
            If InStr(sAliasAnexo, ".") = 0 Then
                sAliasAnexo = sAliasAnexo & ".doc"
            End If
        End If
        Call objEmail.Inicializa(sSMTP, sUsu, sSenha, lPorta, iSSL)
        Email_Enviar = objEmail.Enviar_Email(sEmail, sNomeExibicao, "", sTo, sCc, sCco, sSubject, sTexto, iHtml, sAnexos, sAliasAnexo, iConfirmacaoLeitura, sErro)
        Call objEmail.Finaliza
    End If
    
    Exit Function
Erro_Email_Enviar:
    Email_Enviar = Err
    Select Case Err
        Case Else
            MsgBox "Erro Fornecido Pelo Visual Basic." & sErro, vbOKOnly, Err
    End Select
    Exit Function
End Function

'Funcao Email_Enviar
'Envia um e-mail para destinatario(s) sTo

'Descricao dos parametros:
'
'   sTo -> Destinatario(s)
'   sSubject -> Assunto da mensagem
'   sTexto -> Corpo da mensagem
'   sAnexo -> Path do arquivo a ser anexado
'   sAliasAnexo -> Com qual nome o arquivo saira na mensagem

Public Function Email_Enviar1(ByVal sTo As String, ByVal sSubject As String, ByVal sTexto As String, ByVal sAnexo As String, ByVal sAliasAnexo As String, ByVal sCc As String, ByVal sOutrosAnexos As String, ByVal sCco As String) As Long

Dim MAPISession1 As Object
Dim MAPIMessages1 As Object
Dim iPosToken As Integer
Dim sToAux As String, sCcAux As String, sAnexoAux As String, sCcoAux
Dim sLetraCorrente As String, sSTMPAux As String, lRetorno As Long

On Error GoTo Erro_Email_Enviar1
    
    sSTMPAux = String(255, 0)
    lRetorno = GetPrivateProfileString("Forprint", "MailSMTP", "smtp: ", sSTMPAux, 255, "ADM100.INI")
    sSTMPAux = Left(sSTMPAux, lRetorno)
    If sSTMPAux = "0" Then sSTMPAux = ""
    
    'Se Destinatario não foi preenchido
    If Len(Trim(sTo)) = 0 Then Error 1006
        
    Set MAPISession1 = CreateObject("MSMAPI.MAPISESSION")
    Set MAPIMessages1 = CreateObject("MSMAPI.MAPIMESSAGES")
            
    'Estabelece uma conexao
'------------------------------
    MAPISession1.LogonUI = True
    
    'Impedindo que os e-mails sejam baixados no inicio da conexao
    MAPISession1.DownLoadMail = False
    
    MAPISession1.SignOn
    MAPISession1.NewSession = True
    
'------------------------------
    
    'Apos estabelecida a conexao
    'um identificador da sessao eh
    'criado e armazenado na propriedade
    'SessionID do controle MAPISession.
    'Esse Identificador deve ser passado
    'para o controle MAPIMessages
    
    MAPIMessages1.SessionID = MAPISession1.SessionID
        
    'Criando a mensagem
    MAPIMessages1.MsgIndex = -1 '???
    
    MAPIMessages1.Compose
                     
    '--- a seguir, as propriedades que soh
    'podem ser alteradas em tempo de execucao
    
    MAPIMessages1.RecipIndex = 0 '???
    MAPIMessages1.RecipType = 1 '???

    'Atribuindo o destino
    
    For iPosToken = 1 To Len(sTo)
        sLetraCorrente = Mid(sTo, iPosToken, 1)
        
        If sLetraCorrente = MAILSEPARADOR Then
            MAPIMessages1.RecipAddress = sSTMPAux & sToAux
            MAPIMessages1.RecipDisplayName = MAPIMessages1.RecipAddress
            'MAPIMessages1.ResolveName
            MAPIMessages1.RecipIndex = MAPIMessages1.RecipCount
            sToAux = ""
            sLetraCorrente = ""
        End If
        
        sToAux = sToAux + sLetraCorrente
    Next
    
    MAPIMessages1.RecipAddress = sSTMPAux & sToAux
    MAPIMessages1.RecipDisplayName = MAPIMessages1.RecipAddress
    'MAPIMessages1.ResolveName
    
    If Len(Trim(sCc)) <> 0 Then
        
        'MAPIMessages1.ResolveName
        MAPIMessages1.RecipIndex = MAPIMessages1.RecipCount
        MAPIMessages1.RecipType = 2 '???
        
        For iPosToken = 1 To Len(sCc)
            sLetraCorrente = Mid(sCc, iPosToken, 1)
            
            If sLetraCorrente = MAILSEPARADOR Then
                MAPIMessages1.RecipAddress = sSTMPAux & sCcAux
                MAPIMessages1.RecipDisplayName = MAPIMessages1.RecipAddress
                'MAPIMessages1.ResolveName
                MAPIMessages1.RecipIndex = MAPIMessages1.RecipCount
                MAPIMessages1.RecipType = 2 '???
                sCcAux = ""
                sLetraCorrente = ""
            End If
            
            sCcAux = sCcAux + sLetraCorrente
        Next
    
        MAPIMessages1.RecipAddress = sSTMPAux & sCcAux
        MAPIMessages1.RecipDisplayName = MAPIMessages1.RecipAddress
        
    End If
    
    If Len(Trim(sCco)) <> 0 Then
        
        'MAPIMessages1.ResolveName
        MAPIMessages1.RecipIndex = MAPIMessages1.RecipCount
        MAPIMessages1.RecipType = 3 '???
        
        For iPosToken = 1 To Len(sCco)
            sLetraCorrente = Mid(sCco, iPosToken, 1)
            
            If sLetraCorrente = MAILSEPARADOR Then
                MAPIMessages1.RecipAddress = sSTMPAux & sCcoAux
                MAPIMessages1.RecipDisplayName = MAPIMessages1.RecipAddress
                'MAPIMessages1.ResolveName
                MAPIMessages1.RecipIndex = MAPIMessages1.RecipCount
                MAPIMessages1.RecipType = 3 '???
                sCcoAux = ""
                sLetraCorrente = ""
            End If
            
            sCcoAux = sCcoAux + sLetraCorrente
        Next
    
        MAPIMessages1.RecipAddress = sSTMPAux & sCcoAux
        MAPIMessages1.RecipDisplayName = MAPIMessages1.RecipAddress
        
    End If
    
    '''MAPIMessages1.AddressResolveUI = True '???
    '''MAPIMessages1.ResolveName '???
    'Atribuindo o subject se foi passado
    If Len(Trim(sSubject)) <> 0 Then
        MAPIMessages1.MsgSubject = sSubject
    End If
    
    'Atribuindo o texto se foi passado
    If Len(Trim(sTexto)) <> 0 Then
        MAPIMessages1.MsgNoteText = sTexto
    Else
        If Len(Trim(sSubject)) <> 0 Then
            MAPIMessages1.MsgNoteText = sSubject
        Else
            MAPIMessages1.MsgNoteText = "..."
        End If
    End If
    
    'Se o anexo foi passado...
    If Len(Trim(sAnexo)) <> 0 Then

        MAPIMessages1.AttachmentIndex = 0 '???
        MAPIMessages1.AttachmentPosition = 0
        
'???
'''        'Definir tipo de dados do anexo
'''        MAPIMessages1.AttachmentType = mapData

        'Indicar o path
        MAPIMessages1.AttachmentPathName = sAnexo

        'Nome (Alias) do anexo na mensagem se foi passado
        If Len(Trim(sAliasAnexo)) <> 0 Then
            If InStr(sAliasAnexo, ".") = 0 Then
                sAliasAnexo = sAliasAnexo & ".doc"
            End If

            MAPIMessages1.AttachmentName = sAliasAnexo
        End If
    End If
    
    'Se outros anexos foram passados...
    If Len(Trim(sOutrosAnexos)) <> 0 Then

        MAPIMessages1.AttachmentIndex = MAPIMessages1.AttachmentCount
        MAPIMessages1.AttachmentPosition = MAPIMessages1.AttachmentIndex
        
        For iPosToken = 1 To Len(sOutrosAnexos)
            sLetraCorrente = Mid(sOutrosAnexos, iPosToken, 1)
            
            If sLetraCorrente = MAILSEPARADOR Then
                
                MAPIMessages1.AttachmentPathName = sAnexoAux
                MAPIMessages1.AttachmentName = sAnexoAux
                
                MAPIMessages1.AttachmentIndex = MAPIMessages1.AttachmentCount
                MAPIMessages1.AttachmentPosition = MAPIMessages1.AttachmentIndex
                
                sAnexoAux = ""
                sLetraCorrente = ""
                
            End If
            
            sAnexoAux = sAnexoAux + sLetraCorrente
        Next
    
        MAPIMessages1.AttachmentPathName = sAnexoAux
        MAPIMessages1.AttachmentName = sAnexoAux
        
    End If
    
    MAPIMessages1.MsgReceiptRequested = True
    DoEvents '???
    'Envio Efetivo da mensagem
    MAPIMessages1.Send False
       
    'Fechando a conexao
    MAPISession1.SignOff
   
    MAPISession1.NewSession = False

    Email_Enviar1 = 0
    
    Exit Function

Erro_Email_Enviar1:

    Email_Enviar1 = Err

    Select Case Err

        Case 1006
            MsgBox "Erro - Destinatário não preenchido.", vbOKOnly, Err
        
        Case Else
            MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err
        
    End Select
    
    'Fechando a conexao
    MAPISession1.SignOff
   
    MAPISession1.NewSession = False
    
    Exit Function
    
End Function

Public Function Email_Enviar2(ByVal sSMPTHost As String, ByVal sFrom As String, ByVal sFromName As String, ByVal sUser As String, ByVal sPassword As String, ByVal sTo As String, ByVal sSubject As String, ByVal sTexto As String, ByVal sAnexo As String, ByVal sCc As String, ByVal sCco As String) As Long

Dim poSendMail As Object

On Error GoTo Erro_Email_Enviar2

    Set poSendMail = CreateObject("vbSendMail.clsSendMail")

    With poSendMail
        
        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = MAILSEPARADOR                            ' Optional, default = ";" (semicolon)
        .SMTPHost = sSMPTHost                   ' Required the fist time, optional thereafter
        .From = sFrom                         ' Required the fist time, optional thereafter
        .FromDisplayName = sFromName          ' Optional, saved after first use
        .Recipient = sTo                   ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = sTo       ' Optional, separate multiple entries with delimiter character
        .CcRecipient = sCc                        ' Optional, separate multiple entries with delimiter character
        .CcDisplayName = sCc               ' Optional, separate multiple entries with delimiter character
        .BccRecipient = sCco                    ' Optional, separate multiple entries with delimiter character
        .ReplyToAddress = sFrom            ' Optional, used when different than 'From' address
        .Subject = sSubject                 ' Optional
        .Message = sTexto                      ' Optional
        .Attachment = sAnexo                'Trim(txtAttach.Text)          ' Optional, separate multiple entries with delimiter character
        .AsHTML = True                              ' Optional, default = FALSE, send mail as html or plain text
        .UseAuthentication = True             ' Optional, default = FALSE
        .Username = sUser                    ' Optional, default = Null String
        .Password = sPassword                    ' Optional, default = Null String, value is NOT saved
        
        .Send                                       ' Required
                                                    ' MX look up was used to find a host    End With
    End With
    
    Exit Function

Erro_Email_Enviar2:

    Email_Enviar2 = Err

    Select Case Err

        Case Else
            MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err
        
    End Select

    Exit Function

End Function

Public Function Email_Enviar3(ByVal sSMPTHost As String, ByVal sFrom As String, ByVal sFromName As String, ByVal sUser As String, ByVal sPassword As String, ByVal sTo As String, ByVal sSubject As String, ByVal sTexto As String, ByVal sAnexo As String, ByVal sCc As String, ByVal sCco As String, ByVal sReplyToAddress As String, ByVal bConfirmaEnvio As Boolean, Optional ByVal objClasseTrataEventos As Object, Optional ByVal sNomeProgresso As String, Optional ByVal sNomeFalha As String, Optional ByVal sNomeSucesso As String, Optional ByVal sNomeStatus As String) As Long

Dim objADEnviarEmail As Object

On Error GoTo Erro_Email_Enviar3

    Set objADEnviarEmail = CreateObject("ADEnviarEmail.ClassVBEnvioEmail")
    
    Call objADEnviarEmail.Inicializa(sSMPTHost, sUser, sPassword, objClasseTrataEventos, sNomeProgresso, sNomeFalha, sNomeSucesso, sNomeStatus)

    Email_Enviar3 = objADEnviarEmail.Email_Enviar(sFrom, sFromName, sTo, sSubject, sTexto, sAnexo, sCc, sCco, sReplyToAddress, bConfirmaEnvio)

    Call objADEnviarEmail.Finaliza

    Exit Function

Erro_Email_Enviar3:

    Email_Enviar3 = Err

    Select Case Err

        Case Else
            MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err
        
    End Select
    
    If Not (objADEnviarEmail Is Nothing) Then objADEnviarEmail.Finaliza

    Exit Function
    
End Function

Private Sub Email_Config_Obtem_Dados(ByVal sEmailConfig As String, iPgmEmail As Integer, sSMTP As String, lPorta As Long, sEmail As String, sNomeExibicao As String, sUsu As String, sSenha As String, iSSL As Integer, iConfirmacaoLeitura As Integer)
Dim sAux As String, iPos As Integer
'    sEmailConfig = objEmailConfig.sEmail
'    sEmailConfig = sEmailConfig & "|" & objEmailConfig.sSMTP
'    sEmailConfig = sEmailConfig & "|" & objEmailConfig.sSMTPUsu
'    sEmailConfig = sEmailConfig & "|" & objEmailConfig.sSMTPSenha
'    sEmailConfig = sEmailConfig & "|" & objEmailConfig.sNome
'    sEmailConfig = sEmailConfig & "|" & objEmailConfig.lSMTPPorta
'    sEmailConfig = sEmailConfig & "|" & objEmailConfig.iSSL
'    sEmailConfig = sEmailConfig & "|" & objEmailConfig.iPgmEmail
'    sEmailConfig = sEmailConfig & "|" & objEmailConfig.iConfirmacaoLeitura & "|"

    iPgmEmail = 1
    If Len(Trim(sEmailConfig)) <> 0 Then
        sAux = sEmailConfig
        iPos = InStr(1, sAux, "|")
        sEmail = Left(sAux, iPos - 1)
        
        sAux = Mid(sAux, iPos + 1)
        iPos = InStr(1, sAux, "|")
        sSMTP = Left(sAux, iPos - 1)
    
        sAux = Mid(sAux, iPos + 1)
        iPos = InStr(1, sAux, "|")
        sUsu = Left(sAux, iPos - 1)
        
        sAux = Mid(sAux, iPos + 1)
        iPos = InStr(1, sAux, "|")
        sSenha = Left(sAux, iPos - 1)
        
        sAux = Mid(sAux, iPos + 1)
        iPos = InStr(1, sAux, "|")
        sNomeExibicao = Left(sAux, iPos - 1)
        
        sAux = Mid(sAux, iPos + 1)
        iPos = InStr(1, sAux, "|")
        lPorta = CLng(Left(sAux, iPos - 1))
        
        sAux = Mid(sAux, iPos + 1)
        iPos = InStr(1, sAux, "|")
        iSSL = CInt(Left(sAux, iPos - 1))
        
        sAux = Mid(sAux, iPos + 1)
        iPos = InStr(1, sAux, "|")
        iPgmEmail = CInt(Left(sAux, iPos - 1))
        
        sAux = Mid(sAux, iPos + 1)
        iPos = InStr(1, sAux, "|")
        iConfirmacaoLeitura = CInt(Left(sAux, iPos - 1))
    End If

End Sub
