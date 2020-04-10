VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVBEnvioEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents poSendMail As ClassSendMail
Attribute poSendMail.VB_VarHelpID = -1

Const MAILSEPARADOR = ";"

' misc local vars
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean

Dim gobjClasseTrataEventos As Object
Dim gsNomeProgresso As String
Dim gsNomeFalha As String
Dim gsNomeSucesso As String
Dim gsNomeStatus As String

Public Function Inicializa(ByVal sSMTPHost As String, ByVal sUsuario As String, ByVal sSenha As String, Optional ByVal objClasseTrataEventos As Object, Optional ByVal sNomeProgresso As String, Optional ByVal sNomeFalha As String, Optional ByVal sNomeSucesso As String, Optional ByVal sNomeStatus As String, Optional lPorta As Long)
    Dim bCnx As Boolean
    Dim iIndice As Integer
    Set poSendMail = New ClassSendMail
    
    Set gobjClasseTrataEventos = objClasseTrataEventos
    gsNomeProgresso = sNomeProgresso
    gsNomeFalha = sNomeFalha
    gsNomeSucesso = sNomeSucesso
    gsNomeStatus = sNomeStatus
    
    poSendMail.SMTPHost = sSMTPHost
    poSendMail.Username = sUsuario
    poSendMail.Password = sSenha
    poSendMail.SMTPHostValidation = VALIDATE_HOST_SYNTAX         ' Optional, default = VALIDATE_HOST_DNS
    poSendMail.EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
    poSendMail.Delimiter = MAILSEPARADOR                            ' Optional, default = ";" (semicolon)
    poSendMail.AsHTML = True                              ' Optional, default = FALSE, send mail as html or plain text
    poSendMail.UseAuthentication = True             ' Optional, default = FALSE
    poSendMail.NaoDesconectar = True 'Para não desconectar após cada envio
    
    If lPorta <> 0 Then
        poSendMail.SMTPPort = lPorta
    End If
'    bCnx = False
'    iIndice = 1
'    Do While Not bCnx
'        bCnx = poSendMail.Connect
'        iIndice = iIndice + 1
'        Call poSendMail_Status("Tentativa " & CStr(iIndice))
'        If iIndice > 3 Then Exit Do
'    Loop
End Function

Public Function Finaliza()
    Call poSendMail.Disconnect
    Set poSendMail = Nothing
End Function

Public Function Email_Enviar(ByVal sFrom As String, ByVal sFromName As String, ByVal sTo As String, ByVal sSubject As String, ByVal sTexto As String, ByVal sAnexo As String, ByVal sCc As String, ByVal sCco As String, ByVal sReplyToAddress As String, ByVal bConfirmaEnvio As Boolean) As Long

On Error GoTo Erro_Email_Enviar
    
    With poSendMail
        
        .from = sFrom                         ' Required the fist time, optional thereafter
        .FromDisplayName = sFromName          ' Optional, saved after first use
        .Recipient = sTo                   ' Required, separate multiple entries with delimiter character
        .RecipientDisplayName = sTo       ' Optional, separate multiple entries with delimiter character
        .CcRecipient = sCc                        ' Optional, separate multiple entries with delimiter character
        .CcDisplayName = sCc               ' Optional, separate multiple entries with delimiter character
        .BccRecipient = sCco                    ' Optional, separate multiple entries with delimiter character
        .ReplyToAddress = sReplyToAddress            ' Optional, used when different than 'From' address
        .Subject = sSubject                 ' Optional
        .Message = sTexto                      ' Optional
        .Attachment = sAnexo                'Trim(txtAttach.Text)          ' Optional, separate multiple entries with delimiter character
        .Receipt = bConfirmaEnvio
        
        .Send                                       ' Required
                                                    ' MX look up was used to find a host    End With
    End With
    
    Exit Function

Erro_Email_Enviar:

    Email_Enviar = Err

    Select Case Err

        Case Else
            MsgBox "Erro Fornecido Pelo Visual Basic.", vbOKOnly, Err
        
    End Select

    Exit Function

End Function

Private Sub poSendMail_Progress(lPercentCompete As Long)
    If Not (gobjClasseTrataEventos Is Nothing) Then
        If Len(Trim(gsNomeProgresso)) > 0 Then
            Call CallByName(gobjClasseTrataEventos, gsNomeProgresso, VbMethod, lPercentCompete)
        End If
    End If
End Sub

Private Sub poSendMail_SendFailed(Explanation As String)
    Dim sMgs As String
    sMgs = Traduz_Explanation(Explanation)
    If Not (gobjClasseTrataEventos Is Nothing) Then
        If Len(Trim(gsNomeFalha)) > 0 Then
            Call CallByName(gobjClasseTrataEventos, gsNomeFalha, VbMethod, sMgs)
        End If
    End If
End Sub

Private Sub poSendMail_SendSuccesful()
    If Not (gobjClasseTrataEventos Is Nothing) Then
        If Len(Trim(gsNomeSucesso)) > 0 Then
            Call CallByName(gobjClasseTrataEventos, gsNomeSucesso, VbMethod)
        End If
    End If
End Sub

Private Sub poSendMail_Status(Status As String)
    Dim sMgs As String
    sMgs = Traduz_Status(Status)
    If Not (gobjClasseTrataEventos Is Nothing) Then
        If Len(Trim(gsNomeStatus)) > 0 And Len(Trim(sMgs)) > 0 Then
            Call CallByName(gobjClasseTrataEventos, gsNomeStatus, VbMethod, sMgs)
        End If
    End If
End Sub

Private Function Traduz_Status(Status As String) As String
    Traduz_Status = Trim(Replace(Replace(Replace(Replace(Status, Chr(13), ""), Chr(10), ""), vbNullChar, ""), vbNullString, ""))
End Function

Private Function Traduz_Explanation(Explanation As String) As String
    Traduz_Explanation = Trim(Replace(Replace(Replace(Replace(Explanation, Chr(13), ""), Chr(10), ""), vbNullChar, ""), vbNullString, ""))
End Function
