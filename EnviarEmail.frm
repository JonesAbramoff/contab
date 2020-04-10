VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EnviarEmail 
   Caption         =   "Enviar Email"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoProcurar 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6390
      TabIndex        =   19
      Top             =   3180
      Width           =   495
   End
   Begin VB.TextBox Cco 
      Height          =   285
      Left            =   1920
      MaxLength       =   8000
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2010
      Width           =   4935
   End
   Begin VB.TextBox OutrosAnexos 
      Height          =   285
      Left            =   1920
      MaxLength       =   250
      TabIndex        =   6
      Top             =   3210
      Width           =   4365
   End
   Begin VB.TextBox Cc 
      Height          =   285
      Left            =   1920
      MaxLength       =   8000
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1620
      Width           =   4935
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5160
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   150
      Width           =   1680
      Begin VB.CommandButton Cancelar 
         Height          =   360
         Left            =   1110
         Picture         =   "EnviarEmail.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "EnviarEmail.frx":017E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton Enviar 
         Height          =   360
         Left            =   90
         Picture         =   "EnviarEmail.frx":06B0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Enviar email"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.TextBox Mensagem 
      Height          =   825
      Left            =   180
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3990
      Width           =   6675
   End
   Begin VB.TextBox Arquivo 
      Height          =   285
      Left            =   1920
      MaxLength       =   250
      TabIndex        =   5
      Top             =   2790
      Width           =   4935
   End
   Begin VB.TextBox Assunto 
      Height          =   285
      Left            =   1920
      MaxLength       =   250
      TabIndex        =   4
      Top             =   2385
      Width           =   4935
   End
   Begin VB.TextBox Para 
      Height          =   645
      Left            =   1920
      MaxLength       =   8000
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   1245
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LabelCco 
      Alignment       =   1  'Right Justify
      Caption         =   "Cco:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1455
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   18
      Top             =   2025
      Width           =   405
   End
   Begin VB.Label LabelInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   135
      TabIndex        =   17
      Top             =   150
      Width           =   4920
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Outros Anexos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   390
      TabIndex        =   16
      Top             =   3255
      Width           =   1455
   End
   Begin VB.Label LabelCc 
      Caption         =   "Cc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1545
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   15
      Top             =   1605
      Width           =   330
   End
   Begin VB.Label Label4 
      Caption         =   "Mensagem:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   3705
      Width           =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "Anexo do Relatório:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   9
      ToolTipText     =   "Nome do anexo que conterá a imagem do relatório"
      Top             =   2835
      Width           =   1725
   End
   Begin VB.Label Label2 
      Caption         =   "Assunto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1095
      TabIndex        =   8
      Top             =   2400
      Width           =   765
   End
   Begin VB.Label LabelPara 
      Caption         =   "Para:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1395
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   0
      Top             =   825
      Width           =   510
   End
End
Attribute VB_Name = "EnviarEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Substituido por nova versao em 08/01/2002 (Tulio)
'Comentado por Daniel em 08/01/2002 a pedido de Jones.
'Codificado por Tulio em 07/01/2002

Dim MAPISession As Object
Dim MAPIMessages As Object
Dim MAPIMessagesCC As Object
Dim MAPIMessagesCCo As Object

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio '??? lixo
'Dim gsOpcao As String '??? lixo, só está aqui p/nao dar erro de compa

Dim giPgmEmail As Integer

Private Sub BotaoLimpar_Click()
    Call Limpa_Tela(Me)
End Sub

Private Sub Cancelar_Click()
'Criado por Tulio em 07/01/2002

    '??? gsOpcao = "Cancelar"
    gobjRelOpcoes.bDesistiu = True
    Unload Me
    
End Sub

Private Sub Enviar_Click()
'Criado por Tulio em 07/01/2002

Dim lErro As Long, iPartes As Integer, iIndice As Integer

On Error GoTo Erro_Enviar_Click

    If Len(Trim(Para.Text)) = 0 Then gError 124012
    
    lErro = gobjRelOpcoes.IncluirParametro("TTO_EMAIL", left(Para.Text, 250))
    If lErro <> AD_BOOL_TRUE Then gError 97110
    
    iPartes = Arredonda_ParaCima((0# + Len(Para.Text)) / 250#)
    For iIndice = 1 To (iPartes - 1)
    
        lErro = gobjRelOpcoes.IncluirParametro("TTO_EMAIL" & CStr(iIndice), Mid(Para.Text, 1 + (250 * iIndice), 250))
        If lErro <> AD_BOOL_TRUE Then gError 97110
    
    Next

    lErro = gobjRelOpcoes.IncluirParametro("TCC_EMAIL", left(Cc.Text, 250))
    If lErro <> AD_BOOL_TRUE Then gError 97110
    
    If Len(Cc.Text) > 250 Then
    
        iPartes = Arredonda_ParaCima((0# + Len(Cc.Text)) / 250#)
        For iIndice = 1 To (iPartes - 1)
        
            lErro = gobjRelOpcoes.IncluirParametro("TCC_EMAIL" & CStr(iIndice), Mid(Cc.Text, 1 + (250 * iIndice), 250))
            If lErro <> AD_BOOL_TRUE Then gError 97110
        
        Next
    
    End If
    
    lErro = gobjRelOpcoes.IncluirParametro("TCCO_EMAIL", left(Cco.Text, 250))
    If lErro <> AD_BOOL_TRUE Then gError 97110
    
    If Len(Cco.Text) > 250 Then
    
        iPartes = Arredonda_ParaCima((0# + Len(Cco.Text)) / 250#)
        For iIndice = 1 To (iPartes - 1)
        
            lErro = gobjRelOpcoes.IncluirParametro("TCCO_EMAIL" & CStr(iIndice), Mid(Cco.Text, 1 + (250 * iIndice), 250))
            If lErro <> AD_BOOL_TRUE Then gError 97110
        
        Next
    
    End If
    
    lErro = gobjRelOpcoes.IncluirParametro("TSUBJECT", Assunto.Text)
    If lErro <> AD_BOOL_TRUE Then gError 97111

    lErro = gobjRelOpcoes.IncluirParametro("TALIASATTACH", IIf(Len(Trim(Arquivo.Text)) = 0, "anexo" & gsExtensaoGerRelExp, Arquivo.Text & gsExtensaoGerRelExp))
    If lErro <> AD_BOOL_TRUE Then gError 97112

    lErro = gobjRelOpcoes.IncluirParametro("TOUTROSANEXOS", IIf(Len(Trim(OutrosAnexos.Text)) = 0, "", OutrosAnexos.Text))
    If lErro <> AD_BOOL_TRUE Then gError 97112

    lErro = gobjRelOpcoes.IncluirParametro("TMENSAGEM", Mensagem.Text)
    If lErro <> AD_BOOL_TRUE Then gError 97113
    
    Call CF("EmailConfig_Grava_INI")
    
    '??? gsOpcao = "Enviar"
    
    Unload Me
    
    Exit Sub

Erro_Enviar_Click:

    Select Case gErr
            
        Case 97110 To 97113
                    
        Case 124012
            Call Rotina_Erro(vbOKOnly, "ERRO_EMAIL_SEM_DESTINATARIO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159468)

    End Select
    
    Exit Sub

End Sub

Private Sub Form_Load()

Dim iPgmEmail As Integer

On Error GoTo Erro_Form_Load

    iPgmEmail = -1
    
    lErro = CF("Tabela_Le_Campo", "EmailConfig", "PgmEmail", "Usuario='" & CStr(gsUsuario) & "'", TIPO_CAMPO_INTEGER, iPgmEmail)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If iPgmEmail = -1 Then
    
        lErro = CF("Tabela_Le_Campo", "EmailConfig", "PgmEmail", "Usuario=''", TIPO_CAMPO_INTEGER, iPgmEmail)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    giPgmEmail = iPgmEmail
    
    If giPgmEmail <> 0 Then

        Set MAPISession = CreateObject("MSMAPI.MAPISESSION")
        Set MAPIMessages = CreateObject("MSMAPI.MAPIMESSAGES")
        Set MAPIMessagesCC = CreateObject("MSMAPI.MAPIMESSAGES")
        Set MAPIMessagesCCo = CreateObject("MSMAPI.MAPIMESSAGES")
    
        MAPISession.LogonUI = True
        
        'Impedindo que os e-mails sejam baixados no inicio da conexao
        MAPISession.DownLoadMail = False
        
        MAPISession.SignOn
        MAPISession.NewSession = True
        
        MAPIMessages.SessionID = MAPISession.SessionID
        MAPIMessagesCC.SessionID = MAPISession.SessionID
        MAPIMessagesCCo.SessionID = MAPISession.SessionID
            
        'Criando a mensagem
        MAPIMessages.Compose
    
        'Criando a mensagem
        MAPIMessagesCC.Compose
    
        'Criando a mensagem
        MAPIMessagesCCo.Compose
        
    End If

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 32003 'Erro de falha no logon
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159469)

    End Select

    Exit Sub
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not (MAPISession Is Nothing) Then
    
        MAPISession.SignOff
        MAPISession.NewSession = False
    
        Set MAPISession = Nothing
        Set MAPIMessages = Nothing
        Set MAPIMessagesCC = Nothing
        Set MAPIMessagesCCo = Nothing
        
    End If
    
    Set gobjRelOpcoes = Nothing

End Sub

Private Sub LabelCco_Click()

Dim iIndex As Integer, sRecipAddress As String

On Error GoTo Erro_LabelCco_Click

    If giPgmEmail <> 0 Then
        MAPIMessagesCCo.Show
        
        Cco.Text = ""
        
        iIndex = MAPIMessagesCCo.RecipCount - 1
        Do While iIndex >= 0
            
            MAPIMessagesCCo.RecipIndex = iIndex
            If left(MAPIMessagesCCo.RecipAddress, 5) = "SMTP:" Then
                sRecipAddress = Mid(MAPIMessagesCCo.RecipAddress, 6)
            Else
                sRecipAddress = MAPIMessagesCCo.RecipAddress
            End If
            
            If Cco.Text <> "" Then
                Cco.Text = Cco.Text + ";" + sRecipAddress
            Else
                Cco.Text = sRecipAddress
            End If
            
            iIndex = iIndex - 1
        
        Loop
    End If
    
    Exit Sub

Erro_LabelCco_Click:

    Select Case Err

        Case 32001 'Erro quando cancela o Address Book

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159470)

    End Select

    Exit Sub

End Sub

Private Sub LabelPara_Click()
    
Dim iIndex As Integer, sRecipAddress As String

On Error GoTo Erro_LabelPara_Click

    If giPgmEmail <> 0 Then
        MAPIMessages.Show
        
        Para.Text = ""
        
        iIndex = MAPIMessages.RecipCount - 1
        Do While iIndex >= 0
            
            MAPIMessages.RecipIndex = iIndex
            
            If left(MAPIMessages.RecipAddress, 5) = "SMTP:" Then
                sRecipAddress = Mid(MAPIMessages.RecipAddress, 6)
            Else
                sRecipAddress = MAPIMessages.RecipAddress
            End If
            
            If Para.Text <> "" Then
                Para.Text = Para.Text + ";" + sRecipAddress
            Else
                Para.Text = sRecipAddress
            End If
            
            iIndex = iIndex - 1
        
        Loop
    End If
    
    Exit Sub

Erro_LabelPara_Click:

    Select Case Err

        Case 32001 'Erro quando cancela o Address Book

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159470)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objRelOpcoes As AdmRelOpcoes, sOpcao As String) As Long 'tinha o parametro objRelatorio As AdmRelatorio

Dim lErro As Long
Dim sToEmail As String, sSubject As String, sArquivo As String, sMensagem As String
Dim sCcEmail As String, sCcoEmail As String, sInfoEmail As String

On Error GoTo Erro_Trata_Parametros

    'If Not (gobjRelatorio Is Nothing) Then gError 97114
    
    'Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    '??? gsOpcao = sOpcao

    lErro = objRelOpcoes.ObterParametro("TINFO_EMAIL", sInfoEmail)
    If lErro <> SUCESSO Then gError 124013
    
    lErro = objRelOpcoes.ObterParametro("TTO_EMAIL", sToEmail)
    If lErro <> SUCESSO Then gError 124013
    
    lErro = objRelOpcoes.ObterParametro("TCC_EMAIL", sCcEmail)
    If lErro <> SUCESSO Then gError 124013
    
    lErro = objRelOpcoes.ObterParametro("TCCO_EMAIL", sCcoEmail)
    If lErro <> SUCESSO Then gError 124013
    
    lErro = objRelOpcoes.ObterParametro("TSUBJECT", sSubject)
    If lErro <> SUCESSO Then gError 124014
                
    lErro = objRelOpcoes.ObterParametro("TALIASATTACH", sArquivo)
    If lErro <> SUCESSO Then gError 124015
    
    lErro = objRelOpcoes.ObterParametro("TMENSAGEM", sMensagem)
    If lErro <> SUCESSO Then gError 124016
                
    LabelInfo.Caption = sInfoEmail
    Para.Text = sToEmail
    Cc.Text = sCcEmail
    Cco.Text = sCcoEmail
    Assunto.Text = sSubject
    Arquivo.Text = sArquivo
    Mensagem.Text = sMensagem
    
    '????
    'Preenche com as Opcoes
    'lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    'If lErro <> SUCESSO Then gError 97115

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 124013 To 124016
        
     '   Case 97114
        
     '   Case 97115
     '       lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159471)

    End Select

    Exit Function

End Function

Private Sub LabelCc_Click()

Dim iIndex As Integer, sRecipAddress As String

On Error GoTo Erro_LabelCc_Click

    If giPgmEmail <> 0 Then
        MAPIMessagesCC.Show
        
        Cc.Text = ""
        
        iIndex = MAPIMessagesCC.RecipCount - 1
        Do While iIndex >= 0
            
            MAPIMessagesCC.RecipIndex = iIndex
            If left(MAPIMessagesCC.RecipAddress, 5) = "SMTP:" Then
                sRecipAddress = Mid(MAPIMessagesCC.RecipAddress, 6)
            Else
                sRecipAddress = MAPIMessagesCC.RecipAddress
            End If
            
            If Cc.Text <> "" Then
                Cc.Text = Cc.Text + ";" + sRecipAddress
            Else
                Cc.Text = sRecipAddress
            End If
            
            iIndex = iIndex - 1
        
        Loop
    End If
    
    Exit Sub

Erro_LabelCc_Click:

    Select Case Err

        Case 32001 'Erro quando cancela o Address Book

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159470)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProcurar_Click()

    On Error GoTo Erro_BotaoProcurar_Click

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog1.ShowOpen

    ' Display name of selected file

    If OutrosAnexos.Text = "" Then
        OutrosAnexos.Text = CommonDialog1.FileName
    Else
        OutrosAnexos.Text = OutrosAnexos.Text & ";" & CommonDialog1.FileName
    End If
    
    Exit Sub

Erro_BotaoProcurar_Click:
    'User pressed the Cancel button
    Exit Sub

End Sub


