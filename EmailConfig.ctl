VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl EmailConfigOcx 
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   KeyPreview      =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   6930
   Begin VB.Frame Frame4 
      Caption         =   "Outras Opções"
      Height          =   1035
      Left            =   120
      TabIndex        =   25
      Top             =   4725
      Width           =   6735
      Begin VB.Frame Frame5 
         Caption         =   "Preferência"
         Height          =   480
         Left            =   105
         TabIndex        =   26
         Top             =   465
         Width           =   6510
         Begin VB.OptionButton OptPrefCorp 
            Caption         =   "Envio direto pelo Corporator"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3540
            TabIndex        =   11
            Top             =   195
            Width           =   2835
         End
         Begin VB.OptionButton OptPrefPadrao 
            Caption         =   "Usar o meu pgm padrão de Email"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   225
            TabIndex        =   10
            Top             =   195
            Value           =   -1  'True
            Width           =   3180
         End
      End
      Begin VB.CheckBox Confirmacao 
         Caption         =   "Solicitar confirmação de Leitura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1110
         TabIndex        =   9
         Top             =   225
         Width           =   3330
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informações do Usuário"
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   1530
      Width           =   6735
      Begin VB.TextBox Nome 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1110
         TabIndex        =   3
         Top             =   705
         Width           =   2805
      End
      Begin MSMask.MaskEdBox Email 
         Height          =   315
         Left            =   1110
         TabIndex        =   2
         Top             =   270
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
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
         Left            =   150
         TabIndex        =   24
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
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
         Left            =   150
         TabIndex        =   23
         Top             =   750
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informações do Servidor de Email"
      Height          =   1980
      Left            =   120
      TabIndex        =   16
      Top             =   2745
      Width           =   6735
      Begin VB.CheckBox SSL 
         Caption         =   "Requer uma conexão criptografada (SSL)"
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
         Left            =   1785
         TabIndex        =   6
         Top             =   630
         Width           =   4905
      End
      Begin VB.TextBox SMTPSenha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1110
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1515
         Width           =   2805
      End
      Begin MSMask.MaskEdBox SMTP 
         Height          =   315
         Left            =   1110
         TabIndex        =   4
         Top             =   255
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox SMTPUsu 
         Height          =   315
         Left            =   1110
         TabIndex        =   7
         Top             =   1080
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Porta 
         Height          =   315
         Left            =   1110
         TabIndex        =   5
         Top             =   660
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Porta:"
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
         Left            =   135
         TabIndex        =   21
         Top             =   705
         Width           =   915
      End
      Begin VB.Label LabelSMTPSenha 
         Alignment       =   1  'Right Justify
         Caption         =   "Senha:"
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
         Left            =   150
         TabIndex        =   19
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label LabelSMTPUsu 
         Alignment       =   1  'Right Justify
         Caption         =   "Usuário:"
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
         Left            =   150
         TabIndex        =   18
         Top             =   1110
         Width           =   915
      End
      Begin VB.Label LabelSMTP 
         Alignment       =   1  'Right Justify
         Caption         =   "SMTP:"
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
         Left            =   150
         TabIndex        =   17
         Top             =   285
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuração"
      Height          =   780
      Left            =   105
      TabIndex        =   15
      Top             =   690
      Width           =   6750
      Begin VB.OptionButton OptPorUsu 
         Caption         =   "Para o Usuário"
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
         Left            =   1980
         TabIndex        =   1
         Top             =   330
         Width           =   1815
      End
      Begin VB.OptionButton OptGeral 
         Caption         =   "Geral"
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
         Left            =   420
         TabIndex        =   0
         Top             =   315
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.Label Usuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3945
         TabIndex        =   20
         Top             =   285
         Width           =   2640
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   5790
      ScaleHeight     =   450
      ScaleWidth      =   1005
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   105
      Width           =   1065
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "EmailConfig.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   510
         Picture         =   "EmailConfig.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
End
Attribute VB_Name = "EmailConfigOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Configuração de Envio de Email"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "EmailConfig"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


Private Sub OptGeral_Click()
    Call Traz_EmailConfig_Tela
End Sub

Private Sub OptPorUsu_Click()
    Call Traz_EmailConfig_Tela
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    'Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    'gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202793)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Usuario.Caption = gsUsuario
    
    lErro = Traz_EmailConfig_Tela()
    If lErro <> SUCESSO Then gError 202794

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 202794

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202795)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202796)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objEmailConfig As ClassEmailConfig) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    If OptGeral.Value Then
        objEmailConfig.sUsuario = ""
    Else
        objEmailConfig.sUsuario = Usuario.Caption
    End If
    objEmailConfig.sSMTP = Trim(SMTP.Text)
    objEmailConfig.sSMTPUsu = Trim(SMTPUsu.Text)
    objEmailConfig.sSMTPSenha = Trim(SMTPSenha.Text)
    objEmailConfig.lSMTPPorta = StrParaLong(Porta.Text)
    
    If SSL.Value = vbChecked Then
        objEmailConfig.iSSL = MARCADO
    Else
        objEmailConfig.iSSL = DESMARCADO
    End If
    
    If Confirmacao.Value = vbChecked Then
        objEmailConfig.iConfirmacaoLeitura = MARCADO
    Else
        objEmailConfig.iConfirmacaoLeitura = DESMARCADO
    End If
    
    If OptPrefCorp.Value Then
        objEmailConfig.iPgmEmail = 0
    Else
        objEmailConfig.iPgmEmail = 1
    End If
    
    objEmailConfig.sEmail = Trim(Email.Text)
    objEmailConfig.sNome = Trim(Nome.Text)
    

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202797)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objEmailConfig As New ClassEmailConfig

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Preenche o objEmailConfig
    lErro = Move_Tela_Memoria(objEmailConfig)
    If lErro <> SUCESSO Then gError 202798

    lErro = Trata_Alteracao(objEmailConfig, objEmailConfig.sUsuario)
    If lErro <> SUCESSO Then gError 202799

    'Grava o/a EmailConfig no Banco de Dados
    lErro = CF("EmailConfig_Grava", objEmailConfig)
    If lErro <> SUCESSO Then gError 202800
    
    iAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 202798 To 202800

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202801)

    End Select

    Exit Function

End Function

Function Traz_EmailConfig_Tela() As Long

Dim lErro As Long
Dim objEmailConfig As New ClassEmailConfig

On Error GoTo Erro_Traz_EmailConfig_Tela
    
    If OptGeral.Value Then
        objEmailConfig.sUsuario = ""
    Else
        objEmailConfig.sUsuario = gsUsuario
    End If

    'Lê o EmailConfig que está sendo Passado
    lErro = CF("EmailConfig_Le", objEmailConfig)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 202803

    SMTP.Text = objEmailConfig.sSMTP
    SMTPUsu.Text = objEmailConfig.sSMTPUsu
    SMTPSenha.Text = objEmailConfig.sSMTPSenha
    
    If objEmailConfig.lSMTPPorta <> 0 Then
        Porta.PromptInclude = False
        Porta.Text = CStr(objEmailConfig.lSMTPPorta)
        Porta.PromptInclude = True
    End If
    
    If objEmailConfig.iSSL = MARCADO Then
        SSL.Value = vbChecked
    Else
        SSL.Value = vbUnchecked
    End If
    
    If objEmailConfig.iConfirmacaoLeitura = MARCADO Then
        Confirmacao.Value = vbChecked
    Else
        Confirmacao.Value = vbUnchecked
    End If
    
    If objEmailConfig.iPgmEmail = 0 Then
        OptPrefCorp.Value = True
    Else
        OptPrefPadrao.Value = True
    End If
    
    Email.Text = objEmailConfig.sEmail
    Nome.Text = objEmailConfig.sNome

    iAlterado = 0

    Traz_EmailConfig_Tela = SUCESSO

    Exit Function

Erro_Traz_EmailConfig_Tela:

    Traz_EmailConfig_Tela = gErr

    Select Case gErr

        Case 202803

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202804)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 202805

    Unload Me

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 202805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202806)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202807)

    End Select

    Exit Sub

End Sub

Private Sub SMTP_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SMTP_Validate

    'Verifica se SMTP está preenchida
    If Len(Trim(SMTP.Text)) <> 0 Then

       '#######################################
       'CRITICA SMTP
       '#######################################

    End If

    Exit Sub

Erro_SMTP_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202810)

    End Select

    Exit Sub

End Sub

Private Sub SMTP_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SMTPUsu_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SMTPUsu_Validate

    'Verifica se SMTPUsu está preenchida
    If Len(Trim(SMTPUsu.Text)) <> 0 Then

       '#######################################
       'CRITICA SMTPUsu
       '#######################################

    End If

    Exit Sub

Erro_SMTPUsu_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202811)

    End Select

    Exit Sub

End Sub

Private Sub SMTPUsu_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SMTPSenha_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SMTPSenha_Validate

    'Verifica se SMTPSenha está preenchida
    If Len(Trim(SMTPSenha.Text)) <> 0 Then

       '#######################################
       'CRITICA SMTPSenha
       '#######################################

    End If

    Exit Sub

Erro_SMTPSenha_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202812)

    End Select

    Exit Sub

End Sub

Private Sub SMTPSenha_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Porta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Porta_GotFocus()
    Call MaskEdBox_TrataGotFocus(Porta, iAlterado)
End Sub

Private Sub Email_Validate(Cancel As Boolean)

    If Len(Trim(SMTPUsu.Text)) = 0 Then
        SMTPUsu.Text = Email.Text
    End If
End Sub
