VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TRPUsuarios 
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   5625
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   2895
      Picture         =   "TRPUsuarios.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4350
      Width           =   1005
   End
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   525
      Left            =   1470
      Picture         =   "TRPUsuarios.ctx":0102
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4350
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   "Acesso"
      Height          =   2715
      Left            =   165
      TabIndex        =   10
      Top             =   1575
      Width           =   5250
      Begin VB.TextBox ConfSenha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1785
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1185
         Width           =   2205
      End
      Begin VB.TextBox Senha 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1785
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   735
         Width           =   2205
      End
      Begin VB.ComboBox GrupoAcesso 
         Height          =   315
         Left            =   1785
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2100
         Width           =   2385
      End
      Begin VB.CheckBox AlteraSenhaProxLog 
         Caption         =   "Altera Senha no próximo logon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1770
         TabIndex        =   3
         Top             =   1665
         Width           =   3210
      End
      Begin MSMask.MaskEdBox Login 
         Height          =   315
         Left            =   1785
         TabIndex        =   0
         Top             =   300
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Conf. de Senha:"
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
         Height          =   315
         Left            =   45
         TabIndex        =   14
         Top             =   1230
         Width           =   1605
      End
      Begin VB.Label LabelGrupoAcesso 
         Alignment       =   1  'Right Justify
         Caption         =   "Grupo de Acesso:"
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
         Left            =   45
         TabIndex        =   13
         Top             =   2145
         Width           =   1605
      End
      Begin VB.Label LabelSenha 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   630
         TabIndex        =   12
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label LabelLogin 
         Alignment       =   1  'Right Justify
         Caption         =   "Login:"
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
         Height          =   315
         Left            =   630
         TabIndex        =   11
         Top             =   330
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identidicação"
      Height          =   1320
      Left            =   180
      TabIndex        =   5
      Top             =   120
      Width           =   5235
      Begin VB.Label Codigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1785
         TabIndex        =   9
         Top             =   675
         Width           =   2985
      End
      Begin VB.Label TipoUsuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   345
         Left            =   1785
         TabIndex        =   8
         Top             =   240
         Width           =   2985
      End
      Begin VB.Label LabelCodigo 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   645
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   7
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label LabelTipoUsuario 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   660
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   6
         Top             =   270
         Width           =   990
      End
   End
End
Attribute VB_Name = "TRPUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gobjUsuarioWeb As ClassTRPUsuarios

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Usuários Web"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRPUsuarios"

End Function

Public Sub Show()
    'Parent.Show
    'Parent.SetFocus
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

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set gobjUsuarioWeb = Nothing

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200334)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giRetornoTela = vbCancel

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200335)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTRPUsuarios As ClassTRPUsuarios) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTRPUsuarios Is Nothing) Then

        lErro = Traz_TRPUsuarios_Tela(objTRPUsuarios)
        If lErro <> SUCESSO Then gError 200336

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 200336

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200337)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objTRPUsuarios As ClassTRPUsuarios) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objTRPUsuarios.sLogin = Login.Text
    objTRPUsuarios.sSenha = Senha.Text
    
    If AlteraSenhaProxLog.Value = vbChecked Then
        objTRPUsuarios.iAlteraSenhaProxLog = MARCADO
    Else
        objTRPUsuarios.iAlteraSenhaProxLog = DESMARCADO
    End If
    
    objTRPUsuarios.sGrupoAcesso = GrupoAcesso.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200338)

    End Select

    Exit Function

End Function

Function Traz_TRPUsuarios_Tela(objTRPUsuarios As ClassTRPUsuarios) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_TRPUsuarios_Tela

    Set gobjUsuarioWeb = objTRPUsuarios

    Select Case objTRPUsuarios.iTipoUsuario
    
        Case TRP_USUARIO_CLIENTE
            TipoUsuario.Caption = TRP_USUARIO_CLIENTE_TEXTO
        Case TRP_USUARIO_EMISSOR
            TipoUsuario.Caption = TRP_USUARIO_EMISSOR_TEXTO
        Case TRP_USUARIO_VENDEDOR
            TipoUsuario.Caption = TRP_USUARIO_VENDEDOR_TEXTO
        
    End Select

    Codigo.Caption = objTRPUsuarios.lCodigo & SEPARADOR & objTRPUsuarios.sNome

    Login.Text = objTRPUsuarios.sLogin
    Senha.Text = objTRPUsuarios.sSenha

    If objTRPUsuarios.iAlteraSenhaProxLog = MARCADO Then
        AlteraSenhaProxLog.Value = vbChecked
    Else
        AlteraSenhaProxLog.Value = vbUnchecked
    End If

     Call CF("sCombo_Seleciona2", GrupoAcesso, objTRPUsuarios.sGrupoAcesso)

    iAlterado = 0

    Traz_TRPUsuarios_Tela = SUCESSO

    Exit Function

Erro_Traz_TRPUsuarios_Tela:

    Traz_TRPUsuarios_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200339)

    End Select

    Exit Function

End Function

Private Sub Login_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Login_Validate

    'Verifica se Login está preenchida
    If Len(Trim(Login.Text)) <> 0 Then

       '#######################################
       'CRITICA Login
       '#######################################

    End If

    Exit Sub

Erro_Login_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200340)

    End Select

    Exit Sub

End Sub

Private Sub Login_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Senha_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Senha_Validate

    'Verifica se Senha está preenchida
    If Len(Trim(Senha.Text)) <> 0 Then

       '#######################################
       'CRITICA Senha
       '#######################################

    End If

    Exit Sub

Erro_Senha_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200341)

    End Select

    Exit Sub

End Sub

Private Sub Senha_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub GrupoAcesso_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoCancela_Click()

    Unload Me
    
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim objTRPUsuarios As New ClassTRPUsuarios

On Error GoTo Erro_BotaoOK_Click

    If Len(Trim(Login.Text)) = 0 Then gError 200342
    If Len(Trim(Senha.Text)) = 0 Then gError 200343
    If Len(Trim(ConfSenha.Text)) = 0 Then gError 200344
    If UCase(ConfSenha.Text) <> UCase(Senha.Text) Then gError 200345
    
    objTRPUsuarios.iTipoUsuario = gobjUsuarioWeb.iTipoUsuario
    objTRPUsuarios.lCodigo = gobjUsuarioWeb.lCodigo
    objTRPUsuarios.sLogin = Login.Text
    
    lErro = CF("TRPUsuarios_Testa_Login", objTRPUsuarios)
    If lErro <> SUCESSO Then gError 200346
    
    lErro = Move_Tela_Memoria(gobjUsuarioWeb)
    If lErro <> SUCESSO Then gError 200347

    giRetornoTela = vbOK
    Unload Me
 
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr
    
        Case 200342
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_USUARIO_NAO_INFORMADO", gErr)

        Case 200343
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_USUARIO_NAO_INFORMADA", gErr)

        Case 200344
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_CONFIRMADA", gErr)

        Case 200345
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_USUARIO_NAO_CONFERE", gErr)

        Case 200346, 200347

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 200348)
            
    End Select
    
    Exit Sub
        
End Sub
