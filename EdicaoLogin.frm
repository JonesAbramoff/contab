VERSION 5.00
Begin VB.Form EdicaoLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2655
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4545
   Icon            =   "EdicaoLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1568.662
   ScaleMode       =   0  'User
   ScaleWidth      =   4267.509
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Usuário"
      Height          =   1485
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   4035
      Begin VB.ComboBox ComboUsuario 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   375
         Width           =   2940
      End
      Begin VB.TextBox TextSenha 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   945
         Width           =   2910
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   420
         Width           =   555
      End
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2445
      Picture         =   "EdicaoLogin.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1905
      Width           =   975
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1005
      Picture         =   "EdicaoLogin.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1905
      Width           =   975
   End
End
Attribute VB_Name = "EdicaoLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gobjFlag As AdmGenerico

Function Trata_Parametros(objFlag As AdmGenerico) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    Set gobjFlag = objFlag
    objFlag.vVariavel = False
        
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159203)
    
    End Select
    
    Exit Function

End Function

Private Sub BotaoCancela_Click()
    
    Unload Me
    
End Sub

Private Sub BotaoOk_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOk_Click
    
    'Verificar se o campo Usuario esta preenchido
    If Len(ComboUsuario) = 0 Then Error 41654
    
    'Verificar se a senha esta preenchida
    If Len(TextSenha) = 0 Then Error 41655
    
    'faz login utilizando o codigo do usuario e a senha
    lErro = Sistema_Login(ComboUsuario.Text, TextSenha.Text)
    If lErro <> AD_BOOL_TRUE Then Error 41656
    
    gobjFlag.vVariavel = True
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoOk_Click:

    Select Case Err
    
        Case 41654 'Usuario nao preenchido
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_PREENCHIDO", Err)
            
        Case 41655 'Senha nao preenchida
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SENHA_NAO_PREENCHIDA", Err)
        
        Case 41656 'nao conseguiu fazer login
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 159204)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim colUsuarios As New Collection
Dim objUsuarios As ClassUsuarios

On Error GoTo Erro_Form_Load

    'Le todos os usuarios da tabela usuarios e coloca na colecao
    lErro = CF("Usuarios_Le_Todos", colUsuarios)
    If lErro <> SUCESSO Then Error 41653

    'Coloca todos os Usuarios do "grupo de supervisores"
    'com senha nao expirada na ComboUsuario
    For Each objUsuarios In colUsuarios
        If StrComp(UCase(objUsuarios.sCodGrupo), UCase(GRUPO_SUP), 1) = 0 Then
            ComboUsuario.AddItem objUsuarios.sCodUsuario
        End If
    Next

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err
    
    Select Case Err
    
        Case 41653
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 159205)
    
    End Select
    
    Exit Sub

End Sub
