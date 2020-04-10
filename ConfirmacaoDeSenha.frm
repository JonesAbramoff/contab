VERSION 5.00
Begin VB.Form ConfirmacaoDeSenha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirme a senha digitada"
   ClientHeight    =   1140
   ClientLeft      =   1845
   ClientTop       =   4800
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ConfirmaSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Repita a Senha:"
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
      Left            =   720
      TabIndex        =   3
      Top             =   420
      Width           =   1395
   End
End
Attribute VB_Name = "ConfirmacaoDeSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
                
    Unload Me

End Sub

Private Sub OKButton_Click()

On Error GoTo Erro_OKButton_Click
   
    If Len(ConfirmaSenha.Text) = 0 Then gError 134264
    
    If ConfirmaSenha.Text <> UsuarioTela.Senha Then gError 134265
     
    UsuarioTela.bSenhaAlterada = False
     
    Unload Me
    
    Exit Sub
    
Erro_OKButton_Click:

    Select Case gErr
    
        Case 134264
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_USUARIO_NAO_INFORMADA", gErr)
            
        Case 134265
        
            Call Rotina_Erro(vbOKOnly, "ERRO_SENHA_USUARIO_NAO_CONFERE", gErr)
            ConfirmaSenha.Text = ""
            ConfirmaSenha.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154699)
        
    End Select
    
    Exit Sub

End Sub
