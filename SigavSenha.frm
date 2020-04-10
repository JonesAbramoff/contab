VERSION 5.00
Begin VB.Form SigavSenha 
   Caption         =   "Extração de Dados"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1965
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1425
      TabIndex        =   2
      Top             =   1395
      Width           =   1215
   End
   Begin VB.TextBox Senha 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2295
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   945
      Width           =   1530
   End
   Begin VB.Label Label2 
      Caption         =   "ATENÇÃO: O Sigav deve estar sendo executado nessa máquina para que a extração de dados possa ser feita."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   3705
   End
   Begin VB.Label Label1 
      Caption         =   "Digite a senha no Sigav:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   150
      TabIndex        =   0
      Top             =   975
      Width           =   2370
   End
End
Attribute VB_Name = "SigavSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gobjSenha As Object

Private Sub BotaoOK_Click()
    
Dim lErro As Long
    
On Error GoTo Erro_BotaoOK_Click
    
    'Indica que saiu da tela de forma legal
    giRetornoTela = vbOK
    
    gobjSenha.sSenha = Senha.Text
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192670)

    End Select

    Exit Sub
    
End Sub

Function Trata_Parametros(ByVal objSenha As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjSenha = objSenha

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192669)

    End Select

    Exit Function

End Function

