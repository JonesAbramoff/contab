VERSION 5.00
Begin VB.Form RelDlgExec 
   Caption         =   "Execução de Relatório"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BotaoEmail 
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   3450
      Picture         =   "RelDlgExec.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   195
      Width           =   1290
   End
   Begin VB.CommandButton BotaoImpressora 
      Caption         =   "Impressora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1845
      Picture         =   "RelDlgExec.frx":09A2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   195
      Width           =   1290
   End
   Begin VB.CommandButton BotaoPrevia 
      Caption         =   "Prévia"
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
      Height          =   870
      Left            =   255
      Picture         =   "RelDlgExec.frx":100C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   195
      Width           =   1290
   End
End
Attribute VB_Name = "RelDlgExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gobjRelOpcoes As AdmRelOpcoes

Function Trata_Parametros(objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjRelOpcoes = objRelOpcoes
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166673)

    End Select

    Exit Function

End Function

Private Sub BotaoEmail_Click()
    gobjRelOpcoes.iDispositivoDeSaida = REL_SAIDA_EMAIL
    Unload Me
End Sub

Private Sub BotaoImpressora_Click()
    gobjRelOpcoes.iDispositivoDeSaida = REL_SAIDA_IMPRESSORA
    Unload Me
End Sub

Private Sub BotaoPrevia_Click()
    gobjRelOpcoes.iDispositivoDeSaida = REL_SAIDA_PREVIA
    Unload Me
End Sub

Private Sub Form_Load()
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166674)

    End Select

    Exit Sub
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode <> vbFormCode Then
        gobjRelOpcoes.bDesistiu = True
    End If

End Sub
