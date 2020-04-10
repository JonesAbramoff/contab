VERSION 5.00
Begin VB.Form Arquivamento 
   Caption         =   "Arquivamento"
   ClientHeight    =   2220
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3660
      Picture         =   "Arquivamento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1575
      Width           =   1380
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
      Left            =   1410
      Picture         =   "Arquivamento.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1575
      Width           =   1380
   End
   Begin VB.Label Label2 
      Caption         =   $"Arquivamento.frx":025C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   1620
      TabIndex        =   3
      Top             =   285
      Width           =   4605
   End
   Begin VB.Label Label1 
      Caption         =   "ALERTA !!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   360
      TabIndex        =   2
      Top             =   285
      Width           =   1290
   End
End
Attribute VB_Name = "Arquivamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Form_Load()
    
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    lErro_Chama_Tela = SUCESSO
        
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 162791)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoCancela_Click()
    
    giRetornoTela = vbCancel
    
    Unload Me
    
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim objFrmAguarde As New ClassFrmAguarde
Dim objFrmAguardeTela As New FrmAguarde

On Error GoTo Erro_BotaoOK_Click

    lErro = CF("Arquivamento_Executa", objFrmAguarde, objFrmAguardeTela)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Not (objFrmAguardeTela Is Nothing) Then Set objFrmAguardeTela = Nothing
    
    giRetornoTela = vbOK

    Unload Me
    
    Exit Sub
    
Erro_BotaoOK_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209001)
    
    End Select
    
    If Not (objFrmAguardeTela Is Nothing) Then
        Call objFrmAguardeTela.Trata_Erro
    End If
    
    Exit Sub
    
End Sub

