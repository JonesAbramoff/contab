VERSION 5.00
Begin VB.Form FormMsgAviso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SGE - Forprint"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "FormMsgAviso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox MsgDeErro 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   75
      Width           =   5490
   End
   Begin VB.Frame FrameBotoes 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   1590
      Width           =   5415
      Begin VB.CommandButton BotaoOK 
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
         Height          =   375
         Index           =   0
         Left            =   2153
         TabIndex        =   0
         Top             =   90
         Width           =   1335
      End
   End
   Begin VB.Frame FrameBotoes 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   14
      Top             =   1590
      Width           =   5415
      Begin VB.CommandButton BotaoCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3068
         TabIndex        =   2
         Top             =   90
         Width           =   1335
      End
      Begin VB.CommandButton BotaoOK 
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
         Height          =   375
         Index           =   1
         Left            =   1238
         TabIndex        =   1
         Top             =   90
         Width           =   1335
      End
   End
   Begin VB.Frame FrameBotoes 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   17
      Top             =   1590
      Width           =   5415
      Begin VB.CommandButton BotaoIgnorar 
         Caption         =   "Ignorar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3848
         TabIndex        =   11
         Top             =   90
         Width           =   1335
      End
      Begin VB.CommandButton BotaoAbortar 
         Caption         =   "Abortar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   458
         TabIndex        =   9
         Top             =   90
         Width           =   1335
      End
      Begin VB.CommandButton BotaoRepetir 
         Caption         =   "Repetir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2153
         TabIndex        =   8
         Top             =   90
         Width           =   1335
      End
   End
   Begin VB.Frame FrameBotoes 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   0
      TabIndex        =   10
      Top             =   1590
      Width           =   5415
      Begin VB.CommandButton BotaoNao 
         Caption         =   "Não"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2153
         TabIndex        =   5
         Top             =   90
         Width           =   1335
      End
      Begin VB.CommandButton BotaoSim 
         Caption         =   "Sim"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   458
         TabIndex        =   4
         Top             =   90
         Width           =   1335
      End
      Begin VB.CommandButton BotaoCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3848
         TabIndex        =   3
         Top             =   90
         Width           =   1335
      End
   End
   Begin VB.Frame FrameBotoes 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Index           =   4
      Left            =   0
      TabIndex        =   16
      Top             =   1590
      Width           =   5415
      Begin VB.CommandButton BotaoSim 
         Caption         =   "Sim"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1305
         TabIndex        =   7
         Top             =   90
         Width           =   1335
      End
      Begin VB.CommandButton BotaoNao 
         Caption         =   "Não"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   6
         Top             =   90
         Width           =   1335
      End
   End
   Begin VB.Frame FrameBotoes 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Index           =   5
      Left            =   0
      TabIndex        =   18
      Top             =   1590
      Width           =   5415
      Begin VB.CommandButton BotaoRepetir 
         Caption         =   "Repetir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1305
         TabIndex        =   13
         Top             =   90
         Width           =   1335
      End
      Begin VB.CommandButton BotaoCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   12
         Top             =   90
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FormMsgAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sErro As String
Public MsgBoxTipo As VbMsgBoxStyle
Public MsgBoxResultado As VbMsgBoxResult

Private Sub BotaoAbortar_Click(Index As Integer)
    
    'Passa a escolha e fecha a Tela
    MsgBoxResultado = vbAbort
    
    Unload Me
    
End Sub

Private Sub BotaoCancelar_Click(Index As Integer)

    'Passa a escolha e fecha a Tela
    MsgBoxResultado = vbCancel
    
    Unload Me
    
End Sub

Private Sub BotaoIgnorar_Click(Index As Integer)

    'Passa a escolha e fecha a Tela
    MsgBoxResultado = vbIgnore
    
    Unload Me

End Sub

Private Sub BotaoNao_Click(Index As Integer)

    'Passa a escolha e fecha a Tela
    MsgBoxResultado = vbNo
    
    Unload Me

End Sub

Private Sub BotaoOK_Click(Index As Integer)

    'Passa a escolha e fecha a Tela
    MsgBoxResultado = vbOK
    
    Unload Me

End Sub

Private Sub BotaoRepetir_Click(Index As Integer)
    
    'Passa a escolha e fecha a Tela
    MsgBoxResultado = vbRetry
    
    Unload Me

End Sub

Private Sub BotaoSim_Click(Index As Integer)

    'Passa a escolha e fecha a Tela
    MsgBoxResultado = vbYes
    
    Unload Me

End Sub

Private Sub Form_Load()

Dim iIndice As Integer
    
    'Passa o Erro para a Tela
    MsgDeErro.Text = Replace(sErro, Chr$(10), vbNewLine)
    
    'resposta default
    Select Case MsgBoxTipo
        
        Case vbYesNo
            MsgBoxResultado = vbNo
        
        Case vbOKOnly
            MsgBoxResultado = vbOK
        
        Case Else
            MsgBoxResultado = vbCancel
        
    End Select
    
    'Seleciona o Frame de Botoes de acordo com tipo selecionado
    For iIndice = 0 To 5
            
        'Se for o Tipo passado
        If MsgBoxTipo = iIndice Then
            'Torna o Frame visivel
            FrameBotoes(iIndice).Visible = True
        Else
            FrameBotoes(iIndice).Visible = False
        End If
    
    Next
        
End Sub

Private Sub MsgDeErro_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MsgDeErro, Source, X, Y)
End Sub

Private Sub MsgDeErro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MsgDeErro, Button, Shift, X, Y)
End Sub

