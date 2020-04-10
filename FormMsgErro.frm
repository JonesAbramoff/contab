VERSION 5.00
Begin VB.Form FormMsgErro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensagem de Erro"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "FormMsgErro.frx":0000
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
      Height          =   1485
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   75
      Width           =   5460
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
      Left            =   2153
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label MsgSolucao 
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   5415
   End
End
Attribute VB_Name = "FormMsgErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function HelpHelp_ErroTipo_Carregar Lib "ADHELP01.DLL" (ByVal lTipoErro As Long, ByVal lpMsgErro As String) As Long
Private Declare Function HelpHelp_ErroLocal_Carregar Lib "ADHELP02.DLL" (ByVal lLocalErro As Long, ByVal lpMsgErro As String) As Long

Public sErro As String
Public lLocalErro As Long
Public sTipoErro As String

Private Sub BotaoOK_Click()

    Unload Me
    
End Sub

'Private Sub BotaoSolucao_Click()
'
'Dim sSolucao As String, iTam As Integer
'
'    'tentar obter erro especifico do contexto em que ocorreu
'    sSolucao = String(1024, 0)
'    iTam = HelpHelp_ErroLocal_Carregar(lLocalErro, sSolucao)
'
''??? aguardando conclusao de help do help
'''    If iTam = 0 Then
'''        sSolucao = String(1024, 0)
'''        iTam = HelpHelp_ErroTipo_Carregar(lTipoErro, sSolucao)
'''    End If
'
'    If iTam = 0 Then
'        sSolucao = "Não há informações adicionais disponíveis."
'    Else
'        sSolucao = StringZ(sSolucao)
'    End If
'
'    MsgSolucao.Caption = sSolucao
'    Me.Height = 4215
'    BotaoSolucao.Enabled = False
'
'End Sub

Private Sub Form_Load()
    
    MsgDeErro.Text = Replace(sErro, Chr$(10), vbNewLine)
    
End Sub

'Private Sub MsgSolucao_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(MsgSolucao, Source, X, Y)
'End Sub
'
'Private Sub MsgSolucao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(MsgSolucao, Button, Shift, X, Y)
'End Sub
'
'Private Sub MsgDeErro_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(MsgDeErro, Source, X, Y)
'End Sub
'
'Private Sub MsgDeErro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(MsgDeErro, Button, Shift, X, Y)
'End Sub
'
