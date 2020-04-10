VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Indexador 
   Caption         =   "Cadastro de Indexadores"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Mensagem 
      Height          =   1680
      Left            =   3375
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   690
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3390
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   45
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Indexador.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "Indexador.frx":015A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "Indexador.frx":02E4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "Indexador.frx":0816
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1620
      Picture         =   "Indexador.frx":0994
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Numeração Automática"
      Top             =   720
      Width           =   300
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1035
      TabIndex        =   1
      Top             =   705
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NomeReduzido 
      Height          =   315
      Left            =   1035
      TabIndex        =   3
      Top             =   1155
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Simbolo 
      Height          =   315
      Left            =   1035
      TabIndex        =   6
      Top             =   1620
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      Format          =   "AAA"
      PromptChar      =   " "
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   375
      TabIndex        =   5
      Top             =   1230
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Símbolo:"
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
      Left            =   195
      TabIndex        =   4
      Top             =   1665
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
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
      Left            =   285
      TabIndex        =   2
      Top             =   765
      Width           =   660
   End
End
Attribute VB_Name = "Indexador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub
