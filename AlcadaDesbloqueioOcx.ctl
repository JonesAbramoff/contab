VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl AlcadaDesbloqueioOcx 
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8535
   ScaleHeight     =   3930
   ScaleWidth      =   8535
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Bloqueio"
      Height          =   1200
      Left            =   300
      TabIndex        =   8
      Top             =   120
      Width           =   5535
      Begin MSMask.MaskEdBox Bloqueio 
         Height          =   315
         Left            =   1635
         TabIndex        =   0
         Top             =   315
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelNomeReduzido 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1605
         TabIndex        =   11
         Top             =   720
         Width           =   2475
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nome Reduzido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   810
         Width           =   1410
      End
      Begin VB.Label LabelBloqueio 
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
         Left            =   885
         TabIndex        =   9
         Top             =   345
         Width           =   660
      End
   End
   Begin VB.ListBox Autorizados 
      Columns         =   3
      Height          =   2085
      ItemData        =   "AlcadaDesbloqueioOcx.ctx":0000
      Left            =   330
      List            =   "AlcadaDesbloqueioOcx.ctx":0002
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1710
      Width           =   5160
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6255
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   210
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "AlcadaDesbloqueioOcx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "AlcadaDesbloqueioOcx.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "AlcadaDesbloqueioOcx.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "AlcadaDesbloqueioOcx.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   " Usuários"
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
      Left            =   300
      TabIndex        =   7
      Top             =   1500
      Width           =   810
   End
End
Attribute VB_Name = "AlcadaDesbloqueioOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

