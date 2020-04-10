VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl SuprimentoDinheiro 
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   ScaleHeight     =   4215
   ScaleWidth      =   7605
   Begin VB.Frame FrameSuprimentoDinheiro 
      Caption         =   "Suprimento em Dinheiro"
      Height          =   3825
      Left            =   240
      TabIndex        =   0
      Top             =   200
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton BotaoSair 
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   3720
         TabIndex        =   6
         Top             =   2880
         Width           =   2985
      End
      Begin VB.CommandButton BotaoSuprimento 
         Caption         =   "Suprimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   360
         TabIndex        =   5
         Top             =   2880
         Width           =   2985
      End
      Begin MSMask.MaskEdBox SuprimentoDinheiro 
         Height          =   645
         Left            =   2280
         TabIndex        =   1
         Top             =   1740
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1138
         _Version        =   393216
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label SaldoDinheiro 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   2280
         TabIndex        =   4
         Top             =   720
         Width           =   3765
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   1470
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   600
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   2
         Top             =   1800
         Width           =   1380
      End
   End
End
Attribute VB_Name = "SuprimentoDinheiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

