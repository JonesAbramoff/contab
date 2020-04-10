VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl GrafEvolucaoOrcamentos 
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   ScaleHeight     =   3270
   ScaleMode       =   0  'User
   ScaleWidth      =   6945
   Begin VB.CommandButton BotaoExecutar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4853
      Picture         =   "GrafEvolucaoOrcamentos.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   945
      Width           =   1605
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "GrafEvolucaoOrcamentos.ctx":2DF6
      Left            =   1200
      List            =   "GrafEvolucaoOrcamentos.ctx":2DF8
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   240
      Width           =   2670
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   750
      Left            =   240
      TabIndex        =   10
      Top             =   850
      Width           =   4035
      Begin MSComCtl2.UpDown UpDownDataDe 
         Height          =   300
         Left            =   1650
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataDe 
         Height          =   300
         Left            =   690
         TabIndex        =   12
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAte 
         Height          =   300
         Left            =   3645
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAte 
         Height          =   300
         Left            =   2685
         TabIndex        =   14
         Top             =   285
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelDataDe 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   300
         Width           =   315
      End
      Begin VB.Label LabelDataAte 
         AutoSize        =   -1  'True
         Caption         =   "Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2265
         TabIndex        =   15
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4560
      ScaleHeight     =   495
      ScaleWidth      =   2130
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   2190
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "GrafEvolucaoOrcamentos.ctx":2DFA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   600
         Picture         =   "GrafEvolucaoOrcamentos.ctx":2F54
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "GrafEvolucaoOrcamentos.ctx":30DE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1650
         Picture         =   "GrafEvolucaoOrcamentos.ctx":3610
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame FrameGrafico 
      Caption         =   "Exibir no Gráfico"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   6495
      Begin VB.CheckBox ExibeNumOrcamentos 
         Caption         =   "Nº de Orçamentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox ExibeValorVendas 
         Caption         =   "Valor Orçamentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox ExibeQuantProdutos 
         Caption         =   "Quant. de Produtos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CheckBox DetalhadoCaixa 
      Caption         =   "Detalhar valor dos orçamentos por caixa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
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
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   270
      Width           =   615
   End
End
Attribute VB_Name = "GrafEvolucaoOrcamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

