VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TipoGarantia 
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
   ScaleHeight     =   4635
   ScaleWidth      =   6345
   Begin VB.Frame Frame7 
      Caption         =   "Serviços/Peças"
      Height          =   2880
      Left            =   225
      TabIndex        =   9
      Top             =   1635
      Width           =   5985
      Begin VB.CheckBox GarantiaTotal 
         Caption         =   "Garantia Total"
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
         Left            =   570
         TabIndex        =   18
         Top             =   360
         Width           =   1830
      End
      Begin MSMask.MaskEdBox PrazoValidade 
         Height          =   225
         Left            =   5100
         TabIndex        =   15
         Top             =   1185
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.TextBox DescricaoProduto 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   2055
         MaxLength       =   250
         TabIndex        =   11
         Top             =   1185
         Width           =   3000
      End
      Begin MSMask.MaskEdBox Produto 
         Height          =   225
         Left            =   795
         TabIndex        =   12
         Top             =   1110
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridServicos 
         Height          =   1875
         Left            =   150
         TabIndex        =   10
         Top             =   870
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3307
         _Version        =   393216
         Rows            =   6
         Cols            =   3
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   0
      End
      Begin MSMask.MaskEdBox GarantiaTotalPrazo 
         Height          =   315
         Left            =   4620
         TabIndex        =   16
         Top             =   375
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Prazo (em dias):"
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
         Left            =   3195
         TabIndex        =   17
         Top             =   420
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4110
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   105
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TipoGarantia.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TipoGarantia.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1095
         Picture         =   "TipoGarantia.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TipoGarantia.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Descricao 
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   765
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   50
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1650
      TabIndex        =   1
      Top             =   300
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox PrazoPadraoValidade 
      Height          =   315
      Left            =   2235
      TabIndex        =   13
      Top             =   1215
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "Prazo Padrão(em dias):"
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
      Left            =   180
      TabIndex        =   14
      Top             =   1260
      Width           =   1980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Index           =   0
      Left            =   660
      TabIndex        =   3
      Top             =   810
      Width           =   930
   End
   Begin VB.Label LblTipo 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
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
      Left            =   1140
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   2
      Top             =   345
      Width           =   450
   End
End
Attribute VB_Name = "TipoGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

